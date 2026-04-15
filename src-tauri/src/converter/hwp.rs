use cfb::CompoundFile;
use flate2::read::DeflateDecoder;
use std::io::Read;
use std::path::Path;

// 실제 HWP 5.0 태그 ID (바이너리 분석으로 확인)
const TAG_PARA_HEADER: u16 = 66;   // HWPTAG_BEGIN(16) + 50
const TAG_PARA_TEXT: u16 = 67;     // HWPTAG_BEGIN(16) + 51
const TAG_CTRL_HEADER: u16 = 71;   // HWPTAG_BEGIN(16) + 55
const TAG_LIST_HEADER: u16 = 72;   // HWPTAG_BEGIN(16) + 56

// CTRL_HEADER에서 "tbl " 식별자 (리틀엔디안: 0x20 0x6C 0x62 0x74)
const TBL_CTRL_TYPE: [u8; 4] = [0x20, 0x6C, 0x62, 0x74];

pub fn convert(path: &Path) -> Result<String, String> {
    let file = std::fs::File::open(path)
        .map_err(|e| format!("HWP 파일 열기 실패: {}", e))?;
    let mut comp = CompoundFile::open(file)
        .map_err(|e| format!("HWP CFB 열기 실패: {}", e))?;

    let compressed = is_compressed(&mut comp)?;
    let mut all_elements: Vec<Element> = Vec::new();
    let mut section_idx = 0;

    loop {
        let stream_name = format!("BodyText/Section{}", section_idx);
        let section_data = match read_stream(&mut comp, &stream_name) {
            Ok(data) => data,
            Err(_) => break,
        };
        let decompressed = if compressed { decompress(&section_data)? } else { section_data };
        let elements = parse_section(&decompressed)?;
        all_elements.extend(elements);
        section_idx += 1;
    }

    if section_idx == 0 {
        return Err("BodyText 섹션을 찾을 수 없습니다".to_string());
    }

    Ok(render_markdown(&all_elements))
}

enum Element {
    Paragraph(String),
    Table(TableData),
}

struct TableData {
    cells: Vec<CellData>,
}

struct CellData {
    col: u16,
    row: u16,
    col_span: u16,
    row_span: u16,
    text: String,
}

fn is_compressed(comp: &mut CompoundFile<std::fs::File>) -> Result<bool, String> {
    let data = read_stream(comp, "FileHeader")?;
    if data.len() < 40 { return Err("FileHeader가 너무 짧습니다".to_string()); }
    Ok(u32::from_le_bytes([data[36], data[37], data[38], data[39]]) & 0x01 != 0)
}

fn read_stream(comp: &mut CompoundFile<std::fs::File>, name: &str) -> Result<Vec<u8>, String> {
    let mut stream = comp.open_stream(name)
        .map_err(|e| format!("스트림 '{}' 열기 실패: {}", name, e))?;
    let mut buf = Vec::new();
    stream.read_to_end(&mut buf).map_err(|e| format!("스트림 읽기 실패: {}", e))?;
    Ok(buf)
}

fn decompress(data: &[u8]) -> Result<Vec<u8>, String> {
    let mut decoder = DeflateDecoder::new(data);
    let mut out = Vec::new();
    decoder.read_to_end(&mut out).map_err(|e| format!("압축 해제 실패: {}", e))?;
    Ok(out)
}

struct Record { tag_id: u16, level: u16, data: Vec<u8> }

fn parse_records(data: &[u8]) -> Vec<Record> {
    let mut records = Vec::new();
    let mut pos = 0;
    while pos + 4 <= data.len() {
        let h = u32::from_le_bytes([data[pos], data[pos+1], data[pos+2], data[pos+3]]);
        pos += 4;
        let tag_id = (h & 0x3FF) as u16;
        let level = ((h >> 10) & 0x3FF) as u16;
        let mut size = ((h >> 20) & 0xFFF) as usize;
        if size == 0xFFF {
            if pos + 4 > data.len() { break; }
            size = u32::from_le_bytes([data[pos], data[pos+1], data[pos+2], data[pos+3]]) as usize;
            pos += 4;
        }
        if pos + size > data.len() { break; }
        records.push(Record { tag_id, level, data: data[pos..pos+size].to_vec() });
        pos += size;
    }
    records
}

/// CTRL_HEADER가 테이블인지 확인 (ctrl_type == "tbl ")
fn is_table_ctrl(data: &[u8]) -> bool {
    data.len() >= 4 && data[0..4] == TBL_CTRL_TYPE
}

fn parse_section(data: &[u8]) -> Result<Vec<Element>, String> {
    let records = parse_records(data);
    let mut elements: Vec<Element> = Vec::new();
    let mut i = 0;
    let mut skip_until: Option<usize> = None;

    while i < records.len() {
        if let Some(end) = skip_until {
            if i < end { i += 1; continue; } else { skip_until = None; }
        }

        let rec = &records[i];

        // CTRL_HEADER + "tbl " = 표 시작
        if rec.tag_id == TAG_CTRL_HEADER && is_table_ctrl(&rec.data) {
            let (table, consumed) = parse_table(&records, i);
            elements.push(Element::Table(table));
            skip_until = Some(i + consumed);
            i += 1;
            continue;
        }

        // 일반 텍스트 (표 밖)
        if rec.tag_id == TAG_PARA_TEXT {
            let text = extract_para_text(&rec.data);
            if !text.is_empty() {
                elements.push(Element::Paragraph(text));
            }
        }

        i += 1;
    }

    Ok(elements)
}

fn parse_table(records: &[Record], start: usize) -> (TableData, usize) {
    let ctrl_level = records[start].level;
    let mut cells: Vec<CellData> = Vec::new();
    let mut current_cell_text = String::new();
    let mut in_cell = false;
    let mut cur_col: u16 = 0;
    let mut cur_row: u16 = 0;
    let mut cur_cspan: u16 = 1;
    let mut cur_rspan: u16 = 1;
    let mut i = start + 1;

    while i < records.len() {
        let rec = &records[i];

        // 표 레벨 이하로 나가면 종료
        if rec.level <= ctrl_level {
            break;
        }

        // LIST_HEADER = 새 셀
        if rec.tag_id == TAG_LIST_HEADER && rec.level == ctrl_level + 1 {
            // 이전 셀 저장
            if in_cell {
                cells.push(CellData {
                    col: cur_col, row: cur_row,
                    col_span: cur_cspan, row_span: cur_rspan,
                    text: current_cell_text.trim().to_string(),
                });
            }
            current_cell_text.clear();
            in_cell = true;

            // 셀 위치 파싱 (offset 8: col, 10: row, 12: colSpan, 14: rowSpan)
            if rec.data.len() >= 16 {
                cur_col = u16::from_le_bytes([rec.data[8], rec.data[9]]);
                cur_row = u16::from_le_bytes([rec.data[10], rec.data[11]]);
                cur_cspan = u16::from_le_bytes([rec.data[12], rec.data[13]]);
                cur_rspan = u16::from_le_bytes([rec.data[14], rec.data[15]]);
                // 유효성 검사
                if cur_cspan == 0 { cur_cspan = 1; }
                if cur_rspan == 0 { cur_rspan = 1; }
                if cur_col > 50 || cur_row > 500 { cur_col = 0; cur_row = 0; }
            }
        }

        // 셀 안의 텍스트
        if rec.tag_id == TAG_PARA_TEXT && in_cell && rec.level > ctrl_level {
            let text = extract_para_text(&rec.data);
            if !text.is_empty() {
                if !current_cell_text.is_empty() {
                    current_cell_text.push_str("<br>");
                }
                current_cell_text.push_str(&text);
            }
        }

        i += 1;
    }

    // 마지막 셀
    if in_cell {
        cells.push(CellData {
            col: cur_col, row: cur_row,
            col_span: cur_cspan, row_span: cur_rspan,
            text: current_cell_text.trim().to_string(),
        });
    }

    (TableData { cells }, i - start)
}

fn extract_para_text(data: &[u8]) -> String {
    if data.len() < 2 { return String::new(); }
    let mut text = String::new();
    let mut i = 0;
    while i + 1 < data.len() {
        let ch = u16::from_le_bytes([data[i], data[i+1]]);
        i += 2;
        match ch {
            1..=3 | 11 | 12 => { i += 14; }
            4..=8 => { i += 14; }
            9 => { text.push('\t'); }
            10 => { text.push('\n'); }
            13 | 0 | 14..=23 | 25..=29 => {}
            24 => { text.push('-'); }
            30 | 31 => { text.push(' '); }
            _ => {
                if let Some(c) = char::from_u32(ch as u32) {
                    if !c.is_control() { text.push(c); }
                }
            }
        }
    }
    text.trim().to_string()
}

fn render_markdown(elements: &[Element]) -> String {
    let mut result: Vec<String> = Vec::new();
    let mut prev_blank = false;

    for elem in elements {
        match elem {
            Element::Paragraph(text) => {
                let t = text.trim();
                if t.is_empty() {
                    if !prev_blank && !result.is_empty() { result.push(String::new()); prev_blank = true; }
                } else { prev_blank = false; result.push(t.to_string()); }
            }
            Element::Table(table) => {
                prev_blank = false;
                let md = render_table(table);
                if !md.is_empty() {
                    result.push(String::new());
                    result.push(md);
                    result.push(String::new());
                }
            }
        }
    }
    while result.last().map_or(false, |l| l.is_empty()) { result.pop(); }
    result.join("\n")
}

fn render_table(table: &TableData) -> String {
    if table.cells.is_empty() { return String::new(); }

    // 행/열 수를 셀 위치에서 계산
    let max_row = table.cells.iter().map(|c| c.row + c.row_span).max().unwrap_or(1) as usize;
    let max_col = table.cells.iter().map(|c| c.col + c.col_span).max().unwrap_or(1) as usize;

    if max_col == 0 || max_row == 0 { return String::new(); }

    // 그리드 생성
    let mut grid: Vec<Vec<String>> = vec![vec![String::new(); max_col]; max_row];
    let mut merged: Vec<Vec<bool>> = vec![vec![false; max_col]; max_row];

    for cell in &table.cells {
        let r = cell.row as usize;
        let c = cell.col as usize;
        if r >= max_row || c >= max_col { continue; }

        grid[r][c] = cell.text.clone();

        // 병합 영역 마킹
        let cspan = (cell.col_span as usize).max(1);
        let rspan = (cell.row_span as usize).max(1);
        for dr in 0..rspan {
            for dc in 0..cspan {
                if dr == 0 && dc == 0 { continue; }
                let mr = r + dr;
                let mc = c + dc;
                if mr < max_row && mc < max_col {
                    merged[mr][mc] = true;
                }
            }
        }
    }

    // Markdown 표 생성
    let mut lines: Vec<String> = Vec::new();

    for (row_idx, row) in grid.iter().enumerate() {
        let mut cells_str: Vec<String> = Vec::new();
        for (col_idx, cell_text) in row.iter().enumerate() {
            if merged[row_idx][col_idx] {
                cells_str.push(String::from("  "));
            } else {
                let cleaned = cell_text.replace('|', "\\|");
                cells_str.push(format!(" {} ", cleaned));
            }
        }
        lines.push(format!("|{}|", cells_str.join("|")));

        if row_idx == 0 {
            let sep: Vec<String> = (0..max_col).map(|_| " --- ".to_string()).collect();
            lines.push(format!("|{}|", sep.join("|")));
        }
    }

    lines.join("\n")
}
