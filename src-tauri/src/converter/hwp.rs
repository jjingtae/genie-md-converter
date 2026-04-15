use cfb::CompoundFile;
use flate2::read::DeflateDecoder;
use std::io::Read;
use std::path::Path;

const HWPTAG_BEGIN: u16 = 0x010;
const HWPTAG_PARA_TEXT: u16 = HWPTAG_BEGIN + 51;    // 67
const HWPTAG_TABLE: u16 = HWPTAG_BEGIN + 80;        // 96
const HWPTAG_LIST_HEADER: u16 = HWPTAG_BEGIN + 79;  // 95

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
    rows: u16,
    cols: u16,
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
    let flags = u32::from_le_bytes([data[36], data[37], data[38], data[39]]);
    Ok(flags & 0x01 != 0)
}

fn read_stream(comp: &mut CompoundFile<std::fs::File>, name: &str) -> Result<Vec<u8>, String> {
    let mut stream = comp.open_stream(name).map_err(|e| format!("스트림 '{}' 열기 실패: {}", name, e))?;
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

        if rec.tag_id == HWPTAG_TABLE {
            let (table, consumed) = parse_table(&records, i);
            elements.push(Element::Table(table));
            skip_until = Some(i + consumed);
            i += 1;
            continue;
        }

        if rec.tag_id == HWPTAG_PARA_TEXT {
            let text = extract_para_text(&rec.data);
            if !text.is_empty() {
                elements.push(Element::Paragraph(text));
            }
        }
        i += 1;
    }
    Ok(elements)
}

/// LIST_HEADER에서 셀 위치/병합 정보 추출
fn parse_cell_info(data: &[u8]) -> (u16, u16, u16, u16) {
    // LIST_HEADER 공통: paraCount(2) + flags(4) + unknown(8) = 14바이트
    // 셀 정보: col(2) + row(2) + colSpan(2) + rowSpan(2) = 8바이트
    // 총 최소 22바이트 필요
    if data.len() >= 22 {
        let col = u16::from_le_bytes([data[14], data[15]]);
        let row = u16::from_le_bytes([data[16], data[17]]);
        let col_span = u16::from_le_bytes([data[18], data[19]]);
        let row_span = u16::from_le_bytes([data[20], data[21]]);
        // 유효성 검사: 비정상적 값이면 기본값 반환
        if col < 100 && row < 1000 && col_span > 0 && col_span <= 50 && row_span > 0 && row_span <= 100 {
            return (col, row, col_span, row_span);
        }
    }
    // 파싱 실패 시 기본값 (순차 배치에서 덮어씀)
    (0, 0, 1, 1)
}

fn parse_table(records: &[Record], start: usize) -> (TableData, usize) {
    let table_rec = &records[start];
    let table_level = table_rec.level;

    let (rows, cols) = if table_rec.data.len() >= 8 {
        let r = u16::from_le_bytes([table_rec.data[4], table_rec.data[5]]);
        let c = u16::from_le_bytes([table_rec.data[6], table_rec.data[7]]);
        (r.max(1), c.max(1))
    } else {
        (1, 1)
    };

    let mut cells: Vec<CellData> = Vec::new();
    let mut current_cell_text = String::new();
    let mut in_cell = false;
    let mut current_col: u16 = 0;
    let mut current_row: u16 = 0;
    let mut current_col_span: u16 = 1;
    let mut current_row_span: u16 = 1;
    let mut cell_index: usize = 0;
    let mut i = start + 1;

    while i < records.len() {
        let rec = &records[i];

        if rec.level <= table_level && rec.tag_id != HWPTAG_LIST_HEADER {
            break;
        }

        if rec.tag_id == HWPTAG_LIST_HEADER && rec.level == table_level + 1 {
            // 이전 셀 저장
            if in_cell {
                cells.push(CellData {
                    col: current_col,
                    row: current_row,
                    col_span: current_col_span,
                    row_span: current_row_span,
                    text: current_cell_text.trim().to_string(),
                });
            }
            current_cell_text.clear();
            in_cell = true;

            // LIST_HEADER에서 셀 위치 파싱 시도
            let (parsed_col, parsed_row, parsed_cspan, parsed_rspan) = parse_cell_info(&rec.data);

            // 파싱된 값이 유효하면 사용, 아니면 순차 계산
            if rec.data.len() >= 22 && parsed_col < cols && parsed_row < rows {
                current_col = parsed_col;
                current_row = parsed_row;
                current_col_span = parsed_cspan;
                current_row_span = parsed_rspan;
            } else {
                // 순차 배치: 왼→오, 위→아래
                current_row = (cell_index as u16) / cols;
                current_col = (cell_index as u16) % cols;
                current_col_span = 1;
                current_row_span = 1;
            }
            cell_index += 1;
        }

        if rec.tag_id == HWPTAG_PARA_TEXT && in_cell && rec.level > table_level {
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
            col: current_col,
            row: current_row,
            col_span: current_col_span,
            row_span: current_row_span,
            text: current_cell_text.trim().to_string(),
        });
    }

    (TableData { rows, cols, cells }, i - start)
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
            13 => {}
            0 | 14..=23 | 25..=29 => {}
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
                } else {
                    prev_blank = false;
                    result.push(t.to_string());
                }
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
    let cols = table.cols as usize;
    let rows = table.rows as usize;
    if cols == 0 || rows == 0 || table.cells.is_empty() { return String::new(); }

    // 그리드 생성 (병합 고려)
    let mut grid: Vec<Vec<String>> = vec![vec![String::new(); cols]; rows];
    let mut merged: Vec<Vec<bool>> = vec![vec![false; cols]; rows]; // 병합으로 가려진 셀

    for cell in &table.cells {
        let r = cell.row as usize;
        let c = cell.col as usize;
        if r >= rows || c >= cols { continue; }

        grid[r][c] = cell.text.clone();

        // 셀 병합 처리: 병합된 영역을 마킹
        let cspan = (cell.col_span as usize).max(1);
        let rspan = (cell.row_span as usize).max(1);
        for dr in 0..rspan {
            for dc in 0..cspan {
                if dr == 0 && dc == 0 { continue; } // 원본 셀은 스킵
                let mr = r + dr;
                let mc = c + dc;
                if mr < rows && mc < cols {
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
                // 병합으로 가려진 셀은 빈칸
                cells_str.push(String::from("  "));
            } else {
                let cleaned = cell_text.replace('|', "\\|");
                cells_str.push(format!(" {} ", cleaned));
            }
        }
        lines.push(format!("|{}|", cells_str.join("|")));

        if row_idx == 0 {
            let sep: Vec<String> = (0..cols).map(|_| " --- ".to_string()).collect();
            lines.push(format!("|{}|", sep.join("|")));
        }
    }

    lines.join("\n")
}
