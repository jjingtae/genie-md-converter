use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::Read;
use std::path::Path;

pub fn convert(path: &Path) -> Result<String, String> {
    let file = std::fs::File::open(path)
        .map_err(|e| format!("HWPX 파일 열기 실패: {}", e))?;
    let mut archive = zip::ZipArchive::new(file)
        .map_err(|e| format!("HWPX ZIP 읽기 실패: {}", e))?;

    let mut all_md = String::new();
    let mut section_idx = 0;

    loop {
        let name = format!("Contents/section{}.xml", section_idx);
        match archive.by_name(&name) {
            Ok(mut entry) => {
                let mut buf = String::new();
                entry.read_to_string(&mut buf)
                    .map_err(|e| format!("XML 읽기 실패: {}", e))?;
                let md = parse_section_xml(&buf)?;
                if !md.is_empty() {
                    if !all_md.is_empty() { all_md.push('\n'); }
                    all_md.push_str(&md);
                }
                section_idx += 1;
            }
            Err(_) => break,
        }
    }

    if section_idx == 0 {
        return Err("HWPX 섹션을 찾을 수 없습니다".to_string());
    }

    Ok(all_md)
}

/// 네임스페이스 접두사 제거 (hp:p → p, w:t → t)
fn local_name(name: &[u8]) -> &[u8] {
    if let Some(pos) = name.iter().position(|&b| b == b':') {
        &name[pos + 1..]
    } else {
        name
    }
}

/// 속성에서 local name으로 값 추출
fn get_attr(e: &quick_xml::events::BytesStart, attr_local: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = local_name(attr.key.as_ref());
        if key == attr_local {
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

fn get_attr_u16(e: &quick_xml::events::BytesStart, attr_local: &[u8]) -> u16 {
    get_attr(e, attr_local)
        .and_then(|v| v.parse::<u16>().ok())
        .unwrap_or(0)
}

#[derive(Clone)]
struct CellInfo {
    col: u16,
    row: u16,
    col_span: u16,
    row_span: u16,
    text: String,
}

fn parse_section_xml(xml: &str) -> Result<String, String> {
    let mut reader = Reader::from_str(xml);
    reader.trim_text(false);

    let mut output: Vec<String> = Vec::new();
    let mut current_line = String::new();
    let mut in_paragraph = false;
    let mut is_bold = false;
    let mut is_italic = false;
    let mut in_text = false;
    let mut in_table = false;
    let mut in_cell = false;
    let mut current_cell = String::new();
    let mut table_cells: Vec<CellInfo> = Vec::new();
    let mut cur_col: u16 = 0;
    let mut cur_row: u16 = 0;
    let mut cur_cspan: u16 = 1;
    let mut cur_rspan: u16 = 1;
    let mut buf: Vec<u8> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let ln = local_name(e.name().as_ref());

                match ln {
                    b"p" => {
                        in_paragraph = true;
                        if !in_cell {
                            current_line.clear();
                        }
                    }
                    b"run" => {
                        is_bold = false;
                        is_italic = false;
                    }
                    b"t" => { in_text = true; }
                    b"bold" => { is_bold = true; }
                    b"italic" => { is_italic = true; }
                    b"tbl" => {
                        in_table = true;
                        table_cells.clear();
                    }
                    b"tc" => {
                        in_cell = true;
                        current_cell.clear();
                        cur_col = 0;
                        cur_row = 0;
                        cur_cspan = 1;
                        cur_rspan = 1;
                    }
                    b"cellAddr" => {
                        cur_col = get_attr_u16(e, b"colAddr");
                        cur_row = get_attr_u16(e, b"rowAddr");
                    }
                    b"cellSpan" => {
                        cur_cspan = get_attr_u16(e, b"colSpan").max(1);
                        cur_rspan = get_attr_u16(e, b"rowSpan").max(1);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let ln = local_name(e.name().as_ref());

                match ln {
                    b"bold" => { is_bold = true; }
                    b"italic" => { is_italic = true; }
                    b"cellAddr" => {
                        cur_col = get_attr_u16(e, b"colAddr");
                        cur_row = get_attr_u16(e, b"rowAddr");
                    }
                    b"cellSpan" => {
                        cur_cspan = get_attr_u16(e, b"colSpan").max(1);
                        cur_rspan = get_attr_u16(e, b"rowSpan").max(1);
                    }
                    b"br" | b"lineBreak" => {
                        if in_cell {
                            current_cell.push_str("<br>");
                        } else {
                            current_line.push('\n');
                        }
                    }
                    b"tab" => {
                        if in_cell { current_cell.push('\t'); }
                        else { current_line.push('\t'); }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if !in_text { buf.clear(); continue; }
                let raw = e.unescape().unwrap_or_default().to_string();
                let text = raw.trim().to_string();
                if text.is_empty() { buf.clear(); continue; }

                let formatted = if is_bold && is_italic { format!("***{}***", text) }
                    else if is_bold { format!("**{}**", text) }
                    else if is_italic { format!("*{}*", text) }
                    else { text };

                if in_cell {
                    add_with_space(&mut current_cell, &formatted);
                } else {
                    add_with_space(&mut current_line, &formatted);
                }
            }
            Ok(Event::End(ref e)) => {
                let ln = local_name(e.name().as_ref());

                match ln {
                    b"t" => { in_text = false; }
                    b"p" => {
                        if in_cell {
                            if !current_cell.is_empty() {
                                current_cell.push_str("<br>");
                            }
                        } else if in_paragraph {
                            let line = current_line.trim().to_string();
                            output.push(line);
                        }
                        in_paragraph = false;
                        if !in_cell { current_line.clear(); }
                    }
                    b"tc" => {
                        in_cell = false;
                        // <br> 끝에 남은 거 제거
                        let cell_text = current_cell.trim().to_string();
                        let cell_text = cell_text.trim_end_matches("<br>").to_string();
                        table_cells.push(CellInfo {
                            col: cur_col,
                            row: cur_row,
                            col_span: cur_cspan,
                            row_span: cur_rspan,
                            text: cell_text,
                        });
                        current_cell.clear();
                    }
                    b"tbl" => {
                        in_table = false;
                        if !table_cells.is_empty() {
                            let md = render_table(&table_cells);
                            if !md.is_empty() {
                                output.push(String::new());
                                output.push(md);
                                output.push(String::new());
                            }
                        }
                        table_cells.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("HWPX XML 파싱 오류: {}", e)),
            _ => {}
        }
        buf.clear();
    }

    // 빈 줄 정리
    let mut result: Vec<String> = Vec::new();
    let mut prev_blank = false;
    for line in &output {
        if line.is_empty() {
            if !prev_blank && !result.is_empty() { result.push(String::new()); prev_blank = true; }
        } else { prev_blank = false; result.push(line.clone()); }
    }
    Ok(result.join("\n"))
}

/// 단어 사이 공백 자동 삽입
fn add_with_space(target: &mut String, text: &str) {
    if target.is_empty() || text.is_empty() { target.push_str(text); return; }
    let last = target.chars().last().unwrap_or(' ');
    let first = text.chars().next().unwrap_or(' ');
    if (last.is_alphanumeric() || last == '*' || last == '"' || last == '\u{201D}')
        && (first.is_alphanumeric() || first == '*' || first == '"' || first == '\u{201C}') {
        target.push(' ');
    }
    target.push_str(text);
}

/// 셀 정보로 표 렌더링 (^/> 마커 기반 병합)
fn render_table(cells: &[CellInfo]) -> String {
    if cells.is_empty() { return String::new(); }

    let max_row = cells.iter().map(|c| c.row + c.row_span).max().unwrap_or(1) as usize;
    let max_col = cells.iter().map(|c| c.col + c.col_span).max().unwrap_or(1) as usize;

    if max_col == 0 || max_row == 0 { return String::new(); }

    let mut grid: Vec<Vec<String>> = vec![vec![String::new(); max_col]; max_row];

    for cell in cells {
        let r = cell.row as usize;
        let c = cell.col as usize;
        if r >= max_row || c >= max_col { continue; }

        grid[r][c] = cell.text.clone();

        let cspan = (cell.col_span as usize).max(1);
        let rspan = (cell.row_span as usize).max(1);
        for dr in 0..rspan {
            for dc in 0..cspan {
                if dr == 0 && dc == 0 { continue; }
                let mr = r + dr;
                let mc = c + dc;
                if mr < max_row && mc < max_col {
                    if dr == 0 && dc > 0 {
                        grid[mr][mc] = ">".to_string();
                    } else {
                        grid[mr][mc] = "^".to_string();
                    }
                }
            }
        }
    }

    let mut lines: Vec<String> = Vec::new();
    for (row_idx, row) in grid.iter().enumerate() {
        let cells_str: Vec<String> = row.iter()
            .map(|c| {
                let cleaned = c.replace('|', "\\|");
                format!(" {} ", cleaned)
            })
            .collect();
        lines.push(format!("|{}|", cells_str.join("|")));

        if row_idx == 0 {
            let sep: Vec<String> = (0..max_col).map(|_| " --- ".to_string()).collect();
            lines.push(format!("|{}|", sep.join("|")));
        }
    }

    lines.join("\n")
}
