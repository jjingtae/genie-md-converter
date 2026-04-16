use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::Read;
use std::path::Path;

pub fn convert(path: &Path) -> Result<String, String> {
    let file = std::fs::File::open(path).map_err(|e| format!("HWPX 파일 열기 실패: {}", e))?;
    let mut archive = zip::ZipArchive::new(file).map_err(|e| format!("HWPX ZIP 열기 실패: {}", e))?;

    let mut section_names: Vec<String> = Vec::new();
    for i in 0..archive.len() {
        if let Ok(entry) = archive.by_index(i) {
            let name = entry.name().to_string();
            if name.starts_with("Contents/section") && name.ends_with(".xml") {
                section_names.push(name);
            }
        }
    }
    section_names.sort();

    if section_names.is_empty() {
        return Err("HWPX에서 섹션 파일을 찾을 수 없습니다".to_string());
    }

    let mut all_output: Vec<String> = Vec::new();
    for section_name in &section_names {
        let xml_content = {
            let mut entry = archive.by_name(section_name)
                .map_err(|e| format!("{} 열기 실패: {}", section_name, e))?;
            let mut buf = String::new();
            entry.read_to_string(&mut buf).map_err(|e| format!("XML 읽기 실패: {}", e))?;
            buf
        };
        let parsed = parse_section_xml(&xml_content)?;
        all_output.extend(parsed);
    }

    let mut result: Vec<String> = Vec::new();
    let mut prev_blank = false;
    for line in &all_output {
        if line.is_empty() {
            if !prev_blank && !result.is_empty() { result.push(String::new()); prev_blank = true; }
        } else { prev_blank = false; result.push(line.clone()); }
    }
    while result.last().map_or(false, |l| l.is_empty()) { result.pop(); }
    Ok(result.join("\n"))
}

fn strip_ns(name: &[u8]) -> &[u8] {
    if let Some(pos) = name.iter().position(|&b| b == b':') {
        &name[pos + 1..]
    } else {
        name
    }
}

fn parse_section_xml(xml: &str) -> Result<Vec<String>, String> {
    let mut reader = Reader::from_str(xml);
    reader.trim_text(false);

    let mut output: Vec<String> = Vec::new();
    let mut current_line = String::new();
    let mut in_paragraph = false;
    let mut in_text = false;
    let mut is_bold = false;
    let mut is_italic = false;

    let mut in_cell = false;
    let mut current_cell = String::new();
    let mut table_cells: Vec<CellInfo> = Vec::new();
    let mut cur_col: u16 = 0;
    let mut cur_row: u16 = 0;
    let mut cur_colspan: u16 = 1;
    let mut cur_rowspan: u16 = 1;
    let mut cell_col_counter: u16 = 0;
    let mut cell_row_counter: u16 = 0;

    let mut buf: Vec<u8> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name_bytes = e.name();
                let local = strip_ns(name_bytes.as_ref());
                match local {
                    b"p" => {
                        in_paragraph = true;
                        if !in_cell { current_line.clear(); }
                    }
                    b"run" => { is_bold = false; is_italic = false; }
                    b"t" => { in_text = true; }
                    b"bold" => { is_bold = true; }
                    b"italic" => { is_italic = true; }
                    b"tbl" => {
                        table_cells.clear();
                        cell_row_counter = 0;
                    }
                    b"tr" => { cell_col_counter = 0; }
                    b"tc" => {
                        in_cell = true;
                        current_cell.clear();
                        cur_col = cell_col_counter;
                        cur_row = cell_row_counter;
                        cur_colspan = 1;
                        cur_rowspan = 1;
                    }
                    b"cellSpan" => {
                        for attr in e.attributes().flatten() {
                            let key_bytes = attr.key;
                            let key = strip_ns(key_bytes.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            if key == b"colSpan" { cur_colspan = val.parse::<u16>().unwrap_or(1).max(1); }
                            if key == b"rowSpan" { cur_rowspan = val.parse::<u16>().unwrap_or(1).max(1); }
                        }
                    }
                    b"cellAddr" => {
                        for attr in e.attributes().flatten() {
                            let key_bytes = attr.key;
                            let key = strip_ns(key_bytes.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            if key == b"colAddr" { cur_col = val.parse::<u16>().unwrap_or(0); }
                            if key == b"rowAddr" { cur_row = val.parse::<u16>().unwrap_or(0); }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name_bytes = e.name();
                let local = strip_ns(name_bytes.as_ref());
                match local {
                    b"bold" => { is_bold = true; }
                    b"italic" => { is_italic = true; }
                    b"cellSpan" => {
                        for attr in e.attributes().flatten() {
                            let key_bytes = attr.key;
                            let key = strip_ns(key_bytes.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            if key == b"colSpan" { cur_colspan = val.parse::<u16>().unwrap_or(1).max(1); }
                            if key == b"rowSpan" { cur_rowspan = val.parse::<u16>().unwrap_or(1).max(1); }
                        }
                    }
                    b"cellAddr" => {
                        for attr in e.attributes().flatten() {
                            let key_bytes = attr.key;
                            let key = strip_ns(key_bytes.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            if key == b"colAddr" { cur_col = val.parse::<u16>().unwrap_or(0); }
                            if key == b"rowAddr" { cur_row = val.parse::<u16>().unwrap_or(0); }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if !in_text { buf.clear(); continue; }
                let raw = e.unescape().unwrap_or_default().to_string();
                if raw.is_empty() { buf.clear(); continue; }
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
                let name_bytes = e.name();
                let local = strip_ns(name_bytes.as_ref());
                match local {
                    b"t" => { in_text = false; }
                    b"run" => {}
                    b"p" => {
                        if in_cell {
                            if !current_cell.is_empty() { current_cell.push_str("<br>"); }
                        } else if in_paragraph {
                            let line = current_line.trim().to_string();
                            output.push(line);
                            current_line.clear();
                        }
                        in_paragraph = false;
                    }
                    b"tc" => {
                        in_cell = false;
                        let cell_text = current_cell.trim_end_matches("<br>").trim().to_string();
                        table_cells.push(CellInfo {
                            col: cur_col, row: cur_row,
                            col_span: cur_colspan, row_span: cur_rowspan,
                            text: cell_text,
                        });
                        current_cell.clear();
                        cell_col_counter = cur_col + cur_colspan;
                    }
                    b"tr" => { cell_row_counter += 1; }
                    b"tbl" => {
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

    Ok(output)
}

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

struct CellInfo {
    col: u16,
    row: u16,
    col_span: u16,
    row_span: u16,
    text: String,
}

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