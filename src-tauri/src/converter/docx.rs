use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::Read;
use std::path::Path;

/// DOCX 파일을 Markdown으로 변환
pub fn convert(path: &Path) -> Result<String, String> {
    let file = std::fs::File::open(path)
        .map_err(|e| format!("파일 열기 실패: {}", e))?;

    let mut archive = zip::ZipArchive::new(file)
        .map_err(|e| format!("DOCX ZIP 읽기 실패: {}", e))?;

    let xml_content = {
        let mut entry = archive
            .by_name("word/document.xml")
            .map_err(|_| "word/document.xml을 찾을 수 없습니다".to_string())?;
        let mut buf = String::new();
        entry.read_to_string(&mut buf)
            .map_err(|e| format!("XML 읽기 실패: {}", e))?;
        buf
    };

    let md = parse_document_xml(&xml_content)?;
    Ok(md)
}

fn parse_document_xml(xml: &str) -> Result<String, String> {
    let mut reader = Reader::from_str(xml);
    reader.trim_text(true);

    let mut output: Vec<String> = Vec::new();
    let mut current_line = String::new();
    let mut in_paragraph = false;
    let mut is_bold = false;
    let mut is_italic = false;
    let mut is_heading = false;
    let mut heading_level: u8 = 0;
    let mut in_table = false;
    let mut table_row: Vec<String> = Vec::new();
    let mut table_rows: Vec<Vec<String>> = Vec::new();
    let mut current_cell = String::new();
    let mut in_cell = false;
    let mut list_level: i32 = -1;
    let mut buf: Vec<u8> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                handle_start_tag(e.name().as_ref(), &e, &mut in_paragraph, &mut current_line,
                    &mut is_heading, &mut heading_level, &mut list_level,
                    &mut is_bold, &mut is_italic, &mut in_table, &mut table_rows,
                    &mut table_row, &mut in_cell, &mut current_cell);
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                let name_bytes = name.as_ref();
                if name_bytes == b"w:b" { is_bold = true; }
                if name_bytes == b"w:i" { is_italic = true; }
                if name_bytes == b"w:br" {
                    if in_cell { current_cell.push(' '); }
                    else { current_line.push('\n'); }
                }
                if name_bytes == b"w:pStyle" {
                    check_style(&e, &mut is_heading, &mut heading_level, &mut list_level);
                }
                if name_bytes == b"w:ilvl" {
                    check_ilvl(&e, &mut list_level);
                }
            }
            Ok(Event::Text(ref e)) => {
                let text = e.unescape().unwrap_or_default().to_string();
                if text.is_empty() { buf.clear(); continue; }

                let formatted = if is_bold && is_italic {
                    format!("***{}***", text)
                } else if is_bold {
                    format!("**{}**", text)
                } else if is_italic {
                    format!("*{}*", text)
                } else {
                    text
                };

                if in_cell {
                    current_cell.push_str(&formatted);
                } else {
                    current_line.push_str(&formatted);
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                let name_bytes = name.as_ref();

                if name_bytes == b"w:r" {
                    is_bold = false;
                    is_italic = false;
                }
                if name_bytes == b"w:p" {
                    if in_cell {
                        // 셀 내 단락
                    } else if in_paragraph {
                        let line = current_line.trim().to_string();
                        if is_heading && !line.is_empty() {
                            let prefix = "#".repeat(heading_level as usize);
                            output.push(format!("{} {}", prefix, line));
                        } else if list_level >= 0 && !line.is_empty() {
                            let indent = "  ".repeat(list_level as usize);
                            output.push(format!("{}- {}", indent, line));
                        } else {
                            output.push(line);
                        }
                    }
                    in_paragraph = false;
                    current_line.clear();
                }
                if name_bytes == b"w:tc" {
                    in_cell = false;
                    table_row.push(current_cell.trim().to_string());
                    current_cell.clear();
                }
                if name_bytes == b"w:tr" {
                    table_rows.push(table_row.clone());
                    table_row.clear();
                }
                if name_bytes == b"w:tbl" {
                    in_table = false;
                    if !table_rows.is_empty() {
                        let md_table = render_table(&table_rows);
                        output.push(md_table);
                    }
                    table_rows.clear();
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML 파싱 오류: {}", e)),
            _ => {}
        }
        buf.clear();
    }

    // 빈 줄 정리
    let mut result: Vec<String> = Vec::new();
    let mut prev_blank = false;
    for line in &output {
        if line.is_empty() {
            if !prev_blank {
                result.push(String::new());
                prev_blank = true;
            }
        } else {
            prev_blank = false;
            result.push(line.clone());
        }
    }

    Ok(result.join("\n"))
}

#[allow(clippy::too_many_arguments)]
fn handle_start_tag(
    name_bytes: &[u8],
    e: &quick_xml::events::BytesStart,
    in_paragraph: &mut bool,
    current_line: &mut String,
    is_heading: &mut bool,
    heading_level: &mut u8,
    list_level: &mut i32,
    is_bold: &mut bool,
    is_italic: &mut bool,
    in_table: &mut bool,
    table_rows: &mut Vec<Vec<String>>,
    table_row: &mut Vec<String>,
    in_cell: &mut bool,
    current_cell: &mut String,
) {
    if name_bytes == b"w:p" {
        *in_paragraph = true;
        current_line.clear();
        *is_heading = false;
        *heading_level = 0;
        *list_level = -1;
    }
    if name_bytes == b"w:pStyle" {
        check_style(e, is_heading, heading_level, list_level);
    }
    if name_bytes == b"w:ilvl" {
        check_ilvl(e, list_level);
    }
    if name_bytes == b"w:b" { *is_bold = true; }
    if name_bytes == b"w:i" { *is_italic = true; }
    if name_bytes == b"w:tbl" {
        *in_table = true;
        table_rows.clear();
    }
    if name_bytes == b"w:tr" {
        table_row.clear();
    }
    if name_bytes == b"w:tc" {
        *in_cell = true;
        current_cell.clear();
    }
}

fn check_style(e: &quick_xml::events::BytesStart, is_heading: &mut bool, heading_level: &mut u8, list_level: &mut i32) {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == b"w:val" {
            let val = String::from_utf8_lossy(&attr.value).to_string();
            if val.starts_with("Heading") || val.starts_with("heading") {
                *is_heading = true;
                *heading_level = val
                    .chars()
                    .last()
                    .and_then(|c| c.to_digit(10))
                    .unwrap_or(1) as u8;
                if *heading_level > 6 { *heading_level = 6; }
            } else if val.contains("List") || val.contains("list") {
                *list_level = 0;
            }
        }
    }
}

fn check_ilvl(e: &quick_xml::events::BytesStart, list_level: &mut i32) {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == b"w:val" {
            let val = String::from_utf8_lossy(&attr.value);
            *list_level = val.parse::<i32>().unwrap_or(0);
        }
    }
}

fn render_table(rows: &[Vec<String>]) -> String {
    if rows.is_empty() { return String::new(); }

    let max_cols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    if max_cols == 0 { return String::new(); }

    let mut col_widths = vec![3usize; max_cols];
    for row in rows {
        for (i, cell) in row.iter().enumerate() {
            if i < max_cols {
                col_widths[i] = col_widths[i].max(cell.len()).max(3);
            }
        }
    }

    let mut lines: Vec<String> = Vec::new();

    for (row_idx, row) in rows.iter().enumerate() {
        let mut cells: Vec<String> = Vec::new();
        for i in 0..max_cols {
            let cell = row.get(i).cloned().unwrap_or_default();
            cells.push(format!(" {} ", cell));
        }
        lines.push(format!("|{}|", cells.join("|")));

        if row_idx == 0 {
            let sep: Vec<String> = col_widths
                .iter()
                .map(|w| format!(" {} ", "-".repeat(*w)))
                .collect();
            lines.push(format!("|{}|", sep.join("|")));
        }
    }

    lines.join("\n")
}
