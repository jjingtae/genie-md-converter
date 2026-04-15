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
    // trim_text를 false로 설정하여 공백 보존
    let mut reader = Reader::from_str(xml);
    reader.trim_text(false);

    let mut output: Vec<String> = Vec::new();
    let mut current_line = String::new();
    let mut in_paragraph = false;
    let mut is_bold = false;
    let mut is_italic = false;
    let mut is_heading = false;
    let mut heading_level: u8 = 0;
    let mut _in_table = false;
    let mut table_row: Vec<String> = Vec::new();
    let mut table_rows: Vec<Vec<String>> = Vec::new();
    let mut current_cell = String::new();
    let mut in_cell = false;
    let mut list_level: i32 = -1;
    let mut _in_run = false;
    let mut in_text = false;
    let mut preserve_space = false;
    let mut buf: Vec<u8> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let name_bytes = name.as_ref();

                if name_bytes == b"w:p" {
                    in_paragraph = true;
                    current_line.clear();
                    is_heading = false;
                    heading_level = 0;
                    list_level = -1;
                }
                if name_bytes == b"w:pStyle" {
                    check_style(e, &mut is_heading, &mut heading_level, &mut list_level);
                }
                if name_bytes == b"w:ilvl" {
                    check_ilvl(e, &mut list_level);
                }
                if name_bytes == b"w:r" {
                    _in_run = true;
                    is_bold = false;
                    is_italic = false;
                }
                if name_bytes == b"w:rPr" {
                    // run properties 시작
                }
                if name_bytes == b"w:b" { is_bold = true; }
                if name_bytes == b"w:i" { is_italic = true; }
                if name_bytes == b"w:t" {
                    in_text = true;
                    // xml:space="preserve" 확인
                    preserve_space = false;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"xml:space" {
                            let val = String::from_utf8_lossy(&attr.value);
                            if val == "preserve" {
                                preserve_space = true;
                            }
                        }
                    }
                }
                if name_bytes == b"w:tbl" {
                    _in_table = true;
                    table_rows.clear();
                }
                if name_bytes == b"w:tr" {
                    table_row.clear();
                }
                if name_bytes == b"w:tc" {
                    in_cell = true;
                    current_cell.clear();
                }
                if name_bytes == b"w:tab" || name_bytes == b"w:tab/" {
                    if in_cell {
                        current_cell.push('\t');
                    } else {
                        current_line.push('\t');
                    }
                }
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
                if name_bytes == b"w:tab" {
                    if in_cell { current_cell.push('\t'); }
                    else { current_line.push('\t'); }
                }
                if name_bytes == b"w:pStyle" {
                    check_style(e, &mut is_heading, &mut heading_level, &mut list_level);
                }
                if name_bytes == b"w:ilvl" {
                    check_ilvl(e, &mut list_level);
                }
                // 빈 <w:t/> 처리
                if name_bytes == b"w:t" {
                    // 빈 텍스트 노드 — 공백 추가
                }
            }
            Ok(Event::Text(ref e)) => {
                if !in_text {
                    buf.clear();
                    continue;
                }
                let raw_text = e.unescape().unwrap_or_default().to_string();
                if raw_text.is_empty() { buf.clear(); continue; }

                let text = if preserve_space {
                    raw_text
                } else {
                    raw_text.trim().to_string()
                };

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
                    // 셀 내에서 단어 사이 공백 보장
                    if !current_cell.is_empty() && !current_cell.ends_with(' ') && !current_cell.ends_with('\t') {
                        // 이전 텍스트와 현재 텍스트 사이에 공백이 필요한지 확인
                        let last_char = current_cell.chars().last().unwrap_or(' ');
                        let first_char = formatted.chars().next().unwrap_or(' ');
                        if last_char.is_alphanumeric() && first_char.is_alphanumeric() {
                            current_cell.push(' ');
                        }
                    }
                    current_cell.push_str(&formatted);
                } else {
                    // 단어 사이 공백 보장
                    if !current_line.is_empty() && !current_line.ends_with(' ') && !current_line.ends_with('\n') && !current_line.ends_with('\t') {
                        let last_char = current_line.chars().last().unwrap_or(' ');
                        let first_char = formatted.chars().next().unwrap_or(' ');
                        if last_char.is_alphanumeric() && first_char.is_alphanumeric() {
                            current_line.push(' ');
                        }
                    }
                    current_line.push_str(&formatted);
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                let name_bytes = name.as_ref();

                if name_bytes == b"w:t" {
                    in_text = false;
                    preserve_space = false;
                }
                if name_bytes == b"w:r" {
                    _in_run = false;
                }
                if name_bytes == b"w:p" {
                    if in_cell {
                        // 셀 내 여러 단락이면 공백으로 연결
                        if !current_cell.is_empty() {
                            current_cell.push(' ');
                        }
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
                    _in_table = false;
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

    let mut lines: Vec<String> = Vec::new();

    for (row_idx, row) in rows.iter().enumerate() {
        let mut cells: Vec<String> = Vec::new();
        for i in 0..max_cols {
            let cell = row.get(i).cloned().unwrap_or_default();
            cells.push(format!(" {} ", cell));
        }
        lines.push(format!("|{}|", cells.join("|")));

        if row_idx == 0 {
            let sep: Vec<String> = (0..max_cols)
                .map(|_| " --- ".to_string())
                .collect();
            lines.push(format!("|{}|", sep.join("|")));
        }
    }

    lines.join("\n")
}
