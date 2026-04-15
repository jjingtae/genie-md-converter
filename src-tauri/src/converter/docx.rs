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

    // word/document.xml 읽기
    let xml_content = {
        let mut entry = archive
            .by_name("word/document.xml")
            .map_err(|_| "word/document.xml을 찾을 수 없습니다".to_string())?;
        let mut buf = String::new();
        entry.read_to_string(&mut buf)
            .map_err(|e| format!("XML 읽기 실패: {}", e))?;
        buf
    };

    // 넘버링 정보 읽기 (있으면)
    let numbering_xml = {
        match archive.by_name("word/numbering.xml") {
            Ok(mut entry) => {
                let mut buf = String::new();
                let _ = entry.read_to_string(&mut buf);
                Some(buf)
            }
            Err(_) => None,
        }
    };

    let md = parse_document_xml(&xml_content, &numbering_xml)?;
    Ok(md)
}

/// document.xml을 파싱하여 Markdown 생성
fn parse_document_xml(xml: &str, _numbering: &Option<String>) -> Result<String, String> {
    let mut reader = Reader::from_str(xml);
    reader.trim_text(true);

    let mut output = Vec::new();
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
    let mut in_hyperlink = false;
    let mut list_level: i32 = -1;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                match name.as_str() {
                    "w:p" | "w:p/" => {
                        in_paragraph = true;
                        current_line.clear();
                        is_heading = false;
                        heading_level = 0;
                        list_level = -1;
                    }
                    "w:pStyle" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w:val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val.starts_with("Heading") || val.starts_with("heading") {
                                    is_heading = true;
                                    heading_level = val
                                        .chars()
                                        .last()
                                        .and_then(|c| c.to_digit(10))
                                        .unwrap_or(1) as u8;
                                    if heading_level > 6 { heading_level = 6; }
                                } else if val.contains("List") || val.contains("list") {
                                    list_level = 0;
                                }
                            }
                        }
                    }
                    "w:ilvl" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w:val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                list_level = val.parse::<i32>().unwrap_or(0);
                            }
                        }
                    }
                    "w:b" | "w:b/" => is_bold = true,
                    "w:i" | "w:i/" => is_italic = true,
                    "w:tbl" => {
                        in_table = true;
                        table_rows.clear();
                    }
                    "w:tr" => {
                        table_row.clear();
                    }
                    "w:tc" => {
                        in_cell = true;
                        current_cell.clear();
                    }
                    "w:hyperlink" => {
                        in_hyperlink = true;
                    }
                    "w:br" | "w:br/" => {
                        if in_cell {
                            current_cell.push(' ');
                        } else {
                            current_line.push('\n');
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                let text = e.unescape().unwrap_or_default().to_string();
                if text.is_empty() {
                    buf.clear();
                    continue;
                }

                let mut formatted = text.clone();
                if is_bold && is_italic {
                    formatted = format!("***{}***", text);
                } else if is_bold {
                    formatted = format!("**{}**", text);
                } else if is_italic {
                    formatted = format!("*{}*", text);
                }

                if in_cell {
                    current_cell.push_str(&formatted);
                } else {
                    current_line.push_str(&formatted);
                }
            }
            Ok(Event::End(ref e)) => {
                let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                match name.as_str() {
                    "w:r" => {
                        is_bold = false;
                        is_italic = false;
                    }
                    "w:p" => {
                        if in_cell {
                            if !current_cell.is_empty() {
                                // 셀 내 단락 구분
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
                    "w:tc" => {
                        in_cell = false;
                        table_row.push(current_cell.trim().to_string());
                        current_cell.clear();
                    }
                    "w:tr" => {
                        table_rows.push(table_row.clone());
                        table_row.clear();
                    }
                    "w:tbl" => {
                        in_table = false;
                        // 테이블을 Markdown으로 변환
                        if !table_rows.is_empty() {
                            let md_table = render_table(&table_rows);
                            output.push(md_table);
                        }
                        table_rows.clear();
                    }
                    "w:hyperlink" => {
                        in_hyperlink = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML 파싱 오류: {}", e)),
            _ => {}
        }
        buf.clear();
    }

    // 빈 줄 정리
    let mut result = Vec::new();
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

/// 테이블 행 데이터를 Markdown 테이블로 렌더링
fn render_table(rows: &[Vec<String>]) -> String {
    if rows.is_empty() {
        return String::new();
    }

    // 최대 열 수 계산
    let max_cols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    if max_cols == 0 {
        return String::new();
    }

    // 각 열의 최대 너비 계산
    let mut col_widths = vec![3usize; max_cols];
    for row in rows {
        for (i, cell) in row.iter().enumerate() {
            if i < max_cols {
                col_widths[i] = col_widths[i].max(cell.len()).max(3);
            }
        }
    }

    let mut lines = Vec::new();

    for (row_idx, row) in rows.iter().enumerate() {
        let mut cells: Vec<String> = Vec::new();
        for i in 0..max_cols {
            let cell = row.get(i).cloned().unwrap_or_default();
            cells.push(format!(" {} ", cell));
        }
        lines.push(format!("|{}|", cells.join("|")));

        // 첫 행 다음에 구분선
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
