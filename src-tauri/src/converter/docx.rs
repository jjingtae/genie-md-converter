use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::Read;
use std::path::Path;

pub fn convert(path: &Path) -> Result<String, String> {
    let file = std::fs::File::open(path).map_err(|e| format!("파일 열기 실패: {}", e))?;
    let mut archive = zip::ZipArchive::new(file).map_err(|e| format!("DOCX ZIP 읽기 실패: {}", e))?;

    let xml_content = {
        let mut entry = archive.by_name("word/document.xml")
            .map_err(|_| "word/document.xml을 찾을 수 없습니다".to_string())?;
        let mut buf = String::new();
        entry.read_to_string(&mut buf).map_err(|e| format!("XML 읽기 실패: {}", e))?;
        buf
    };

    parse_document_xml(&xml_content)
}

fn parse_document_xml(xml: &str) -> Result<String, String> {
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
    let mut table_row: Vec<CellInfo> = Vec::new();
    let mut table_rows: Vec<Vec<CellInfo>> = Vec::new();
    let mut current_cell = String::new();
    let mut in_cell = false;
    let mut list_level: i32 = -1;
    let mut in_text = false;
    let mut preserve_space = false;
    // 셀 병합 정보
    let mut current_grid_span: u16 = 1;
    let mut current_vmerge: VMerge = VMerge::None;
    let mut buf: Vec<u8> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let nb = name.as_ref();

                if nb == b"w:p" {
                    in_paragraph = true;
                    current_line.clear();
                    is_heading = false;
                    heading_level = 0;
                    list_level = -1;
                }
                if nb == b"w:pStyle" { check_style(e, &mut is_heading, &mut heading_level, &mut list_level); }
                if nb == b"w:ilvl" { check_ilvl(e, &mut list_level); }
                if nb == b"w:r" { is_bold = false; is_italic = false; }
                if nb == b"w:b" { is_bold = true; }
                if nb == b"w:i" { is_italic = true; }
                if nb == b"w:t" {
                    in_text = true;
                    preserve_space = false;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"xml:space" {
                            if String::from_utf8_lossy(&attr.value) == "preserve" { preserve_space = true; }
                        }
                    }
                }
                if nb == b"w:tbl" {
                    _in_table = true;
                    table_rows.clear();
                }
                if nb == b"w:tr" {
                    table_row.clear();
                }
                if nb == b"w:tc" {
                    in_cell = true;
                    current_cell.clear();
                    current_grid_span = 1;
                    current_vmerge = VMerge::None;
                }
                if nb == b"w:tcPr" {
                    // 셀 속성 시작 — gridSpan, vMerge는 하위에서 처리
                }
                if nb == b"w:gridSpan" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"w:val" {
                            let val = String::from_utf8_lossy(&attr.value);
                            current_grid_span = val.parse::<u16>().unwrap_or(1);
                        }
                    }
                }
                if nb == b"w:vMerge" {
                    let mut is_restart = false;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"w:val" {
                            let val = String::from_utf8_lossy(&attr.value);
                            if val == "restart" { is_restart = true; }
                        }
                    }
                    current_vmerge = if is_restart { VMerge::Restart } else { VMerge::Continue };
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                let nb = name.as_ref();
                if nb == b"w:b" { is_bold = true; }
                if nb == b"w:i" { is_italic = true; }
                if nb == b"w:br" {
                    if in_cell { current_cell.push(' '); }
                    else { current_line.push('\n'); }
                }
                if nb == b"w:tab" {
                    if in_cell { current_cell.push('\t'); }
                    else { current_line.push('\t'); }
                }
                if nb == b"w:pStyle" { check_style(e, &mut is_heading, &mut heading_level, &mut list_level); }
                if nb == b"w:ilvl" { check_ilvl(e, &mut list_level); }
                if nb == b"w:gridSpan" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"w:val" {
                            let val = String::from_utf8_lossy(&attr.value);
                            current_grid_span = val.parse::<u16>().unwrap_or(1);
                        }
                    }
                }
                if nb == b"w:vMerge" {
                    // Empty <w:vMerge/> = continue
                    let mut is_restart = false;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"w:val" {
                            if String::from_utf8_lossy(&attr.value) == "restart" { is_restart = true; }
                        }
                    }
                    current_vmerge = if is_restart { VMerge::Restart } else { VMerge::Continue };
                }
            }
            Ok(Event::Text(ref e)) => {
                if !in_text { buf.clear(); continue; }
                let raw = e.unescape().unwrap_or_default().to_string();
                if raw.is_empty() { buf.clear(); continue; }
                let text = if preserve_space { raw } else { raw.trim().to_string() };
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
                let name = e.name();
                let nb = name.as_ref();

                if nb == b"w:t" { in_text = false; preserve_space = false; }
                if nb == b"w:r" { /* bold/italic 이미 리셋 불필요 — 다음 run에서 다시 설정 */ }
                if nb == b"w:p" {
                    if in_cell {
                        if !current_cell.is_empty() { current_cell.push(' '); }
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
                if nb == b"w:tc" {
                    in_cell = false;
                    table_row.push(CellInfo {
                        text: current_cell.trim().to_string(),
                        grid_span: current_grid_span,
                        vmerge: current_vmerge,
                    });
                    current_cell.clear();
                    current_grid_span = 1;
                    current_vmerge = VMerge::None;
                }
                if nb == b"w:tr" {
                    table_rows.push(table_row.clone());
                    table_row.clear();
                }
                if nb == b"w:tbl" {
                    _in_table = false;
                    if !table_rows.is_empty() {
                        let md = render_table_with_merge(&table_rows);
                        output.push(md);
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
            if !prev_blank { result.push(String::new()); prev_blank = true; }
        } else { prev_blank = false; result.push(line.clone()); }
    }
    Ok(result.join("\n"))
}

/// 단어 사이 공백 자동 삽입
fn add_with_space(target: &mut String, text: &str) {
    if target.is_empty() || text.is_empty() { target.push_str(text); return; }
    let last = target.chars().last().unwrap_or(' ');
    let first = text.chars().next().unwrap_or(' ');
    // 알파벳/숫자 연결 시 공백 삽입
    if (last.is_alphanumeric() || last == '*' || last == '"' || last == '\u{201D}')
        && (first.is_alphanumeric() || first == '*' || first == '"' || first == '\u{201C}') {
        target.push(' ');
    }
    target.push_str(text);
}

#[derive(Clone, Copy, PartialEq)]
enum VMerge { None, Restart, Continue }

#[derive(Clone)]
struct CellInfo {
    text: String,
    grid_span: u16,
    vmerge: VMerge,
}

/// 셀 병합을 고려한 표 렌더링
fn render_table_with_merge(rows: &[Vec<CellInfo>]) -> String {
    if rows.is_empty() { return String::new(); }

    // 실제 열 수 계산 (gridSpan 고려)
    let max_cols = rows.iter()
        .map(|r| r.iter().map(|c| c.grid_span as usize).sum::<usize>())
        .max()
        .unwrap_or(0);
    if max_cols == 0 { return String::new(); }

    // 2D 그리드 구성
    let mut grid: Vec<Vec<String>> = Vec::new();

    for (row_idx, row) in rows.iter().enumerate() {
        let mut grid_row = vec![String::new(); max_cols];
        let mut col_pos = 0;

        for cell in row {
            if col_pos >= max_cols { break; }

            match cell.vmerge {
                VMerge::Continue => {
                    // 위 셀의 내용 연속 — 빈칸으로
                    let span = cell.grid_span as usize;
                    col_pos += span;
                }
                _ => {
                    // None 또는 Restart — 내용 배치
                    grid_row[col_pos] = cell.text.clone();
                    let span = cell.grid_span as usize;
                    // gridSpan > 1이면 나머지 열은 빈칸
                    col_pos += span;
                }
            }
        }
        grid.push(grid_row);
    }

    // Markdown 표 생성
    let mut lines: Vec<String> = Vec::new();
    for (row_idx, row) in grid.iter().enumerate() {
        let cells: Vec<String> = row.iter()
            .map(|c| {
                let cleaned = c.replace('|', "\\|");
                format!(" {} ", cleaned)
            })
            .collect();
        lines.push(format!("|{}|", cells.join("|")));

        if row_idx == 0 {
            let sep: Vec<String> = (0..max_cols).map(|_| " --- ".to_string()).collect();
            lines.push(format!("|{}|", sep.join("|")));
        }
    }

    lines.join("\n")
}

fn check_style(e: &quick_xml::events::BytesStart, is_heading: &mut bool, heading_level: &mut u8, list_level: &mut i32) {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == b"w:val" {
            let val = String::from_utf8_lossy(&attr.value).to_string();
            if val.starts_with("Heading") || val.starts_with("heading") {
                *is_heading = true;
                *heading_level = val.chars().last().and_then(|c| c.to_digit(10)).unwrap_or(1) as u8;
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
            *list_level = String::from_utf8_lossy(&attr.value).parse::<i32>().unwrap_or(0);
        }
    }
}
