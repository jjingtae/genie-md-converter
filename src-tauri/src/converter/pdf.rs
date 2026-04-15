use std::path::Path;

/// PDF 파일에서 텍스트를 추출하여 Markdown으로 변환
pub fn convert(path: &Path) -> Result<String, String> {
    let bytes = std::fs::read(path)
        .map_err(|e| format!("파일 읽기 실패: {}", e))?;

    let text = pdf_extract::extract_text_from_mem(&bytes)
        .map_err(|e| format!("PDF 텍스트 추출 실패: {}", e))?;

    let md = post_process(&text);
    Ok(md)
}

/// 추출된 텍스트를 정리하여 Markdown으로 변환
fn post_process(raw: &str) -> String {
    let mut lines: Vec<String> = Vec::new();
    let mut prev_blank = false;

    for line in raw.lines() {
        let trimmed = line.trim();

        // 빈 줄 중복 제거
        if trimmed.is_empty() {
            if !prev_blank && !lines.is_empty() {
                lines.push(String::new());
                prev_blank = true;
            }
            continue;
        }
        prev_blank = false;

        // 페이지 번호 패턴 제거 (숫자만 있는 줄)
        if trimmed.chars().all(|c| c.is_ascii_digit() || c == '-' || c == ' ') 
           && trimmed.len() <= 10 
        {
            continue;
        }

        // 제목 추정: 짧고 마침표 없으면 ## 처리
        if trimmed.len() <= 60 
           && !trimmed.ends_with('.') 
           && !trimmed.ends_with(',')
           && !trimmed.contains('.')
           && trimmed.chars().filter(|c| c.is_whitespace()).count() <= 8
        {
            lines.push(format!("## {}", trimmed));
        } else {
            lines.push(trimmed.to_string());
        }
    }

    // 끝 빈줄 제거
    while lines.last().map_or(false, |l| l.is_empty()) {
        lines.pop();
    }

    lines.join("\n")
}
