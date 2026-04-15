pub mod pdf;
pub mod docx;
pub mod hwp;

use std::path::Path;

/// 파일 확장자에 따라 적절한 변환기를 호출
pub fn convert_to_markdown(path: &str) -> Result<String, String> {
    let path = Path::new(path);

    if !path.exists() {
        return Err("파일이 존재하지 않습니다".to_string());
    }

    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_default();

    match ext.as_str() {
        "pdf" => pdf::convert(path),
        "docx" => docx::convert(path),
        "hwp" => hwp::convert(path),
        "hwpx" => {
            // HWPX는 ZIP 기반 — 내부 content.hpf를 파싱
            Err("HWPX 지원은 추후 추가 예정입니다".to_string())
        }
        "txt" | "md" => {
            // 텍스트 파일은 그대로 반환
            std::fs::read_to_string(path)
                .map_err(|e| format!("파일 읽기 실패: {}", e))
        }
        _ => Err(format!("지원하지 않는 형식입니다: .{}", ext)),
    }
}
