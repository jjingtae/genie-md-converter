#![cfg_attr(
    all(not(debug_assertions), target_os = "windows"),
    windows_subsystem = "windows"
)]

mod converter;

use std::path::Path;

/// 단일 파일을 Markdown으로 변환
#[tauri::command]
fn convert_file(path: String) -> Result<String, String> {
    converter::convert_to_markdown(&path)
}

/// 텍스트를 파일로 저장 (이미 변환된 MD 텍스트 저장용)
#[tauri::command]
fn save_text(text: String, path: String) -> Result<String, String> {
    std::fs::write(&path, &text)
        .map_err(|e| format!("저장 실패: {}", e))?;
    Ok("저장 완료".to_string())
}

/// 파일을 변환하고 .md 파일로 저장
#[tauri::command]
fn convert_and_save(input_path: String, output_path: String) -> Result<String, String> {
    let md = converter::convert_to_markdown(&input_path)?;
    std::fs::write(&output_path, &md)
        .map_err(|e| format!("저장 실패: {}", e))?;
    Ok(format!("저장 완료: {}", output_path))
}

/// 여러 파일을 일괄 변환
#[tauri::command]
fn convert_batch(
    paths: Vec<String>,
    output_dir: String,
) -> Result<Vec<BatchResult>, String> {
    let out_dir = Path::new(&output_dir);
    if !out_dir.exists() {
        std::fs::create_dir_all(out_dir)
            .map_err(|e| format!("출력 폴더 생성 실패: {}", e))?;
    }

    let mut results = Vec::new();
    for input_path in &paths {
        let p = Path::new(input_path);
        let stem = p.file_stem().and_then(|s| s.to_str()).unwrap_or("output");
        let out_path = out_dir.join(format!("{}.md", stem));

        match converter::convert_to_markdown(input_path) {
            Ok(md) => match std::fs::write(&out_path, &md) {
                Ok(_) => results.push(BatchResult {
                    file: input_path.clone(),
                    success: true,
                    message: out_path.to_string_lossy().to_string(),
                }),
                Err(e) => results.push(BatchResult {
                    file: input_path.clone(),
                    success: false,
                    message: format!("저장 실패: {}", e),
                }),
            },
            Err(e) => results.push(BatchResult {
                file: input_path.clone(),
                success: false,
                message: e,
            }),
        }
    }

    Ok(results)
}

#[derive(serde::Serialize)]
struct BatchResult {
    file: String,
    success: bool,
    message: String,
}

fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![
            convert_file,
            save_text,
            convert_and_save,
            convert_batch
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
