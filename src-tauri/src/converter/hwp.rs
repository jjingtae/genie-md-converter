use cfb::CompoundFile;
use flate2::read::DeflateDecoder;
use std::io::Read;
use std::path::Path;

const HWPTAG_BEGIN: u16 = 0x010; // 16
const HWPTAG_PARA_TEXT: u16 = HWPTAG_BEGIN + 51; // 67

/// HWP 파일을 Markdown으로 변환
pub fn convert(path: &Path) -> Result<String, String> {
    let file = std::fs::File::open(path)
        .map_err(|e| format!("HWP 파일 열기 실패: {}", e))?;
    let mut comp = CompoundFile::open(file)
        .map_err(|e| format!("HWP CFB 열기 실패: {}", e))?;

    // FileHeader에서 압축 여부 확인
    let compressed = is_compressed(&mut comp)?;

    // BodyText 섹션들 읽기
    let mut all_text = Vec::new();
    let mut section_idx = 0;

    loop {
        let stream_name = format!("BodyText/Section{}", section_idx);
        let section_data = match read_stream(&mut comp, &stream_name) {
            Ok(data) => data,
            Err(_) => break, // 더 이상 섹션 없음
        };

        let decompressed = if compressed {
            decompress(&section_data)?
        } else {
            section_data
        };

        let paragraphs = parse_section(&decompressed)?;
        all_text.extend(paragraphs);
        section_idx += 1;
    }

    if section_idx == 0 {
        return Err("BodyText 섹션을 찾을 수 없습니다".to_string());
    }

    let md = post_process(&all_text);
    Ok(md)
}

/// FileHeader에서 압축 플래그 확인
fn is_compressed(comp: &mut CompoundFile<std::fs::File>) -> Result<bool, String> {
    let data = read_stream(comp, "FileHeader")?;
    if data.len() < 40 {
        return Err("FileHeader가 너무 짧습니다".to_string());
    }
    // offset 36: 속성 플래그, bit 0 = 압축
    let flags = u32::from_le_bytes([data[36], data[37], data[38], data[39]]);
    Ok(flags & 0x01 != 0)
}

/// OLE 스트림 읽기
fn read_stream(comp: &mut CompoundFile<std::fs::File>, name: &str) -> Result<Vec<u8>, String> {
    let mut stream = comp
        .open_stream(name)
        .map_err(|e| format!("스트림 '{}' 열기 실패: {}", name, e))?;
    let mut buf = Vec::new();
    stream
        .read_to_end(&mut buf)
        .map_err(|e| format!("스트림 읽기 실패: {}", e))?;
    Ok(buf)
}

/// zlib(deflate) 압축 해제
fn decompress(data: &[u8]) -> Result<Vec<u8>, String> {
    let mut decoder = DeflateDecoder::new(data);
    let mut decompressed = Vec::new();
    decoder
        .read_to_end(&mut decompressed)
        .map_err(|e| format!("압축 해제 실패: {}", e))?;
    Ok(decompressed)
}

/// 섹션 바이너리 데이터에서 레코드를 파싱하여 텍스트 추출
fn parse_section(data: &[u8]) -> Result<Vec<String>, String> {
    let mut paragraphs = Vec::new();
    let mut pos = 0;

    while pos + 4 <= data.len() {
        // 레코드 헤더: 4바이트
        let header = u32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]);
        pos += 4;

        let tag_id = (header & 0x3FF) as u16;        // bits 0-9
        let _level = ((header >> 10) & 0x3FF) as u16; // bits 10-19
        let mut size = ((header >> 20) & 0xFFF) as usize; // bits 20-31

        // 크기가 0xFFF이면 다음 4바이트가 실제 크기
        if size == 0xFFF {
            if pos + 4 > data.len() {
                break;
            }
            size = u32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]])
                as usize;
            pos += 4;
        }

        if pos + size > data.len() {
            break;
        }

        let record_data = &data[pos..pos + size];
        pos += size;

        // PARA_TEXT 레코드에서 텍스트 추출
        if tag_id == HWPTAG_PARA_TEXT {
            let text = extract_para_text(record_data);
            if !text.is_empty() {
                paragraphs.push(text);
            }
        }
    }

    Ok(paragraphs)
}

/// PARA_TEXT 레코드의 UTF-16LE 데이터에서 텍스트 추출
fn extract_para_text(data: &[u8]) -> String {
    if data.len() < 2 {
        return String::new();
    }

    let mut text = String::new();
    let mut i = 0;

    while i + 1 < data.len() {
        let ch = u16::from_le_bytes([data[i], data[i + 1]]);
        i += 2;

        match ch {
            // 확장 컨트롤 문자 (8 wchar = 16바이트, 첫 2바이트는 이미 읽음 → 14바이트 건너뛰기)
            1 | 2 | 3 | 11 | 12 => {
                i += 14; // 나머지 7 wchar 건너뛰기
            }
            // 인라인 확장 컨트롤 (8 wchar)
            4 | 5 | 6 | 7 | 8 => {
                i += 14;
            }
            // 탭
            9 => {
                text.push('\t');
            }
            // 줄바꿈
            10 => {
                text.push('\n');
            }
            // 단락 끝
            13 => {
                // 단락 구분 (무시 — 단락 단위로 이미 분리됨)
            }
            // 기타 컨트롤 문자 무시
            0 | 14..=23 | 25..=29 => {}
            // 하이픈
            24 => {
                text.push('-');
            }
            // 비파괴 공백
            30 => {
                text.push(' ');
            }
            // 고정폭 공백
            31 => {
                text.push(' ');
            }
            // 일반 문자
            _ => {
                if let Some(c) = char::from_u32(ch as u32) {
                    text.push(c);
                }
            }
        }
    }

    text.trim().to_string()
}

/// 추출된 텍스트를 정리하여 Markdown으로 후처리
fn post_process(paragraphs: &[String]) -> String {
    let mut result = Vec::new();
    let mut prev_blank = false;

    for para in paragraphs {
        let trimmed = para.trim();
        if trimmed.is_empty() {
            if !prev_blank && !result.is_empty() {
                result.push(String::new());
                prev_blank = true;
            }
            continue;
        }
        prev_blank = false;
        result.push(trimmed.to_string());
    }

    // 끝 빈줄 제거
    while result.last().map_or(false, |l| l.is_empty()) {
        result.pop();
    }

    result.join("\n")
}
