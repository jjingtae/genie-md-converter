# 지니 MD 변환기

문서(PDF, DOCX, HWP)를 Markdown으로 변환하는 데스크톱 도구입니다.

## 지원 형식

| 입력 | 설명 |
|------|------|
| PDF  | 텍스트 기반 PDF |
| DOCX | Word 문서 (표, 제목, 리스트 포함) |
| HWP  | 한글 문서 (텍스트 추출) |
| TXT  | 텍스트 파일 (그대로 반환) |

## 다운로드

[Releases](../../releases) 또는 Actions 탭에서 최신 빌드를 받을 수 있습니다.

## 기술 스택

- **Rust** — 변환 엔진 (pdf-extract, quick-xml, cfb)
- **Tauri** — 데스크톱 앱 프레임워크 (~5MB)
- **HTML/JS** — UI
