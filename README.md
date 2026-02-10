# Markdown Converter

A web-based document converter that transforms PDF, DOC, and DOCX files into Markdown format with image extraction support.

> PDF, DOC, DOCX 문서를 마크다운으로 변환하는 웹 기반 변환기입니다. 이미지 추출을 지원합니다.

## Features / 주요 기능

- **PDF to Markdown** - Powered by [Marker](https://github.com/VikParuchuri/marker) with GPU acceleration (CUDA)
- **DOC/DOCX to Markdown** - Powered by [Mammoth](https://github.com/mwilliamson/python-mammoth) + [markdownify](https://github.com/matthewwithanm/python-markdownify)
- **Web UI** - Drag & drop interface for single file or batch conversion (Flask)
- **Image extraction** - Saves images to separate folders with relative path references in Markdown
- **Batch conversion** - Convert entire folders with optional recursive subdirectory support

## Screenshots / 스크린샷

The web interface provides:
- Single file conversion with drag & drop / 드래그 앤 드롭으로 단일 파일 변환
- Batch folder conversion / 폴더 일괄 변환
- Real-time progress & logs / 실시간 진행률 및 로그

## Requirements / 요구사항

- Python 3.10+
- CUDA-compatible GPU (optional, for faster PDF conversion / PDF 변환 속도 향상을 위해 권장)
- Microsoft Word (required only for `.doc` format conversion / `.doc` 형식 변환 시에만 필요)

## Installation / 설치

```bash
pip install -r requirements.txt
```

### Dependencies / 의존성

| Package | Purpose |
|---------|---------|
| `marker-pdf` | PDF to Markdown conversion engine |
| `flask` | Web server & UI |
| `mammoth` | DOCX to HTML conversion |
| `markdownify` | HTML to Markdown conversion |
| `doc2docx` | Legacy DOC to DOCX conversion (requires Word) |
| `tqdm` | Progress bars |

## Usage / 사용법

### Web UI

```bash
python main.py
```

Opens the web interface at `http://127.0.0.1:5000`.

브라우저에서 `http://127.0.0.1:5000` 으로 접속합니다.

### Custom output directory / 출력 경로 지정

```bash
python main.py --output-dir C:\Output
```

### Windows shortcut scripts / Windows 실행 스크립트

- `pdf2md_conv.bat` - Batch file launcher
- `pdf2md_conv.vbs` - VBScript launcher (uses script location as output directory)

## Output Structure / 출력 구조

```
{output_dir}/
  {filename}/
    {filename}.md              # Converted Markdown / 변환된 마크다운
    {filename}_images/         # Extracted images / 추출된 이미지
      _page_0_Picture_1.jpeg
      ...
```

## Project Structure / 프로젝트 구조

```
markdown_converter/
  converter/
    __init__.py              # Module exports
    marker_converter.py      # PDF conversion engine (Marker)
    docx_converter.py        # DOC/DOCX conversion engine (Mammoth)
  templates/
    index.html               # Web UI template (standalone version)
  main.py                    # Flask server with inline HTML
  pdf2md_conv.bat            # Windows batch launcher
  pdf2md_conv.vbs            # Windows VBScript launcher
  requirements.txt           # Python dependencies
```

## License

MIT
