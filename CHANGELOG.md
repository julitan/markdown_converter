# PDF to Markdown Converter - 변경 이력

## 2026-02-09

### 1. 변환 속도 최적화
- **PdfConverter 인스턴스 캐싱**: 매 변환마다 새로 생성하던 PdfConverter를 전역으로 캐싱하여 재사용 (`converter/marker_converter.py`)
  - 특히 배치 변환 시 2번째 파일부터 모델 초기화 시간 절약
- **폴링 간격 조정**: 프론트엔드 상태 폴링 500ms → 2000ms로 변경 (`main.py`, `templates/index.html`)
  - 서버 부하 75% 감소, 변환 스레드에 더 많은 리소스 할당

### 2. 이미지 관리 방식 변경
- **이미지 별도 폴더 저장**: 변환 시 이미지를 `{PDF파일명}_images/` 폴더에 저장
- **마크다운 내 상대경로**: 이미지 경로를 `{PDF파일명}_images/이미지파일.jpeg` 형태로 수정
- **특수문자 경로 처리**: 파일명에 공백, 중괄호`{}`, 괄호`()` 등이 포함된 경우 CommonMark 표준 꺾쇠 괄호`<>`로 감싸서 처리
  - 예: `![](<IEC 60601-1-2{ed4.0}_(english only)_images/_page_0_Picture_1.jpeg>)`
  - VSCode 마크다운 미리보기에서 정상 표시
- **(취소) Base64 임베딩**: 이미지를 data URI로 마크다운에 직접 삽입하는 방식 시도 후, 파일 관리 편의성을 위해 별도 폴더 방식으로 최종 결정

### 3. 출력 경로 설정
- **`--output-dir` 인자 추가** (`main.py`): 명령줄에서 출력 폴더 지정 가능
  - 미지정 시 기본값: `pdf_to_markdown/output/`
- **`pdf2md_conv.vbs` 수정**: VBS 파일이 위치한 폴더를 `--output-dir`로 자동 전달
  - VBS를 아무 폴더에 복사 후 실행하면 해당 폴더에 결과 저장
- **PDF 파일 제외**: 업로드된 PDF는 임시 폴더에 저장 후 변환 완료 시 자동 삭제
  - 출력 폴더에는 `.md` 파일과 `_images/` 폴더만 남음

### 4. GPU 확인 로그
- 서버 시작 시 CUDA 사용 가능 여부 및 GPU 장치명 출력 (`converter/marker_converter.py`)

### 5. Git 초기화
- `.gitignore` 추가: `__pycache__/`, `*.pyc`, `output/`, `nul` 제외
- 초기 커밋 생성 (`8a580f0`)

---

## 파일 구조

```
pdf_to_markdown/
  converter/
    __init__.py              # convert_pdf 함수 export
    marker_converter.py      # Marker 기반 변환 엔진
  templates/
    index.html               # 웹 UI (별도 파일 버전)
  main.py                    # Flask 서버 + 인라인 HTML
  pdf2md_conv.bat            # 실행 스크립트 (BAT)
  pdf2md_conv.vbs            # 실행 스크립트 (VBS, 출력 경로 전달)
  requirements.txt           # 의존성 목록
  .gitignore
  CHANGELOG.md               # 이 파일
```

## 출력 결과 구조

```
{출력폴더}/
  {PDF파일명}/
    {PDF파일명}.md            # 변환된 마크다운
    {PDF파일명}_images/       # 추출된 이미지
      _page_0_Picture_1.jpeg
      _page_0_Picture_5.jpeg
      ...
```
