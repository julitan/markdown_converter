"""
Mammoth + markdownify 기반 DOC/DOCX → Markdown 변환기
"""
import os
import base64
import hashlib
from pathlib import Path
from typing import Optional

import mammoth
from markdownify import markdownify as md

from .marker_converter import ConversionResult


def _convert_doc_to_docx(doc_path: Path) -> Path:
    """
    DOC(구형 Word) 파일을 DOCX로 변환 (doc2docx 사용, Word 필요)

    Returns:
        변환된 DOCX 파일 경로
    """
    from doc2docx import convert

    docx_path = doc_path.with_suffix(".docx")
    convert(str(doc_path), str(docx_path))
    return docx_path


def convert_docx(
    docx_path: str | Path,
    output_dir: Optional[str | Path] = None,
    save_images: bool = True,
) -> ConversionResult:
    """
    단일 DOC/DOCX 파일을 Markdown으로 변환

    Args:
        docx_path: DOC 또는 DOCX 파일 경로
        output_dir: 출력 디렉토리 (None이면 원본과 같은 위치)
        save_images: 이미지 저장 여부

    Returns:
        ConversionResult: 변환 결과
    """
    docx_path = Path(docx_path)

    if not docx_path.exists():
        return ConversionResult(
            success=False,
            markdown="",
            images={},
            metadata={},
            error=f"파일을 찾을 수 없습니다: {docx_path}",
        )

    if output_dir is None:
        output_dir = docx_path.parent
    else:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

    # DOC → DOCX 전처리
    temp_docx = None
    original_stem = docx_path.stem
    if docx_path.suffix.lower() == ".doc":
        try:
            temp_docx = _convert_doc_to_docx(docx_path)
            docx_path = temp_docx
        except Exception as e:
            return ConversionResult(
                success=False,
                markdown="",
                images={},
                metadata={},
                error=f"DOC → DOCX 변환 실패 (Microsoft Word 필요): {e}",
            )

    try:
        images_saved = {}

        # 이미지 저장 디렉토리 준비
        images_dir_name = f"{original_stem}_images"
        images_dir = output_dir / images_dir_name
        if save_images:
            images_dir.mkdir(exist_ok=True)

        # mammoth 이미지 핸들러: 이미지를 파일로 저장하고 경로 반환
        def handle_image(image):
            with image.open() as img_stream:
                img_data = img_stream.read()

            # 파일명 생성: content_type + hash 기반
            content_type = image.content_type or "image/png"
            ext = content_type.split("/")[-1]
            if ext == "jpeg":
                ext = "jpg"
            img_hash = hashlib.md5(img_data).hexdigest()[:8]
            img_name = f"image_{len(images_saved) + 1}_{img_hash}.{ext}"

            if save_images:
                img_path = images_dir / img_name
                img_path.write_bytes(img_data)
                images_saved[img_name] = img_data
                rel_path = f"{images_dir_name}/{img_name}"
                return {"src": rel_path}
            else:
                # 이미지 저장하지 않을 경우 base64 인라인
                b64 = base64.b64encode(img_data).decode("ascii")
                return {"src": f"data:{content_type};base64,{b64}"}

        # mammoth으로 DOCX → HTML 변환
        with open(str(docx_path), "rb") as f:
            result = mammoth.convert_to_html(
                f,
                convert_image=mammoth.images.img_element(handle_image),
            )

        html_content = result.value

        # markdownify로 HTML → Markdown 변환
        markdown_text = md(
            html_content,
            heading_style="ATX",
            strip=["script", "style"],
        )

        # 마크다운 파일 저장
        md_filename = original_stem + ".md"
        md_path = output_dir / md_filename
        md_path.write_text(markdown_text, encoding="utf-8")

        return ConversionResult(
            success=True,
            markdown=markdown_text,
            images=images_saved,
            metadata={},
        )

    except Exception as e:
        return ConversionResult(
            success=False,
            markdown="",
            images={},
            metadata={},
            error=str(e),
        )
    finally:
        # DOC에서 변환된 임시 DOCX 파일 정리
        if temp_docx and temp_docx.exists():
            try:
                temp_docx.unlink()
            except OSError:
                pass


def convert_docx_batch(
    input_dir: str | Path,
    output_dir: Optional[str | Path] = None,
    recursive: bool = False,
    save_images: bool = True,
) -> list[tuple[Path, ConversionResult]]:
    """
    폴더 내 모든 DOC/DOCX 파일을 Markdown으로 변환

    Args:
        input_dir: 입력 디렉토리
        output_dir: 출력 디렉토리 (None이면 입력과 같은 위치)
        recursive: 하위 폴더 포함 여부
        save_images: 이미지 저장 여부

    Returns:
        list[tuple[Path, ConversionResult]]: (파일경로, 변환결과) 리스트
    """
    input_dir = Path(input_dir)

    if output_dir is None:
        output_dir = input_dir
    else:
        output_dir = Path(output_dir)

    # DOC/DOCX 파일 목록 수집
    doc_files = []
    for ext in ("*.doc", "*.docx"):
        if recursive:
            doc_files.extend(input_dir.rglob(ext))
        else:
            doc_files.extend(input_dir.glob(ext))

    results = []

    for doc_path in doc_files:
        # 하위 폴더 구조 유지
        if recursive:
            relative_path = doc_path.parent.relative_to(input_dir)
            current_output_dir = output_dir / relative_path
        else:
            current_output_dir = output_dir

        result = convert_docx(doc_path, current_output_dir, save_images)
        results.append((doc_path, result))

    return results
