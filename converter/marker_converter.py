"""
Marker 기반 PDF → Markdown 변환기
"""
import math
import os
import shutil
import tempfile
from concurrent.futures import ThreadPoolExecutor, TimeoutError
from pathlib import Path
from dataclasses import dataclass
from typing import Optional

from pypdf import PdfReader, PdfWriter

# Windows MKL 메모리 누수 방지 (KMeans 후처리 멈춤 해결)
os.environ["OMP_NUM_THREADS"] = "1"

# GPU 사용 (CUDA 가능 시)
import torch
DEVICE = "cuda" if torch.cuda.is_available() else "cpu"
os.environ["TORCH_DEVICE"] = DEVICE
print(f"[GPU 확인] CUDA 사용 가능: {torch.cuda.is_available()}, 장치: {torch.cuda.get_device_name(0) if torch.cuda.is_available() else 'CPU'}")

from marker.converters.pdf import PdfConverter
from marker.models import create_model_dict
from marker.output import text_from_rendered


@dataclass
class ConversionResult:
    """변환 결과를 담는 데이터 클래스"""
    success: bool
    markdown: str
    images: dict  # {filename: image_data}
    metadata: dict
    error: Optional[str] = None


# 모델·컨버터를 전역으로 캐싱 (최초 1회만 로드)
_model_dict = None
_converter = None


def _get_models():
    """모델 딕셔너리를 가져오거나 생성"""
    global _model_dict
    if _model_dict is None:
        _model_dict = create_model_dict(device=DEVICE)
    return _model_dict


def _get_converter():
    """PdfConverter 인스턴스를 가져오거나 생성 (재사용)"""
    global _converter
    if _converter is None:
        _converter = PdfConverter(
            artifact_dict=_get_models(),
            config={
                "recognition_batch_size": 32,
                "ray_batch_size": 32,
                "drop_repeated_text": True,
                "drop_repeated_table_text": True,
            },
        )
    return _converter


# --- 대용량 PDF 분할 변환 헬퍼 ---

# 분할 기준 (바이트)
_SPLIT_THRESHOLD_2 = 5 * 1024 * 1024    # 5MB 이상 → 2분할
_SPLIT_THRESHOLD_4 = 10 * 1024 * 1024   # 10MB 이상 → 4분할

# 파트별 변환 타임아웃 (초)
_PART_TIMEOUT = 10 * 60  # 10분


def _get_split_count(file_size: int) -> int:
    """파일 크기에 따라 분할 수를 반환 (1, 2, 4)"""
    if file_size >= _SPLIT_THRESHOLD_4:
        return 4
    elif file_size >= _SPLIT_THRESHOLD_2:
        return 2
    return 1


def _split_pdf(pdf_path: Path, num_splits: int) -> tuple[list[Path], str]:
    """
    PDF를 num_splits 등분하여 임시 디렉토리에 저장

    Returns:
        (분할 PDF 경로 리스트, 임시 디렉토리 경로)
    """
    reader = PdfReader(str(pdf_path))
    total_pages = len(reader.pages)
    pages_per_split = math.ceil(total_pages / num_splits)

    tmp_dir = tempfile.mkdtemp(prefix="pdf_split_")
    split_paths = []

    for i in range(num_splits):
        start = i * pages_per_split
        end = min(start + pages_per_split, total_pages)
        if start >= total_pages:
            break

        writer = PdfWriter()
        for page_idx in range(start, end):
            writer.add_page(reader.pages[page_idx])

        part_path = Path(tmp_dir) / f"{pdf_path.stem}_part{i + 1}.pdf"
        with open(part_path, "wb") as f:
            writer.write(f)
        split_paths.append(part_path)

        print(f"[분할] 파트 {i + 1}/{num_splits}: 페이지 {start + 1}~{end} ({end - start}페이지)")

    return split_paths, tmp_dir


def _convert_single_pdf(pdf_path: Path, timeout: int = _PART_TIMEOUT) -> tuple[str, dict]:
    """
    단일 PDF를 Marker로 변환하여 (markdown_text, images_dict) 반환
    파일 저장은 하지 않음 (호출자가 처리)
    timeout초 내에 완료되지 않으면 TimeoutError 발생
    """
    def _run():
        converter = _get_converter()
        rendered = converter(str(pdf_path))
        markdown_text, _, images = text_from_rendered(rendered)
        return markdown_text, images or {}

    with ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(_run)
        try:
            return future.result(timeout=timeout)
        except TimeoutError:
            raise TimeoutError(
                f"{timeout // 60}분 타임아웃 초과 ({pdf_path.name})"
            )


def convert_pdf(
    pdf_path: str | Path,
    output_dir: Optional[str | Path] = None,
    save_images: bool = True,
) -> ConversionResult:
    """
    단일 PDF 파일을 Markdown으로 변환
    5MB 이상이면 자동으로 분할 변환 후 통합

    Args:
        pdf_path: PDF 파일 경로
        output_dir: 출력 디렉토리 (None이면 PDF와 같은 위치)
        save_images: 이미지 저장 여부

    Returns:
        ConversionResult: 변환 결과
    """
    pdf_path = Path(pdf_path)

    if not pdf_path.exists():
        return ConversionResult(
            success=False,
            markdown="",
            images={},
            metadata={},
            error=f"파일을 찾을 수 없습니다: {pdf_path}"
        )

    if output_dir is None:
        output_dir = pdf_path.parent
    else:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

    file_size = pdf_path.stat().st_size
    num_splits = _get_split_count(file_size)

    if num_splits == 1:
        return _convert_pdf_single(pdf_path, output_dir, save_images)
    else:
        return _convert_pdf_split(pdf_path, output_dir, save_images, num_splits)


def _convert_pdf_single(
    pdf_path: Path, output_dir: Path, save_images: bool
) -> ConversionResult:
    """분할 없이 기존 방식으로 변환"""
    try:
        converter = _get_converter()
        print(f"[후처리] converter 호출 시작: {pdf_path.name}")
        rendered = converter(str(pdf_path))
        print(f"[후처리] converter 완료, text_from_rendered 시작")

        markdown_text, _, images = text_from_rendered(rendered)
        print(f"[후처리] text_from_rendered 완료 (이미지 {len(images) if images else 0}개)")

        if save_images and images:
            images_dir_name = f"{pdf_path.stem}_images"
            images_dir = output_dir / images_dir_name
            images_dir.mkdir(exist_ok=True)

            for img_name, img_obj in images.items():
                img_path = images_dir / img_name
                img_obj.save(str(img_path))
                rel_path = f"{images_dir_name}/{img_name}"
                markdown_text = markdown_text.replace(
                    f"({img_name})", f"(<{rel_path}>)"
                )

        md_filename = pdf_path.stem + ".md"
        md_path = output_dir / md_filename
        md_path.write_text(markdown_text, encoding="utf-8")

        return ConversionResult(
            success=True,
            markdown=markdown_text,
            images=images,
            metadata={},
        )

    except Exception as e:
        return ConversionResult(
            success=False, markdown="", images={}, metadata={}, error=str(e)
        )


def _convert_pdf_split(
    pdf_path: Path, output_dir: Path, save_images: bool, num_splits: int
) -> ConversionResult:
    """대용량 PDF를 분할 변환 후 통합. 실패한 파트는 건너뛰고 성공한 파트만 저장"""
    file_size_mb = pdf_path.stat().st_size / (1024 * 1024)
    print(f"[분할 변환] {pdf_path.name} ({file_size_mb:.1f}MB) → {num_splits}분할 변환 시작")

    tmp_dir = None
    try:
        # 1) PDF 분할
        split_paths, tmp_dir = _split_pdf(pdf_path, num_splits)

        # 2) 각 파트 변환 (실패 시 건너뛰기)
        md_parts = []
        all_images = {}
        images_dir_name = f"{pdf_path.stem}_images"
        success_count = 0
        fail_count = 0

        for i, part_path in enumerate(split_paths):
            part_num = i + 1
            print(f"[분할 변환] 파트 {part_num}/{len(split_paths)} 변환 중...")

            try:
                md_text, images = _convert_single_pdf(part_path)
            except Exception as e:
                fail_count += 1
                print(f"[분할 변환] 파트 {part_num}/{len(split_paths)} 실패: {e}")
                continue

            success_count += 1
            print(f"[분할 변환] 파트 {part_num}/{len(split_paths)} 완료 (이미지 {len(images)}개)")

            # 이미지 키에 파트 접두사 추가하여 충돌 방지
            if save_images and images:
                images_dir = output_dir / images_dir_name
                images_dir.mkdir(exist_ok=True)

                for img_name, img_obj in images.items():
                    new_img_name = f"part{part_num}_{img_name}"
                    img_path = images_dir / new_img_name
                    img_obj.save(str(img_path))
                    all_images[new_img_name] = img_obj

                    # 마크다운 내 이미지 경로 수정
                    rel_path = f"{images_dir_name}/{new_img_name}"
                    md_text = md_text.replace(
                        f"({img_name})", f"(<{rel_path}>)"
                    )

            md_parts.append(md_text)

        # 3) 결과 확인
        if not md_parts:
            return ConversionResult(
                success=False, markdown="", images={}, metadata={},
                error=f"모든 파트({len(split_paths)}개) 변환 실패"
            )

        # 4) MD 통합 및 저장
        merged_markdown = "\n\n---\n\n".join(md_parts)

        md_filename = pdf_path.stem + ".md"
        md_path = output_dir / md_filename
        md_path.write_text(merged_markdown, encoding="utf-8")

        status = "완료" if fail_count == 0 else "부분 완료"
        print(f"[분할 변환] {status} → {md_filename} "
              f"(성공 {success_count}/{len(split_paths)}, 이미지 총 {len(all_images)}개)")

        return ConversionResult(
            success=True,
            markdown=merged_markdown,
            images=all_images,
            metadata={
                "split_count": len(split_paths),
                "success_count": success_count,
                "fail_count": fail_count,
            },
        )

    except Exception as e:
        return ConversionResult(
            success=False, markdown="", images={}, metadata={}, error=str(e)
        )
    finally:
        # 임시 분할 파일 정리
        if tmp_dir and os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir, ignore_errors=True)


def convert_batch(
    input_dir: str | Path,
    output_dir: Optional[str | Path] = None,
    recursive: bool = False,
    save_images: bool = True,
) -> list[tuple[Path, ConversionResult]]:
    """
    폴더 내 모든 PDF 파일을 Markdown으로 변환

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

    # PDF 파일 목록 수집
    if recursive:
        pdf_files = list(input_dir.rglob("*.pdf"))
    else:
        pdf_files = list(input_dir.glob("*.pdf"))

    results = []

    for pdf_path in pdf_files:
        # 하위 폴더 구조 유지
        if recursive:
            relative_path = pdf_path.parent.relative_to(input_dir)
            current_output_dir = output_dir / relative_path
        else:
            current_output_dir = output_dir

        result = convert_pdf(pdf_path, current_output_dir, save_images)
        results.append((pdf_path, result))

    return results
