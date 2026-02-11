"""
Marker 기반 PDF → Markdown 변환기
"""
import os
from pathlib import Path
from dataclasses import dataclass
from typing import Optional

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


def convert_pdf(
    pdf_path: str | Path,
    output_dir: Optional[str | Path] = None,
    save_images: bool = True,
) -> ConversionResult:
    """
    단일 PDF 파일을 Markdown으로 변환

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

    try:
        # Marker 변환 수행 (캐싱된 컨버터 재사용)
        converter = _get_converter()
        print(f"[후처리] converter 호출 시작: {pdf_path.name}")
        rendered = converter(str(pdf_path))
        print(f"[후처리] converter 완료, text_from_rendered 시작")

        # 마크다운 텍스트 추출 (marker-pdf 1.10+ API)
        markdown_text, _, images = text_from_rendered(rendered)
        print(f"[후처리] text_from_rendered 완료 (이미지 {len(images) if images else 0}개)")

        # 이미지 별도 폴더에 저장 + 마크다운 내 경로를 폴더 상대경로로 수정
        if save_images and images:
            images_dir_name = f"{pdf_path.stem}_images"
            images_dir = output_dir / images_dir_name
            images_dir.mkdir(exist_ok=True)

            for img_name, img_obj in images.items():
                img_path = images_dir / img_name
                img_obj.save(str(img_path))
                # 마크다운 내 이미지 경로를 상대경로로 수정 (꺾쇠로 감싸 특수문자 처리)
                rel_path = f"{images_dir_name}/{img_name}"
                markdown_text = markdown_text.replace(
                    f"({img_name})", f"(<{rel_path}>)"
                )

        # 마크다운 파일 저장
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
            success=False,
            markdown="",
            images={},
            metadata={},
            error=str(e)
        )


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
