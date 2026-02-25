from pathlib import Path

from pptx.shapes.picture import Picture

from src.models import ImageData


def extract_image_data(picture: Picture, slot: str) -> ImageData:
    """Picture Shape에서 바이너리 데이터와 확장자를 추출한다."""
    image = picture.image
    ext = image.content_type.split("/")[-1].lower()

    # 포맷 정규화
    ext_map = {"jpeg": "jpg", "x-ms-bmp": "bmp"}
    ext = ext_map.get(ext, ext)

    # MPO 등 미지원 포맷 → jpg로 저장
    supported = {"bmp", "gif", "jpg", "png", "tiff", "wmf"}
    if ext not in supported:
        ext = "jpg"

    return ImageData(slot=slot, blob=image.blob, ext=ext)


def save_image(image_data: ImageData, base_name: str, output_path: Path) -> Path:
    """이미지 데이터를 슬롯 폴더에 저장한다."""
    folder = output_path / image_data.slot
    filename = f"{base_name}_{image_data.slot}.{image_data.ext}"
    filepath = folder / filename

    # 파일명 충돌 처리
    counter = 1
    while filepath.exists():
        filename = f"{base_name}({counter})_{image_data.slot}.{image_data.ext}"
        filepath = folder / filename
        counter += 1

    filepath.write_bytes(image_data.blob)
    return filepath


def cleanup_images(saved_files: list[Path]) -> None:
    """추출 실패 시 이미 저장된 이미지를 삭제한다."""
    for filepath in saved_files:
        if filepath.exists():
            filepath.unlink()
            print(f"  [정리] {filepath.name} 삭제")
