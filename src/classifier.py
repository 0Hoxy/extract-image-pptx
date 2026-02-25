from pptx.shapes.picture import Picture

from src.config import MIN_IMAGE_WIDTH, MIN_IMAGE_HEIGHT


def filter_images(slide) -> list:
    """슬라이드에서 실제 사진만 필터링한다 (아이콘/로고 제외)."""
    return [
        shape for shape in slide.shapes
        if isinstance(shape, Picture)
        and shape.width >= MIN_IMAGE_WIDTH
        and shape.height >= MIN_IMAGE_HEIGHT
    ]


def classify_slots(images: list) -> dict[str, Picture]:
    """이미지 5개를 좌표 기반으로 MAIN, SUB1~4로 분류한다."""
    # left 기준으로 정렬하여 가장 왼쪽 = MAIN
    sorted_by_left = sorted(images, key=lambda s: s.left)
    main_image = sorted_by_left[0]
    sub_images = sorted_by_left[1:]

    # 나머지 4개를 top 기준으로 상위/하위 분리
    sorted_by_top = sorted(sub_images, key=lambda s: s.top)
    top_row = sorted(sorted_by_top[:2], key=lambda s: s.left)
    bottom_row = sorted(sorted_by_top[2:], key=lambda s: s.left)

    return {
        "MAIN": main_image,
        "SUB1": top_row[0],
        "SUB2": top_row[1],
        "SUB3": bottom_row[0],
        "SUB4": bottom_row[1],
    }
