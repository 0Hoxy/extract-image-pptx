import re

from src.config import TEXT_PATTERN_PAREN, TEXT_PATTERN_QUOTE
from src.models import SlideInfo


def parse_slide_text(slide, slide_idx: int) -> SlideInfo:
    """슬라이드 하단 텍스트에서 인물 정보를 파싱한다."""
    text_shapes = [
        shape for shape in slide.shapes
        if shape.has_text_frame and shape.text.strip()
    ]

    if not text_shapes:
        print(f"  [경고] 슬라이드 {slide_idx}: 텍스트를 찾을 수 없습니다. 기본 파일명 사용.")
        return SlideInfo(index=slide_idx, name=f"slide_{slide_idx}", year="", height="")

    # 하단 우선, 같은 높이면 좌측 우선 (우측 특이사항보다 좌측 이름이 먼저 선택)
    text_shapes.sort(key=lambda s: (-s.top, s.left))

    for shape in text_shapes:
        text = shape.text.strip()
        info = _try_parse(text, slide_idx)
        if info:
            return info

    # 파싱 실패 시 가장 하단 텍스트를 파일명으로 사용
    fallback = text_shapes[0].text.strip()
    fallback = re.sub(r'[\\/:*?"<>|]', '_', fallback)
    print(f"  [경고] 슬라이드 {slide_idx}: 텍스트 파싱 실패. '{fallback}' 사용.")
    return SlideInfo(index=slide_idx, name=fallback, year="", height="")


def _try_parse(text: str, slide_idx: int) -> SlideInfo | None:
    """텍스트에서 정규식 매칭을 시도한다."""
    # 형식1: "이름 (99) 178cm"
    match = TEXT_PATTERN_PAREN.search(text)
    if match:
        return SlideInfo(
            index=slide_idx,
            name=match.group(1).strip(),
            year=match.group(2),
            height=match.group(3),
        )

    # 형식2: "VINCENT 1997's 179cm 프랑스"
    match = TEXT_PATTERN_QUOTE.search(text)
    if match:
        return SlideInfo(
            index=slide_idx,
            name=match.group(1).strip(),
            year=match.group(2)[-2:],
            height=match.group(3),
        )

    return None
