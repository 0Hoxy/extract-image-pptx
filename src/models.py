from dataclasses import dataclass


@dataclass
class SlideInfo:
    """슬라이드에서 파싱한 인물 정보."""
    index: int
    name: str
    year: str
    height: str

    @property
    def base_name(self) -> str:
        return f"{self.name}({self.year}){self.height}cm"


@dataclass
class ImageData:
    """추출된 이미지 데이터."""
    slot: str
    blob: bytes
    ext: str
