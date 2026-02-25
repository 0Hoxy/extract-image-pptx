import re

# 슬롯 이름 정의
SLOT_NAMES = ["MAIN", "SUB1", "SUB2", "SUB3", "SUB4"]

# 작은 이미지 필터 기준 (100pt x 100pt, 1pt = 12700 EMU)
MIN_IMAGE_WIDTH = 100 * 12700
MIN_IMAGE_HEIGHT = 100 * 12700

# 텍스트 파싱 정규식
# 형식1: "이름 (99) 178cm"
TEXT_PATTERN_PAREN = re.compile(r"(.+?)\s*\((\d{2})\)\s*(\d{2,3})\s*(?:cm)?")
# 형식2: "VINCENT 1997's 179cm 프랑스"
TEXT_PATTERN_QUOTE = re.compile(r"(.+?)\s+(\d{4})(?:['\u2019]s)?\s+(\d{2,3})\s*cm")
