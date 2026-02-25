# PPTX 이미지 추출기

PPTX 파일에서 슬라이드별 이미지를 위치 기반으로 분류(MAIN, SUB1~4)하고, 슬라이드 텍스트에서 파싱한 정보로 파일명을 자동 생성하여 저장하는 도구.

## 지원하는 PPTX 형식

### 슬라이드 레이아웃

각 슬라이드에 **5장의 사진**이 아래와 같이 배치되어 있어야 합니다.

```
              SUB1    SUB2
  MAIN
              SUB3    SUB4

  [이름 정보 텍스트]
```

- **MAIN**: 좌측에 배치된 큰 사진 (1장)
- **SUB1~4**: 우측에 2x2 격자로 배치된 사진 (4장)
- 100pt x 100pt 이하의 아이콘/로고는 자동 무시

### 텍스트 형식

슬라이드 하단 좌측에 인물 정보 텍스트가 있어야 합니다. 두 가지 형식을 지원합니다.

| 형식 | 예시 | 파일명 변환 |
|------|------|------------|
| `이름 (년생2자리) 키cm` | `홍길동 (99) 178cm` | `홍길동(99)178cm` |
| `이름 년생4자리's 키cm 국적` | `VINCENT 1997's 179cm 프랑스` | `VINCENT(97)179cm` |

### 출력 결과

```
output/
└── PPTX파일명/
    ├── MAIN/
    │   └── VINCENT(97)179cm_MAIN.png
    ├── SUB1/
    │   └── VINCENT(97)179cm_SUB1.png
    ├── SUB2/
    ├── SUB3/
    └── SUB4/
```

- 동일 파일명 충돌 시 `(1)`, `(2)` suffix 자동 추가
- 이미지 5장이 아닌 슬라이드는 스킵 (마지막에 제외 목록 출력)
- 5장 중 하나라도 추출 실패 시 해당 슬라이드 전체 폐기

## 시작하기

### 사전 요구사항

- Python 3.12+

### 설치 및 실행

```bash
git clone https://github.com/0Hoxy/extract-image-pptx.git
cd extract-image-pptx
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### CLI 실행

```bash
python main.py <PPTX파일경로>

# 출력 폴더 지정
python main.py <PPTX파일경로> --output ./result
```

### GUI 실행

```bash
python gui.py
```

> macOS에서 tkinter 오류 발생 시: `brew install python-tk@3.12`

## Windows exe 배포

### GitHub Actions 자동 빌드

태그 푸시 시 자동으로 Windows exe가 빌드되어 Release에 업로드됩니다.

```bash
git tag v1.0.0
git push origin --tags
```

### 로컬 빌드 (Windows)

```bash
pip install -r requirements.txt
pyinstaller build.spec
# → dist/PPTX_이미지_추출기.exe
```

## 프로젝트 구조

```
extract-image-pptx/
├── src/
│   ├── config.py        # 상수, 정규식 패턴
│   ├── models.py        # SlideInfo, ImageData 데이터 모델
│   ├── parser.py        # 슬라이드 텍스트 파싱
│   ├── classifier.py    # 이미지 필터링 및 슬롯 분류
│   ├── storage.py       # 파일 저장/삭제/충돌 처리
│   └── extractor.py     # 추출 오케스트레이터
├── main.py              # CLI 진입점
├── gui.py               # GUI 진입점
├── build.spec           # PyInstaller 빌드 설정
├── requirements.txt
└── .github/workflows/
    └── build.yml        # Windows exe 자동 빌드
```

## 기술 스택

- **Python 3.12+**
- **python-pptx**: PPTX 파일 파싱
- **tkinter**: GUI (Python 내장)
- **PyInstaller**: Windows exe 패키징
