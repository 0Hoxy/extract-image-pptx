import argparse
import sys
import time

from extractor import extract_images_from_pptx


def main():
    parser = argparse.ArgumentParser(
        description="PPTX 파일에서 슬롯별(MAIN, SUB1~4) 이미지를 추출합니다."
    )
    parser.add_argument("pptx_file", help="추출할 PPTX 파일 경로")
    parser.add_argument(
        "--output", "-o",
        default="./output",
        help="출력 디렉토리 (기본값: ./output)",
    )

    args = parser.parse_args()

    start = time.time()

    try:
        extract_images_from_pptx(args.pptx_file, args.output)
    except FileNotFoundError as e:
        print(f"[오류] {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"[오류] 처리 중 문제가 발생했습니다: {e}", file=sys.stderr)
        sys.exit(1)

    elapsed = time.time() - start
    print(f"\n소요 시간: {elapsed:.2f}초")


if __name__ == "__main__":
    main()
