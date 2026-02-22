#!/bin/bash
# 02_gonmun_example.sh — 공문 템플릿으로 문서 빌드 예제
#
# 공문 템플릿의 placeholder를 실제 내용으로 교체한 section0.xml을 작성하여 빌드한다.

set -euo pipefail

SKILL_DIR="$(cd "$(dirname "$0")/.." && pwd)"
VENV="${VENV:-$(cd "$SKILL_DIR/../.." && pwd)/.venv/bin/activate}"
source "$VENV"

# 공문 템플릿으로 빌드 (기본 placeholder 포함)
OUTPUT="/tmp/gonmun_example.hwpx"
python3 "$SKILL_DIR/scripts/build_hwpx.py" \
  --template gonmun \
  --title "출판 콘텐츠 협의 건" \
  --creator "골든래빗" \
  --output "$OUTPUT"

# 검증
python3 "$SKILL_DIR/scripts/validate.py" "$OUTPUT"

# 텍스트 추출 확인
echo ""
echo "=== 추출된 텍스트 ==="
python3 "$SKILL_DIR/scripts/text_extract.py" "$OUTPUT"

echo ""
echo "생성 완료: $OUTPUT"
