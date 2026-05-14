# HWPX Raw XML 편집 함정 모음 (Mac 한글 무한루프 방지)

이 문서는 HWPX를 AI/스크립트로 직접 raw XML 편집할 때 반복적으로 마주친 실패 사례와 해결책을 정리한다. **`validate.py`와 `page_guard.py`만 통과해서는 충분하지 않다**. Mac 한글(Hancom Office HWP 12.x)이 거부하거나 CPU 100% 무한루프에 빠지는 OWPML 위반 패턴이 있다.

## 진단 우선순위 (Mac 한글에서 안 열림/멈춤 발생 시 점검 순서)

1. **단락 id 중복** — `<hp:p>` 단위 id가 모두 unique한가?
2. **hp:linesegarray 누락** — 텍스트 있는 모든 `<hp:p>`에 linesegarray가 있는가?
3. **rowSpan/colSpan 점유 위반** — rowSpan>1 셀이 점유한 행/열 자리에 빈 `<hp:tc>`가 또 있는가?
4. **cellSz width 합 ≠ 표 width** — 각 행의 cellSz width 합계가 `<hp:tbl>`의 sz width와 같은가?
5. **borderFillIDRef·paraPrIDRef·charPrIDRef 미정의** — header.xml에 정의되지 않은 ID를 참조하는가?
6. **lxml 직렬화 부작용** — 전체 트리를 `etree.tostring`으로 다시 직렬화하면 namespace 선언 순서·attribute 표기가 변해 Mac 한글이 거부할 수 있다.

## 함정 1: 단락 id 중복

### 증상
Mac 한글이 문서를 열 때 CPU 100%로 멈춤. validate.py와 page_guard.py는 통과.

### 원인
AI가 새 `<hp:p>` 단락을 추가할 때 id 속성에 placeholder 값(예: `2147483648` = 2^31, `0`)을 그대로 두면 동일 id가 수십~수백 개 발생. 한글은 단락 id로 cross-reference(주석·하이퍼링크·줄세그먼트 캐시 등)를 계산하므로 중복이 있으면 거부.

### 발견 방법
```python
import zipfile
from lxml import etree
from collections import Counter
NS = {"hp": "http://www.hancom.co.kr/hwpml/2011/paragraph"}
with zipfile.ZipFile(path) as zf:
    s = etree.fromstring(zf.read("Contents/section0.xml"))
ids = [p.get("id") for p in s.xpath(".//hp:p", namespaces=NS)]
dups = [i for i, c in Counter(ids).items() if c > 1]
print(f"total={len(ids)} unique={len(set(ids))} dups={dups}")
```

### 수정 (raw 바이트 patch — lxml 직렬화 사용 금지)
1. 모든 `<hp:p ... id="VALUE" ...>` 매치
2. 중복 id 검출 후, 첫 등장은 유지하고 두 번째 등장부터 새 unique id(예: 5000000000번대)로 raw 바이트만 교체

## 함정 2: hp:linesegarray 누락

### 증상
Mac 한글이 문서를 열 때 "열다가 멈춤"처럼 느려짐. 작은 문서면 결국 열리지만 큰 표가 있으면 CPU 100% 무한루프.

### 원인
한글은 각 `<hp:p>`의 `<hp:linesegarray>`를 줄 위치 렌더링 캐시로 사용. 텍스트 있는 단락에 linesegarray가 없으면 한글이 처음 열 때 줄 위치를 모두 재계산. 누락 단락이 80개 넘어가면 사실상 안 열림.

### 발견 방법
```python
for p in s.xpath(".//hp:p", namespaces=NS):
    t_len = sum(len("".join(t.itertext())) for t in p.xpath(".//hp:t", namespaces=NS))
    if t_len > 0 and p.find("hp:linesegarray", namespaces=NS) is None:
        print(f"MISSING lineseg: id={p.get('id')}")
```

### 수정 (raw 바이트 patch)
각 누락 hp:p의 `</hp:p>` 닫는 태그 직전에 placeholder lineseg array 삽입:
```xml
<hp:linesegarray>
  <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000"
              baseline="850" spacing="600" horzpos="0" horzsize="{HORZ}" flags="393216"/>
</hp:linesegarray>
```

`horzsize` 값:
- 본문 단락: 본문폭(예: 37420 또는 42520 HWPUNIT)
- 표 셀 단락: 부모 `<hp:tc>`의 `<hp:cellSz>` width

한글이 첫 편집/저장 시 정확한 값으로 다시 그림.

## 함정 3: rowSpan/colSpan 점유 자리에 빈 셀 존재

### 증상
표가 들어 있는 문서를 한글이 열 때 CPU 100% 무한루프. 다른 모든 검증(ID·lineseg·cellSz width)은 통과.

### 원인 (OWPML 규칙)
정상 HWPX 표에서 rowSpan>1 또는 colSpan>1 셀이 점유한 자리에는 **빈 `<hp:tc>` 셀이 존재하지 않는다**. 그 행의 cell 수가 줄어든다.

예: 26x3 표에서 r1의 첫 셀이 rowSpan=4라면 → r2, r3, r4는 각각 2 cells (colAddr=1, 2만). r1은 3 cells (colAddr=0, 1, 2). 점유된 (0, r2), (0, r3), (0, r4) 위치에 빈 셀을 두면 안 됨.

AI가 표를 생성할 때 "11x8 모든 행에 8 cells"로 균등 출력하면서 rowSpan을 설정하면 이 규칙을 어김.

### 발견 방법
```python
for tbl in s.xpath(".//hp:tbl", namespaces=NS):
    rc = int(tbl.get("rowCnt", 0))
    cc = int(tbl.get("colCnt", 0))
    # 각 행의 cell 수와 rowSpan 점유 관계 점검
    occupied = {}  # (col, row) -> True
    for ri, tr in enumerate(tbl.xpath("./hp:tr", namespaces=NS)):
        for tc in tr.xpath("./hp:tc", namespaces=NS):
            addr = tc.find("hp:cellAddr", namespaces=NS)
            cs = tc.find("hp:cellSpan", namespaces=NS)
            col = int(addr.get("colAddr"))
            row = int(addr.get("rowAddr"))
            rspan = int(cs.get("rowSpan", 1))
            cspan = int(cs.get("colSpan", 1))
            if (col, row) in occupied:
                print(f"VIOLATION: ({col},{row}) already occupied by rowSpan/colSpan")
            for dr in range(rspan):
                for dc in range(cspan):
                    occupied[(col+dc, row+dr)] = True
```

### 수정 (raw 바이트 patch)
점유 자리에 있는 빈 `<hp:tc>` 요소 전체를 raw에서 식별하여 제거. 식별은 `<hp:cellAddr colAddr="X" rowAddr="Y"/>` 마커로 위치를 찾은 뒤 그 마커를 감싸는 `<hp:tc>...</hp:tc>` 범위를 추출.

### 정상 표 패턴 예 (v17 26x3 교차 안정 자질 표)
```
r0: 3 cells (헤더)              | (0,0) | (1,0) | (2,0) |
r1: 3 cells (rowSpan=4 시작)    | (0,1) rs=4 | (1,1) | (2,1) |
r2: 2 cells (col 0 점유됨)                  | (1,2) | (2,2) |
r3: 2 cells                                | (1,3) | (2,3) |
r4: 2 cells                                | (1,4) | (2,4) |
r5: 3 cells (rowSpan=6 시작)    | (0,5) rs=6 | (1,5) | (2,5) |
...
```

## 함정 3.5: 양쪽정렬(JUSTIFY) 단락에 단일 lineseg placeholder = 자간 폭주

### 증상
한글에서 문서가 열리지만 일부 본문 단락의 **자간이 비정상적으로 넓게** 표시됨. 글자 사이가 텅 비어 보임.

### 원인
- paraPr가 `<hh:align horizontal="JUSTIFY"/>` (양쪽정렬)인 단락은 한글이 각 줄을 본문폭에 맞춰 자간을 자동 조절
- linesegarray가 정확하게 줄별로 분할된 lineseg를 가지면 한글이 각 줄의 텍스트 양에 맞게 자간 산정
- **placeholder로 textpos=0 단일 lineseg**만 두면 한글이 "전체 텍스트가 한 줄"로 해석 → 200자, 500자, 1000자를 한 줄 폭(예: horzsize=37420)에 맞추려고 자간을 극단적으로 늘림

### 발견 방법 (일반화된 조건 — 권장)

**시그니처(textpos=0 vertpos=0)에 의존하지 말 것**. AI마다 placeholder vertpos를 다르게 둬서 시그니처 매칭으로는 누락이 발생. 대신 **lineseg 수가 텍스트 길이로 추정한 줄 수보다 적은 단락**을 잡는다.

```python
import math
# JUSTIFY paraPr 식별
justify_ppr = set()
for ppr in h.xpath(".//hh:paraPr", namespaces=NS):
    al = ppr.find("hh:align", namespaces=NS)
    if al is not None and al.get("horizontal") == "JUSTIFY":
        justify_ppr.add(ppr.get("id"))

# 위반 단락 찾기: 양쪽정렬 본문(표 외부) + lineseg 수 부족
# 한국어 본문은 horzsize=37420 폭에서 한 줄 약 50~60자
# 보수적으로 60자/줄 가정 → 줄 수 = ceil(text_len/60)
for p in s.xpath(".//hp:p", namespaces=NS):
    if is_in_table_cell(p): continue
    if p.get("paraPrIDRef") not in justify_ppr: continue
    t_len = sum(len("".join(t.itertext())) for t in p.xpath(".//hp:t", namespaces=NS))
    if t_len < 50: continue
    lsa = p.find("hp:linesegarray", namespaces=NS)
    if lsa is None: continue  # 이미 제거됨 → 한글이 재계산할 것
    n_lseg = len(lsa.findall("hp:lineseg", namespaces=NS))
    expected_min = max(1, math.ceil(t_len / 60))
    if n_lseg < expected_min:
        print(f"VIOLATION: id={p.get('id')} text_len={t_len} lineseg={n_lseg} expected≥{expected_min}")
```

### 수정 (raw 바이트 patch)
**해당 단락의 `<hp:linesegarray>...</hp:linesegarray>`를 통째로 제거**. 한글이 첫 열기 시 정확한 줄 분할로 재계산 → 정상 자간 복원.

표 셀 안의 단락은 lineseg를 유지해야 함(셀이 자동 양쪽정렬되지만 줄 분할이 셀 폭에 따라 분명함). **표 밖 본문 단락 + JUSTIFY + 단일 placeholder lineseg + 30자 이상**일 때만 제거.

### 트레이드오프
- placeholder lineseg를 추가한 이유: lineseg 누락 단락이 80개 넘으면 한글 첫 열기에서 무한루프
- 자간 문제 회피를 위해 양쪽정렬 본문 단락만 placeholder 제거
- 결과: 표 셀 단락은 placeholder 유지(자간 영향 미미), 본문 단락은 한글이 재계산(자간 정확)

### 짧은 단락은 영향 없음
30자 미만 단락은 placeholder 단일 lineseg를 두어도 한 줄 안에 들어가서 자간 영향이 거의 없음. 자간 문제는 100자+ 단락에서 두드러짐(특히 500자+ 단락은 극단적).

## 함정 4: lxml 직렬화로 인한 OWPML 호환성 손실

### 증상
raw 바이트로 단순 텍스트 치환은 잘 되는데, lxml로 트리를 파싱하고 `etree.tostring`으로 다시 직렬화한 결과는 Mac 한글이 거부.

### 원인 (추정)
lxml의 직렬화는:
- namespace 선언 순서를 정규화
- attribute 순서를 정규화
- self-closing 태그(`<foo/>`) vs 빈 태그(`<foo></foo>`) 형식이 입력과 달라질 수 있음
- xml declaration의 `standalone` 속성 처리

이런 미세한 차이를 Mac 한글이 거부할 수 있음. Windows 한글은 더 관대.

### 권장 (강력)
**원본 OWPML 구조를 100% 보존하려면 raw 바이트 patch만 사용**. lxml은 진단(어디에 무엇이 있는지)에만 쓰고, 수정은 raw 바이트로:
- attribute 값 교체: 정규식으로 정확한 위치만 치환
- element 추가: 닫는 태그 직전에 raw XML 문자열 삽입
- element 제거: 시작 태그부터 닫는 태그까지 byte 범위 추출 후 삭제

## 함정 5: Mac 한글 무한루프 진입 시 후속 처리

### 증상
잘못된 HWPX를 한 번 열면 한글이 CPU 100%로 멈춤. 강제 종료 후 재실행해도 자동복구가 같은 파일을 다시 열려고 시도 → 또 무한루프.

### 해결 절차
```bash
# 1. 한글 강제 종료
pkill -9 -f "Hancom Office HWP"

# 2. Saved Application State 삭제 (자동복구 차단)
rm -rf ~/Library/Saved\ Application\ State/com.hancom.*

# 3. Temp 디렉터리 정리
rm -rf /private/var/folders/*/T/Hwp120

# 4. (선택) Caches 정리
rm -rf ~/Library/Caches/com.hancom.office.hwp12.mac.general

# 5. 한글 재실행
```

### Containers·Preferences는 보존
- `~/Library/Containers/com.hancom.*.HwpQLExtPreview` — QuickLook 미리보기 (본문 열기와 무관)
- `~/Library/Preferences/com.hancom.*.plist` — 환경설정 (초기화 시 사용자 설정 손실)

## 함정 6: validate.py와 page_guard.py의 한계

이 두 스크립트는 **구조적 무결성과 텍스트 길이 안정성**만 검사한다. 다음 항목은 검사하지 않으므로 별도로 점검 필요:
- 단락 id 중복 여부
- linesegarray 존재 여부
- rowSpan/colSpan 점유 충돌
- cellSz width 합과 표 width 일치
- IDRef가 header.xml에 정의되어 있는지
- lxml 직렬화로 인한 형식 변화

**최종 검증은 항상 Mac 한글 + Windows 한글 두 환경에서 실제 열어볼 것**.

## 안전한 HWPX 편집 워크플로우 (권장)

1. **원본 ZIP을 그대로 두고 `Contents/section0.xml`만 raw 바이트 patch**
   - ZIP entry 순서·압축·timestamp 보존
   - 다른 모든 파일(header.xml, content.hpf, masterpage0.xml 등)은 손대지 않음

2. **편집은 정규식 또는 정확한 바이트 범위 삽입/삭제**
   - 전체 트리를 lxml로 재직렬화하지 말 것
   - lxml은 진단 도구로만 사용

3. **패치 후 검증 6단계** (이 문서 함정 1~6 모두 통과)

4. **사고 방지 체크리스트**
   - [ ] hp:p id 모두 unique
   - [ ] 텍스트 있는 hp:p 100%에 linesegarray 존재
   - [ ] 표의 rowSpan/colSpan 점유 자리에 빈 tc 없음
   - [ ] 표의 cellSz width 합 = 표 sz width
   - [ ] 모든 IDRef가 header.xml에 정의됨
   - [ ] Mac 한글 + Windows 한글 양쪽에서 실제 열기 성공

## 사고 이력

- **2026-05-12 v18 → v23 사고**:
  - v18 (다른 세션 AI가 raw XML로 11×8 표 추가): placeholder ID 178+30개, lineseg 누락 83개, rowSpan 점유 자리에 빈 셀 5개 — 3중 위반
  - v19 (raw로 ID만 fix): 여전히 lineseg 누락 + rowSpan 위반 → 안 열림
  - v20 (lxml로 ID + lineseg 직렬화): lxml 직렬화 부작용 추가 → 안 열림
  - v21 (raw로 ID + lineseg fix): rowSpan 위반 여전 → 안 열림 (CPU 100% 무한루프)
  - v22 (raw로 ID + lineseg + 빈 tc 제거): 열리지만 양쪽정렬 본문 단락에 placeholder 단일 lineseg → 자간 폭주
  - v23 (양쪽정렬 단락 중 시그니처 textpos=0 vertpos=0인 6개만 제거): 일부 자간 정상화. **그러나 다른 vertpos placeholder 6개 추가로 누락**
  - v24 (일반화된 조건: lineseg 수 < ceil(text_len/60) 으로 12개 모두 제거): 안전 후보. **시그니처 매칭 대신 줄 수 추정으로 일반화하는 것이 핵심 교훈**
