# 설계 계약

이 스킬은 런타임 reference PPTX에 의존하지 않고, 안정적인 내부 설계 계약을 사용해 PDF 원고 페이지에서 개념 익히기 PPTX 덱을 생성해야 한다.

## 파이프라인

1. 원본 PDF와 요청 페이지를 LLM 검토용 페이지 이미지로 렌더링한다.
2. LLM은 대상 `개념 익히기` 블록과 제외된 요청 페이지를 포함한 `llm-ir.json`을 작성한다.
3. LLM이 작성한 각 블록은 `PracticeBlock` IR로 정규화된다.
   - 원고 페이지
   - 개념 번호와 제목
   - 익히기 번호
   - 프롬프트
   - 순서가 있는 하위 문항
   - 리포트/디버깅용 원고 추출 줄
   - 레이아웃 유형: 현재 `vertical_list`, `two_column_grid`, `worked_stack`, `unknown`
   - 문항 번호, 식 텍스트, 사각 빈칸 수, 괄호형 빈칸 여부, 분수/지수 플래그, PDF 원고 bounding box, 행 index, 열 index를 포함한 문항 인벤토리
   - 문항별 사각 빈칸은 기본 위치 `[ ]` 또는 지수 위치 `^[ ]` 토큰으로 표현한다. `blank_count`가 있으면 전체 빈칸 토큰 수와 일치해야 한다.
   - 문항별 `source_lines`는 여러 줄 원고의 PDF 줄 구조를 보존한다.
   - 문항별 `display_lines`는 `source_lines`에서 파생하며, PDF의 별도 줄이 편집 가능한 PPTX 텍스트/수식 행이 되게 한다.
   - 문항별 `display_segments`는 `math`, `marker`, `korean_label`, `number_line` 등으로 타입을 지정해 혼합 행이 세그먼트별 올바른 폰트와 도형을 사용하게 한다. PDF 시각 간격을 보존해야 하면 `gap_after_in`을 포함할 수 있다. `a>b이면` 같은 수식/한국어 혼합 조건 행은 `display_lines` 텍스트만 남기지 말고 여기서 표현해야 한다.
   - 문항별 `layout_shapes`는 여러 표시 행을 묶는 왼쪽 브래킷처럼 텍스트가 아닌 레이아웃 전용 원고 선화를 담는다.
4. 렌더러는 LLM이 작성한 IR만 받는다. 결정론적 PyMuPDF 문항 추출은 이 스킬 패키지의 일부가 아니다.
5. 렌더러는 `assets/design/style-map.json`을 적용해 16:9 편집 가능 PPTX 슬라이드를 만든다.
6. 수식 렌더러는 일반 slash 분수와 superscript 텍스트를 편집 가능한 텍스트 상자, 분수선, 빈칸 도형으로 변환한다.
7. PPTX 작성 뒤 style-only OOXML 자산을 주입한다. 대상은 theme, slide masters, slide layouts, presentation properties, view properties, table styles다. 생성된 제목 슬라이드는 내장 제목 레이아웃으로, 문제 슬라이드는 내장 문제 레이아웃으로 연결하되 실제 슬라이드 내용과 순서는 보존한다. 배경과 콘텐츠 패널은 master/layout 파트에서 오며, 슬라이드 로컬 중복 overlay 도형으로 만들지 않는다.
8. 나타내기 애니메이션은 그룹을 대상으로 하는 `p:set` timing node로 주입한다.
9. 검증은 편집 가능한 텍스트, 내장 OOXML style part, 그룹 애니메이션, 보이는 페이지 라벨 없음, 화면 밖 객체 없음, 빈칸 도형 요구사항을 확인한다.
10. 리포트 검증은 요청 페이지가 모두 설명되었는지, 포함/제외 페이지 집합이 일관적인지, slide trace가 정규화된 문항 인벤토리를 모두 덮는지, 페이지 나눔 chunk가 연속적인지, 품질 경고가 해결되었거나 명시적으로 허용되었는지 확인한다.
11. 표준 LLM-IR build runner는 모든 검증 리포트를 작성하고 generation-report, editable-deck, OOXML-style, reference-pattern 검증이 모두 통과하지 않으면 실패한다.
12. 스킬 패키지 self-check는 공유 전에 필수 파일, JSON 파일, Python 소스 컴파일, bundled asset hash, 제거된 runtime dependency 부재를 검증한다.

## 런타임 설계 소스

- `assets/design/style-map.json`은 공통 런타임 스타일 소스다.
- `assets/concept-practice-style/`에는 이전 reference 분석에서 추출한 재사용 가능 제목/header 이미지와 style-only OOXML이 들어 있다. 이 자산은 특정 페이지 범위 런타임 의존성이 아니라 일반 개념 익히기 스타일 자산이다.
- `assets/concept-practice-style/ooxml-style-parts/`는 런타임에 생성된 PPTX 패키지에 적용된다. 스킬 공유 전 `style-parts-manifest.json`으로 검증해야 한다.
- 생성 슬라이드는 주입된 master/layout 배경과 패널에 의존해야 한다. style part가 이미 제공하는 경우, 전체 캔버스 배경 사각형이나 큰 문제 패널 도형을 슬라이드 XML에 직접 추가하지 않는다.
- 런타임에 raw reference PPTX 파일이 필요하면 안 된다.
- 로컬 기대 출력 폴더는 평가 보조 자료일 뿐이며, 스킬 지침이나 런타임 스크립트에서 참조하면 안 된다.

## 확장 규칙

- 반복되는 레이아웃 계열이 발견되면 측정값을 `assets/design/style-map.json`에 추가한다.
- 렌더러 전용 hack을 추가하기 전에 LLM IR 스키마에 필드를 추가한다.
- 원고에 보이는 사각 빈칸은 `raw_text`에서 `[ ]` 또는 `^[ ]` 토큰으로 보존한다. 보이는 모든 사각 빈칸은 정확히 하나의 IR 빈칸 토큰과 하나의 렌더링된 빈칸 도형에 대응해야 한다.
- 지수 빈칸, 같은 행 조각, 읽기 어려운 수식 정규화처럼 모호한 구조는 LLM-only 추출로 분류한다. 변환 IR을 만들기 위해 결정론적 PyMuPDF 문항 추출을 추가하지 않는다.
- 원고 PDF 레이아웃 유형을 먼저 보존한다. 한 열 목록, 2열 그리드, 연쇄 등식, 표, 수직선, 도형, 혼합 레이아웃을 구분한다.
- 렌더링 전에 문항 수준 원고 줄 구조를 보존한다. 여러 줄 문항은 `source_lines`를 유지하고, `display_lines`를 파생한 뒤, 렌더링 전에 그 줄을 타입이 있는 `display_segments`로 나눈다. 수식 세그먼트는 편집 가능한 수식으로, 한국어 라벨 세그먼트는 프롬프트와 같은 한국어 폰트로 유지한다. `a>b이면`, `a<b일 때`, `x+2는 5 이하이다`처럼 한 시각 행에 수식과 한국어 문법/라벨이 섞인 경우도 혼합 행이며, `math`와 `korean_label` 세그먼트로 표현해야 한다. 예를 들어 두 줄 원고 `(3) a>b이면 / a-1/2 ○ b-1/2`는 `line_index: 0`에 `a>b`와 `이면`, `line_index: 1`에 수식 `a-1/2 ○ b-1/2`를 둔다. 등호가 여러 개 있다는 이유만으로 한 줄 reference-planned 식을 여러 행으로 쪼개지 않는다.
- 프롬프트 텍스트는 글자 수 기준 수동 줄바꿈이나 강제 개행을 하지 않는다. 측정된 reference 텍스트 상자 폭 안에서 PowerPoint가 줄바꿈하도록 하나의 편집 가능 텍스트 문자열로 렌더링한다.
- 원고의 보이는 마커 도형은 `display_segments` marker 항목으로 보존한다. 단원 1 reference 덱에서 `분모의 소인수:` 앞 마커는 약 `0.400in x 0.292in`의 `rightArrow` auto shape이며, theme `bg1` fill과 `lumMod=50000`, outline 없음으로 표현된다.
- 원고의 보이는 연결선과 브래킷 선화는 `display_segments` 텍스트나 원고 크롭이 아니라 문항 수준 `layout_shapes`로 보존한다. 세 행을 잇는 왼쪽 브래킷은 `{"kind":"brace_connector","shape":"left_square_bracket","line_start":0,"line_end":2,"tick_lines":[0,1,2],"x_anchor":"before_display_lines","x_offset_in":-0.18,"width_in":0.16,"stroke_pt":1.0,"stroke_color":"#111111"}`를 사용한다. 렌더러는 이를 문항 행과 같은 reveal 그룹 안의 편집 가능한 선/사각형 도형으로 그려야 한다.
- 원고의 곡선 화살표가 문제 풀이 흐름의 일부라면 문항 수준 `layout_shapes`로 보존한다. `{"kind":"curved_arrow","shape":"cubic_bezier","name":"curved-arrow-1","x1_in":4.2,"y1_in":3.2,"control1_x_in":3.2,"control1_y_in":3.8,"control2_x_in":2.5,"control2_y_in":3.8,"x2_in":2.1,"y2_in":3.4,"stroke_color":"#878787","stroke_dash":"round_dot","stroke_pt":1.6,"arrowhead":"triangle"}`를 사용한다. 렌더러는 이를 connector, typed arrow, image crop이 아니라 같은 문항 그룹 안의 커스텀 cubic Bezier geometry를 가진 편집 가능한 PowerPoint freeform shape(`p:sp`)로 그려야 한다.
- 의미 있는 원고 가로 간격은 `display_segments[].gap_after_in`으로 보존한다. `3/4=3÷4=` 뒤에 멀리 떨어진 `(        )` 답란이 있는 행처럼, `raw_text`의 반복 공백만으로 간격을 표현하지 않는다.
- 축 domain, 눈금, 라벨, 경계점, 끝점 개폐, 구간 방향으로 재구성할 수 있는 수직선 그래픽은 `display_segments`의 number-line 항목으로 보존한다. 이는 공식 `pdf2concept-problem-pptx.llm-ir.v1` display segment kind다.
  - `kind`: `number_line`
  - `line_index`: 수직선이 나타날 표시 행
  - `text`: 세미콜론으로 구분된 key/value spec
  - 지원하는 `text` key: `min`, `max`, `point`, `direction`, `closed`, `ticks`, `labels`, `blank`
  - `min`/`max`: 보이는 축 domain의 수치 범위
  - `point`: 구간 수직선의 경계값
  - `direction`: `point`에서 시작하는 구간 방향인 `left` 또는 `right`
  - `closed`: 닫힌 점은 `true`, 열린 점은 `false`
  - `ticks`: 쉼표로 구분한 눈금 위치
  - `labels`: 표시할 눈금 라벨이며 `ticks`와 정렬한다.
  - `blank=true`: 학생이 표시할 수 있도록 꾸미지 않은 축/눈금/라벨만 그린다. 이때 `point`, `direction`, `closed`는 생략한다.
- 열린 끝점이 `-3`이고 오른쪽으로 칠한 원고 구간 예: `{"kind":"number_line","text":"min=-4;max=1;point=-3;direction=right;closed=false;ticks=-3;labels=-3","line_index":0}`. `x≥4`를 위한 빈 축 예: `{"kind":"number_line","text":"min=1;max=5;blank=true;ticks=1,2,3,4,5;labels=1,2,3,4,5","line_index":1}`.
- 수직선 geometry는 레이아웃 데이터이지 문제 텍스트가 아니다. 문항 `raw_text`는 의미 중심으로 유지하고, 원고에 실제 사각 빈칸이 따로 있지 않다면 `blank_count`는 `0`으로 둔다. 수직선은 문항 reveal 그룹 안의 편집 가능한 PowerPoint 도형으로 렌더링한다.
- 2열 블록은 추출 정규화 뒤 시각적 행 우선 순서로 렌더링하고 애니메이션을 적용한다.
- source box가 없고 명시적인 오른쪽 열 문항도 없는 2열 블록은 기본 `column_index=0`을 그대로 믿지 말고 표시 순서로 grid 위치를 만든다. `column = index % 2`, `row = index // 2`.
- 여러 줄 2열 worked block은 과밀을 피하기 위해 시각 행 단위로 페이지를 나눈다. 이어지는 슬라이드에서는 첫 행이 표준 첫 행 y 위치에서 시작하도록 로컬 행 index를 정규화한다.
- 리포트에는 PPTX를 열지 않고도 레이아웃과 수식 추출 실패를 감사할 수 있도록 정규화된 block/item 인벤토리가 포함되어야 한다.
- 리포트에는 애니메이션 순서와 페이지 나눔을 PPTX 없이 감사할 수 있도록 실제 문제 슬라이드 번호, block chunk index, item number, 기본 표시 item number, click reveal item number가 포함되어야 한다.
- 리포트에는 source block/item count, title/problem slide count, source blank/fraction/exponent/multiline count, 의도적 정규화, 의심스러운 추출 조건에 대한 non-fatal warning이 들어 있는 `quality_summary`가 포함되어야 한다.
- 수식 렌더링은 편집 가능해야 한다. equation screenshot 대신 텍스트 상자, 사각형 bar, superscript run, 빈칸 도형을 사용한다.
- 기준 `[ ]` 빈칸은 쌓은 분수 분모 안의 빈칸을 포함해 full-size no-fill 둥근 사각형이다. 측정된 style-map 값인 약 `0.551in x 0.551in`, line width `31750` EMU, theme `bg2` outline과 `lumMod=75000`, round cap, round join을 사용한다.
- 분모 빈칸에 보이는 지수가 있으면 full-size 분모 빈칸과 그 오른쪽 위의 별도 작은 지수 텍스트 상자로 렌더링한다. 단원 1 reference 덱에서 `[ ]²`, `[ ]³`의 지수는 빈칸 top보다 약 `0.112in` 위, 빈칸 오른쪽 끝보다 약 `0.023in` 뒤에 놓이며, 곱셈 기호는 분자 baseline이 아니라 빈칸 midline에 놓인다.
- 쌓은 분수 geometry는 `assets/design/style-map.json`의 bundled reference PPTX 측정값을 따른다. 중앙 정렬된 분자/분모 텍스트 상자, 수평 bar, 숫자 분수용 고정 digit slot 폭, `×`와 지수 run이 있는 곱셈 분모용 더 넓은 폭, 대칭적으로 너무 멀리 떨어지지 않고 분수선에 가까운 분모 상자를 사용한다.
- 지수 위치 빈칸은 LLM IR에서 `^[ ]`로 작성하고, 보이는 caret 없이 작게 올라간 빈칸 도형으로 렌더링한다. 연산자 뒤의 기본 위치 빈칸은 full-size `[ ]` 빈칸으로 유지한다. `a[ ]` 문맥 추론에 의존하지 않는다.
- visual refinement loop 리포트에는 각 iteration의 source PNG, PPTX slide PNG, comparison PNG, diagnostics, pass/fail, design adjustment가 포함되어야 한다.
- inline 숫자 수식 토큰은 기본 character-width 추정 뒤 style-map numeric width allowance(`math.inline_numeric_width_scale`, `math.inline_numeric_width_padding`, `math.inline_numeric_min_w`)를 사용해야 한다. 이렇게 해야 `0.15` 같은 굵은 `BT수식M` 소수 텍스트가 editable text box를 넘지 않는다.
- 수식은 편집 가능하게 유지한다. 재구성하기 어려운 도형이나 그림에만 atomic image를 사용한다.
- 생성 덱이 기대 시각 패턴과 다르면 reference deck copy에 의존하게 만들지 말고 design map과 renderer를 업데이트한다.
