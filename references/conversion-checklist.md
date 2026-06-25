# 변환 체크리스트

## 입력

- 원본 PDF 경로
- 요청 페이지 범위. 사용자가 인쇄 페이지 번호를 지정하지 않은 한 PDF viewer의 페이지 번호를 기준으로 한다.
- Reference PPTX 경로. 사용자가 보정용으로 명시적으로 제공한 경우에만 사용한다. 일반 런타임 생성은 복사된 reference deck이 아니라 스킬 내부 설계 계약을 사용해야 한다.
- 출력 PPTX 경로
- 문제 전용 슬라이드, 풀이 슬라이드, 정답 요약 슬라이드, 또는 전체 포함 여부
- 공유 전 또는 스킬 패키지를 변경한 뒤 `../scripts/self_check_skill.py --report <self-check-report.json>`를 실행하고 `"passes": true`를 요구한다.
- 일반 페이지 범위 변환에서는 `../scripts/export_llm_extraction_packet.py --pdf <input.pdf> --pages <start-end> --output-dir <packet-dir>`를 실행하고, 렌더링된 페이지 이미지를 확인한 뒤 `<packet-dir>/llm-ir.json`을 LLM-only IR로 작성한다.
- `../scripts/build_concept_practice_deck_from_llm_ir.py --pdf <input.pdf> --ir <packet-dir>/llm-ir.json --output <output.pptx> --report-dir <report-dir>`로 빌드한다. 이 기본 build runner는 LLM이 작성한 `PracticeBlock` IR을 로드하고, `../assets/design/style-map.json`을 적용하고, 내장 OOXML style part를 주입하고, reveal animation이 있는 편집 가능한 그룹 슬라이드를 작성하고, 표준 검증 체인을 실행한다.
- LLM-IR build runner는 `generation-report.json`, `generation-report-validation.json`, `editable-report.json`, `ooxml-style-report.json`, `pattern-report.json`, `build-summary.json`을 작성한다. `build-summary.json`의 `"passes": true`를 요구한다.
- 결정론적 PDF 텍스트 추출이나 embedded page plan을 사용하지 않는다. 특정 page family에 특별 처리가 필요하면 LLM IR과 renderer/design-map 규칙으로 표현한다.
- 생성 뒤 generator report의 `quality_summary`와 `slide_trace`를 확인한다. `quality_summary.warning_count`는 자동 실패는 아니지만, 각 warning은 전달 전 생성 덱과 대조해 확인해야 한다. `quality_summary.normalizations`는 2열 원고 추출의 행 우선 정렬 같은 의도적 보정을 설명해야 한다.

## PDF 추출

1. 대량 처리 전에 페이지 범위를 시각적으로 확인한다.
2. `개념익히기`, `개념 익히기`, `개념 확인` 또는 교재별 알려진 변형 heading을 찾는다.
3. 대상 페이지를 이미지로 렌더링하고, LLM이 그 이미지에서 직접 IR을 추출하게 한다. 결정론적 PDF 텍스트 parsing으로 변환 IR을 만들지 않는다.
4. 스캔 PDF도 같은 LLM-image workflow를 사용한 뒤, 슬라이드는 편집 가능한 텍스트, 수식 텍스트, 도형으로 재구성한다.
5. 이미지 crop은 원자적 도형, 그래프, QR 유사 아이콘, 삽화에만 사용한다. 사용자가 이미지 전용 슬라이드를 명시적으로 요청하지 않는 한 전체 문제 블록 crop을 슬라이드 본문으로 쓰지 않는다.
6. 각 `개념 익히기` 블록을 레이아웃 전에 프롬프트와 하위 문항으로 나눈다.
7. 각 문제 블록의 원고 레이아웃을 기록한다. 한 열 세로 목록, 2열 그리드, 표, worked stack, 수직선, 도형 레이아웃, 혼합 레이아웃을 구분한다. 재구성할 때 행 baseline, 왼쪽 수식 열, 오른쪽 답란 열의 관계를 유지한다.
8. 인벤토리에서 문항별 원고 줄 구조를 보존한다. 하위 문항이 여러 원고 줄에 걸치면 그 줄을 문항과 함께 유지하고, 하나의 reveal 그룹 안에 여러 편집 가능 수식 행으로 렌더링한다. 등호가 여러 개 있다는 이유만으로 한 줄 reference-planned 식을 쪼개지 않는다.
   - 원고 행에 수식 표기와 한국어 연결어/라벨이 섞이면 `display_lines`만 믿지 말고 명시적인 `display_segments`를 기록한다. 예를 들어 `(3) a>b이면 / a-1/2 ○ b-1/2`는 `line_index: 0`에 math `a>b`와 `korean_label` `이면`, `line_index: 1`에 math `a-1/2 ○ b-1/2`를 둔다.
   - 원고 행이나 시각 행이 연립방정식을 보이면 `kind: "equation_system"`인 명시적 `display_segments` 항목을 기록한다. 이는 brace-stacked system과 `{x+2y=5, 3x-2y=1}` 같은 압축 텍스트에 필수다. `display_lines`만 남기면 슬라이드가 한 줄 inline 수식으로 렌더링된다.
   - 브래킷으로 묶인 판단 문항에서 첫 행에만 수식이 있으면 `kind: "spacer"` 세그먼트와 긴 `korean_label`의 `width_in`을 사용해 2~3행이 수식 열이 아니라 문장 열에 맞춰지게 한다.
9. 여러 행을 묶는 왼쪽 브래킷 연결선처럼 원고에 보이는 선화는 문항 수준 `layout_shapes`로 보존한다. `kind: "brace_connector"`, `shape: "left_square_bracket"`, `line_start`, `line_end`, `tick_lines` 행 index를 사용한다. 이를 텍스트 glyph나 원고 이미지 crop으로 인코딩하지 않는다.
10. 각 원고 문제에 실제 사각 빈칸이 있는지 표시하고, 각 빈칸의 의미 anchor를 기록한다. 예: "after equals sign", "exponent position". 보이는 모든 사각 빈칸은 문항 `raw_text`에 정확히 하나의 빈칸 토큰으로 넣는다. 기본 위치는 `[ ]`, 지수 위치는 `^[ ]`다. `blank_count`가 있으면 합산 빈칸 토큰 수와 일치해야 한다. 빈칸 안에 빨간 guide answer가 있으면 빈칸 도형은 유지하고 빨간 텍스트만 제거한다. 원고에 없는 사각 빈칸을 만들지 않는다.
11. 수직선 원고 그래픽은 `kind: "number_line"`인 `display_segments` 항목으로 geometry를 기록한다. 공식 세미콜론 spec은 `min`, `max`, `point`, `direction=left|right`, `closed=true|false`, `ticks`, `labels`, `blank=true`다. 예: `{"kind":"number_line","text":"min=-4;max=1;point=-3;direction=right;closed=false;ticks=-3;labels=-3","line_index":0}`. 학생이 표시할 빈 축은 `blank=true`를 사용하고 `point`, `direction`, `closed`를 생략한다.
12. 수직선 문항의 `raw_text`는 `(1) 수직선: 열린 점 -3, 오른쪽 영역` 또는 `(5) x≥4`처럼 의미 중심으로 유지한다. 그림을 typed arrow나 사각 빈칸 토큰으로 인코딩하지 않는다. 원고가 수직선 밖에 실제 사각 빈칸을 포함하지 않는 한 `blank_count`는 `0`으로 둔다.
13. 원고 값 표와 연립방정식 행은 편집 가능한 `display_segments`를 사용한다. 표는 `kind: "value_table"`, 보이는 모든 연립방정식은 `kind: "equation_system"`, 필요한 최종 사각 답란은 `kind: "blank_shape"`를 사용한다. `{식1, 식2}` 같은 compact IR 요약은 `raw_text`에 남아도 되지만, 표시되는 시스템은 반드시 `equation_system` 세그먼트로 분해한다. 빨간 원고 답 값은 보이는 표 텍스트가 아니라 문제 전용 hidden answer로 유지한다. 표 분수 `1/2`는 IR에서 분수 텍스트로 보존하고 `½`나 일반 slash 텍스트가 아니라 compact stacked fraction으로 렌더링한다.
14. 추출이 의심스러우면 페이지 이미지를 다시 열고 `llm-ir.json`을 수정한다. 모호함을 해결하려고 결정론적 item extraction으로 전환하지 않는다.
15. 명시적으로 요청받지 않은 `개념 다지기`, `개념 마무리`, `학교 시험 준비하기`는 제외한다.
16. 리포트에는 problem id, 원고 페이지, 추출 소스, 원고 레이아웃 유형, 원고 줄, 생성된 편집 가능 슬라이드 객체, 슬라이드 번호를 추적할 수 있는 표가 있어야 한다. 슬라이드 캔버스에 원고 페이지 번호를 보이지 않는다.
17. report slide trace에 실제 슬라이드 번호, chunk index, item number, 기본 표시 item number, click-reveal item number가 들어 있는지 확인한다.

## 설계 계약

1. 렌더러 구조를 변경하기 전에 `design-contract.md`를 읽는다.
2. `../assets/design/style-map.json`을 공통 슬라이드 좌표, 폰트, 색, 애니메이션 의도의 런타임 소스로 취급한다.
3. 런타임 스크립트나 스킬 지침에서 로컬 기대 출력 파일이나 raw reference PPTX 파일을 참조하지 않는다.
4. 반복되는 레이아웃 계열에 더 정확한 fidelity가 필요하면 측정값을 design map에 추가한다.

## 출력 품질

- 서로 관련 없는 디자인 스타일을 섞지 않는다.
- reference slide item grid를 적용하기 전에 PDF 원고 문제 본문 레이아웃을 맞춘다. 원고 문제가 한 열 세로 목록이면, 한 행에 한 문항씩 두고 원고와 비슷한 행 간격과 오른쪽 답란 열 정렬을 유지한다.
- 문제 전용 덱에서는 원고 답/예제 답을 숨긴다. 빨간 채움 값, 빨간 소수, `무한소수` 같은 빨간 답 단어는 문제 텍스트가 아니라 답으로 취급한다.
- reference 스타일이 편집 가능한 텍스트와 그룹 reveal 객체를 기대하는 경우, 전체 페이지 스크린샷이나 큰 문제 블록 crop을 사용하지 않는다.
- 명시적으로 요청받지 않은 비-`개념익히기` 문제를 포함하지 않는다.
- 사용자가 명시적으로 요청하지 않은 `개념 다지기`와 `개념 마무리` 페이지는 제외한다.
- 모든 슬라이드가 충분한 대비를 가지며 텍스트 겹침이 없는지 확인한다.
- 미디어 파일이 포함되어 있고 누락된 로컬 경로에 link되어 있지 않은지 확인한다.
- 문제 슬라이드는 생성 PNG/JPEG만이 아니라 편집 가능한 텍스트/도형 콘텐츠를 포함해야 한다.
- 슬라이드 위 reference-style reveal animation이 그룹을 대상으로 하며 누락된 `p:spTgt` id가 없는지 확인한다.
- 생성된 문항 객체가 최종 편집을 위해 그룹화되어 있는지 확인한다. 외부 item group은 reveal 대상이고, 여러 줄/브래킷 문항은 번호, 브래킷 연결선, 각 렌더링 행에 해당하는 식별 가능한 child group을 포함한다.
- 슬라이드의 사각 빈칸은 typed `□` 문자가 아니라 도형 객체여야 한다.
- 사각 빈칸 도형은 원고 PDF에 실제 사각 빈칸이 있던 곳에만 나타나야 한다.
- `[ ]` 빈칸 도형은 reference deck 스타일과 맞아야 한다. no-fill rounded rectangle, 약 `0.551in x 0.551in`, `31750` EMU theme `bg2` outline, `lumMod=75000`, round cap, round join. 쌓은 분수 분모 안의 기본 위치 빈칸은 full-size를 유지하고, 명시적 지수 빈칸 `^[ ]`만 compact해야 한다.
- `a^[ ]`, `b^[ ]`, `(b³)^[ ]` 같은 지수 위치 빈칸은 보이는 caret 없이 compact raised blank로 렌더링되는지 확인한다. `×[ ]` 같은 연산자 뒤 빈칸은 normal baseline blank로 남아야 한다.
- `--visual-refine-loop`를 사용할 때는 `visual-refinement-report.json`에 iteration PNG, diagnostics, pass/fail, adjustment 값이 기록되는지 확인한다.
- generator report에 quality warning이 없어야 한다. 있으면 페이지 이미지를 확인하고 전달 전에 `llm-ir.json`을 수정한다.
- 사각 빈칸 도형은 PDF와 같은 수식 위치에 있어야 한다. 특히 변환 연습에서 `=`, `÷`, 마지막 `=` 사이의 빈칸 위치를 확인한다.
- PDF에서 빨간 guide answer가 들어 있던 사각 빈칸은 생성된 문제 슬라이드에서 비어 있어야 한다.
- 요청 결과가 정답/풀이 슬라이드를 포함하지 않는 한 빨간 원고 답 또는 예제 답이 보이면 안 된다.
- 원고 페이지 번호가 슬라이드 캔버스에 보이면 안 된다.
- 사용자가 의도적 bleed 객체를 요청하지 않는 한 생성된 슬라이드 객체가 슬라이드 밖에 완전히 또는 부분적으로 놓이면 안 된다.
- 사용 가능한 경우 프롬프트 텍스트는 `나눔스퀘어라운드 ExtraBold`, 수식은 `BT수식M`을 사용하고, PDF가 쌓은 분수를 보이면 수평 분수선을 사용한다.
- inline 수식/한국어 혼합 행은 `math`와 `korean_label`을 분리한 `display_segments`로 표현되어야 한다. 특히 `a>b이면` 조건 행 뒤의 수식 행을 확인한다.
- 원고의 보이는 마커 도형은 `display_segments` marker 항목으로 표현되고 편집 가능한 도형으로 렌더링되어야 한다. 단원 1에서 `분모의 소인수:` 앞 마커는 typed arrow나 누락된 장식이 아니라 `bg1` fill과 `lumMod=50000`을 가진 reference-style right arrow여야 한다.
- 원고의 보이는 브래킷 연결선은 item `layout_shapes`로 표현되고 item reveal group 안의 편집 가능한 line/rectangle 도형으로 렌더링되어야 한다. typed bracket glyph나 pasted line crop으로 나타나면 안 된다.
- 의미 있는 원고 가로 간격은 IR의 `display_segments[].gap_after_in`으로 표현되고 실제 객체 간격으로 렌더링되어야 한다. 특히 마지막 `=`와 멀리 떨어진 괄호형 답란을 확인한다.
- 원고 수직선은 `kind: "number_line"` display segment로 표현되고, 축, 화살촉, 눈금, 라벨, 선택적 구간 band, 선택적 끝점, 선택적 구간 화살표가 편집 가능 도형으로 렌더링되어야 한다. 열린 끝점은 속이 빈 원, 닫힌 끝점은 채운 원이어야 한다. 빈 축은 점이나 구간 band를 보이면 안 된다.
- 원고 값 표는 스크린샷이 아니라 편집 가능한 grouped cell/text로 렌더링되고, 빨간 원고 답 셀은 문제 전용 덱에서 숨겨져야 한다. 표 분수는 수평 분수선을 가진 compact stacked fraction으로 렌더링되어야 한다. 연립방정식과 최종 사각 답란도 같은 item group 안의 편집 가능 객체로 렌더링되어야 한다.
- number-line segment spec은 원고 구간 방향과 라벨을 보존해야 한다. 예를 들어 `-3`에서 열린 점이 오른쪽으로 칠해져 있으면 `point=-3;direction=right;closed=false`, `5`에서 열린 점이 왼쪽으로 칠해져 있으면 `point=5;direction=left;closed=false`를 사용한다.
- 모든 stacked fraction은 가장 긴 분자/분모가 들어갈 만큼 넓고 no-wrap text box를 사용해야 한다. `19/10`은 `19` over `10`으로 렌더링되어야 하며 `1` over `9` over `0`처럼 세로로 깨지면 안 된다.
- stacked fraction은 reference PPTX geometry를 따라야 한다. 숫자 분수는 측정된 digit-slot 폭을 사용하고, `2²×3×5` 같은 곱셈 분모는 superscript run이 있는 하나의 denominator textbox에 남아야 하며, 분모 텍스트/빈칸은 분수선 가까이에 있어야 한다. `11/([ ]²×[ ]²)`, `1/[ ]³` 같은 지수 있는 분모 빈칸은 full-size 빈칸을 유지하고 지수를 별도 오른쪽 위 textbox에 배치해야 한다.
- 프롬프트/하위 문항 간격, 줄 간격, 문항 간격은 가독성에 영향을 주는 경우 PDF 원고 레이아웃을 따르면서 reference deck의 시각 스타일과 애니메이션 패턴을 유지해야 한다.
- 2열 worked-solution block은 raw PDF text extraction 순서가 아니라 시각적 행 우선 순서로 렌더링되고 reveal되어야 한다.
- 여러 줄 문항이 있는 빽빽한 2열 worked-solution block은 필요할 때 시각 행 기준으로 나누고, continuation slide는 행 y 위치를 중간에서 시작하지 말고 reset해야 한다.
- 여러 줄 worked-solution 문항은 모든 행을 포함한 하나의 reveal step으로 그룹화되어야 한다.
- generator report에는 missing item number, duplicate item number, unknown layout, empty prompt, missing source line, slide trace count mismatch 같은 unexplained quality warning이 없어야 한다.
- `validate_generation_report.py`가 통과해야 한다. 이는 요청 페이지가 모두 설명되고, 포함/제외 페이지 집합이 겹치지 않고, slide trace가 모든 inventory item을 덮고, chunk index가 연속이며, default/reveal item order가 내부적으로 일관됨을 증명한다.
- 최종 덱 전달 전 렌더링하거나 시각 점검한다. 문제를 수정하고 최종 수정 PPTX만 제공한다.
- reference의 화면 밖 보조/template 객체는 분석 중 무시되고 생성 슬라이드에서 제거되었는지 확인한다.
- 커밋 또는 스킬 폴더 공유 전 `self_check_skill.py`가 통과해야 한다. 이 검사는 필수 파일, JSON 유효성, Python compile 상태, asset/style-part manifest hash, 제거된 runtime dependency 부재를 검증한다.
- 최종 덱은 `.pptx`로 저장한다.
