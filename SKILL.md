---
name: pdf2concept-problem-pptx
description: 수학 개념서 PDF의 지정된 페이지에서 "개념익히기" 문제만 추출해 16:9 강의용 PPTX로 재구성한다. 문제는 편집 가능한 텍스트, 수식 텍스트, 도형, 애니메이션으로 만들고, 스킬에 내장된 기준 스타일을 따른다. "PDF 10~50쪽에서 개념익히기 문제만 뽑아서 같은 패턴의 강의용 PPTX로 만들어줘" 같은 중고등 수학 교재 변환 요청에 사용한다. 모호한 수식, 빈칸, 레이아웃을 LLM-only 방식으로 PracticeBlock IR에 추출하거나 디버깅할 때도 사용한다. 사용자가 명시적으로 이미지 전용 슬라이드를 요청하지 않는 한, 전체 문제 스크린샷이나 문제 블록 크롭 이미지를 슬라이드 본문으로 사용하지 않는다.
---

# PDF 개념 문제 PPTX 변환

이 스킬은 PDF의 개념 익히기 페이지를 편집 가능하고 애니메이션이 있는 PPTX 슬라이드로 변환한다. 이 파일은 짧은 운영 안내서로 유지하고, 세부 규칙은 필요한 경우에만 아래 reference 문서를 읽는다.

## 절대 규칙

1. LLM-only 추출 경로를 사용한다. PDF 페이지를 이미지로 렌더링하고, 이미지를 직접 확인해 `llm-ir.json`을 작성한 뒤 그 IR에서 PPTX를 빌드한다. 결정론적 PDF 텍스트 추출이나 하드코딩된 페이지 계획을 추가하지 않는다.
2. 요청된 페이지 범위에서 `개념 익히기` / `개념익히기` 블록만 포함한다. 사용자가 명시적으로 요청하지 않는 한 예제, 설명, 답안, `개념 다지기`, `개념 마무리`, `학교 시험 준비하기`는 제외한다.
3. 원고에 보이는 모든 사각형/직사각형 답란은 IR의 사각 빈칸 토큰으로 보존한다. 기본 빈칸은 `[ ]`, 지수 위치 빈칸은 `^[ ]`를 쓴다. `blank_count`가 있으면 보이는 빈칸 토큰 수와 일치해야 한다. 원고의 사각형/직사각형 답란을 절대 `(        )`로 바꾸지 않는다.
4. `(        )`는 원고가 실제로 괄호형 빈칸을 보일 때만 사용한다. 밑줄 답란과 연결선은 괄호 빈칸이 아니라 레이아웃으로 다룬다.
5. 원고에 보이는 연립방정식은 반드시 `kind: "equation_system"`인 `display_segments`로 명시한다. `{x+2y=5, 3x-2y=1}` 같은 중괄호+쉼표 텍스트를 일반 `math` 한 줄로 남기지 않는다.
6. 문제 전용 덱에서는 원고의 답과 빨간 가이드 텍스트를 숨긴다. 빨간 답이 원고 빈칸 안에 들어 있으면 빨간 텍스트만 제거하고 빈칸 도형은 유지한다.
7. 슬라이드 본문은 편집 가능한 PowerPoint 텍스트, 수식 텍스트, 도형, 그룹, 애니메이션으로 만든다. 전체 페이지나 전체 문제 블록을 슬라이드 본문 스크린샷으로 붙이지 않는다.
8. 공통 슬라이드 스타일을 적용하기 전에 원래 수학적 의미와 원고 문제 본문의 레이아웃을 보존한다. 빽빽한 내용은 글자를 줄여 겹치게 하지 말고 슬라이드를 나눈다.
9. 슬라이드에 원고 페이지 번호를 표시하지 않는다. 생성된 슬라이드에는 불필요한 화면 밖 보조 객체가 없어야 한다.

## 규칙 문서 선택

추출이나 편집 전에 관련 reference 문서를 읽는다.

- `references/conversion-checklist.md`: 실제 PDF-to-PPTX 변환 작업에 사용한다.
- `references/blank-rules.md`: 빈칸, 답란, 빨간 답이 들어간 빈칸, 밑줄 답란, `[ ]` / `^[ ]` 판단에 사용한다.
- `references/layout-rules.md`: 세로 목록, 2열 그리드, 풀이 스택, 프롬프트, 원고 줄 구조, 마커, 연결선, 곡선 화살표, 표, 수직선에 사용한다.
- `references/math-rendering.md`: 분수, 지수, 수학/한국어 혼합 행, 연립방정식, 값 표, 간격 힌트, 편집 가능한 수식 렌더링에 사용한다.
- `references/design-contract.md`: 렌더러 구조, IR 스키마, style-map 동작, OOXML 스타일 주입, 애니메이션 주입, 검증 스크립트를 변경하기 전에 읽는다.

## 작업 흐름

1. 원본 PDF, 요청 페이지 범위, 출력 PPTX 경로, 문제 전용/정답 포함/풀이 포함 여부를 확인한다.
2. LLM 추출 패킷을 내보낸다.

```bash
python scripts/export_llm_extraction_packet.py --pdf <input.pdf> --pages <start-end> --output-dir <packet-dir>
```

3. `<packet-dir>/llm-extraction-packet.json`과 `<packet-dir>/pages/` 아래의 렌더링된 페이지 이미지를 확인한다.
4. `pdf2concept-problem-pptx.llm-ir.v1` 스키마로 `<packet-dir>/llm-ir.json`을 작성한다. 대상 블록을 모두 포함하고 제외된 요청 페이지는 `excluded_pages`에 기록한다.
5. 빌드하고 검증한다.

```bash
python scripts/build_concept_practice_deck_from_llm_ir.py --pdf <input.pdf> --ir <packet-dir>/llm-ir.json --output <output.pptx> --report-dir <report-dir>
```

지수 빈칸, 괄호/브래킷 정렬, 답란 열 위치, 행 간격처럼 이미 렌더링된 객체의 기하 조정이 필요할 때만 `--visual-refine-loop --max-iterations 3`을 사용한다. 사용자가 요청하지 않는 한 표준 runner 외의 추가 검증 관문을 만들지 않는다.

## 출력 계약

- 런타임 스타일 소스는 `assets/design/style-map.json`이다.
- 생성 덱은 `assets/concept-practice-style/ooxml-style-parts`의 내장 style-only OOXML 파트를 사용해야 한다.
- 제목 슬라이드와 문제 슬라이드는 주입된 master/layout 배경과 패널에 의존해야 하며, 슬라이드 위에 중복된 전체 배경 도형을 그리지 않는다.
- 나타내기 애니메이션은 그룹 객체(`p:grpSp`)를 대상으로 해야 하며, 클릭 순서를 리포트에 보존해야 한다.
- 이미지는 편집 재구성이 어려운 원자적 도형, 그래프, 아이콘, 삽화에만 사용한다.

## 검증

공유 전 또는 스킬 패키지를 변경한 뒤에는 다음을 실행한다.

```bash
python scripts/self_check_skill.py --report <self-check-report.json>
```

비-정확복제 덱을 생성한 뒤에는 위 표준 빌드 runner를 실행하고 `<report-dir>/build-summary.json`의 `"passes": true`를 요구한다. 수동으로 검증할 때는 다음을 실행한다.

```bash
python scripts/validate_generation_report.py <report.json> --report <generation-report-validation.json>
python scripts/validate_editable_deck.py <output.pptx> --report <editable-report.json>
python scripts/validate_ooxml_style_parts.py <output.pptx> --report <ooxml-style-report.json>
python scripts/validate_reference_pattern.py <output.pptx> --require-animations --require-group-animation --forbid-visible-page-labels --forbid-off-slide-objects --report <pattern-report.json>
```

원고 문제에 실제 사각형/직사각형 빈칸이 있을 때만 `--require-blank-shapes`를 추가한다.

## 최종 전달

빌드가 통과하고, 생성된 PPTX를 충분히 시각 점검해 16:9 크기, 한국어/수식 렌더링, 겹침/잘림 없음, 페이지 라벨 없음, 누락 미디어 없음, 편집 가능한 슬라이드 본문, 기준 스타일의 그룹형 나타내기 동작을 확인한 뒤에만 전달한다.
