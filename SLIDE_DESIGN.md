# AI Contest HTML Slide Design

## 0. One-Line Concept

**문서를 자동화 자산으로 바꾸는 AI 활용법**

워드, 엑셀, PPT는 사람이 손으로 고치는 파일처럼 보이지만, 실제로는 OOXML 기반의 구조화된 데이터다. 이 사실을 이해하면 AI는 단순히 문서를 "써주는" 도구가 아니라, 반복 문서 업무를 자동화하는 파이썬 코드를 만들어주는 도구가 된다. 그리고 그 코드를 비개발자도 안전하게 실행할 수 있게 만들면, 개인의 AI 활용이 사내 업무 자산으로 확장된다.

## 1. Presentation Positioning

### Speaker Persona

발표자는 문서 작업을 많이 하는 실무자를 대변하는 사람이 아니다.

발표자의 강점은 **문서 파일을 업무 화면이 아니라 구조화된 데이터로 다시 보게 만드는 것**이다.

따라서 오프닝에서 "여러분 문서 작업 많이 힘드시죠?"라고 말하면 약하다. 대신 이렇게 시작한다.

> 저는 사실 문서 작업을 많이 하는 사람은 아닙니다.  
> 여러분은 코딩을 안 한다고 생각하셨겠지만, 사실 워드와 엑셀과 파워포인트는 여러분 대신 뒤에서 XML을 열심히 쓰고 있었습니다.  
> 우리가 문서라고 부르는 것들은 사실 코드가 읽고 쓸 수 있는 구조화된 데이터였습니다.  
> 그렇다면 질문이 바뀝니다.  
> "AI로 문서를 잘 쓰게 할 수 있을까?"가 아니라,  
> "AI로 반복되는 문서 작업 자체를 자동화할 수 있지 않을까?"

### Audience

- 개발자: "문서 자동화가 생각보다 구현 가능한 영역이구나."
- 비개발자: "내가 매번 하는 문서 작업도 자동화 후보일 수 있구나."
- 관리자/심사위원: "개인 생산성 팁이 아니라 사내 확산 가능한 시스템이구나."

### Core Promise

발표를 본 사람이 가져가야 하는 생각은 하나다.

> 반복되는 문서 작업은 수작업이 아니라 자동화 후보입니다.

## 2. Strategic Narrative

발표는 기술 소개가 아니라 인식 전환이어야 한다.

### Narrative Arc

1. 나는 문서 전문가가 아니다.
2. 그런데 워드, 엑셀, 파워포인트는 사용자 대신 뒤에서 XML을 쓰고 있었다.
3. `.docx`, `.xlsx`, `.pptx`의 `x`는 OOXML 세계로 가는 힌트다.
4. 문서가 구조화되어 있으면 코드가 읽고 쓸 수 있다.
5. 이제 AI가 그 코드를 만들어줄 수 있다.
6. 하지만 비개발자는 AI가 만든 코드를 실행하기 어렵다.
7. 그래서 "문서 자동화 스크립트 공유소 + 안전한 실행기"가 필요하다.
8. 이것이 개인의 AI 활용을 사내 업무 능력으로 확장하는 방법이다.

### What This Presentation Is Not

- OOXML 문법 강의가 아니다.
- 파이썬 라이브 코딩쇼가 아니다.
- PPT 대신 HTML이 좋다는 취향 발표가 아니다.
- "AI가 다 해줍니다" 식의 과장 발표가 아니다.

### What This Presentation Is

- 문서를 자동화 가능한 데이터로 보는 관점 제안.
- AI를 문서 작성기가 아니라 자동화 코드 생성기로 쓰는 방법 제안.
- 사내 직원이 안전하게 자동화를 공유하고 실행하는 제품 제안.

## 3. Visual Design Direction

### Overall Style

- Format: 16:9 browser-based HTML slides.
- Base: reveal.js recommended.
- Tone: clean, bright, confident, product-demo 느낌.
- Background: light background by default.
- Accent colors: blue/green for automation, amber/red only for risk/warning.
- Avoid: dark hacker aesthetic, random neon, overdone gradient, tiny code blocks.

### Typography

- Korean font: Pretendard, Noto Sans KR, or system sans-serif.
- Title: 48-64px.
- Main statement: 36-44px.
- Body: 28-34px.
- Captions: 20-24px minimum.
- Code: 22-28px, only 3-6 important lines.

### Motion Principles

Use animation only when it explains structure.

Good animation:

- `.pptx` transforms into `.zip`.
- A file opens into folders: `ppt/slides`, `ppt/media`, `ppt/theme`.
- Manual workflow collapses into automated pipeline.
- AI-generated script moves from "developer-only" to "safe runner".

Bad animation:

- Random floating objects.
- Overly decorative transitions.
- 3D effects that do not explain the concept.
- Long code typing animation.

### Three Hero Slides

These three slides must look excellent. The rest can be simple.

1. **OOXML reveal slide**
   - A `.pptx` card becomes a folder tree.
   - Purpose: "문서는 구조화된 데이터다."

2. **Automation pipeline slide**
   - Excel input -> AI-generated Python -> Word/PPT output.
   - Purpose: "문서 작업은 자동화할 수 있다."

3. **Script library + safe runner slide**
   - A product-like UI mockup.
   - Purpose: "비개발자도 안전하게 실행할 수 있게 만들었다."

## 4. Technical Build Strategy

### Recommended Stack

- HTML slide framework: reveal.js.
- Styling: custom CSS.
- Charts: Chart.js only if needed.
- Icons: lucide icons or simple inline SVG icons.
- Deployment: GitHub Pages.
- Backup: PDF export if time permits.

### Submission Strategy Under Time Pressure

Because only a few hours remain, submission should optimize for a live link.

Minimum viable deliverable:

- GitHub Pages link works.
- Title slide exists.
- Core message is visible.
- Slide navigation works.
- At least 6-8 slides communicate the story.

Tonight's improvement:

- Add animations.
- Add product mockup.
- Add charts.
- Add screenshots.
- Polish copy.
- Add backup PDF.

### Repository Structure

Use the simplest structure that GitHub Pages can serve.

```txt
/
  index.html
  README.md
  assets/
    images/
    fonts/
    screenshots/
  src/
    styles.css
    slides.css
    interactions.js
```

If using Vite or a framework, build to `dist/`, but for the fastest submission, plain static files are safer.

## 5. Slide-by-Slide Plan

### Slide 1. Title

**Title**

문서를 자동화 자산으로 바꾸는 AI 활용법

**Subtitle**

OOXML, AI-generated Python, and a safe runner for everyone

**Screen**

Large centered title. Under it, three small pills:

- OOXML
- AI Automation
- Safe Runner

**Speaker Note**

오늘 발표는 AI가 문서를 대신 써준다는 이야기가 아닙니다. 문서 작업 자체를 자동화 가능한 업무 자산으로 바꾸는 이야기입니다.

**Priority**

MVP 필수.

### Slide 2. Honest Opening

**Title**

여러분 대신 XML을 쓰고 있었습니다

**Main Copy**

여러분은 코딩을 안 한다고 생각하셨겠지만,  
사실 워드와 엑셀과 파워포인트는 여러분 대신 뒤에서 XML을 열심히 쓰고 있었습니다.

**Visual**

Left: 사용자가 보는 문서 화면: Word/Excel/PPT.  
Right: 뒤에서 만들어지는 구조: XML tags and folder tree.

**Speaker Note**

이 농담의 목적은 코딩을 한다고 우기는 것이 아닙니다. 문서 프로그램이 뒤에서 구조화된 데이터를 만들고 있다는 사실을 가볍게 보여주는 것입니다.

**Priority**

MVP 필수. 발표자의 진정성을 살리는 핵심 슬라이드.

### Slide 3. The Question Shift

**Title**

질문을 바꾸면 가능성이 보입니다

**Before**

AI가 문서를 잘 써줄 수 있을까?

**After**

AI가 반복 문서 작업을 자동화할 수 있을까?

**Visual**

Two large cards. First card fades down. Second card grows larger.

**Speaker Note**

AI 활용을 문장 생성으로만 보면 한계가 있습니다. 하지만 문서를 데이터로 보면, AI는 자동화 코드를 만들어주는 도구가 됩니다.

**Priority**

MVP 필수.

### Slide 4. The `x` in Office Files

**Title**

`.docx`, `.xlsx`, `.pptx`의 `x`

**Main Copy**

이 `x`는 단순한 버전명이 아니라, XML 기반 문서 형식이라는 힌트입니다.

**Visual**

Three file cards:

- report.docx
- sales.xlsx
- deck.pptx

The `x` character glows or gets highlighted.

**Speaker Note**

워드, 엑셀, 파워포인트 파일 뒤에 붙는 x는 OOXML 계열 문서라는 힌트입니다. 여기서 중요한 건 XML 문법을 외우는 것이 아니라, 문서가 구조화된 데이터라는 사실입니다.

**Priority**

MVP 필수.

### Slide 5. OOXML Reveal

**Title**

문서 파일은 사실 패키지입니다

**Main Copy**

하나의 파일처럼 보이지만, 안에는 XML, 이미지, 스타일, 관계 정보가 들어 있습니다.

**Visual**

Animated sequence:

1. `presentation.pptx`
2. arrow: rename/open as zip
3. folder tree:

```txt
ppt/
  slides/
    slide1.xml
    slide2.xml
  media/
    image1.png
    image2.jpg
  theme/
  _rels/
```

**Speaker Note**

이 장면은 길게 설명하지 않습니다. 딱 한 번만 보여주면 됩니다. "아, PPT 안에 이런 구조가 있구나"라는 인식 전환이 목적입니다.

**Priority**

Hero 필수. 가장 예쁘게 만들 것.

### Slide 6. Why AI Can Modify Documents

**Title**

왜 AI가 문서를 고칠 수 있을까?

**Bullets**

- XML은 태그와 계층 구조로 의미를 표시한다.
- 그래서 코드가 "제목", "표", "슬라이드", "이미지" 같은 위치를 찾을 수 있다.
- LLM은 이 구조를 이해해 조작 방법을 설명하거나 스크립트를 만들 수 있다.
- Codex나 Claude 같은 AI 에이전트는 보통 Python/JavaScript 스크립트나 문서 라이브러리를 실행해 파일을 수정한다.
- 즉, AI가 마우스로 문서를 고치는 것이 아니라 문서 구조를 코드로 바꾸는 것이다.

**Visual**

Three-layer diagram:

```txt
OOXML structure
  "여기에 제목이 있고, 여기에 표가 있고, 여기에 이미지가 있다"
        ↓
LLM understanding
  "이 구조를 어떻게 바꿀지 코드로 설명할 수 있다"
        ↓
Agent/script execution
  Python/JavaScript or document libraries modify the file
```

**Speaker Note**

XML이 자동화에 유리한 이유는 구조가 명시되어 있기 때문입니다. 예를 들어 슬라이드의 텍스트, 워드의 문단, 엑셀의 셀은 각각 특정 XML 위치나 라이브러리 API로 접근할 수 있습니다. LLM은 그 구조를 텍스트로 이해하고, 실제 수정은 Python 스크립트나 `python-docx`, `openpyxl`, `python-pptx` 같은 라이브러리를 통해 수행합니다. 그래서 Codex나 Claude 같은 AI 에이전트가 문서를 수정할 때도 본질적으로는 "문서를 열고 클릭"하는 게 아니라, 구조화된 파일을 읽고 쓰는 코드를 실행하는 방식에 가깝습니다.

**Priority**

MVP 필수.

### Slide 7. Past vs Now

**Title**

과거에는 전문가의 일이었습니다

**Before**

OOXML 구조 이해  
라이브러리 학습  
파이썬 코드 작성  
오류 디버깅

**Now**

AI에게 자동화 코드를 요청  
샘플 파일로 검증  
반복 업무에 재사용

**Visual**

Timeline or two-column comparison.

**Speaker Note**

예전에도 가능했습니다. 다만 비용이 높았습니다. AI가 바꾼 것은 가능 여부가 아니라 시도 비용입니다.

**Priority**

MVP 필수.

### Slide 8. Prompt That Changes the Game

**Title**

AI에게 이렇게 요청할 수 있습니다

**Prompt Card**

```txt
같은 양식의 xlsx 파일들이 여러 개 있습니다.
각 파일의 "Summary" 시트에서 총액, 부서명, 작성일을 읽고
하나의 CSV로 합치는 Python 스크립트를 만들어주세요.
openpyxl을 사용하고, 오류 파일은 로그로 남겨주세요.
```

**Visual**

Prompt card on left, generated Python script preview on right.

**Speaker Note**

막연히 "엑셀 처리해줘"라고 하는 것보다, 파일 구조와 원하는 입력/출력을 말하면 AI가 훨씬 좋은 자동화 코드를 만들어냅니다.

**Priority**

MVP 필수.

### Slide 9. Manual Work Becomes a Pipeline

**Title**

반복 문서 작업은 파이프라인이 됩니다

**Pipeline**

Input files -> AI-generated script -> validation -> generated report -> human review

**Visual**

Horizontal pipeline with icons:

- folder upload
- Python
- check mark
- Word/PPT output
- human approval

**Speaker Note**

중요한 건 사람이 사라지는 게 아닙니다. 사람이 매번 복붙하던 일을 코드가 처리하고, 사람은 검토와 판단에 집중하게 되는 것입니다.

**Priority**

Hero 필수.

### Slide 10. Example Use Cases

**Title**

업무에서 바로 떠올릴 수 있는 예시

**Cards**

1. 엑셀 취합
   - 여러 부서 파일을 하나의 CSV로 병합

2. 워드 보고서 생성
   - 표준 양식에 값만 채워 월간 보고서 생성

3. PPT 업데이트
   - 최신 지표로 차트와 수치를 자동 갱신

4. 문서 검증
   - 필수 항목 누락, 숫자 불일치, 양식 오류 탐지

**Speaker Note**

여기서 중요한 것은 직무별 상상력입니다. 각자 자기 업무에서 "매번 비슷하게 반복되는 문서 작업"을 떠올리면 됩니다.

**Priority**

MVP 가능. 시간이 없으면 4개 카드만.

### Slide 11. HTML Slides as Proof

**Title**

그리고 이 발표자료도 문서입니다

**Main Copy**

PPT 대신 HTML/CSS/JavaScript로 만들면, 발표자료도 코드가 되고 데이터와 상호작용할 수 있습니다.

**Visual**

Current slide interface lightly reveals browser frame. Show:

- animation
- chart
- interaction
- deployment link

**Speaker Note**

PPT가 나쁘다는 뜻은 아닙니다. 다만 발표자료의 목적이 데이터를 설득력 있게 보여주는 것이라면, HTML은 더 자유로운 선택지가 될 수 있습니다.

**Priority**

중요하지만 메인 주인공은 아님. 1-2분 이내.

### Slide 12. HTML Presentation Advantages

**Title**

HTML 발표자료가 유리한 순간

**Bullets**

- 데이터 차트와 애니메이션이 많을 때
- 인터랙티브 데모가 필요할 때
- 웹으로 바로 공유하고 싶을 때
- GitHub Pages로 배포하고 싶을 때
- 발표 후에도 살아있는 자료로 남기고 싶을 때

**Visual**

Browser window mockup with URL bar:

`https://username.github.io/ai-contest/`

**Speaker Note**

이 발표자료 자체가 예시입니다. 발표자료가 파일 첨부물이 아니라 링크로 공유되는 지식 자산이 됩니다.

**Priority**

MVP 선택. 시간이 없으면 Slide 11에 합쳐도 됨.

### Slide 13. The Execution Gap

**Title**

하지만 여기서 막힙니다

**Main Copy**

AI가 코드를 만들어줘도, 모두가 그 코드를 실행할 수 있는 것은 아닙니다.

**Pain Points**

- Python 설치가 어렵다.
- 패키지 설치에서 막힌다.
- 파일 경로를 모른다.
- 오류 메시지를 해석하기 어렵다.
- 이 코드가 안전한지 모른다.

**Visual**

AI-generated code card on left.  
Confused user / blocked state on right.

**Speaker Note**

이 지점이 중요합니다. AI가 코드를 만들어준다는 것만으로는 사내 전체의 업무 능력이 올라가지 않습니다. 실행 장벽을 해결해야 합니다.

**Priority**

MVP 필수. 솔루션으로 넘어가는 다리.

### Slide 14. Product Reveal

**Title**

문서 자동화 스크립트 공유소 + 안전한 실행기

**Main Copy**

승인된 자동화 스크립트를 등록하고, 누구나 파일을 넣어 안전하게 실행할 수 있는 사내 도구.

**Visual**

Product UI mockup:

Left sidebar:

- 자동화 목록
- 즐겨찾기
- 실행 기록
- 승인 대기

Main panel:

- "월간 보고서 생성기"
- description
- input requirements
- upload area
- Run button
- Preview result

**Speaker Note**

제가 만들고 싶은 것은 단순히 코드 모음이 아닙니다. 개발자나 파워유저가 만든 자동화를 사내 직원들이 안전하게 실행할 수 있는 환경입니다.

**Priority**

Hero 필수. 가장 제품처럼 보여야 함.

### Slide 15. How It Works

**Title**

작동 방식

**Flow**

1. 개발자/파워유저가 스크립트를 등록한다.
2. 관리자가 입력 형식과 권한을 검토한다.
3. 사용자는 파일을 업로드한다.
4. 안전 실행기가 격리 환경에서 실행한다.
5. 결과물과 로그를 제공한다.

**Visual**

Swimlane diagram:

- Creator
- Reviewer
- User
- Runner

**Speaker Note**

사내 도구가 되려면 "돌아간다"보다 "안전하게 돌아간다"가 중요합니다. 그래서 등록, 검토, 실행, 로그가 하나의 흐름으로 묶여야 합니다.

**Priority**

MVP 필수.

### Slide 16. Safety Model

**Title**

AI 자동화에는 안전장치가 필요합니다

**Safety Features**

- 승인된 스크립트만 실행
- 입력 파일 타입 제한
- 네트워크 접근 제한
- 실행 시간 제한
- 결과 미리보기
- 실행 로그 기록
- 민감정보 업로드 경고

**Visual**

Shield icon in center, safety features orbiting around it.

**Speaker Note**

AI와 자동화는 편하지만, 사내에서는 보안과 신뢰가 핵심입니다. 그래서 자유 실행기가 아니라 안전 실행기여야 합니다.

**Priority**

MVP 필수. 사내 발표 신뢰도를 높임.

### Slide 17. What Changes for Employees

**Title**

직원의 AI 활용 방식이 바뀝니다

**Before**

AI에게 문서 초안을 부탁한다.

**After**

AI로 반복 업무 자동화 코드를 만들고, 팀이 재사용한다.

**Visual**

Maturity ladder:

1. Ask
2. Draft
3. Automate
4. Share
5. Govern

**Speaker Note**

AI 활용 능력의 다음 단계는 프롬프트를 더 예쁘게 쓰는 것이 아닙니다. 반복되는 일을 자동화하고, 그 자동화를 팀 자산으로 만드는 것입니다.

**Priority**

MVP 필수.

### Slide 18. Demo Plan

**Title**

데모: 파일을 넣으면 결과물이 나온다

**Demo Scenario**

Sample input:

- `sales_jan.xlsx`
- `sales_feb.xlsx`
- `sales_mar.xlsx`

Automation:

- 매출 데이터 취합
- 요약 표 생성
- 보고서 문구 생성
- PPT/HTML 요약 슬라이드 생성

Output:

- `summary.csv`
- `monthly_report.docx`
- `executive_summary.pptx` or HTML section

**Visual**

Interactive fake runner is enough if real software is unfinished:

- upload card
- progress bar
- result list
- download buttons

**Speaker Note**

실제 백엔드가 완성되지 않아도 발표용으로는 작동 흐름을 보여주는 시뮬레이션이 효과적입니다. 핵심은 "어떤 경험을 제공할 것인가"입니다.

**Priority**

MVP 선택. 소프트웨어가 미완성이면 mock demo로 대체.

### Slide 19. Why This Matters to the Company

**Title**

개인의 자동화를 회사의 자산으로

**Main Points**

- 반복 문서 업무 감소
- 부서별 자동화 재사용
- 승인된 방식으로 안전하게 실행
- 업무 지식이 스크립트로 축적
- AI 활용 격차 완화

**Visual**

Small automations from individuals flow into shared repository, then spread to teams.

**Speaker Note**

한 사람이 만든 자동화가 자기 PC 안에서 끝나면 개인 생산성입니다. 공유되고 안전하게 실행되면 조직 생산성이 됩니다.

**Priority**

MVP 필수.

### Slide 20. Closing

**Title**

문서는 더 이상 최종 결과물만이 아닙니다

**Main Copy**

문서는 데이터이고,  
AI는 자동화 코드를 만들 수 있고,  
안전한 실행기는 그 자동화를 모두의 업무 능력으로 바꿉니다.

**Final Line**

반복되는 문서 작업은 이제 자동화 후보입니다.

**Visual**

Three-layer stack:

- OOXML: document as data
- AI: code generation
- Safe Runner: shared execution

**Speaker Note**

오늘 이후로 문서 작업을 볼 때, "이걸 더 잘 작성할 방법"뿐 아니라 "이걸 반복하지 않을 방법"도 같이 떠올리면 좋겠습니다.

**Priority**

MVP 필수.

## 6. MVP Slide Set for 3-Hour Deadline

If only a minimal version can be shipped before deadline, build these 9 slides first.

1. Title
2. Honest Opening
3. The Question Shift
4. The `x` in Office Files
5. OOXML Reveal
6. Past vs Now
7. Execution Gap
8. Product Reveal
9. Closing

This version is enough to submit a GitHub Pages link that communicates the idea.

## 7. Tonight Polish Plan

After the link is submitted, improve in this order.

### Phase 1. Make It Not Break

- Ensure GitHub Pages URL opens.
- Remove external dependencies if blocked by company network.
- Test in Chrome, Edge, and phone browser.
- Make keyboard navigation obvious.

### Phase 2. Make the Three Hero Slides Excellent

- OOXML reveal animation.
- Automation pipeline animation.
- Product UI mockup.

### Phase 3. Add Credibility

- Add fake but realistic sample data.
- Add screenshot or mock demo of safe runner.
- Add safety model slide.
- Add 1-2 concrete use cases.

### Phase 4. Add Delight

- Subtle transitions.
- Interactive chart.
- Download/result simulation.
- Progress animation.

### Phase 5. Backup

- Export PDF.
- Add README.
- Add fallback `index.html` note.

## 8. Copy Bank

Use these lines across the deck.

### Strong Opening Lines

- 여러분은 코딩을 안 한다고 생각하셨겠지만, 사실 워드와 엑셀과 파워포인트는 여러분 대신 뒤에서 XML을 열심히 쓰고 있었습니다.
- 문서 프로그램이 대신 XML을 쓰고 있었다면, 문서 파일은 사람이 보는 화면인 동시에 코드가 다룰 수 있는 데이터입니다.
- 우리가 문서라고 부르는 것들은 사실 코드가 읽고 쓸 수 있는 데이터였습니다.
- AI 활용의 질문을 바꿔보겠습니다. "문서를 잘 쓰게 할까?"가 아니라 "문서 작업을 반복하지 않게 할 수 있을까?"

### OOXML Lines

- `.docx`, `.xlsx`, `.pptx`는 하나의 파일처럼 보이지만 내부에는 XML, 이미지, 스타일, 관계 정보가 들어 있습니다.
- XML은 태그와 계층 구조로 의미를 표시하기 때문에 코드가 원하는 위치를 찾아 읽고 쓸 수 있습니다.
- LLM은 XML 구조를 텍스트로 이해하고, 실제 수정은 AI 에이전트가 생성한 Python/JavaScript 스크립트나 문서 라이브러리가 수행합니다.
- Codex나 Claude 같은 AI 에이전트가 문서를 수정할 때도 본질적으로는 문서를 클릭하는 것이 아니라 구조화된 파일을 읽고 쓰는 코드를 실행하는 방식에 가깝습니다.
- 문서가 구조화되어 있으면 코드가 읽고 쓸 수 있고, 이제 AI가 그 코드를 만들어줄 수 있습니다.

### Automation Lines

- 반복되는 문서 작업은 수작업이 아니라 자동화 후보입니다.
- AI가 바꾼 것은 가능 여부가 아니라 시도 비용입니다.
- 사람이 사라지는 것이 아니라, 사람이 복붙 대신 검토와 판단에 집중하게 됩니다.

### Safe Runner Lines

- AI가 코드를 만들어줘도 모두가 코드를 실행할 수 있는 것은 아닙니다.
- 자동화가 사내에 퍼지려면 실행 장벽과 보안 장벽을 같이 해결해야 합니다.
- 개인의 스크립트가 공유소와 안전 실행기를 만나면 조직의 업무 자산이 됩니다.

### Closing Lines

- 문서는 더 이상 최종 결과물만이 아닙니다. 자동화 가능한 데이터입니다.
- AI 활용 능력의 다음 단계는 프롬프트를 잘 쓰는 것이 아니라, 반복 업무를 자동화 자산으로 바꾸는 것입니다.
- 오늘 이후로 문서 파일을 보면 한 가지 질문을 떠올려주세요. "이 작업, 반복하지 않을 수 있지 않을까?"

## 9. Risk Control

### Do Not Emphasize

- 엑셀 시트 비밀번호 해제.
- 문서에 실행 파일 숨기기.
- 보안 유출을 가볍게 농담하는 표현.
- "상남자", "몰래", "강제 해제" 같은 자극적인 표현.
- XML 문법 상세 강의.

### Safer Alternatives

Instead of:

> 시트 암호를 강제로 해제할 수 있습니다.

Say:

> 파일 구조를 이해하면 문서 복구, 검증, 변환 같은 업무 자동화가 가능해집니다.

Instead of:

> 이상한 파일도 숨길 수 있습니다.

Say:

> 문서 파일 안에는 생각보다 다양한 리소스와 메타데이터가 들어갈 수 있기 때문에 보안 검토가 중요합니다.

Instead of:

> 외부 변환 사이트 쓰지 마세요.

Say:

> 사내 문서는 외부 변환 사이트보다 내부 자동화 도구로 처리하는 것이 보안상 안전합니다.

## 10. Demo Design

### Best Demo If Software Is Ready

1. Open safe runner.
2. Select "월간 보고서 생성기".
3. Upload sample Excel.
4. Click Run.
5. Show progress.
6. Download generated report.
7. Show execution log.

### Best Demo If Software Is Not Ready

Use a simulated UI inside the slide.

The simulation can still be persuasive if it shows:

- what the user uploads,
- what script runs,
- what safety checks happen,
- what output is generated.

Label it honestly as "컨셉 데모" or "프로토타입 화면" if needed.

### Demo Script

> 예를 들어 매달 같은 양식의 엑셀 파일을 취합해서 보고서를 만든다고 해보겠습니다.  
> 사용자는 스크립트나 파이썬을 몰라도 됩니다.  
> 승인된 자동화를 선택하고, 파일을 넣고, 실행하면 됩니다.  
> 실행기는 입력 파일을 검사하고, 제한된 환경에서 스크립트를 실행하고, 결과물과 로그를 남깁니다.

## 11. HTML Slide Interaction Ideas

### Easy Interactions

- Click to expand `.pptx` into folder tree.
- Click "Run" in fake runner to animate progress bar.
- Click tabs: Before / After.
- Toggle "Developer View" and "Employee View".

### Medium Interactions

- Chart showing manual hours vs automated hours.
- File cards flowing into output documents.
- Script approval status changing from Draft -> Reviewed -> Approved.

### Hard Interactions

- Real file upload parsing.
- Real Python backend.
- 3D visualization.

Avoid hard interactions before submission unless already implemented.

## 12. Visual Components

### File Card

Use for `.docx`, `.xlsx`, `.pptx`, `.zip`, `.py`, `.csv`.

Properties:

- icon
- file name
- short description
- status badge

### Folder Tree

Use monospace text with highlighted important files.

Example:

```txt
ppt/
  slides/
    slide1.xml   <- slide content
  media/
    image1.png   <- original image
  theme/
  _rels/
```

### Pipeline

Use horizontal steps:

```txt
Input Files -> AI Prompt -> Python Script -> Safe Runner -> Output Docs
```

### Product Mockup

Must include:

- script list
- script details
- upload area
- run button
- safety checklist
- result area

## 13. Suggested 5-Minute Talk Track

### 0:00-0:30

여러분은 코딩을 안 한다고 생각하셨겠지만, 사실 워드와 엑셀과 파워포인트는 여러분 대신 뒤에서 XML을 열심히 쓰고 있었습니다. 오늘은 이 가벼운 농담에서 출발해서, 문서 작업을 자동화 자산으로 바꾸는 방법을 이야기해보겠습니다.

### 0:30-1:20

워드, 엑셀, PPT는 하나의 파일처럼 보이지만 사실 OOXML 기반의 구조화된 데이터입니다. XML은 태그와 계층 구조로 의미를 표시하기 때문에, 코드가 원하는 문단, 표, 셀, 슬라이드를 찾아 읽고 쓸 수 있습니다.

### 1:20-2:10

이 사실을 알면 질문이 바뀝니다. AI로 문서를 잘 쓰게 할 수 있을까가 아니라, AI로 반복 문서 작업을 자동화할 수 있을까가 됩니다.

### 2:10-3:00

AI에게 구체적인 입력과 출력을 말하면 엑셀 취합, 보고서 생성, PPT 업데이트 같은 자동화 코드를 만들 수 있습니다.

### 3:00-4:10

하지만 AI가 코드를 만들어줘도 모든 직원이 실행할 수 있는 것은 아닙니다. 그래서 문서 자동화 스크립트 공유소와 안전한 실행기가 필요합니다.

### 4:10-5:00

승인된 자동화를 선택하고, 파일을 업로드하고, 결과물을 받는 방식으로 만들면 개인의 AI 활용이 사내 업무 자산이 됩니다.

## 14. Suggested 10-Minute Talk Track

### 0:00-1:00 Opening

- Honest opening.
- Question shift.

### 1:00-2:30 OOXML

- `x` in file extensions.
- package/folder reveal.
- document as data.

### 2:30-4:00 AI Automation

- prompt example.
- generated Python concept.
- pipeline.

### 4:00-5:00 HTML Slides

- this deck is also code.
- interactive documents.

### 5:00-7:00 Execution Gap

- AI code is not enough.
- non-developer execution barrier.
- developer reuse problem.

### 7:00-9:00 Safe Runner

- product mockup.
- workflow.
- safety model.

### 9:00-10:00 Closing

- employee impact.
- company impact.
- final message.

## 15. README Submission Copy

Use this in `README.md`.

```md
# AI Contest Presentation

## Title
문서를 자동화 자산으로 바꾸는 AI 활용법

## Live Presentation
https://<username>.github.io/<repo-name>/

## Summary
이 발표는 OOXML 기반 문서 파일(.docx, .xlsx, .pptx)을 구조화된 데이터로 이해하고,
AI가 생성한 Python 자동화 코드와 안전한 실행기를 통해 반복 문서 작업을 사내 업무 자산으로 전환하는 방법을 제안합니다.

## Recommended Viewing
- Chrome or Edge
- 16:9 display
- Fullscreen mode

## Backup
If the live page is unavailable, open `index.html` locally or refer to the exported PDF.
```

## 16. Final Judgement

The winning angle is not:

> OOXML is interesting.

The winning angle is:

> OOXML helps people realize documents are automatable, AI lowers the cost of creating automation, and the safe runner makes that automation usable by everyone in the company.

If time is short, protect this story above everything else.

## 17. Slide Media Plan

Each slide should have one visual job. Do not add media just to make the page look full. Use images, diagrams, UI mockups, motion, and short simulations only when they make the idea easier to understand.

### Slide 1. Title

**Best Media**

- Clean abstract background made from document/file cards.
- Subtle floating file icons: `.docx`, `.xlsx`, `.pptx`, `.py`, `.html`.

**Why**

The audience should immediately understand that this is about documents becoming automation assets.

**Avoid**

- Generic AI brain images.
- Robot hand / glowing circuit stock imagery.

### Slide 2. Honest Opening

**Best Media**

- Split-screen visual:
  - left: normal office document screen,
  - right: XML tag snippets and folder tree behind it.
- Light animation where Word/Excel/PPT icons briefly reveal XML tags behind them.

**Why**

This makes the joke land: "you were not coding directly, but the apps were writing structured data for you."

### Slide 3. The Question Shift

**Best Media**

- Two big cards:
  - "AI가 문서를 잘 써줄까?"
  - "AI가 문서 작업을 반복하지 않게 할 수 있을까?"
- Animate the first card shrinking and the second card becoming the main focus.

**Why**

The audience needs to feel the frame shift, not just read it.

### Slide 4. The `x` in Office Files

**Best Media**

- Three oversized file cards:
  - `report.docx`
  - `sales.xlsx`
  - `deck.pptx`
- Highlight only the final `x`.

**Why**

People remember one small visual hook. The `x` is that hook.

### Slide 5. OOXML Reveal

**Best Media**

- Animated unzip sequence:
  - `deck.pptx`
  - `deck.zip`
  - folder tree
  - `ppt/slides/slide1.xml`
  - `ppt/media/image1.png`
- If time allows, include one real screenshot of an unzipped sample PPT folder.

**Why**

This is the proof moment. People need to see that a PPT is not a mysterious binary blob.

**Avoid**

- Long XML scrolls.
- Too many folders at once.

### Slide 6. Why AI Can Modify Documents

**Best Media**

- Three-layer diagram:
  - OOXML structure
  - LLM understanding
  - script/library execution
- Small code preview:

```python
from pptx import Presentation
prs = Presentation("deck.pptx")
prs.slides[0].shapes.title.text = "Updated by AI"
prs.save("updated.pptx")
```

**Why**

This makes the mechanism concrete: AI does not click PowerPoint. It creates or runs code that edits the document structure.

### Slide 7. Past vs Now

**Best Media**

- Before/after timeline.
- Before side: "OOXML docs, library docs, debugging".
- Now side: "prompt, generated script, sample validation".

**Why**

The point is not that automation became newly possible. The point is that the cost of trying became much lower.

### Slide 8. Prompt That Changes the Game

**Best Media**

- Chat-style prompt bubble on the left.
- Generated script preview on the right.
- Highlight input, output, library, error handling with different colors.

**Why**

People need to learn what a useful automation prompt looks like.

### Slide 9. Manual Work Becomes a Pipeline

**Best Media**

- Animated pipeline:
  - folder of input files,
  - AI prompt,
  - Python script,
  - validation,
  - generated document,
  - human review.
- File cards should move through the pipeline.

**Why**

This converts "AI writes code" into "my workflow can become a repeatable process."

### Slide 10. Example Use Cases

**Best Media**

- Four compact cards with icons:
  - merge spreadsheets,
  - generate reports,
  - update slides,
  - validate documents.
- Optional: tiny before/after thumbnail inside each card.

**Why**

Different job functions need different entry points. This slide should maximize "내 업무에도 있겠는데?" reactions.

### Slide 11. HTML Slides as Proof

**Best Media**

- Show the current slide inside a browser frame.
- Add small labels:
  - HTML
  - CSS
  - JavaScript
  - GitHub Pages

**Why**

The presentation itself becomes evidence for the message: documents can be code-backed interactive assets.

### Slide 12. HTML Presentation Advantages

**Best Media**

- Mini carousel of capabilities:
  - animated chart,
  - interactive filter,
  - 3D object placeholder,
  - shareable URL.
- Use very short looping animations if possible.

**Why**

HTML slides are easier to understand by seeing the extra interaction, not by listing features.

### Slide 13. The Execution Gap

**Best Media**

- A generated Python code block on the left.
- On the right, five blockers as warning chips:
  - Python install
  - package error
  - file path
  - permission
  - security
- Optional: progress stops at 20% with "ModuleNotFoundError".

**Why**

This makes the non-developer execution barrier obvious without mocking the user.

### Slide 14. Product Reveal

**Best Media**

- High-fidelity product mockup of the safe runner.
- Must include:
  - automation list,
  - selected script detail,
  - upload zone,
  - safety checklist,
  - run button,
  - result area.

**Why**

This is the main original contribution. It should feel like a real product, not just a concept.

### Slide 15. How It Works

**Best Media**

- Swimlane diagram:
  - creator,
  - reviewer,
  - employee,
  - safe runner.
- Use arrows to show script registration, approval, execution, logging.

**Why**

This proves the solution has an operating model, not just a button.

### Slide 16. Safety Model

**Best Media**

- Shield diagram or checklist panel.
- Show safety controls as concrete toggles:
  - approved script only,
  - allowed file types,
  - no network,
  - time limit,
  - logs,
  - preview.

**Why**

For an internal AI contest, safety is credibility.

### Slide 17. What Changes for Employees

**Best Media**

- Maturity ladder:
  - ask,
  - draft,
  - automate,
  - share,
  - govern.
- Each step gets a small icon and one short phrase.

**Why**

This shows that the proposal is about raising the AI maturity of the organization.

### Slide 18. Demo Plan

**Best Media**

- Clickable prototype simulation:
  - upload files,
  - click run,
  - progress bar,
  - generated results.
- If real software is incomplete, use a clearly polished concept demo.

**Why**

People believe workflows when they see state changes.

### Slide 19. Why This Matters to the Company

**Best Media**

- Network effect diagram:
  - individual scripts flow into shared repository,
  - teams reuse approved automations,
  - logs and governance wrap the system.
- Optional chart: repeated manual hours decreasing over time.

**Why**

The contest value is organizational leverage, not a single cool script.

### Slide 20. Closing

**Best Media**

- Three-layer final stack:
  - OOXML: document as data,
  - AI: code generation,
  - Safe Runner: shared execution.
- Final phrase appears last:
  - "반복되는 문서 작업은 이제 자동화 후보입니다."

**Why**

The audience should leave with one repeatable sentence.

### Asset Priority Under Time Pressure

If there is not enough time, make only these visuals excellent:

1. OOXML unzip/folder reveal.
2. AI modifies documents through script/library diagram.
3. Manual work to automation pipeline.
4. Safe runner product mockup.
5. Clickable demo simulation.

Everything else can be clean typography and simple icons.
