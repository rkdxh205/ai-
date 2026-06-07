import streamlit as st
from openai import OpenAI, AuthenticationError, RateLimitError
import os
from datetime import datetime
from dotenv import load_dotenv
from docx import Document
from fpdf import FPDF
import re
import io
import base64

load_dotenv()

st.set_page_config(
    page_title="AI 해커톤 심사 시스템",
    page_icon="🏆",
    layout="wide"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@400;500;600;700;800;900&display=swap');

    html, body, [class*="css"] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, 'Apple SD Gothic Neo', 'Noto Sans KR', sans-serif;
    }
    /* 로고 컬러: 네이비 #0D2E52 · 파랑 #1276B8 · 하늘 #45AEDD */
    .stApp { background-color: #EEF4FA; }
    [data-testid="stSidebar"] { display: none; }
    [data-testid="collapsedControl"] { display: none; }

    /* 히어로 */
    .hero {
        background: white;
        border-radius: 20px;
        padding: 44px 40px 40px;
        text-align: center;
        margin-bottom: 24px;
        border: 1px solid #D8E8F4;
    }
    .hero-logo-wrap {
        margin-bottom: 20px;
    }
    .hero-logo-wrap img {
        height: 72px;
        object-fit: contain;
    }
    .hero-divider {
        width: 40px; height: 2px;
        background: linear-gradient(90deg, #45AEDD, #1276B8);
        margin: 0 auto 22px;
        border-radius: 2px;
    }
    .hero-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #EEF4FA;
        color: #1276B8;
        border-radius: 100px;
        padding: 5px 16px;
        font-size: 0.73rem;
        font-weight: 700;
        letter-spacing: 0.06em;
        text-transform: uppercase;
        margin-bottom: 18px;
        border: 1px solid #C5DEF0;
    }
    .hero h1 {
        margin: 0 0 10px;
        font-size: 2rem;
        font-weight: 900;
        color: #0D2E52;
        letter-spacing: -0.04em;
        line-height: 1.25;
    }
    .hero p {
        margin: 0;
        font-size: 0.88rem;
        color: #7A9AB8;
        letter-spacing: -0.01em;
    }
    .hero-stats {
        display: flex;
        justify-content: center;
        gap: 40px;
        margin-top: 28px;
        padding-top: 24px;
        border-top: 1px solid #C8DEF0;
    }
    .hero-stat { text-align: center; }
    .hero-stat .stat-val {
        font-size: 1.6rem;
        font-weight: 900;
        color: #1276B8;
        letter-spacing: -0.04em;
        line-height: 1;
    }
    .hero-stat .stat-lbl {
        font-size: 0.71rem;
        color: #9DBAD4;
        margin-top: 4px;
        font-weight: 600;
        letter-spacing: 0.02em;
    }

    /* 채점 기준 그리드 */
    .criteria-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 12px;
        margin-bottom: 24px;
    }
    .criteria-card {
        background: white;
        border-radius: 18px;
        padding: 28px 30px;
        border: 1px solid #D8E8F4;
        display: flex;
        flex-direction: column;
    }
    .c-title-row {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 18px;
    }
    .criteria-card .c-title {
        font-size: 1.25rem;
        font-weight: 800;
        color: #0D2E52;
        letter-spacing: -0.02em;
        line-height: 1.3;
        margin: 0;
    }
    .criteria-card .pts-badge {
        display: inline-block;
        background: #EEF4FA;
        color: #1276B8;
        border-radius: 100px;
        padding: 4px 16px;
        font-size: 0.9rem;
        font-weight: 700;
        border: 1px solid #C5DEF0;
        white-space: nowrap;
        flex-shrink: 0;
    }
    .sub-scores {
        border-top: 1px solid #C8DEF0;
        padding-top: 12px;
        margin-bottom: 14px;
        flex: 0 0 96px;
        overflow: hidden;
    }
    .sub-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 6px 0;
        font-size: 0.97rem;
    }
    .sub-row .sub-label { color: #7A9AB8; }
    .sub-row .sub-pts {
        font-weight: 700;
        color: #0D2E52;
        font-size: 0.95rem;
    }
    .warn-box {
        background: #FFFBF5;
        border: 1px solid #FFE4B5;
        border-radius: 10px;
        padding: 12px 16px;
        margin-top: 12px;
        flex: 1;
        box-sizing: border-box;
    }
    /* 창의성 전폭 카드는 고정 해제 */
    .criteria-card:last-child .sub-scores {
        flex: 0 0 auto;
        overflow: visible;
    }
    .criteria-card:last-child .warn-box {
        flex: 0 0 auto;
    }
    .warn-box .warn-title {
        font-size: 0.92rem;
        font-weight: 700;
        color: #B45309;
        margin-bottom: 8px;
    }
    .warn-box ul { margin: 0; padding-left: 16px; list-style: disc; }
    .warn-box ul li { font-size: 0.9rem; color: #92400E; line-height: 1.7; }

    /* 업로드 카드 박스 */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: linear-gradient(180deg, #F2F8FE 0%, #FFFFFF 35%) !important;
        border-radius: 20px !important;
        border: 1.5px solid #C8DEF0 !important;
        box-shadow: 0 4px 24px rgba(13,46,82,0.09) !important;
        margin-bottom: 0 !important;
        transition: box-shadow 0.2s !important;
        height: 100% !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:hover {
        box-shadow: 0 8px 32px rgba(13,46,82,0.14) !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"] > div {
        padding: 24px 20px 20px !important;
        height: 100% !important;
        box-sizing: border-box !important;
    }

    /* 업로드 카드 헤더 */
    .upload-header {
        text-align: center;
        padding-bottom: 16px;
        border-bottom: 1px solid #C8DEF0;
        margin-bottom: 16px;
    }
    .upload-icon-wrap {
        font-size: 2.2rem;
        line-height: 1;
        margin-bottom: 10px;
    }
    .upload-title {
        font-size: 1.05rem;
        font-weight: 800;
        color: #0D2E52;
        letter-spacing: -0.03em;
        margin-bottom: 6px;
    }
    .upload-desc {
        font-size: 0.72rem;
        color: #9DBAD4;
        font-weight: 500;
        line-height: 1.6;
    }

    /* 업로드 카드 등고선 */
    [data-testid="stHorizontalBlock"] > div {
        align-self: stretch !important;
    }
    [data-testid="stHorizontalBlock"] > div > div {
        height: 100% !important;
    }

    /* 파일 업로더 드롭존 */
    [data-testid="stFileUploader"] {
        display: flex;
        flex-direction: column;
        align-items: stretch;
    }
    [data-testid="stFileUploaderDropzone"] {
        border: 2px dashed #B8D4EA !important;
        border-radius: 14px !important;
        background: #F8FBFE !important;
        padding: 20px 14px !important;
        text-align: center !important;
        transition: border-color 0.2s, background 0.2s !important;
        cursor: pointer !important;
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
        min-height: 80px !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: #1276B8 !important;
        background: #EEF4FA !important;
    }
    /* 파일형식·용량 텍스트 숨김 (버튼 내부 span은 유지) */
    [data-testid="stFileUploaderDropzoneInstructions"],
    [data-testid="stFileUploaderDropzoneInstructions"] span,
    [data-testid="stFileUploaderDropzone"] small {
        display: none !important;
    }
    /* 업로드 버튼 — 단독 노출 및 중앙 정렬 */
    [data-testid="stFileUploaderDropzone"] button {
        display: block !important;
        background: linear-gradient(135deg, #1276B8, #0D2E52) !important;
        color: white !important;
        border: none !important;
        border-radius: 100px !important;
        font-weight: 700 !important;
        font-size: 0.82rem !important;
        padding: 10px 32px !important;
        width: auto !important;
        margin: 0 auto !important;
    }
    /* 업로드된 파일 리스트 */
    [data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] {
        border-radius: 8px !important;
        border: 1px solid #D8E8F4 !important;
        background: white !important;
    }

    /* 섹션 타이틀 (채점 기준용) */
    .section-title {
        font-size: 0.88rem;
        font-weight: 800;
        color: #0D2E52;
        letter-spacing: -0.02em;
        margin: 0 0 10px 0;
        padding-bottom: 10px;
        border-bottom: 1px solid #C8DEF0;
    }

    /* 점수 히어로 카드 */
    .score-hero {
        max-width: 900px;
        margin: 24px auto 0;
        background: linear-gradient(135deg, #1276B8, #0D2E52);
        border-radius: 24px;
        padding: 44px 40px;
        text-align: center;
    }
    .score-hero .sh-label {
        font-size: 0.95rem;
        font-weight: 700;
        color: rgba(255,255,255,0.7);
        letter-spacing: 0.1em;
        text-transform: uppercase;
        margin-bottom: 16px;
    }
    .score-hero .sh-score {
        font-size: 6rem;
        font-weight: 900;
        color: white;
        line-height: 1;
        letter-spacing: -0.04em;
    }
    .score-hero .sh-denom {
        font-size: 1.5rem;
        color: rgba(255,255,255,0.55);
        margin-top: 12px;
        font-weight: 600;
    }

    /* 결과 박스 */
    .result-box {
        background: white;
        border-radius: 16px;
        padding: 36px;
        margin: 20px auto;
        max-width: 900px;
        border: 1px solid #D8E8F4;
        line-height: 1.8;
        word-break: keep-all;
        overflow-wrap: break-word;
    }

    /* 버튼 */
    div[data-testid="stButton"] > button[kind="primary"] {
        background: linear-gradient(135deg, #1276B8, #0D2E52);
        border: none;
        border-radius: 100px;
        padding: 14px 0;
        font-size: 1rem;
        font-weight: 700;
        letter-spacing: -0.01em;
        color: white;
        box-shadow: 0 4px 16px rgba(18,118,184,0.3);
        transition: all 0.15s;
    }
    div[data-testid="stButton"] > button[kind="primary"]:hover {
        box-shadow: 0 6px 20px rgba(18,118,184,0.45);
        transform: translateY(-1px);
    }

    /* 탭 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        background: #EEF4FA;
        border-radius: 10px;
        padding: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        font-size: 0.85rem;
        font-weight: 600;
        padding: 6px 18px;
        color: #7A9AB8;
    }
    .stTabs [aria-selected="true"] {
        color: #0D2E52 !important;
        background: white !important;
    }

    hr { border-color: #AABFD4; margin: 24px 0; }

    /* 스피너 하단 중앙 고정 */
    [data-testid="stSpinner"],
    .stSpinner,
    [data-testid="stStatusWidget"] {
        position: fixed !important;
        left: 50% !important;
        bottom: 36px !important;
        transform: translateX(-50%) !important;
        right: auto !important;
        background: white !important;
        padding: 10px 24px !important;
        border-radius: 100px !important;
        border: 1px solid #D8E8F4 !important;
        box-shadow: 0 4px 20px rgba(13,46,82,0.12) !important;
        z-index: 9999 !important;
        width: auto !important;
        white-space: nowrap !important;
    }

    /* 완료 박스 */
    .done-box {
        background: linear-gradient(180deg, #F2F8FE 0%, #FFFFFF 35%);
        border-radius: 20px;
        border: 1.5px solid #C8DEF0;
        box-shadow: 0 4px 24px rgba(13,46,82,0.09);
        padding: 28px 40px;
        text-align: center;
        font-size: 1.05rem;
        font-weight: 700;
        color: #1276B8;
        letter-spacing: -0.01em;
        margin-top: 20px;
        margin-bottom: 24px;
    }
    .done-box .done-sub {
        font-size: 0.85rem;
        font-weight: 500;
        color: #9DBAD4;
        margin-top: 6px;
    }
</style>
""", unsafe_allow_html=True)

SYSTEM_PROMPT = """당신은 [AI 활용 업무 혁신 해커톤]의 수석 심사위원 AI입니다. 현업 직원의 AI 역량 강화 및 업무 자동화·효율화를 목표로 하는 본 대회에서, 제출된 활동계획서(기획)와 앱 소스코드를 면밀히 분석하여 공정하고 전문적인 심사 결과를 제시해야 합니다.

## 언어 스타일 원칙
- 모든 심사 결과는 반드시 '합니다', '있습니다', '됩니다', '하였습니다' 등 정중한 합쇼체로 작성하세요.
- '~함', '~있음', '~됨', '~임', '~함.' 등 명사형·축약형 종결 표현은 절대 사용하지 마세요.
- 문장을 짧게 줄이더라도 반드시 서술어를 완전한 경어체로 마무리하세요.

## 핵심 심사 원칙
1. **증거 기반 평가**: 모든 점수는 제출물에 실제로 존재하는 내용(코드 함수명, 프롬프트 텍스트, 기획서 문구 등)을 인용하여 근거를 제시합니다. 없는 내용을 추정하거나 가정하지 않습니다.
2. **절대 기준 적용**: 해커톤이라는 이유로 관대하게 평가하지 않습니다. IT 현업 전문가 수준의 절대 기준으로 채점합니다.
3. **일관성**: 동일 제출물은 항상 동일한 점수를 받아야 합니다. 루브릭 밴드를 엄격히 준수하세요.
4. **냉철한 전문성**: 시니어 개발자 + 비즈니스 컨설턴트의 시각으로 장점과 결함을 동등하게 서술합니다.

## 심사 절차 (반드시 이 순서로 수행)
1. 제출 파일 목록 및 전체 구조 파악
2. 기획서에서 핵심 문제 정의·해결책·기대 기능 추출
3. 코드의 실제 구현 내용 분석 (함수, 프롬프트, 예외처리 등)
4. 기획-코드 기능 매핑 비교 (기획한 것 vs 실제 구현된 것)
5. 아래 루브릭에 따라 항목별 점수 산정
6. 출력 형식에 맞춰 심사 결과 작성

## 심사 기준 및 세부 루브릭 (총점 100점)

### 1. 업무 유용성 및 실용성 (25점)

**[문제 정의 명확성 — 13점]**
- 12~13점: 특정 직군·부서의 실제 업무 고통점을 데이터·사례와 함께 구체적으로 제시. 누구나 읽고 즉시 이해 가능한 수준.
- 9~11점: 문제가 명확하고 현업 적용 가능하나 구체성 다소 부족.
- 6~8점: 문제를 언급하나 추상적이거나 특정 직군에 한정되지 않은 수준.
- 3~5점: 문제 정의가 불분명하거나 AI 활용과 연관성 낮음.
- 0~2점: 문제 정의 없음 또는 현업과 완전히 무관함.

**[현장 도입 가능성 — 12점]**
- 11~12점: 즉시 현장 투입 가능한 수준. 정량적 시간 절감 효과 또는 생산성 지표가 명시됨.
- 8~10점: 소규모 수정 후 현장 적용 가능. 효과가 합리적으로 추정 가능.
- 5~7점: 현장 적용까지 상당한 개선 필요. 효과 불분명.
- 3~4점: 현장 투입 어려움. 실용성 낮음.
- 0~2점: 현장 도입 불가 수준.

### 2. AI 기술 활용 및 구현력 (15점)

**[AI 기술 적절성 — 8점]**
- 7~8점: 문제에 최적화된 AI 기술 선택. LLM 특성을 충분히 활용하는 설계로 과제 해결에 최적화됨.
- 5~6점: 적절한 AI 기술 선택. 일반적인 수준의 활용이나 개선 여지 있음.
- 3~4점: AI 기술 선택이 다소 부적절하거나 활용이 형식적인 수준.
- 0~2점: AI 기술 미활용 또는 완전히 부적절한 방식.

**[구현 완성도 — 7점]**
- 6~7점: 안정적으로 작동. 프롬프트 설계가 세밀하고 예외처리 완비. 에이전트·파이프라인 등 고도화 구현 포함.
- 4~5점: 대체로 작동. 프롬프트 설계 적절. 소규모 개선 필요.
- 2~3점: 작동하나 불안정하거나 프롬프트 설계 미흡. API 단순 호출 수준.
- 0~1점: 구현이 불완전하거나 작동 불가 수준.

### 3. 코드 완성도 및 구동 안정성 (10점)

**[코드 구조·가독성 — 10점]**
- 9~10점: 함수 분리, 명확한 변수명, 환경변수 관리(.env) 완비.
- 7~8점: 대체로 구조적. 소규모 개선 필요.
- 5~6점: 작동하나 구조·가독성 문제 다수.
- 3~4점: 하드코딩 다수, 함수 분리 없음, 유지보수 어려움.
- 0~2점: 코드 실행 불가 수준 또는 구조 없음.

### 4. 기획-구현 일치성 및 기획서 논리성 (15점)

**[기획서 논리성 — 8점]**
- 7~8점: 문제→원인→해결책→기대효과 논리 흐름 완벽. 타인이 읽고 즉시 이해 가능.
- 5~6점: 논리 흐름 대체로 명확. 소규모 보완 필요.
- 3~4점: 논리 흐름 있으나 비약 또는 누락 존재.
- 1~2점: 논리 불명확. 이해하기 어려움.
- 0점: 기획서 없음 또는 논리 없음.

**[구현 일치도 — 7점]**
- 6~7점: 기획서의 모든 핵심 기능이 코드에 구현됨. 추가 구현 기능도 존재.
- 4~5점: 주요 기능 구현. 1~2개 소기능 누락.
- 2~3점: 일부 핵심 기능 누락. 기획-구현 괴리 존재.
- 1점: 기획 대비 구현이 절반 이하.
- 0점: 기획서와 코드 간 연관성 없음.

### 5. 창의성 및 발표 (35점)

**[아이디어 혁신성 — 10점]**
- 9~10점: 기존 방식과 명확히 차별화된 독창적 아이디어. 업무 혁신 관점에서 새로운 가치를 창출함.
- 7~8점: 아이디어에 참신함이 있으나 유사 사례가 존재함.
- 5~6점: 일반적인 수준의 아이디어. 독창성 부족.
- 3~4점: 기존 방식의 단순 반복. 혁신성 거의 없음.
- 0~2점: 독창성 전무. 진부하거나 범용적.

**[개발 난이도 — 10점]**
- 9~10점: 복잡한 AI 파이프라인·에이전트 구성 또는 다중 API 연동 등 높은 기술 수준을 요구하는 구현.
- 7~8점: 적절한 기술적 도전이 있으며 단순 API 호출 이상의 구현을 포함.
- 5~6점: 기본적인 API 연동 수준. 기술적 도전 요소 미흡.
- 3~4점: 최소한의 구현만 존재. 난이도 매우 낮음.
- 0~2점: 기술 구현이 없거나 템플릿 복사 수준.

**[디자인 — 5점]**
- 5점: 사용자 친화적 UI/UX. 직관적 흐름과 시각적 완성도가 높음.
- 4점: 전반적으로 정돈된 디자인. 소규모 개선 필요.
- 3점: 기본 수준의 UI. 사용성 문제 일부 존재.
- 1~2점: 디자인 고려가 미흡. 사용자 경험 저해.
- 0점: 디자인 없음 또는 구동 불가.

**[발표자료 완성도 — 5점]**
- 5점: 문제·해결책·기대효과·데모가 논리적으로 구성된 발표자료. 설득력 높음.
- 4점: 전반적으로 완성된 발표자료. 일부 내용 보완 필요.
- 3점: 기본 내용은 포함. 구성·설득력 미흡.
- 1~2점: 발표자료가 불완전하거나 내용이 빈약.
- 0점: 발표자료 없음.

**[파급효과 — 5점]**
- 5점: 조직 전체 또는 업계에 적용 가능한 광범위한 파급효과 명시. 정량적 근거 포함.
- 4점: 특정 팀·부서 수준의 파급효과가 구체적으로 기술됨.
- 3점: 파급효과를 언급하나 범위·근거가 불분명.
- 1~2점: 파급효과 서술 매우 빈약.
- 0점: 파급효과 기술 없음.

## 출력 형식 (엄격히 준수)

반드시 아래 마크다운 형식으로 출력하세요. 각 항목의 ✅ 잘한 점과 ❌ 부족한 점에는 반드시 제출물의 구체적 내용(코드·기획서 인용)을 포함하세요. 부족한 점은 최소 2개 이상 제시하세요.

---
## 📊 심사 결과 요약
- **최종 총점:** [X] / 100점

### 항목별 점수 및 심사평

**1. 업무 유용성 및 실용성 — [X] / 25점**
- 문제 정의 명확성: [X] / 13점
- 현장 도입 가능성: [X] / 12점

✅ **잘한 점**
- [구체적 장점 — 기획서/코드 근거 포함]
- [구체적 장점 — 기획서/코드 근거 포함]

❌ **부족한 점 (감점 요인)**
- [구체적 부족점 — 어떤 부분이 왜 아쉬운지 근거 명시]
- [구체적 부족점 — 어떤 부분이 왜 아쉬운지 근거 명시]

---

**2. AI 기술 활용 및 구현력 — [X] / 15점**
- AI 기술 적절성: [X] / 8점
- 구현 완성도: [X] / 7점

✅ **잘한 점**
- [구체적 장점 — 프롬프트/코드 근거 포함]
- [구체적 장점]

❌ **부족한 점 (감점 요인)**
- [구체적 부족점 — 최소 2개]
- [구체적 부족점]

---

**3. 코드 완성도 및 구동 안정성 — [X] / 10점**
- 코드 구조·가독성: [X] / 10점

✅ **잘한 점**
- [구체적 장점]

❌ **부족한 점 (감점 요인)**
- [구체적 부족점 — 최소 2개]
- [구체적 부족점]

---

**4. 기획-구현 일치성 — [X] / 15점**
- 기획서 논리성: [X] / 8점
- 구현 일치도: [X] / 7점

✅ **잘한 점**
- [구체적 장점]

❌ **부족한 점 (감점 요인)**
- [구체적 부족점 — 최소 2개]
- [구체적 부족점]

---

**5. 창의성 및 발표 — [X] / 35점**
- 아이디어 혁신성: [X] / 10점
- 개발 난이도: [X] / 10점
- 디자인: [X] / 5점
- 발표자료 완성도: [X] / 5점
- 파급효과: [X] / 5점

✅ **잘한 점**
- [구체적 장점]

❌ **부족한 점 (감점 요인)**
- [구체적 부족점 — 최소 2개]
- [구체적 부족점]

---

### 💡 종합 평가
- [이 프로젝트의 핵심 가치, 차별점, 현업 적용 시 가장 큰 기여 요소 요약]

### 🚀 고도화 제언
1. [기술적 완성도를 높이기 위한 코드·아키텍처 개선 방향]
2. [프롬프트 튜닝 또는 AI 워크플로우 고도화 방안]
3. [현업 적용 확장을 위한 추가 기능 또는 배포 전략]
---"""


def read_uploaded_file(file) -> str:
    try:
        if file.name.lower().endswith(".docx"):
            doc = Document(io.BytesIO(file.read()))
            return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        content = file.read()
        for encoding in ("utf-8", "cp949", "euc-kr", "latin-1"):
            try:
                return content.decode(encoding)
            except (UnicodeDecodeError, AttributeError):
                continue
        return content.decode("utf-8", errors="replace")
    except Exception as e:
        return f"[파일 읽기 오류: {e}]"


def build_user_message(activity_plan: str, code_content: str, extra_content: str) -> str:
    extra_section = f"\n---\n## 추가 제출 파일 (requirements.txt · README · 프롬프트 등)\n{extra_content}" if extra_content.strip() else ""
    return f"""다음 해커톤 제출물을 심사해주세요.

---
## 활동계획서 및 기획 문서
{activity_plan if activity_plan.strip() else "(제출 없음)"}

---
## 앱 소스코드
{code_content if code_content.strip() else "(제출 없음)"}
{extra_section}

---
위 내용을 바탕으로 심사 기준에 따라 엄격하게 평가하고, 지정된 마크다운 형식으로 심사 결과를 작성해주세요."""


def generate_pdf(markdown_text: str) -> bytes:
    pdf = FPDF()
    pdf.set_margins(left=20, top=20, right=20)
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()
    _base = os.path.dirname(os.path.abspath(__file__))
    pdf.add_font("Malgun", style="", fname=os.path.join(_base, "fonts", "NanumGothic-Regular.ttf"))
    pdf.add_font("Malgun", style="B", fname=os.path.join(_base, "fonts", "NanumGothic-Bold.ttf"))

    pw = pdf.w - pdf.l_margin - pdf.r_margin

    def lx():
        pdf.set_x(pdf.l_margin)

    def strip_inline(text: str) -> str:
        text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
        text = re.sub(r'\*(.+?)\*', r'\1', text)
        text = re.sub(r'`(.+?)`', r'\1', text)
        text = re.sub(r'~~(.+?)~~', r'\1', text)
        text = text.replace('✅', '[O]').replace('❌', '[X]')
        for ch in ('📊', '💡', '🚀', '📝', '🏆'):
            text = text.replace(ch, '')
        return text.strip()

    for line in markdown_text.split('\n'):
        stripped = line.strip()

        if stripped == '---':
            pdf.ln(3)
            pdf.set_draw_color(180, 180, 180)
            y = pdf.get_y()
            pdf.line(pdf.l_margin, y, pdf.w - pdf.r_margin, y)
            pdf.ln(5)

        elif stripped.startswith('## '):
            pdf.ln(5)
            lx()
            pdf.set_font("Malgun", "B", 16)
            pdf.multi_cell(pw, 10, strip_inline(stripped[3:]))
            pdf.ln(2)

        elif stripped.startswith('### '):
            pdf.ln(4)
            lx()
            pdf.set_font("Malgun", "B", 13)
            pdf.multi_cell(pw, 8, strip_inline(stripped[4:]))
            pdf.ln(1)

        elif stripped.startswith('#### '):
            pdf.ln(2)
            lx()
            pdf.set_font("Malgun", "B", 11)
            pdf.multi_cell(pw, 7, strip_inline(stripped[5:]))

        elif stripped.startswith('✅') or stripped.startswith('[O]'):
            pdf.ln(3)
            lx()
            pdf.set_font("Malgun", "B", 11)
            pdf.multi_cell(pw, 7, strip_inline(stripped))
            pdf.ln(1)

        elif stripped.startswith('❌') or stripped.startswith('[X]'):
            pdf.ln(3)
            lx()
            pdf.set_font("Malgun", "B", 11)
            pdf.multi_cell(pw, 7, strip_inline(stripped))
            pdf.ln(1)

        elif stripped.startswith(('- ', '* ', '• ')):
            lx()
            pdf.set_x(pdf.l_margin + 5)
            pdf.set_font("Malgun", "", 10)
            pdf.multi_cell(pw - 5, 7, '• ' + strip_inline(stripped[2:]))

        elif re.match(r'^\d+\.', stripped):
            pdf.ln(1)
            lx()
            pdf.set_font("Malgun", "", 10)
            pdf.multi_cell(pw, 7, strip_inline(stripped))

        elif stripped == '':
            pdf.ln(4)

        else:
            lx()
            pdf.set_font("Malgun", "", 10)
            pdf.multi_cell(pw, 7, strip_inline(stripped))

    return bytes(pdf.output())


def run_evaluation(api_key: str, activity_plan: str, code_content: str, extra_content: str):
    client = OpenAI(api_key=api_key)
    user_message = build_user_message(activity_plan, code_content, extra_content)

    stream = client.chat.completions.create(
        model="gpt-4o",
        max_tokens=8096,
        temperature=0,
        seed=42,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_message},
        ],
        stream=True,
    )
    for chunk in stream:
        delta = chunk.choices[0].delta.content
        if delta is not None:
            yield delta


# ─── UI ──────────────────────────────────────────────────────────────────────

api_key = os.environ.get("OPENAI_API_KEY", "")

# 로고 base64 인코딩
_logo_path = os.path.join(os.path.dirname(__file__), "이노커브AI 로고.png")
with open(_logo_path, "rb") as _f:
    _logo_b64 = base64.b64encode(_f.read()).decode()

st.markdown(f"""
<div class="hero">
  <div class="hero-logo-wrap">
    <img src="data:image/png;base64,{_logo_b64}" alt="이노커브AI 로고">
  </div>
  <div class="hero-divider"></div>
  <h1>AI 활용 업무 혁신<br>해커톤 심사 시스템</h1>
  <p>AI 기반 전문 심사 &nbsp;·&nbsp; 5개 평가 영역 &nbsp;·&nbsp; 100점 만점 정밀 분석</p>
  <div class="hero-stats">
    <div class="hero-stat"><div class="stat-val">100</div><div class="stat-lbl">총점</div></div>
    <div class="hero-stat"><div class="stat-val">5</div><div class="stat-lbl">평가 기준</div></div>
    <div class="hero-stat"><div class="stat-val">11</div><div class="stat-lbl">세부 항목</div></div>
  </div>
</div>
""", unsafe_allow_html=True)

# 채점 기준 카드
st.markdown("""
<div class="criteria-grid">

  <div class="criteria-card">
    <div class="c-title-row">
      <div class="c-title">업무 유용성 및 실용성</div>
      <span class="pts-badge">25점</span>
    </div>
    <div class="sub-scores">
      <div class="sub-row"><span class="sub-label">문제 정의 명확성</span><span class="sub-pts">13점</span></div>
      <div class="sub-row"><span class="sub-label">현장 도입 가능성</span><span class="sub-pts">12점</span></div>
    </div>
    <div class="warn-box">
      <div class="warn-title">주요 감점 포인트</div>
      <ul>
        <li>문제가 추상적이거나 특정 직군에 한정되지 않은 경우</li>
        <li>실제 업무 흐름과 동떨어진 시나리오</li>
        <li>시간 절감 효과가 불분명한 경우</li>
      </ul>
    </div>
  </div>

  <div class="criteria-card">
    <div class="c-title-row">
      <div class="c-title">AI 기술 활용 및 구현력</div>
      <span class="pts-badge">15점</span>
    </div>
    <div class="sub-scores">
      <div class="sub-row"><span class="sub-label">AI 기술 적절성</span><span class="sub-pts">8점</span></div>
      <div class="sub-row"><span class="sub-label">구현 완성도</span><span class="sub-pts">7점</span></div>
    </div>
    <div class="warn-box">
      <div class="warn-title">주요 감점 포인트</div>
      <ul>
        <li>단순 API 호출만 하고 AI 활용이 형식적인 경우</li>
        <li>AI가 핵심 역할이 아닌 단순 장식 수준인 경우</li>
      </ul>
    </div>
  </div>

  <div class="criteria-card">
    <div class="c-title-row">
      <div class="c-title">코드 완성도 및 구동 안정성</div>
      <span class="pts-badge">10점</span>
    </div>
    <div class="sub-scores">
      <div class="sub-row"><span class="sub-label">코드 구조·가독성</span><span class="sub-pts">10점</span></div>
    </div>
    <div class="warn-box">
      <div class="warn-title">주요 감점 포인트</div>
      <ul>
        <li>하드코딩된 API 키·경로</li>
        <li>불필요한 중복 코드·비효율 구조</li>
      </ul>
    </div>
  </div>

  <div class="criteria-card">
    <div class="c-title-row">
      <div class="c-title">기획-구현 일치성</div>
      <span class="pts-badge">15점</span>
    </div>
    <div class="sub-scores">
      <div class="sub-row"><span class="sub-label">기획서 논리성</span><span class="sub-pts">8점</span></div>
      <div class="sub-row"><span class="sub-label">구현 일치도</span><span class="sub-pts">7점</span></div>
    </div>
    <div class="warn-box">
      <div class="warn-title">주요 감점 포인트</div>
      <ul>
        <li>기획서에 명시된 기능이 코드에 없는 경우</li>
        <li>문서가 부실하거나 타인이 이해하기 어려운 경우</li>
        <li>기획 범위를 크게 벗어난 구현</li>
      </ul>
    </div>
  </div>

  <div class="criteria-card" style="grid-column: 1 / -1;">
    <div class="c-title-row">
      <div class="c-title">창의성 및 발표</div>
      <span class="pts-badge">35점</span>
    </div>
    <div class="sub-scores">
      <div class="sub-row"><span class="sub-label">아이디어 혁신성</span><span class="sub-pts">10점</span></div>
      <div class="sub-row"><span class="sub-label">개발 난이도</span><span class="sub-pts">10점</span></div>
      <div class="sub-row"><span class="sub-label">디자인</span><span class="sub-pts">5점</span></div>
      <div class="sub-row"><span class="sub-label">발표자료 완성도</span><span class="sub-pts">5점</span></div>
      <div class="sub-row"><span class="sub-label">파급효과</span><span class="sub-pts">5점</span></div>
    </div>
    <div class="warn-box">
      <div class="warn-title">주요 감점 포인트</div>
      <ul>
        <li>기존 방식의 단순 반복으로 독창성이 없는 경우</li>
        <li>발표자료가 불완전하거나 파급효과 근거가 빈약한 경우</li>
        <li>UI 완성도가 낮거나 사용자 경험이 저해되는 경우</li>
      </ul>
    </div>
  </div>

</div>
""", unsafe_allow_html=True)

activity_plan_text = ""
code_text_extra = ""

col1, col2, col3 = st.columns(3)

# ── 활동계획서 ──
with col1:
    with st.container(border=True):
        st.markdown("""
<div class="upload-header">
  <div class="upload-icon-wrap">📋</div>
  <div class="upload-title">활동계획서 (기획)</div>
  <div class="upload-desc">문제 정의 · 솔루션 · 기대 효과</div>
</div>""", unsafe_allow_html=True)
        plan_files = st.file_uploader(
            "활동계획서 업로드",
            type=["docx", "txt", "md"],
            accept_multiple_files=True,
            key="plan_upload",
            label_visibility="collapsed",
        )
        if plan_files:
            st.caption(f"✓ {len(plan_files)}개 업로드됨")

# ── 앱 소스코드 ──
with col2:
    with st.container(border=True):
        st.markdown("""
<div class="upload-header">
  <div class="upload-icon-wrap">💻</div>
  <div class="upload-title">앱 소스코드</div>
  <div class="upload-desc">구현된 AI 앱 · 스크립트 · 설정</div>
</div>""", unsafe_allow_html=True)
        code_files = st.file_uploader(
            "소스코드 업로드",
            type=["py", "md", "txt", "json", "yaml", "yml", "toml"],
            accept_multiple_files=True,
            key="code_upload",
            label_visibility="collapsed",
        )
        if code_files:
            st.caption(f"✓ {len(code_files)}개 업로드됨")

# ── 추가 파일 ──
with col3:
    with st.container(border=True):
        st.markdown("""
<div class="upload-header">
  <div class="upload-icon-wrap">📎</div>
  <div class="upload-title">추가 파일</div>
  <div class="upload-desc">프롬프트 · README</div>
</div>""", unsafe_allow_html=True)
        extra_files = st.file_uploader(
            "추가 파일 업로드",
            type=["txt", "md", "py", "json", "toml", "yaml", "yml"],
            accept_multiple_files=True,
            key="extra_upload",
            label_visibility="collapsed",
        )
        if extra_files:
            st.caption(f"✓ {len(extra_files)}개 업로드됨")

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

start_btn = st.button("심사 시작", type="primary", use_container_width=True)

if start_btn:
    st.session_state.pop('eval_result', None)
    errors = []
    if not api_key:
        errors.append("OPENAI_API_KEY 환경변수가 설정되지 않았습니다.")

    # 활동계획서 수집
    plan_parts = []
    if activity_plan_text.strip():
        plan_parts.append(activity_plan_text.strip())
    for f in (plan_files or []):
        text = read_uploaded_file(f)
        plan_parts.append(f"### [{f.name}]\n{text}")
    activity_plan = "\n\n".join(plan_parts)

    # 소스코드 수집
    code_parts = []
    for f in (code_files or []):
        text = read_uploaded_file(f)
        code_parts.append(f"### [{f.name}]\n```\n{text}\n```")
    code_content = "\n\n".join(code_parts)

    # 추가 파일 수집
    extra_parts = []
    for f in (extra_files or []):
        text = read_uploaded_file(f)
        extra_parts.append(f"### [{f.name}]\n```\n{text}\n```")
    if code_text_extra.strip():
        extra_parts.append(f"### [직접 입력]\n```\n{code_text_extra.strip()}\n```")
    extra_content = "\n\n".join(extra_parts)

    if not activity_plan and not code_content and not extra_content:
        errors.append("활동계획서 또는 소스코드를 최소 하나 이상 입력하세요.")

    if errors:
        for e in errors:
            st.error(e)
    else:
        st.markdown("---")

        _, result_col, _ = st.columns([1, 4, 1])
        with result_col:
            st.markdown('<h4 style="text-align:center;color:#0D2E52;font-weight:800;letter-spacing:-0.02em;margin-bottom:4px">📝 심사 결과</h4>', unsafe_allow_html=True)
            score_slot = st.empty()
            score_slot.markdown("""
<div class="score-hero">
  <div class="sh-label">최종 총점</div>
  <div class="sh-score" style="opacity:0.3;font-size:3rem;letter-spacing:0">채점 중...</div>
  <div class="sh-denom">AI가 심사하고 있습니다!</div>
</div>""", unsafe_allow_html=True)
            result_placeholder = st.empty()

        full_result = ""

        with result_col:
            with st.spinner("AI 심사 진행 중..."):
                try:
                    for chunk in run_evaluation(api_key, activity_plan, code_content, extra_content):
                        full_result += chunk
                        result_placeholder.markdown(
                            f'<div class="result-box">{full_result}</div>',
                            unsafe_allow_html=True,
                        )

                    score_match = re.search(r'최종\s*총점[^\d]*(\d+)', full_result)
                    score = score_match.group(1) if score_match else "—"
                    score_slot.markdown(f"""
<div class="score-hero">
  <div class="sh-label">최종 총점</div>
  <div class="sh-score">{score}</div>
  <div class="sh-denom">/ 100점</div>
</div>""", unsafe_allow_html=True)

                    filtered_lines = [
                        l for l in full_result.split('\n')
                        if '심사 결과 요약' not in l
                        and not re.match(r'^\s*-\s*\*\*최종\s*총점', l)
                    ]
                    processed_lines = []
                    for line in filtered_lines:
                        s = line.strip()
                        if s == '---':
                            s = '<hr style="border:none;border-top:2px solid #AABFD4;margin:24px 0">'
                        elif s.startswith('#### '):
                            s = f'<h4 style="font-size:1.05rem;font-weight:700;color:#1276B8;margin:14px 0 6px">{s[5:]}</h4>'
                        elif s.startswith('### '):
                            s = f'<h3 style="font-size:1.35rem;font-weight:800;color:#0D2E52;margin:20px 0 8px">{s[4:]}</h3>'
                        elif s.startswith('## '):
                            s = f'<h2 style="font-size:1.5rem;font-weight:900;color:#0D2E52;margin:24px 0 10px">{s[3:]}</h2>'
                        elif re.match(r'^\*\*\d+\.', s):
                            clean = re.sub(r'\*\*(.+?)\*\*', r'\1', s)
                            s = f'<p style="font-size:1.2rem;font-weight:800;color:#0D2E52;margin:16px 0 6px 0">{clean}</p>'
                        elif s.startswith(('- ', '* ', '• ')):
                            content = s[2:]
                            content = re.sub(r'~~(.+?)~~', r'\1', content)
                            content = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', content)
                            s = f'<p style="font-size:1.05rem;color:#374151;margin:6px 0 6px 14px;line-height:1.75">• {content}</p>'
                        elif re.match(r'^\d+\.', s):
                            content = re.sub(r'~~(.+?)~~', r'\1', s)
                            content = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', content)
                            s = f'<p style="font-size:1.05rem;color:#374151;margin:10px 0 6px 0;line-height:1.7">{content}</p>'
                        processed_lines.append(s)
                    display_result = '\n'.join(processed_lines).strip()
                    result_placeholder.markdown(
                        f'<div class="result-box">{display_result}</div>',
                        unsafe_allow_html=True,
                    )

                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    pdf_bytes = generate_pdf(full_result)

                    st.session_state['eval_result'] = {
                        'display_result': display_result,
                        'score': score,
                        'pdf_bytes': pdf_bytes,
                        'timestamp': timestamp,
                    }

                    st.markdown("""
<div class="done-box">
  수고하셨습니다! 모든 심사가 완료되었습니다.
  <div class="done-sub">아래 버튼에서 심사 결과 PDF를 다운로드할 수 있습니다.</div>
</div>""", unsafe_allow_html=True)
                    _, btn_c, _ = st.columns([1, 2, 1])
                    with btn_c:
                        st.download_button(
                            label="📥 심사 결과 다운로드 (.pdf)",
                            data=pdf_bytes,
                            file_name=f"심사결과_{timestamp}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                        )

                except AuthenticationError:
                    st.error("API Key가 올바르지 않습니다.")
                except RateLimitError:
                    st.error("API 요청 한도를 초과했습니다. 잠시 후 다시 시도하세요.")
                except Exception as e:
                    st.error(f"오류가 발생했습니다: {e}")

if not start_btn and 'eval_result' in st.session_state:
    er = st.session_state['eval_result']
    st.markdown("---")
    _, result_col, _ = st.columns([1, 4, 1])
    with result_col:
        st.markdown('<h4 style="text-align:center;color:#0D2E52;font-weight:800;letter-spacing:-0.02em;margin-bottom:4px">📝 심사 결과</h4>', unsafe_allow_html=True)
        st.markdown(f"""
<div class="score-hero">
  <div class="sh-label">최종 총점</div>
  <div class="sh-score">{er['score']}</div>
  <div class="sh-denom">/ 100점</div>
</div>""", unsafe_allow_html=True)
        st.markdown(
            f'<div class="result-box">{er["display_result"]}</div>',
            unsafe_allow_html=True,
        )
        st.markdown("""
<div class="done-box">
  수고하셨습니다! 모든 심사가 완료되었습니다.
  <div class="done-sub">아래 버튼에서 심사 결과 PDF를 다운로드할 수 있습니다.</div>
</div>""", unsafe_allow_html=True)
        _, btn_c, _ = st.columns([1, 2, 1])
        with btn_c:
            st.download_button(
                label="📥 심사 결과 다운로드 (.pdf)",
                data=er['pdf_bytes'],
                file_name=f"심사결과_{er['timestamp']}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
