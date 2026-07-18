# CFPS Cost Calculator

**Cell-Free Protein Synthesis (CFPS) 반응당 단가 계산기**

실험실 연구자를 위한 Streamlit 웹앱. Excel/CSV 프로토콜 파일을 업로드하면 Claude AI가 시약 목록을 자동 파싱하고, 반응당 비용을 계산하며, 시약 가격을 웹에서 실시간 조회합니다.

---

## 주요 기능

- **AI 에이전트**: Excel/CSV/PDF/프로토콜 텍스트 업로드 → 시약 목록 자동 파싱
- **CFPS 전용 파서**: 3-시트 형식 Excel 자동 인식 (buffer composition, reagent info, reaction condition)
- **가격 자동 조회**: Claude + web_search로 각 시약 시판가 검색 (조회 URL·날짜 기록)
- **채팅 수정**: "ATP 사용량 0.3 µL로 변경" 등 자연어로 테이블 편집
- **비용 분석**: 카테고리별 파이차트, 누적 막대 차트, KPI 요약
- **CSV 내보내기**: 공급사, CAS No., 가격 출처 URL 포함

## 스택

| 구분 | 사용 기술 |
|------|-----------|
| 웹 프레임워크 | Streamlit 1.57 |
| AI | Anthropic Claude API (`claude-haiku-4-5` 파싱/채팅, `claude-sonnet-4-6` 가격조회) |
| 웹 검색 | `web_search_20250305` tool |
| 차트 | Plotly |
| 파일 파싱 | pandas, pdfplumber, python-docx |
| 디자인 | IBM Carbon + Ant Design + Radix UI |

## 설치 및 실행

```bash
git clone https://github.com/oygoet0211-max/cfps-cost-calc.git
cd cfps-cost-calc
pip install -r requirements.txt

# API 키 설정
export ANTHROPIC_API_KEY="sk-ant-..."
# 또는 .streamlit/secrets.toml 에:
# ANTHROPIC_API_KEY = "sk-ant-..."

streamlit run app.py
```

## CFPS Excel 형식

다음 3개 시트가 포함된 Excel을 업로드하면 자동 파싱됩니다:

| 시트명 | 내용 |
|--------|------|
| `6X reaction buffer composition` | 버퍼 조성 및 5 mL 스톡 기준 볼륨 |
| `Reagent information` | 공급사·카탈로그·CAS No. |
| `Cell-free condition` | 최종 반응 조성 (총 볼륨, 각 성분 부피) |

## 환경 변수

| 변수 | 설명 |
|------|------|
| `ANTHROPIC_API_KEY` | Anthropic API 키 (필수) |

## 라이선스

MIT
