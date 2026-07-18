"""Claude API 기반 CFPS 시약 파싱·수정·가격조회 에이전트"""
import json
import io
import datetime
import anthropic
import pandas as pd

CATEGORIES = ["추출물", "에너지", "NTPs", "아미노산", "버퍼/염류", "보조인자", "tRNA", "첨가물", "DNA 주형", "기타"]

SYSTEM_PROMPT = f"""당신은 CFPS (Cell-Free Protein Synthesis) 실험 비용 계산 전문 에이전트입니다.
사용자가 업로드한 파일(Excel, CSV, 프로토콜 텍스트, 구매 영수증 등)에서 시약 정보를 추출하거나,
채팅 명령으로 시약 테이블을 수정합니다.

시약 테이블 필드 정의:
- 성분: 시약 이름 (한국어 또는 영어)
- 카테고리: 반드시 아래 중 하나 → {CATEGORIES}
  * 추출물: E. coli 세포추출물
  * 에너지: PEP, 크레아틴인산 등 에너지원
  * NTPs: ATP, GTP, CTP, UTP
  * 아미노산: 아미노산 혼합물
  * 버퍼/염류: HEPES, Mg2+, K+, Tris 등
  * 보조인자: NAD, CoA, cAMP, spermidine, putrescine, folinic acid 등
  * tRNA: tRNA 제제
  * 첨가물: PEG, 계면활성제 등
  * DNA 주형: 플라스미드, 선형 DNA
  * 기타: 위에 해당 없는 것
- 반응당 사용량 (µL): 반응 1회당 사용하는 부피 (마이크로리터, 숫자)
- 시약 총량 (mL): 구매한 용기의 총량 (밀리리터, 숫자)
- 시약 총 비용 (₩): 구매 총 금액 (원화, 숫자)
- 포함: true (포함) 또는 false (제외)

규칙:
1. 파일에서 추출할 때: 모든 시약을 set_reagents 도구 한 번으로 일괄 설정.
2. 채팅 수정 요청: update_reagent / add_reagent / remove_reagent 중 적합한 것을 사용.
3. 금액 정보가 없으면 합리적인 추정값을 넣고 "(추정)"을 성분명 뒤에 표시.
4. 부피 단위가 µL가 아닌 경우(mL, µg 등) µL로 변환.
5. 응답은 반드시 한국어로.
"""

TOOLS = [
    {
        "name": "set_reagents",
        "description": "시약 테이블 전체를 새 데이터로 교체. 파일 파싱 후 최초 설정 시 사용.",
        "input_schema": {
            "type": "object",
            "required": ["reagents"],
            "properties": {
                "reagents": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "required": ["성분", "카테고리", "반응당 사용량 (µL)", "시약 총량 (mL)", "시약 총 비용 (₩)"],
                        "properties": {
                            "성분": {"type": "string"},
                            "카테고리": {"type": "string"},
                            "반응당 사용량 (µL)": {"type": "number"},
                            "시약 총량 (mL)": {"type": "number"},
                            "시약 총 비용 (₩)": {"type": "number"},
                            "포함": {"type": "boolean"},
                        },
                    },
                }
            },
        },
    },
    {
        "name": "update_reagent",
        "description": "특정 시약의 값 수정. 성분명으로 행을 찾아 지정한 필드만 업데이트.",
        "input_schema": {
            "type": "object",
            "required": ["성분"],
            "properties": {
                "성분": {"type": "string", "description": "수정할 시약 이름 (기존 테이블의 성분명과 일치해야 함)"},
                "카테고리": {"type": "string"},
                "반응당 사용량 (µL)": {"type": "number"},
                "시약 총량 (mL)": {"type": "number"},
                "시약 총 비용 (₩)": {"type": "number"},
                "포함": {"type": "boolean"},
            },
        },
    },
    {
        "name": "add_reagent",
        "description": "새 시약을 테이블에 추가.",
        "input_schema": {
            "type": "object",
            "required": ["성분", "카테고리", "반응당 사용량 (µL)", "시약 총량 (mL)", "시약 총 비용 (₩)"],
            "properties": {
                "성분": {"type": "string"},
                "카테고리": {"type": "string"},
                "반응당 사용량 (µL)": {"type": "number"},
                "시약 총량 (mL)": {"type": "number"},
                "시약 총 비용 (₩)": {"type": "number"},
                "포함": {"type": "boolean"},
            },
        },
    },
    {
        "name": "remove_reagent",
        "description": "시약을 테이블에서 삭제.",
        "input_schema": {
            "type": "object",
            "required": ["성분"],
            "properties": {
                "성분": {"type": "string"}
            },
        },
    },
]


# ── 파일 → 텍스트 변환 ────────────────────────────────────────────────────────

def file_to_text(uploaded_file) -> str:
    """Streamlit UploadedFile → Claude에게 넘길 텍스트"""
    name = uploaded_file.name.lower()
    raw = uploaded_file.read()

    if name.endswith((".xlsx", ".xls")):
        try:
            sheets = pd.read_excel(io.BytesIO(raw), sheet_name=None)
            parts = []
            for sheet_name, df in sheets.items():
                parts.append(f"=== 시트: {sheet_name} ===\n{df.to_string(index=False)}")
            return "\n\n".join(parts)
        except Exception as e:
            return f"[Excel 파싱 오류: {e}]"

    if name.endswith(".csv"):
        try:
            for enc in ("utf-8-sig", "cp949", "utf-8"):
                try:
                    df = pd.read_csv(io.BytesIO(raw), encoding=enc)
                    return df.to_string(index=False)
                except UnicodeDecodeError:
                    continue
            return "[CSV 인코딩 오류]"
        except Exception as e:
            return f"[CSV 파싱 오류: {e}]"

    if name.endswith(".pdf"):
        try:
            import pdfplumber
            text_parts = []
            with pdfplumber.open(io.BytesIO(raw)) as pdf:
                for page in pdf.pages:
                    text_parts.append(page.extract_text() or "")
            return "\n".join(text_parts)
        except ImportError:
            return "[PDF 지원: pdfplumber 설치 필요 → pip install pdfplumber]"
        except Exception as e:
            return f"[PDF 파싱 오류: {e}]"

    if name.endswith(".docx"):
        try:
            from docx import Document
            doc = Document(io.BytesIO(raw))
            return "\n".join(p.text for p in doc.paragraphs)
        except ImportError:
            return "[DOCX 지원: python-docx 설치 필요 → pip install python-docx]"
        except Exception as e:
            return f"[DOCX 파싱 오류: {e}]"

    # plain text / md / tsv
    for enc in ("utf-8-sig", "cp949", "utf-8"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return "[텍스트 디코딩 실패]"


# ── 에이전트 클래스 ───────────────────────────────────────────────────────────

PRICE_LOOKUP_TOOL = {
    "name": "set_prices",
    "description": "웹에서 조회한 시약 가격 정보를 저장합니다.",
    "input_schema": {
        "type": "object",
        "required": ["prices"],
        "properties": {
            "prices": {
                "type": "array",
                "items": {
                    "type": "object",
                    "required": ["성분", "시약 총 비용 (₩)", "시약 총량 (mL)", "가격 출처 URL"],
                    "properties": {
                        "성분":            {"type": "string", "description": "시약 이름 (기존 테이블과 일치)"},
                        "시약 총 비용 (₩)": {"type": "number", "description": "구매 총액 (원화). USD면 1380 곱해서 변환"},
                        "시약 총량 (mL)":   {"type": "number", "description": "구매 단위 총량 mL (고체면 g를 mL로 기록)"},
                        "가격 출처 URL":    {"type": "string", "description": "가격 확인한 웹 페이지 URL"},
                        "메모":            {"type": "string", "description": "가격 단위·조건·통화 등 부연 설명"},
                    },
                },
            }
        },
    },
}


class CFPSAgent:
    def __init__(self, api_key: str):
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = "claude-haiku-4-5-20251001"  # 파싱·채팅용
        self.search_model = "claude-sonnet-4-6"    # 웹 검색용

    def parse_file(self, file_text: str, filename: str) -> tuple[list[dict] | None, str]:
        """파일 텍스트 → (시약 목록, 에이전트 메시지)"""
        user_msg = (
            f"다음은 업로드된 파일({filename})의 내용입니다. "
            f"시약 정보를 추출해서 set_reagents 도구로 테이블을 설정해주세요.\n\n"
            f"```\n{file_text[:8000]}\n```"
        )
        return self._run(user_msg, current_table=None)

    def chat(self, user_message: str, current_df: pd.DataFrame) -> tuple[list[dict] | None, str]:
        """채팅 명령 → (변경된 시약 목록 또는 None, 에이전트 메시지)"""
        table_str = current_df.to_string(index=False)
        user_msg = (
            f"현재 시약 테이블:\n```\n{table_str}\n```\n\n"
            f"사용자 요청: {user_message}"
        )
        return self._run(user_msg, current_table=current_df)

    def _run(self, user_msg: str, current_table) -> tuple[list[dict] | None, str]:
        messages = [{"role": "user", "content": user_msg}]
        response = self.client.messages.create(
            model=self.model,
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            tools=TOOLS,
            messages=messages,
        )

        tool_calls = []
        text_parts = []

        for block in response.content:
            if block.type == "text":
                text_parts.append(block.text)
            elif block.type == "tool_use":
                tool_calls.append({"name": block.name, "input": block.input})

        agent_text = "\n".join(text_parts).strip()

        if not tool_calls:
            return None, agent_text or "명령을 처리했습니다."

        # 도구 호출 적용
        updated_df = current_table.copy() if current_table is not None else pd.DataFrame(
            columns=["포함", "성분", "카테고리", "반응당 사용량 (µL)", "시약 총량 (mL)", "시약 총 비용 (₩)", "반응당 비용 (₩)"]
        )

        for call in tool_calls:
            name, inp = call["name"], call["input"]

            if name == "set_reagents":
                rows = []
                for r in inp["reagents"]:
                    rows.append({
                        "포함": r.get("포함", True),
                        "성분": r["성분"],
                        "카테고리": r["카테고리"],
                        "반응당 사용량 (µL)": float(r["반응당 사용량 (µL)"]),
                        "시약 총량 (mL)": float(r["시약 총량 (mL)"]),
                        "시약 총 비용 (₩)": float(r["시약 총 비용 (₩)"]),
                        "반응당 비용 (₩)": 0.0,
                    })
                updated_df = pd.DataFrame(rows)

            elif name == "update_reagent":
                mask = updated_df["성분"] == inp["성분"]
                if not mask.any():
                    # 이름이 약간 다를 수 있으니 포함 검색
                    mask = updated_df["성분"].str.contains(inp["성분"], case=False, na=False)
                for field, val in inp.items():
                    if field != "성분" and field in updated_df.columns:
                        updated_df.loc[mask, field] = val

            elif name == "add_reagent":
                new_row = {
                    "포함": inp.get("포함", True),
                    "성분": inp["성분"],
                    "카테고리": inp["카테고리"],
                    "반응당 사용량 (µL)": float(inp["반응당 사용량 (µL)"]),
                    "시약 총량 (mL)": float(inp["시약 총량 (mL)"]),
                    "시약 총 비용 (₩)": float(inp["시약 총 비용 (₩)"]),
                    "반응당 비용 (₩)": 0.0,
                }
                updated_df = pd.concat([updated_df, pd.DataFrame([new_row])], ignore_index=True)

            elif name == "remove_reagent":
                mask = updated_df["성분"] == inp["성분"]
                if not mask.any():
                    mask = updated_df["성분"].str.contains(inp["성분"], case=False, na=False)
                updated_df = updated_df[~mask].reset_index(drop=True)

        # 반응당 비용 재계산
        updated_df["반응당 비용 (₩)"] = (
            updated_df["반응당 사용량 (µL)"]
            / (updated_df["시약 총량 (mL)"] * 1000)
            * updated_df["시약 총 비용 (₩)"]
        ).round(2)

        return updated_df.to_dict(orient="records"), agent_text or f"✅ {len(tool_calls)}개 작업 완료"

    # ── 가격 조회 ────────────────────────────────────────────────────────────

    def lookup_prices(
        self, df: pd.DataFrame, *, callback=None
    ) -> tuple[pd.DataFrame, str]:
        """
        가격이 0인 시약의 현재 시장가를 웹 검색으로 조회.

        Parameters
        ----------
        df       : 현재 시약 테이블 (가격=0인 행 대상)
        callback : lookup 중 진행 상황을 전달하는 함수 callback(msg: str)

        Returns
        -------
        (updated_df, summary_message)
        """
        today = datetime.date.today().isoformat()
        targets = df[df["시약 총 비용 (₩)"] == 0].copy()
        if targets.empty:
            return df, "모든 시약에 이미 가격이 입력되어 있습니다."

        # 조회 대상 목록 생성
        lines = []
        for _, r in targets.iterrows():
            supplier = r.get("공급사", "") if "공급사" in r.index else ""
            cas      = r.get("CAS No.", "")  if "CAS No." in r.index else ""
            lines.append(
                f"- 성분: {r['성분']}  공급사/카탈로그: {supplier}  CAS: {cas}"
            )

        user_msg = (
            "다음 CFPS 시약들의 **현재 온라인 판매가**를 조사해 주세요.\n"
            "Sigma-Aldrich (sigmaaldrich.com/KR), Thermo Fisher, 한국 총판 등에서 검색합니다.\n"
            "USD 가격은 1 USD = 1380 KRW로 환산하세요.\n\n"
            + "\n".join(lines)
            + "\n\n조회가 완료되면 set_prices 도구로 결과를 저장하세요."
        )

        if callback:
            callback(f"{len(targets)}개 시약 가격 조회 시작...")

        try:
            response = self.client.messages.create(
                model=self.search_model,
                max_tokens=8192,
                tools=[
                    {"type": "web_search_20250305", "name": "web_search", "max_uses": len(targets) + 3},
                    PRICE_LOOKUP_TOOL,
                ],
                messages=[{"role": "user", "content": user_msg}],
            )
        except Exception as e:
            return df, f"가격 조회 중 오류 발생: {e}"

        # 응답 파싱
        price_updates = []
        summary_parts = []
        for block in response.content:
            if block.type == "text" and block.text:
                summary_parts.append(block.text)
            elif block.type == "tool_use" and block.name == "set_prices":
                price_updates = block.input.get("prices", [])

        if not price_updates:
            msg = "\n".join(summary_parts) or "가격 정보를 찾지 못했습니다."
            return df, msg

        # DataFrame 업데이트
        updated = df.copy()
        if "가격 출처 URL" not in updated.columns:
            updated["가격 출처 URL"] = ""
        if "가격 조회일" not in updated.columns:
            updated["가격 조회일"] = ""
        if "가격 메모" not in updated.columns:
            updated["가격 메모"] = ""

        found = []
        for p in price_updates:
            name = p.get("성분", "")
            mask = updated["성분"] == name
            if not mask.any():
                mask = updated["성분"].str.contains(
                    name[:15], case=False, na=False, regex=False
                )
            if mask.any():
                updated.loc[mask, "시약 총 비용 (₩)"] = float(p.get("시약 총 비용 (₩)", 0))
                updated.loc[mask, "시약 총량 (mL)"]   = float(p.get("시약 총량 (mL)", 1.0))
                updated.loc[mask, "가격 출처 URL"]     = p.get("가격 출처 URL", "")
                updated.loc[mask, "가격 조회일"]        = today
                updated.loc[mask, "가격 메모"]          = p.get("메모", "")
                found.append(name)

        # 반응당 비용 재계산
        updated["반응당 비용 (₩)"] = (
            updated["반응당 사용량 (µL)"]
            / (updated["시약 총량 (mL)"] * 1000)
            * updated["시약 총 비용 (₩)"]
        ).round(2)

        summary = "\n".join(summary_parts) if summary_parts else ""
        msg = f"가격 조회 완료: {len(found)}/{len(targets)}종 업데이트\n{summary}"
        return updated, msg
