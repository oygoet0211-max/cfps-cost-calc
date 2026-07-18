"""
CFPS Excel 파일 전용 파서
시트 구조:
  - '6X reaction buffer composition' : 버퍼 조성 및 볼륨 (5 mL 스톡 기준)
  - 'Reagent information'            : 공급사·카탈로그 정보
  - 'Cell-free condition'            : 최종 반응 조성
"""
import io
import re
import pandas as pd

# ── 카테고리 추론 키워드 맵 ────────────────────────────────────────────────────
_CATEGORY_KW = {
    "추출물":   ["extract", "s12", "s30"],
    "NTPs":    ["atp", "gtp", "ctp", "utp", "amp", "cmp", "gmp", "ump",
                "nucleotide", "nucleoside"],
    "에너지":   ["creatine phosphate", "phosphoenolpyruvate", " pep ", " cp "],
    # 버퍼/염류를 아미노산보다 먼저 — potassium glutamate 등이 아미노산으로 오분류되는 것 방지
    "버퍼/염류": ["hepes", "tris", "magnesium", "potassium", "ammonium",
                 "glutamate", "acetate", "phosphate buffer", "kcl", "mgcl",
                 "potassium salt", "k(glu)", "k(oac)", "mg(oac)", "nh4(oac)",
                 "k-glut", "kglut"],
    "아미노산": ["amino acid", "alanine", "arginine", "asparagine", "aspartic",
                "cysteine", "cystein", "glutamic", "glutamine", "glycine",
                "histidine", "isoleucine", "leucine", "lysine", "methionine",
                "phenylalanine", "proline", "serine", "threonine",
                "tryptophan", "tyrosine", "valine"],
    "보조인자":  ["nad", "coa", "camp", "spermidine", "putrescine",
                 "folinic", "dtt", "2-me", "creatine kinase", " ck "],
    "tRNA":    ["trna", "t-rna"],
    "첨가물":   ["peg", "polyethylene glycol", "glycerol", "bsa"],
    "DNA 주형": ["dna", "plasmid", "template", "linear"],
}

def _infer_category(name: str) -> str:
    nl = " " + name.lower() + " "
    for cat, kws in _CATEGORY_KW.items():
        for kw in kws:
            if kw in nl:
                return cat
    return "기타"


def _to_float(val) -> float | None:
    if pd.isna(val):
        return None
    try:
        return float(str(val).replace(",", "").strip())
    except ValueError:
        return None


def _extract_ul(text) -> float | None:
    """'2.5 uL', '4 µL', 'up to 15 uL' 등에서 숫자 추출"""
    if pd.isna(text):
        return None
    m = re.search(r"([\d.]+)\s*[uµ]", str(text), re.IGNORECASE)
    if m:
        return float(m.group(1))
    m = re.search(r"[\d.]+", str(text))
    if m:
        return float(m.group())
    return None


def _parse_condition_sheet(xl: pd.ExcelFile) -> dict:
    """Cell-free condition 시트 → 반응 조성 dict"""
    df = xl.parse("Cell-free condition", header=None)
    result = {"total_vol": 15.0, "components": {}}
    for _, row in df.iterrows():
        name = str(row.iloc[1]).strip().lower() if pd.notna(row.iloc[1]) else ""
        vol_raw = row.iloc[2] if len(row) > 2 else None
        vol = _extract_ul(vol_raw)
        if not name or name in ("nan", ""):
            continue
        if "total" in name:
            if vol:
                result["total_vol"] = vol
        elif "6x" in name or "reaction mixture" in name:
            result["components"]["buffer_stock"] = vol or 2.5
        elif "peg" in name:
            result["components"]["peg"] = vol or 0.75
        elif "ck" in name or "creatine kinase" in name:
            result["components"]["ck"] = vol or 0.1
        elif "extract" in name or "s12" in name or "s30" in name:
            result["components"]["cell_extract"] = vol or 4.0
        elif "dna" in name or "template" in name:
            result["components"]["dna"] = vol
    return result


def _parse_buffer_sheet(xl: pd.ExcelFile, buffer_vol_in_rxn: float) -> list[dict]:
    """
    6X reaction buffer composition 시트 파싱.
    col 9 = µL added to 5 mL stock → per-rxn vol = col9 / 5000 * buffer_vol_in_rxn
    """
    df = xl.parse("6X reaction buffer composition", header=None)
    reagents = []
    current_name = None
    current_abbr = None
    current_ul_in_5ml = None

    for idx, row in df.iterrows():
        num     = _to_float(row.iloc[0])    # 번호 (있으면 새 시약)
        name    = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        abbr    = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        ul_5ml  = _to_float(row.iloc[9]) if len(row) > 9 else None

        # 헤더/주석 행 건너뜀
        if name.lower() in ("stock name", "nan", "") and num is None:
            continue
        if "protocol" in name.lower() or "stock" in name.lower()[:5] and num is None:
            continue

        if num is not None and name and name.lower() not in ("nan",):
            # 새 주 시약 시작
            current_name = name
            current_abbr = abbr
            current_ul_in_5ml = ul_5ml  # None일 수 있음

            if ul_5ml is not None:
                vol_per_rxn = ul_5ml / 5000 * buffer_vol_in_rxn
                reagents.append({
                    "성분": name,
                    "약자": abbr,
                    "반응당 사용량 (µL)": round(vol_per_rxn, 4),
                    "카테고리": _infer_category(name),
                })
        # NaN 번호 = 서브 항목 (예: CMP, GMP, UMP, 개별 아미노산)
        # 이미 상위 항목 볼륨에 포함되므로 별도 추가 안 함

    return reagents


def _parse_reagent_info_sheet(xl: pd.ExcelFile) -> dict:
    """
    Reagent information 시트 → abbr → {full_name, supplier, catalog, cas, mw}
    """
    df = xl.parse("Reagent information", header=None)
    info = {}
    for _, row in df.iterrows():
        abbr = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if not abbr or abbr.lower() in ("abbreviation", "nan", "amino acids"):
            continue
        full   = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        purch  = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
        cas    = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ""
        mw     = _to_float(row.iloc[5]) if len(row) > 5 else None
        info[abbr.lower()] = {
            "full_name": full, "purchase": purch, "cas": cas, "mw": mw
        }
        # 전체 이름으로도 등록
        if full:
            info[full.lower()[:30]] = info[abbr.lower()]
    return info


def _match_reagent_info(name: str, abbr: str, info_map: dict) -> dict:
    """시약 이름/약자로 Reagent information 시트와 매칭"""
    # 정확 매칭 우선
    for key in [abbr.lower(), name.lower()[:30]]:
        if key in info_map:
            return info_map[key]
    # 부분 일치 (키 포함 또는 이름 포함 양방향)
    name_lower = name.lower()
    for k, v in info_map.items():
        if not k:
            continue
        if k in name_lower or name_lower[:len(k)+3].startswith(k[:4]):
            return v
    # 약자 앞 4자 매칭 (typo 보완: cystein ↔ cysteine)
    abbr4 = abbr.lower()[:5]
    for k, v in info_map.items():
        if k[:5] == abbr4:
            return v
    return {}


def parse_cfps_excel(file) -> dict:
    """
    Parameters
    ----------
    file : str | Path | UploadedFile
        Excel 파일 경로 또는 Streamlit UploadedFile

    Returns
    -------
    dict
        {
          "total_rxn_vol": float,          # µL
          "reagents": list[dict],          # 비용 계산기용
          "errors": list[str],
        }
    """
    try:
        if hasattr(file, "read"):
            raw = file.read()
            xl = pd.ExcelFile(io.BytesIO(raw))
        else:
            xl = pd.ExcelFile(file)
    except Exception as e:
        return {"total_rxn_vol": 15.0, "reagents": [], "errors": [str(e)]}

    errors = []

    # ── 1. 반응 조성 파악 ────────────────────────────────────────────────────
    try:
        cond = _parse_condition_sheet(xl)
    except Exception as e:
        cond = {"total_vol": 15.0, "components": {}}
        errors.append(f"Cell-free condition 파싱 오류: {e}")

    total_vol    = cond["total_vol"]
    buffer_vol   = cond["components"].get("buffer_stock", 2.5)
    peg_vol      = cond["components"].get("peg", 0.75)
    ck_vol       = cond["components"].get("ck", 0.1)
    extract_vol  = cond["components"].get("cell_extract", 4.0)

    # ── 2. 버퍼 조성 파싱 ────────────────────────────────────────────────────
    try:
        buf_reagents = _parse_buffer_sheet(xl, buffer_vol)
    except Exception as e:
        buf_reagents = []
        errors.append(f"Buffer composition 파싱 오류: {e}")

    # ── 3. Reagent information 파싱 ──────────────────────────────────────────
    try:
        info_map = _parse_reagent_info_sheet(xl)
    except Exception as e:
        info_map = {}
        errors.append(f"Reagent information 파싱 오류: {e}")

    # ── 4. 버퍼 성분에 공급사 정보 매칭 ────────────────────────────────────
    reagents = []
    for r in buf_reagents:
        ri = _match_reagent_info(r["성분"], r["약자"], info_map)
        reagents.append({
            "포함":           True,
            "성분":           r["성분"],
            "카테고리":       r["카테고리"],
            "반응당 사용량 (µL)": r["반응당 사용량 (µL)"],
            "시약 총량 (mL)":  1.0,
            "시약 총 비용 (₩)": 0,
            "반응당 비용 (₩)": 0.0,
            "공급사":          ri.get("purchase", ""),
            "CAS No.":        ri.get("cas", ""),
            "가격 출처 URL":   "",
            "가격 조회일":     "",
        })

    # ── 5. Cell-free condition 추가 성분 ────────────────────────────────────
    # PEG-8000 (40% stock)
    peg_info = _match_reagent_info("PEG", "PEG", info_map)
    reagents.append({
        "포함": True, "성분": "PEG-8000 (40% stock)", "카테고리": "첨가물",
        "반응당 사용량 (µL)": peg_vol, "시약 총량 (mL)": 1.0,
        "시약 총 비용 (₩)": 0, "반응당 비용 (₩)": 0.0,
        "공급사": peg_info.get("purchase", ""), "CAS No.": "",
        "가격 출처 URL": "", "가격 조회일": "",
    })

    # Creatine Kinase (0.2% stock)
    ck_info = _match_reagent_info("Creatine kinase", "CK", info_map)
    reagents.append({
        "포함": True, "성분": "Creatine Kinase (0.2% stock)", "카테고리": "보조인자",
        "반응당 사용량 (µL)": ck_vol, "시약 총량 (mL)": 1.0,
        "시약 총 비용 (₩)": 0, "반응당 비용 (₩)": 0.0,
        "공급사": ck_info.get("purchase", "Roche(10127566001)"), "CAS No.": "",
        "가격 출처 URL": "", "가격 조회일": "",
    })

    # Cell extract (S12)
    reagents.append({
        "포함": True, "성분": "Cell extract (S12)", "카테고리": "추출물",
        "반응당 사용량 (µL)": extract_vol, "시약 총량 (mL)": 1.0,
        "시약 총 비용 (₩)": 0, "반응당 비용 (₩)": 0.0,
        "공급사": "Lab-made", "CAS No.": "",
        "가격 출처 URL": "", "가격 조회일": "",
    })

    return {
        "total_rxn_vol": total_vol,
        "reagents": reagents,
        "errors": errors,
    }
