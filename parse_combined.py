# -*- coding: utf-8 -*-
"""
두 파일을 결합해 CFPS 통합 단가 분석 CSV 생성
- Excel: CFPS 반응 조성 (cfps_parser.py 사용)
- DOCX : S12 세포추출물 제조 프로토콜 (배지/세척/용해 버퍼 조성 + 배치 수율)
  → 추출물 제조 비용을 배치당 반응 수로 나눠 반응당 비용에 포함
"""
import sys, io, os, shutil
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import pandas as pd
from docx import Document
from cfps_parser import parse_cfps_excel

# ── 경로 ──────────────────────────────────────────────────────────────────────
XL_PATH   = r'C:\Users\user\Desktop\CFPS_workflow\cell-free test components info_for E. coli cell extract_JEJ_250311.xlsx'
DOCX_PATH = r'C:\Users\user\Desktop\CFPS_workflow\S12 cell extract preparation_260318_YK.docx'
OUT_DIR   = 'data'
os.makedirs(OUT_DIR, exist_ok=True)

# DOCX 파일 → data/ 복사
shutil.copy2(DOCX_PATH, os.path.join(OUT_DIR, 'S12_extract_preparation_YK_260318.docx'))
print('[1] DOCX copied to data/')

# ── 1. Excel → CFPS 반응 성분 (v1) ───────────────────────────────────────────
print('[2] Parsing Excel...')
excel_result = parse_cfps_excel(XL_PATH)
df_v1 = pd.DataFrame(excel_result['reagents'])
df_v1['출처'] = 'CFPS 반응 (Excel)'
print(f'    → {len(df_v1)}개 시약, 반응 부피 {excel_result["total_rxn_vol"]} µL')

# ── 2. DOCX → 세포추출물 제조 성분 파싱 ──────────────────────────────────────
print('[3] Parsing DOCX...')
doc = Document(DOCX_PATH)

# DOCX에서 추출한 정보:
# - 2L 배양 → wet cell ~20 g → 최종 상등액 ~12 mL
# - 15 µL/rxn, 4 µL extract → ~3000 rxns per 2L batch
RXNS_PER_BATCH = 3000  # 2 L E. coli 배양 → ~3000 반응 분량 (DOCX: 12 mL / 4 µL × rxn)
CULTURE_VOLUME_L = 2

# Table 0: 2x YT-PG 배지 (1 L 기준) → 2L 배양
MEDIUM_COMPONENTS = [
    # (이름, g/L_of_medium × 2L, 카테고리, 공급사)
    ('Tryptone', 16 * CULTURE_VOLUME_L, '기타', 'BD Difco'),
    ('Yeast extract', 10 * CULTURE_VOLUME_L, '기타', 'BD Difco'),
    ('NaCl', 5 * CULTURE_VOLUME_L, '버퍼/염류', 'Sigma'),
    ('Glucose (10X)', 18.02 * CULTURE_VOLUME_L, '에너지', 'Sigma'),
]

# DOCX 원문 기준 실제 필요량:
#   Wash buffer : "Cell 2 L culture 당 1 L 필요" → 3회 세척 합산 1 L
#   Lysis buffer: "Cell 2 L culture 당 약 25 mL 필요"
WASH_TOTAL_ML  = 1000   # 1,000 mL / 2L 배양 (3회 wash 합산)
LYSIS_TOTAL_ML = 25     # 25 mL / 2L 배양

# DTT: 100 mM in 10 mL per 1L wash → moles per mL
# K(OAc): 6M in 10 mL per 1L wash, Mg(OAc)2: 1.4M in 10 mL, Tris-OAc: 1M in 10 mL
# 2-ME: 14.3 M (pure) → 2.5 mL per 1L wash
# 가격은 별도 조회 필요 → 0 으로 초기화
PREP_BUFFER_COMPONENTS = [
    # (이름, 약자, 카테고리, wash_mL_per_L, lysis_mL_per_L, unit, 공급사)
    ('DTT (100 mM)', 'DTT', '보조인자', 10, 1, 'mL', 'Sigma'),
    ('Potassium acetate (6 M)', 'K(OAc)', '버퍼/염류', 10, 1, 'mL', 'Sigma(P1190)'),
    ('Magnesium acetate (1.4 M)', 'Mg(OAc)2', '버퍼/염류', 10, 1, 'mL', 'Sigma(M5661)'),
    ('Tris-acetate (1 M, pH 8.2)', 'TrisOAc', '버퍼/염류', 10, 1, 'mL', 'Sigma'),
    ('2-Mercaptoethanol (2-ME)', '2-ME', '보조인자', 2.5, 0, 'mL', 'Sigma'),
]

# ── 3. 추출물 제조 성분 → 반응당 볼륨으로 환산 ────────────────────────────────
# 배지 성분은 g 단위이므로 mL 환산 불가 → 반응당 비용만 계산 가능 (가격 입력 후)
# 여기서는 '원 배치 기준 총량'만 기록하고 반응당 사용량은 '상징적 값 0'으로 표시
extract_prep_rows = []

for name, g_per_batch, cat, supplier in MEDIUM_COMPONENTS:
    extract_prep_rows.append({
        '포함': True,
        '성분': f'{name} (배지, 배치당 {g_per_batch}g)',
        '카테고리': cat,
        '반응당 사용량 (µL)': 0.0,   # 고체 원료 → µL 변환 불가
        '시약 총량 (mL)': 1000.0,    # 임시값 (실제 구매 용량 입력 필요)
        '시약 총 비용 (₩)': 0,
        '반응당 비용 (₩)': 0.0,
        '공급사': supplier,
        'CAS No.': '',
        '가격 출처 URL': '',
        '가격 조회일': '',
        '출처': '추출물 제조 (DOCX)',
        '비고': f'2L 배양 기준 {g_per_batch}g 필요 / {RXNS_PER_BATCH}rxn 배치',
    })

for name, abbr, cat, wash_ml_per_L, lysis_ml_per_L, unit, supplier in PREP_BUFFER_COMPONENTS:
    # 배치 내 총 사용량 (mL) = 버퍼 조성비(mL/L) × 버퍼 총량(L) + 용해 버퍼 기여분
    wash_used_ml  = wash_ml_per_L  * (WASH_TOTAL_ML  / 1000)
    lysis_used_ml = lysis_ml_per_L * (LYSIS_TOTAL_ML / 1000)
    total_ml = wash_used_ml + lysis_used_ml
    # 반응당 등가 부피 (µL) = 배치 총 사용량 / 배치 반응 수
    vol_per_rxn_ul = total_ml / RXNS_PER_BATCH * 1000
    extract_prep_rows.append({
        '포함': True,
        '성분': f'{name} (추출물 제조용)',
        '카테고리': cat,
        '반응당 사용량 (µL)': round(vol_per_rxn_ul, 4),
        '시약 총량 (mL)': 1000.0,
        '시약 총 비용 (₩)': 0,
        '반응당 비용 (₩)': 0.0,
        '공급사': supplier,
        'CAS No.': '',
        '가격 출처 URL': '',
        '가격 조회일': '',
        '출처': '추출물 제조 (DOCX)',
        '비고': f'2L 배양 세척({wash_ml_per_L}mL/L×3회)+용해({lysis_ml_per_L}mL/L) / {RXNS_PER_BATCH}rxn',
    })

df_prep = pd.DataFrame(extract_prep_rows)
print(f'    → 추출물 제조 성분 {len(df_prep)}개 추가')

# ── 4. v1 CSV (Excel CFPS 성분만) 저장 ───────────────────────────────────────
v1_path = os.path.join(OUT_DIR, 'cfps_parsed_reagents_v1_rxn_only.csv')
df_v1.to_csv(v1_path, index=False, encoding='utf-8-sig')
print(f'[4] v1 CSV saved: {v1_path}')

# ── 5. v2 CSV (CFPS 반응 + 추출물 제조 통합) 저장 ────────────────────────────
# v1 df에 출처·비고 컬럼 추가
for col in ['출처', '비고']:
    if col not in df_v1.columns:
        df_v1[col] = ''
df_v1['출처'] = 'CFPS 반응 (Excel)'

df_v2 = pd.concat([df_v1, df_prep], ignore_index=True)
v2_path = os.path.join(OUT_DIR, 'cfps_parsed_reagents_v2_with_extract_prep.csv')
df_v2.to_csv(v2_path, index=False, encoding='utf-8-sig')
print(f'[5] v2 CSV saved: {v2_path}  ({len(df_v2)} rows)')

# ── 6. 요약 출력 ──────────────────────────────────────────────────────────────
print('\n' + '='*60)
print('CFPS 반응 성분 (v1 — Excel)')
print('='*60)
for _, r in df_v1.iterrows():
    print(f'  [{r["카테고리"]:10s}] {r["성분"][:42]:42s}  {r["반응당 사용량 (µL)"]:.4f} µL')

print('\n' + '='*60)
print(f'추출물 제조 성분 (DOCX — {RXNS_PER_BATCH}rxn 배치 기준 반응당 환산)')
print('='*60)
for _, r in df_prep.iterrows():
    print(f'  [{r["카테고리"]:10s}] {r["성분"][:42]:42s}  {r["반응당 사용량 (µL)"]:.4f} µL')

print(f'\n총 성분 수: v1={len(df_v1)}, v2={len(df_v2)}')
print(f'v2 CSV: {v2_path}')
print('\n* 가격이 0인 항목은 웹 앱에서 가격 자동 조회 기능으로 입력하세요.')
