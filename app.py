import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide", page_title="마케팅본부 월간 보고서")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;700;900&display=swap');
* { font-family: 'Noto Sans KR', 'Malgun Gothic', sans-serif; }

.report-header {
    background: linear-gradient(135deg, #1a3a6b 0%, #2d5fa8 100%);
    color: white; padding: 16px 28px; border-radius: 0;
    font-size: 22px; font-weight: 900; margin-bottom: 24px;
    letter-spacing: -0.5px;
}
.box-title {
    display: inline-block;
    border: 2px solid #c0392b; border-radius: 20px;
    padding: 4px 18px; font-size: 15px; font-weight: 800;
    color: #c0392b; margin-bottom: 10px; margin-top: 20px;
}
.note-right { font-size: 12px; color: #555; text-align: right; margin-bottom: 4px; }

/* 표 공통 */
table { border-collapse: collapse; width: 100%; font-size: 13px; }
th, td { border: 1px solid #bbb; padding: 6px 8px; text-align: center; }
thead th { background: #2d5fa8; color: white; font-weight: 700; }
.th-sub { background: #4a7bc4 !important; }
.th-group { background: #1a3a6b !important; }
.td-label { background: #e8edf7; font-weight: 700; text-align: center; }
.td-sub-label { background: #f0f4fc; font-weight: 600; }
.rate-red { color: #c0392b; font-weight: 700; }

.industry-table table { font-size: 12px; }
.upload-guide {
    background: #e8f4fd; border: 1px solid #2d5fa8;
    border-radius: 6px; padding: 10px 16px; margin-bottom: 12px; font-size: 13px;
}
@media print {
    .stButton, .stFileUploader, .upload-guide, 
    [data-testid="stSidebar"], header { display: none !important; }
    .report-header { -webkit-print-color-adjust: exact; }
}
</style>
""", unsafe_allow_html=True)

# ── 헬퍼 ──────────────────────────────────────────────
def fmt(v, dec=0):
    if v is None or (isinstance(v, float) and pd.isna(v)): return "-"
    try:
        f = float(v)
        if dec > 0: return f"{f:,.{dec}f}"
        return f"{int(round(f)):,}"
    except: return str(v)

def pct(v):
    if v is None or (isinstance(v, float) and pd.isna(v)): return "-"
    try:
        f = float(v)
        val = f * 100 if abs(f) <= 5 else f
        return f"{val:.1f}%"
    except: return str(v)

def safe(df, r, c):
    try:
        v = df.iloc[r, c]
        return None if pd.isna(v) else v
    except: return None

# ── 데이터 로드 함수 ────────────────────────────────────
@st.cache_data
def load_xlsm(b, fname):
    buf = io.BytesIO(b)
    sheets = {}
    for s in ['공급전 계획(2026년)', '공급량 계획_MJ(2026년)', '영업현황']:
        try:
            sheets[s] = pd.read_excel(buf, sheet_name=s, header=None, engine='openpyxl')
            buf.seek(0)
        except: pass
    return sheets

@st.cache_data
def load_dev(b, fname):
    buf = io.BytesIO(b)
    try:
        detail = pd.read_excel(buf, sheet_name='Sheet1', header=0)
        detail.columns = ['공급신청번호','시공업체','번지순번','주소','신청명','계약구분','건물구분',
                          '신청일','용도','업종','상품','등급','월사용예정량','공급승인일','공급일',
                          '서비스센터','설치장소주소','계량기번호','특정여부','월간개발량','공동주택명']
        return detail
    except: return None

# ════════════════════════════════════════════════════════
# UI
# ════════════════════════════════════════════════════════
st.markdown('<div class="report-header">📊 마케팅본부 _ 월간 영업현황 보고서</div>',
            unsafe_allow_html=True)

st.markdown("""
<div class="upload-guide">
📁 <b>매월 아래 2개 파일 업로드</b> + 총공급량 실적 입력하면 자동으로 보고서가 생성됩니다.<br>
&nbsp;&nbsp;① 영업일보_월간_YYYYMM.xlsm &nbsp;|&nbsp; ② 신규개발량_YYYYMM.xlsx
</div>
""", unsafe_allow_html=True)

col_u1, col_u2 = st.columns(2)
with col_u1:
    f_xlsm = st.file_uploader("① 영업일보 파일 (xlsm)", type=['xlsm','xlsx'])
with col_u2:
    f_dev  = st.file_uploader("② 신규개발량 파일 (xlsx)", type=['xlsx'])

# 수기 입력
st.markdown("---")
col_i1, col_i2, col_i3 = st.columns(3)
with col_i1:
    sel_month = st.selectbox("보고 월", list(range(1,13)), index=5)
with col_i2:
    total_actual_input = st.number_input(
        f"총공급량 당월 실적 (GJ) — 고객지원시스템 확인",
        value=0, step=1000, format="%d")
with col_i3:
    prev_cum_input = st.number_input(
        f"전월 총공급량 누계 실적 (GJ) — 자동계산용",
        value=0, step=1000, format="%d")

# 누계 자동계산
total_cum_calc = prev_cum_input + total_actual_input

if not f_xlsm:
    st.info("👆 영업일보 파일을 먼저 업로드해주세요.")
    st.stop()

# ── 데이터 파싱 ────────────────────────────────────────
sheets  = load_xlsm(f_xlsm.read(), f_xlsm.name)
f_xlsm.seek(0)

detail_df = None
if f_dev:
    detail_df = load_dev(f_dev.read(), f_dev.name)

df_plan = sheets.get('공급전 계획(2026년)')
df_vol  = sheets.get('공급량 계획_MJ(2026년)')
df_biz  = sheets.get('영업현황')

m = sel_month - 1  # 0-based index

# ── 공급전 계획 (신규개발전) ──
# row9=합계계획, col4~15=1~12월
dev_plan_monthly   = [safe(df_plan, 9, c) for c in range(4, 16)] if df_plan is not None else [None]*12
dev_plan_m         = dev_plan_monthly[m] or 0
dev_cum_plan_list  = [safe(df_plan, 12, c) for c in range(4, 16)] if df_plan is not None else [None]*12
dev_cum_plan_m     = dev_cum_plan_list[m] or 0  # 누계계획

# 폐전 계획: 영업현황시트 row7(당월폐전), row6(폐전누계)
# → 공급전 계획 시트에는 없으므로 영업현황 시트의 7행(당월폐전)과 6행(폐전누계) 사용
폐전_실적      = safe(df_biz, 7, 9) if df_biz is not None else None   # 당월폐전 합계
폐전_누계_실적  = safe(df_biz, 6, 9) if df_biz is not None else None   # 폐전누계 합계

# 신규개발전 실적 (row25), 목표 (row26)
cats_biz = ['공동주택','단독주택','소계','일반용','업무용','산업용','열병합','합계']
act, pln = {}, {}
if df_biz is not None:
    for i, cat in enumerate(cats_biz):
        act[cat] = safe(df_biz, 25, i+2)
        pln[cat] = safe(df_biz, 26, i+2)

dev_actual_m = act.get('합계') or 0
dev_rate_m   = dev_actual_m / dev_plan_m * 100 if dev_plan_m else 0

# 폐전 계획
폐전_plan_m   = safe(df_plan, None, None)   # 공급전계획 시트에 폐전 행 없음 → 영업현황에서
# 영업현황 시트: row7=당월폐전, col2~9
if df_biz is not None:
    폐전_act_m = safe(df_biz, 7, 9)  # 당월폐전 합계
    폐전_cum   = safe(df_biz, 6, 9)  # 폐전누계

# 공급량 계획 (MJ→GJ)
vol_plan_monthly = []
vol_cum_plan_monthly = []
if df_vol is not None:
    vol_plan_monthly     = [v/1000 if v else 0 for v in [safe(df_vol, 5, c) for c in range(3, 15)]]
    vol_cum_plan_monthly = [v/1000 if v else 0 for v in [safe(df_vol, 6, c) for c in range(3, 15)]]

vol_plan_m   = vol_plan_monthly[m]   if vol_plan_monthly   else 0
vol_cum_plan = vol_cum_plan_monthly[m] if vol_cum_plan_monthly else 0

# 연간계획
dev_annual   = safe(df_plan, 9, 3) if df_plan is not None else None   # col3=전년말 → 사실 연간계획
# 실제로 공급전 계획 시트: col17=년간계
dev_annual   = safe(df_plan, 9, 17) if df_plan is not None else None
vol_annual   = safe(df_vol,  5, 16) if df_vol  is not None else None  # col16=년간계(MJ)
if vol_annual: vol_annual = vol_annual / 1000

# 산업용 신규 업체
industry_df = pd.DataFrame()
if detail_df is not None:
    industry_df = detail_df[detail_df['용도'].astype(str).str.contains('산업', na=False)].copy()

# ── 순증가 계산 ──
순증가_plan_m  = dev_plan_m  - (safe(df_biz, 7, 9) or 0)  # 개발전계획 - 폐전실적(임시)
순증가_act_m   = dev_actual_m - (safe(df_biz, 7, 9) or 0) if df_biz is not None else 0

# ════════════════════════════════════════════════════════
# 화면 출력 — 2번째 사진과 동일한 구성
# ════════════════════════════════════════════════════════
st.markdown(f"### 📅 **2026년 {sel_month}월** 영업현황 보고")

# ── 표1: 공급전 및 공급량 현황 ──────────────────────────
st.markdown(f'<div class="box-title">📋 공급전 및 공급량 현황</div>', unsafe_allow_html=True)

def r(v): return f'<span class="rate-red">{v}</span>'

# 달성률 색상 처리
def rate_html(numerator, denominator):
    try:
        val = float(numerator) / float(denominator) * 100
        s = f"{val:.1f}%"
        return r(s) if val < 100 else f"<b>{s}</b>"
    except: return "-"

rows_html = f"""
<table>
  <thead>
    <tr>
      <th rowspan="2" colspan="2">구 분</th>
      <th rowspan="2">연간계획</th>
      <th colspan="3">{sel_month}월 (당월)</th>
      <th colspan="3">{sel_month}월 (누계)</th>
    </tr>
    <tr>
      <th class="th-sub">계획</th>
      <th class="th-sub">실적</th>
      <th class="th-sub">달성률</th>
      <th class="th-sub">계획</th>
      <th class="th-sub">실적</th>
      <th class="th-sub">달성률</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="td-label" rowspan="3">공급전<br>(전)</td>
      <td class="td-sub-label">신규개발전</td>
      <td>{fmt(dev_annual)}</td>
      <td>{fmt(dev_plan_m)}</td>
      <td>{fmt(dev_actual_m)}</td>
      <td>{rate_html(dev_actual_m, dev_plan_m)}</td>
      <td>{fmt(dev_cum_plan_m)}</td>
      <td>{fmt(safe(df_biz, 21, 9) if df_biz is not None else None)}</td>
      <td>{rate_html(safe(df_biz, 21, 9) if df_biz is not None else 0, dev_cum_plan_m)}</td>
    </tr>
    <tr>
      <td class="td-sub-label">폐전</td>
      <td>-</td>
      <td>-</td>
      <td>{fmt(safe(df_biz, 7, 9) if df_biz is not None else None)}</td>
      <td>-</td>
      <td>-</td>
      <td>{fmt(safe(df_biz, 6, 9) if df_biz is not None else None)}</td>
      <td>-</td>
    </tr>
    <tr>
      <td class="td-sub-label">순증가</td>
      <td>-</td>
      <td>-</td>
      <td>{fmt(순증가_act_m)}</td>
      <td>-</td>
      <td>-</td>
      <td>{fmt((safe(df_biz, 21, 9) or 0) - (safe(df_biz, 6, 9) or 0) if df_biz is not None else None)}</td>
      <td>-</td>
    </tr>
    <tr>
      <td class="td-label" rowspan="2">공급량<br>(GJ)</td>
      <td class="td-sub-label">신규개발량</td>
      <td>-</td>
      <td>{fmt(vol_plan_m)}</td>
      <td>-</td>
      <td>-</td>
      <td>{fmt(vol_cum_plan)}</td>
      <td>-</td>
      <td>-</td>
    </tr>
    <tr>
      <td class="td-sub-label">총공급량</td>
      <td>{fmt(vol_annual)}</td>
      <td>{fmt(vol_plan_m)}</td>
      <td>{fmt(total_actual_input) if total_actual_input > 0 else "입력필요"}</td>
      <td>{rate_html(total_actual_input, vol_plan_m) if total_actual_input > 0 else "-"}</td>
      <td>{fmt(vol_cum_plan)}</td>
      <td>{fmt(total_cum_calc) if total_cum_calc > 0 else "입력필요"}</td>
      <td>{rate_html(total_cum_calc, vol_cum_plan) if total_cum_calc > 0 else "-"}</td>
    </tr>
  </tbody>
</table>
"""
st.markdown(rows_html, unsafe_allow_html=True)

# ── 표2: 신규개발전 상세 현황 ────────────────────────────
st.markdown(f'<div class="box-title">📋 {sel_month}월 신규개발전 상세 현황</div>',
            unsafe_allow_html=True)
st.markdown('<div class="note-right">(단위 : 전)</div>', unsafe_allow_html=True)

# 상세표: 행=계획/실적/달성률/증감, 열=공동주택/단독주택/소계/일반용/업무용/산업용/열병합/합계
# 공동주택/단독주택 계획: 공급전계획 시트 row18,19 col(4+m)
공동_plan = safe(df_plan, 18, 4+m) if df_plan is not None else None
단독_plan = safe(df_plan, 19, 4+m) if df_plan is not None else None
소계_plan = (공동_plan or 0) + (단독_plan or 0)
일반_plan = safe(df_plan, 5,  4+m) if df_plan is not None else None
업무_plan = safe(df_plan, 6,  4+m) if df_plan is not None else None
산업_plan = safe(df_plan, 7,  4+m) if df_plan is not None else None
열병_plan = safe(df_plan, 8,  4+m) if df_plan is not None else None
합계_plan = safe(df_plan, 9,  4+m) if df_plan is not None else None

공동_act = act.get('공동주택'); 단독_act = act.get('단독주택')
소계_act = act.get('소계');     일반_act = act.get('일반용')
업무_act = act.get('업무용');   산업_act = act.get('산업용')
열병_act = act.get('열병합');   합계_act = act.get('합계')

# 누계 계획 (공급전계획 row22~27)
공동_cum_plan = safe(df_plan, 22, 4+m) if df_plan is not None else None
단독_cum_plan = safe(df_plan, 23, 4+m) if df_plan is not None else None
소계_cum_plan = (공동_cum_plan or 0) + (단독_cum_plan or 0)
일반_cum_plan = safe(df_plan, 24, 4+m) if df_plan is not None else None
업무_cum_plan = safe(df_plan, 25, 4+m) if df_plan is not None else None
산업_cum_plan = safe(df_plan, 26, 4+m) if df_plan is not None else None
열병_cum_plan = safe(df_plan, 27, 4+m) if df_plan is not None else None
합계_cum_plan = dev_cum_plan_m

# 누계 실적 (영업현황 row21)
공급전_누계_실적 = safe(df_biz, 21, 9) if df_biz is not None else None

def inc(a, p):
    try: return fmt((a or 0) - (p or 0))
    except: return "-"

def d_rate(a, p):
    try:
        val = float(a or 0) / float(p or 1) * 100
        s = f"{val:.1f}%"
        return r(s) if val < 100 else f"<b>{s}</b>"
    except: return "-"

detail_html = f"""
<table>
  <thead>
    <tr>
      <th rowspan="2">구 분</th>
      <th colspan="3">주택용</th>
      <th rowspan="2">일반용</th>
      <th rowspan="2">업무용</th>
      <th rowspan="2">산업용</th>
      <th rowspan="2">열병합</th>
      <th rowspan="2">합계</th>
    </tr>
    <tr>
      <th class="th-sub">공동주택</th>
      <th class="th-sub">단독주택</th>
      <th class="th-sub">소계</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="td-label">계획</td>
      <td>{fmt(공동_plan)}</td><td>{fmt(단독_plan)}</td><td>{fmt(소계_plan)}</td>
      <td>{fmt(일반_plan)}</td><td>{fmt(업무_plan)}</td><td>{fmt(산업_plan)}</td>
      <td>{fmt(열병_plan)}</td><td>{fmt(합계_plan)}</td>
    </tr>
    <tr>
      <td class="td-label">실적</td>
      <td>{fmt(공동_act)}</td><td>{fmt(단독_act)}</td><td>{fmt(소계_act)}</td>
      <td>{fmt(일반_act)}</td><td>{fmt(업무_act)}</td><td>{fmt(산업_act)}</td>
      <td>{fmt(열병_act)}</td><td>{fmt(합계_act)}</td>
    </tr>
    <tr>
      <td class="td-label">달성률</td>
      <td>{d_rate(공동_act,공동_plan)}</td><td>{d_rate(단독_act,단독_plan)}</td><td>{d_rate(소계_act,소계_plan)}</td>
      <td>{d_rate(일반_act,일반_plan)}</td><td>{d_rate(업무_act,업무_plan)}</td><td>{d_rate(산업_act,산업_plan)}</td>
      <td>{d_rate(열병_act,열병_plan)}</td><td>{d_rate(합계_act,합계_plan)}</td>
    </tr>
    <tr>
      <td class="td-label">증감</td>
      <td>{inc(공동_act,공동_plan)}</td><td>{inc(단독_act,단독_plan)}</td><td>{inc(소계_act,소계_plan)}</td>
      <td>{inc(일반_act,일반_plan)}</td><td>{inc(업무_act,업무_plan)}</td><td>{inc(산업_act,산업_plan)}</td>
      <td>{inc(열병_act,열병_plan)}</td><td>{inc(합계_act,합계_plan)}</td>
    </tr>
  </tbody>
</table>
<p style="font-size:12px;color:#555;margin-top:4px;">※ (괄호)는 누계 기준임.</p>
"""
st.markdown(detail_html, unsafe_allow_html=True)

# ── 산업용 신규 업체 ────────────────────────────────────
if not industry_df.empty:
    st.markdown(f'<div class="box-title">🏭 {sel_month}월 산업용 신규 업체 현황</div>',
                unsafe_allow_html=True)
    total_m3 = industry_df['월사용예정량'].sum()
    total_gj = industry_df['월간개발량'].sum()

    ind_rows = ""
    for _, row in industry_df.iterrows():
        공일 = str(row['공급일'])[:10] if pd.notna(row['공급일']) else "-"
        ind_rows += f"""
        <tr>
          <td>{row['신청명']}</td>
          <td>{row['업종']}</td>
          <td>{fmt(row['월사용예정량'])} ㎥</td>
          <td>{fmt(row['월간개발량'],2)} GJ</td>
          <td>{공일}</td>
          <td style="font-size:11px;text-align:left">{row['주소']}</td>
        </tr>"""

    ind_html = f"""
    <table>
      <thead>
        <tr>
          <th>업체명</th><th>업종</th>
          <th>월사용예정량</th><th>열량(GJ)</th>
          <th>공급일</th><th>주소</th>
        </tr>
      </thead>
      <tbody>
        {ind_rows}
        <tr style="background:#e8f4fd;font-weight:700;">
          <td colspan="2">합 계</td>
          <td>{fmt(total_m3)} ㎥</td>
          <td>{fmt(total_gj,2)} GJ</td>
          <td colspan="2"></td>
        </tr>
      </tbody>
    </table>"""
    st.markdown(ind_html, unsafe_allow_html=True)
elif f_dev:
    st.info(f"✅ {sel_month}월 신규 산업용 업체 없음")

# ── PDF 출력 안내 ────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style='text-align:center;padding:12px;color:#555;font-size:13px;'>
    🖨️ <b>PDF 출력</b>: Ctrl+P → 대상: PDF로 저장 → 레이아웃: 가로 → 저장
</div>
""", unsafe_allow_html=True)
