import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide", page_title="마케팅본부 월간 보고서")

st.markdown("""
<style>
* { font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; }
.report-header {
    background: linear-gradient(135deg, #1a3a6b 0%, #2d5fa8 100%);
    color: white; padding: 16px 28px; font-size: 22px; font-weight: 900;
    margin-bottom: 20px; border-radius: 6px;
}
.box-title {
    display: inline-block; border: 2px solid #c0392b; border-radius: 20px;
    padding: 4px 18px; font-size: 15px; font-weight: 800;
    color: #c0392b; margin: 20px 0 10px 0;
}
.note-right { font-size: 12px; color: #555; text-align: right; margin-bottom: 4px; }
table { border-collapse: collapse; width: 100%; font-size: 13px; margin-bottom: 12px; }
th, td { border: 1px solid #aaa; padding: 6px 10px; text-align: center; }
thead th { background: #1e3a6b; color: white; font-weight: 700; }
.th-sub { background: #2d5fa8 !important; }
.td-label { background: #dce6f5; font-weight: 700; }
.td-sub-label { background: #eef2fa; font-weight: 600; }
.rate-red { color: #c0392b; font-weight: 700; }
.rate-ok  { color: #1a7a1a; font-weight: 700; }
.upload-box {
    background: #eef4fc; border: 1px solid #2d5fa8;
    border-radius: 6px; padding: 10px 16px; margin-bottom: 14px; font-size: 13px;
}
@media print {
    .stButton, .stFileUploader, .stSelectbox, .stNumberInput,
    .upload-box, header, [data-testid='stSidebar'] { display:none!important; }
}
</style>
""", unsafe_allow_html=True)

# ── 헬퍼 ─────────────────────────────────────────────
def fmt(v, dec=0):
    if v is None or (isinstance(v, float) and pd.isna(v)): return "-"
    try:
        f = float(v)
        return f"{f:,.{dec}f}" if dec > 0 else f"{int(round(f)):,}"
    except: return str(v)

def rate_html(a, p):
    try:
        val = float(a or 0) / float(p or 1) * 100
        s = f"{val:.1f}%"
        cls = "rate-ok" if val >= 100 else "rate-red"
        return f'<span class="{cls}">{s}</span>'
    except: return "-"

def safe(df, r, c):
    try:
        v = df.iloc[r, c]
        return None if (isinstance(v, float) and pd.isna(v)) else v
    except: return None

def inc(a, p):
    try: return fmt((float(a or 0)) - (float(p or 0)))
    except: return "-"

# ── 캐시 로드 ─────────────────────────────────────────
@st.cache_data
def load_xlsm(b, fn):
    buf = io.BytesIO(b)
    out = {}
    for s in ['공급전 계획(2026년)', '공급량 계획_MJ(2026년)', '영업현황']:
        try: out[s] = pd.read_excel(buf, sheet_name=s, header=None, engine='openpyxl'); buf.seek(0)
        except: pass
    return out

@st.cache_data
def load_new2(b, fn):
    buf = io.BytesIO(b)
    out = {}
    for s in ['3_1. 개발량 계획']:
        try: out[s] = pd.read_excel(buf, sheet_name=s, header=None); buf.seek(0)
        except: pass
    return out

@st.cache_data
def load_dev_detail(b, fn):
    buf = io.BytesIO(b)
    try:
        df = pd.read_excel(buf, sheet_name='Sheet1', header=0)
        df.columns = ['공급신청번호','시공업체','번지순번','주소','신청명','계약구분','건물구분',
                      '신청일','용도','업종','상품','등급','월사용예정량','공급승인일','공급일',
                      '서비스센터','설치장소주소','계량기번호','특정여부','월간개발량','공동주택명']
        return df
    except: return None

# ════════════════════════════════════════════════════
# UI
# ════════════════════════════════════════════════════
st.markdown('<div class="report-header">📊 마케팅본부 _ 월간 영업현황 보고서</div>',
            unsafe_allow_html=True)

st.markdown("""
<div class="upload-box">
📁 <b>매월 3개 파일 업로드</b> + 총공급량 수기 입력 → 자동 보고서 생성<br>
① 영업일보_월간_YYYYMM.xlsm &nbsp;|&nbsp; ② new_2_신규개발량.xlsx &nbsp;|&nbsp; ③ 신규개발량_YYYYMM.xlsx
</div>
""", unsafe_allow_html=True)

col_u1, col_u2, col_u3 = st.columns(3)
with col_u1: f_xlsm = st.file_uploader("① 영업일보 (xlsm)", type=['xlsm','xlsx'])
with col_u2: f_new2 = st.file_uploader("② new_2 신규개발량 (xlsx)", type=['xlsx'])
with col_u3: f_dev  = st.file_uploader("③ 신규개발량 상세 (xlsx)", type=['xlsx'])

# ── 수기 입력 + 확인 버튼 ──
st.markdown("---")
col_s1, col_s2, col_s3, col_s4 = st.columns([1,2,2,1])
with col_s1: sel_month = st.selectbox("보고 월", list(range(1,13)), index=5)
with col_s2: total_actual = st.number_input("총공급량 당월 실적 (GJ)", value=0, step=1000, format="%d",
                                             help="고객지원시스템에서 확인한 값 입력")
with col_s3: prev_cum    = st.number_input("전월 총공급량 누계 실적 (GJ)", value=0, step=1000, format="%d",
                                            help="직전 월 누계 실적값 입력 → 당월누계 자동계산")
with col_s4:
    st.markdown("<br>", unsafe_allow_html=True)
    confirmed = st.button("✅ 확인 / 적용", use_container_width=True)

total_cum = prev_cum + total_actual
if confirmed:
    st.success(f"✅ 총공급량 당월 실적: **{total_actual:,} GJ** | 누계 실적: **{total_cum:,} GJ** 적용됨")

if not f_xlsm:
    st.info("👆 영업일보 파일을 먼저 업로드해주세요.")
    st.stop()

# ── 데이터 파싱 ────────────────────────────────────────
sheets = load_xlsm(f_xlsm.read(), f_xlsm.name); f_xlsm.seek(0)
df_plan = sheets.get('공급전 계획(2026년)')
df_vol  = sheets.get('공급량 계획_MJ(2026년)')
df_biz  = sheets.get('영업현황')

new2_sheets, dev_detail = None, None
if f_new2:
    new2_sheets = load_new2(f_new2.read(), f_new2.name); f_new2.seek(0)
if f_dev:
    dev_detail = load_dev_detail(f_dev.read(), f_dev.name); f_dev.seek(0)

df_n2 = new2_sheets.get('3_1. 개발량 계획') if new2_sheets else None
m = sel_month - 1  # 0-based (col offset: col2=1월, col3=2월, ...)
mc = m + 2  # new2 시트 기준 컬럼 인덱스

# ════════════════════════════════════════
# 데이터 추출
# ════════════════════════════════════════

# ── new_1 기준 (연간계획, 당월계획, 당월실적, 누계계획, 누계실적) ──
# row6=신규개발전, row7=폐전, row8=순증가, row9=신규개발량, row10=총공급량
# col2=연간계획, col3=당월계획, col4=당월실적, col5=달성률, col6=누계계획, col7=누계실적, col8=누계달성률

def get_new1_row(row_idx):
    """영업일보 공급전 계획 시트에서 행별 추출"""
    if df_plan is None: return {}
    return {
        '연간': safe(df_plan, row_idx, 17),  # col17=년간계
        '당월계획': safe(df_plan, row_idx, 4+m),
        '당월누계계획': safe(df_plan, row_idx+8, 4+m) if row_idx < 10 else None,
    }

# new_1 파일의 회의자료 시트 직접 사용
try:
    df_rpt = pd.read_excel(f_xlsm if not hasattr(f_xlsm,'seek') else f_xlsm,
                           sheet_name='(회의자료 입력용)공급전 및 공급량 현황', header=None)
    f_xlsm.seek(0)
except:
    df_rpt = None

def rpt(r, c):
    return safe(df_rpt, r, c) if df_rpt is not None else None

# new_1 회의자료 시트 위치:
# row6=신규개발전: col2=연간, col3=당월계획, col4=당월실적, col5=달성률, col6=누계계획, col7=누계실적
# row7=폐전, row8=순증가, row9=신규개발량, row10=총공급량
신규_연간  = rpt(6,2);  신규_당계  = rpt(6,3);  신규_당실 = rpt(6,4)
신규_누계  = rpt(6,6);  신규_누실  = rpt(6,7)
폐전_연간  = rpt(7,2);  폐전_당계  = rpt(7,3);  폐전_당실 = rpt(7,4)
폐전_누계  = rpt(7,6);  폐전_누실  = rpt(7,7)
순증_연간  = rpt(8,2);  순증_당계  = rpt(8,3);  순증_당실 = rpt(8,4)
순증_누계  = rpt(8,6);  순증_누실  = rpt(8,7)
개발량_연간= rpt(9,2);  개발량_당계= rpt(9,3);  개발량_당실= rpt(9,4)
개발량_누계= rpt(9,6);  개발량_누실= rpt(9,7)
총공_연간  = rpt(10,2); 총공_당계  = rpt(10,3)
총공_누계  = rpt(10,6)

# ── 영업현황 시트 신규개발전 실적/목표 ──
cats_biz = ['공동주택','단독주택','소계','일반용','업무용','산업용','열병합','합계']
act, pln = {}, {}
if df_biz is not None:
    for i, cat in enumerate(cats_biz):
        act[cat] = safe(df_biz, 25, i+2)
        pln[cat] = safe(df_biz, 26, i+2)

# ── new_2 파일: 폐전/순증가 연간계획, 상세 월별 계획 ──
# 폐전 합계: row34, col2~13 / 순증가=신규개발전-폐전
# 신규개발전 합계 계획: row15, col(mc)
# 폐전 합계 계획: row34, col(mc)
신규_당계_n2  = safe(df_n2, 15, mc) if df_n2 is not None else None
신규_누계_n2  = safe(df_n2, 16, mc) if df_n2 is not None else None
신규_연간_n2  = safe(df_n2, 15, 14) if df_n2 is not None else None
폐전_당계_n2  = safe(df_n2, 34, mc) if df_n2 is not None else None
폐전_누계_n2  = safe(df_n2, 35, mc) if df_n2 is not None else None
폐전_연간_n2  = safe(df_n2, 34, 14) if df_n2 is not None else None
순증_당계_n2  = (신규_당계_n2 or 0) - (폐전_당계_n2 or 0)
순증_누계_n2  = (신규_누계_n2 or 0) - (폐전_누계_n2 or 0)
순증_연간_n2  = (신규_연간_n2 or 0) - (폐전_연간_n2 or 0)
# 신규개발량 계획 (MJ→GJ): row103=합계, row104=누계
개발량_당계_n2 = safe(df_n2, 103, mc)
개발량_누계_n2 = safe(df_n2, 104, mc)
개발량_연간_n2 = safe(df_n2, 103, 14)
if 개발량_당계_n2:  개발량_당계_n2  /= 1000
if 개발량_누계_n2:  개발량_누계_n2  /= 1000
if 개발량_연간_n2:  개발량_연간_n2  /= 1000

# new_2 상세 (신규개발전 당월 계획 by 용도): row15=합계, 개별 row2~13
dev_cats_idx = {'공동주택':2,'단독주택':4,'가정용':6,'일반용':7,'업무용':9,'산업용':11,'열병합용':13,'합계':15}
dev_plan_m = {}
dev_cum_plan_m = {}
if df_n2 is not None:
    for cat, ri in dev_cats_idx.items():
        dev_plan_m[cat]     = safe(df_n2, ri,   mc)
        dev_cum_plan_m[cat] = safe(df_n2, ri+1, mc) if ri < 15 else safe(df_n2, 16, mc)

# 공급량 계획 (영업일보 공급량 계획 시트)
vol_plan_m   = safe(df_vol, 5, m+3)/1000 if df_vol is not None and safe(df_vol,5,m+3) else None
vol_cum_plan = safe(df_vol, 6, m+3)/1000 if df_vol is not None and safe(df_vol,6,m+3) else None
vol_annual   = safe(df_vol, 5, 16)/1000  if df_vol is not None and safe(df_vol,5,16)  else None

# 우선순위: new_1 회의자료 > new_2 > 영업일보
def pick(*args):
    for v in args:
        if v is not None and not (isinstance(v, float) and pd.isna(v)):
            try:
                if float(v) != 0: return v
            except: return v
    return None

# ════════════════════════════════════════
# 화면 출력
# ════════════════════════════════════════
st.markdown(f"### 📅 **2026년 {sel_month}월** 영업현황 보고")

# ── 표1: 공급전 및 공급량 현황 ──
st.markdown('<div class="box-title">📋 공급전 및 공급량 현황</div>', unsafe_allow_html=True)

def tr(label1, label2, 연간, 당계, 당실, 당달, 누계, 누실, 누달):
    rowspan = ' rowspan="3"' if label1 else ''
    label1_td = f'<td class="td-label"{rowspan}>{label1}</td>' if label1 else ''
    return f"""
    <tr>
      {label1_td}
      <td class="td-sub-label">{label2}</td>
      <td>{fmt(연간)}</td>
      <td>{fmt(당계)}</td>
      <td>{fmt(당실)}</td>
      <td>{당달}</td>
      <td>{fmt(누계)}</td>
      <td>{fmt(누실)}</td>
      <td>{누달}</td>
    </tr>"""

# 값 결정
신규_연간_f  = pick(신규_연간, 신규_연간_n2)
신규_당계_f  = pick(신규_당계, 신규_당계_n2, dev_plan_m.get('합계'))
신규_당실_f  = pick(신규_당실, act.get('합계'))
신규_누계_f  = pick(신규_누계, 신규_누계_n2, dev_cum_plan_m.get('합계'))
신규_누실_f  = pick(신규_누실)

폐전_연간_f  = pick(폐전_연간, 폐전_연간_n2)
폐전_당계_f  = pick(폐전_당계, 폐전_당계_n2)
폐전_당실_f  = pick(폐전_당실, safe(df_biz, 7, 9) if df_biz is not None else None)
폐전_누계_f  = pick(폐전_누계, 폐전_누계_n2)
폐전_누실_f  = pick(폐전_누실, safe(df_biz, 6, 9) if df_biz is not None else None)

순증_연간_f  = pick(순증_연간, 순증_연간_n2)
순증_당계_f  = pick(순증_당계, 순증_당계_n2)
순증_당실_f  = pick(순증_당실, (float(신규_당실_f or 0) - float(폐전_당실_f or 0)) if 신규_당실_f else None)
순증_누계_f  = pick(순증_누계, 순증_누계_n2)
순증_누실_f  = pick(순증_누실, (float(신규_누실_f or 0) - float(폐전_누실_f or 0)) if 신규_누실_f else None)

개발량_연간_f = pick(개발량_연간, 개발량_연간_n2, vol_annual)
개발량_당계_f = pick(개발량_당계, 개발량_당계_n2, vol_plan_m)
개발량_당실_f = pick(개발량_당실)
개발량_누계_f = pick(개발량_누계, 개발량_누계_n2, vol_cum_plan)
개발량_누실_f = pick(개발량_누실)

총공_연간_f  = pick(총공_연간, vol_annual)
총공_당계_f  = pick(총공_당계, vol_plan_m)
총공_당실_f  = total_actual if total_actual > 0 else None
총공_누계_f  = pick(총공_누계, vol_cum_plan)
총공_누실_f  = total_cum if total_cum > 0 else None

html1 = f"""
<table>
  <thead>
    <tr>
      <th colspan="2" rowspan="2">구 분</th>
      <th rowspan="2">연간계획</th>
      <th colspan="3">{sel_month}월 (당월)</th>
      <th colspan="3">{sel_month}월 (누계)</th>
    </tr>
    <tr>
      <th class="th-sub">계획</th><th class="th-sub">실적</th><th class="th-sub">달성률</th>
      <th class="th-sub">계획</th><th class="th-sub">실적</th><th class="th-sub">달성률</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="td-label" rowspan="3">공급전<br>(전)</td>
      <td class="td-sub-label">신규개발전</td>
      <td>{fmt(신규_연간_f)}</td><td>{fmt(신규_당계_f)}</td><td>{fmt(신규_당실_f)}</td>
      <td>{rate_html(신규_당실_f, 신규_당계_f)}</td>
      <td>{fmt(신규_누계_f)}</td><td>{fmt(신규_누실_f)}</td>
      <td>{rate_html(신규_누실_f, 신규_누계_f)}</td>
    </tr>
    <tr>
      <td class="td-sub-label">폐전</td>
      <td>{fmt(폐전_연간_f)}</td><td>{fmt(폐전_당계_f)}</td><td>{fmt(폐전_당실_f)}</td>
      <td>{rate_html(폐전_당실_f, 폐전_당계_f)}</td>
      <td>{fmt(폐전_누계_f)}</td><td>{fmt(폐전_누실_f)}</td>
      <td>{rate_html(폐전_누실_f, 폐전_누계_f)}</td>
    </tr>
    <tr>
      <td class="td-sub-label">순증가</td>
      <td>{fmt(순증_연간_f)}</td><td>{fmt(순증_당계_f)}</td><td>{fmt(순증_당실_f)}</td>
      <td>{rate_html(순증_당실_f, 순증_당계_f)}</td>
      <td>{fmt(순증_누계_f)}</td><td>{fmt(순증_누실_f)}</td>
      <td>{rate_html(순증_누실_f, 순증_누계_f)}</td>
    </tr>
    <tr>
      <td class="td-label" rowspan="2">공급량<br>(GJ)</td>
      <td class="td-sub-label">신규개발량</td>
      <td>{fmt(개발량_연간_f)}</td><td>{fmt(개발량_당계_f)}</td><td>{fmt(개발량_당실_f)}</td>
      <td>{rate_html(개발량_당실_f, 개발량_당계_f)}</td>
      <td>{fmt(개발량_누계_f)}</td><td>{fmt(개발량_누실_f)}</td>
      <td>{rate_html(개발량_누실_f, 개발량_누계_f)}</td>
    </tr>
    <tr>
      <td class="td-sub-label">총공급량</td>
      <td>{fmt(총공_연간_f)}</td><td>{fmt(총공_당계_f)}</td>
      <td>{'<b>'+fmt(총공_당실_f)+'</b>' if 총공_당실_f else '<span style="color:#888">입력필요</span>'}</td>
      <td>{rate_html(총공_당실_f, 총공_당계_f) if 총공_당실_f else '-'}</td>
      <td>{fmt(총공_누계_f)}</td>
      <td>{'<b>'+fmt(총공_누실_f)+'</b>' if 총공_누실_f else '<span style="color:#888">입력필요</span>'}</td>
      <td>{rate_html(총공_누실_f, 총공_누계_f) if 총공_누실_f else '-'}</td>
    </tr>
  </tbody>
</table>"""
st.markdown(html1, unsafe_allow_html=True)

# ── 표2: 신규개발전 상세 현황 ──
st.markdown(f'<div class="box-title">📋 {sel_month}월 신규개발전 상세 현황</div>', unsafe_allow_html=True)
st.markdown('<div class="note-right">(단위 : 전)</div>', unsafe_allow_html=True)

# 상세 계획 (new_2 기준)
공동_p = dev_plan_m.get('공동주택');  단독_p = dev_plan_m.get('단독주택')
소계_p = (공동_p or 0)+(단독_p or 0); 일반_p = dev_plan_m.get('일반용')
업무_p = dev_plan_m.get('업무용');    산업_p = dev_plan_m.get('산업용')
열병_p = dev_plan_m.get('열병합용');  합계_p = dev_plan_m.get('합계') or pick(신규_당계, 신규_당계_n2)

공동_a = act.get('공동주택'); 단독_a = act.get('단독주택')
소계_a = act.get('소계');    일반_a = act.get('일반용')
업무_a = act.get('업무용');  산업_a = act.get('산업용')
열병_a = act.get('열병합');  합계_a = act.get('합계')

# 누계 계획 (new_2 row16=합계누계, dev_cum_plan_m)
공동_cp = dev_cum_plan_m.get('공동주택'); 단독_cp = dev_cum_plan_m.get('단독주택')
소계_cp = (공동_cp or 0)+(단독_cp or 0); 일반_cp = dev_cum_plan_m.get('일반용')
업무_cp = dev_cum_plan_m.get('업무용');  산업_cp = dev_cum_plan_m.get('산업용')
열병_cp = dev_cum_plan_m.get('열병합용'); 합계_cp = dev_cum_plan_m.get('합계') or pick(신규_누계, 신규_누계_n2)

def d_inc(a, p):
    try: return fmt(float(a or 0)-float(p or 0))
    except: return "-"
def d_rate(a, p): return rate_html(a, p)

html2 = f"""
<table>
  <thead>
    <tr>
      <th rowspan="2">구 분</th>
      <th colspan="3">주택용</th>
      <th rowspan="2">일반용</th><th rowspan="2">업무용</th>
      <th rowspan="2">산업용</th><th rowspan="2">열병합</th><th rowspan="2">합계</th>
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
      <td>{fmt(공동_p)}</td><td>{fmt(단독_p)}</td><td>{fmt(소계_p)}</td>
      <td>{fmt(일반_p)}</td><td>{fmt(업무_p)}</td><td>{fmt(산업_p)}</td>
      <td>{fmt(열병_p)}</td><td>{fmt(합계_p)}</td>
    </tr>
    <tr>
      <td class="td-label">실적</td>
      <td>{fmt(공동_a)}</td><td>{fmt(단독_a)}</td><td>{fmt(소계_a)}</td>
      <td>{fmt(일반_a)}</td><td>{fmt(업무_a)}</td><td>{fmt(산업_a)}</td>
      <td>{fmt(열병_a)}</td><td>{fmt(합계_a)}</td>
    </tr>
    <tr>
      <td class="td-label">달성률</td>
      <td>{d_rate(공동_a,공동_p)}</td><td>{d_rate(단독_a,단독_p)}</td><td>{d_rate(소계_a,소계_p)}</td>
      <td>{d_rate(일반_a,일반_p)}</td><td>{d_rate(업무_a,업무_p)}</td><td>{d_rate(산업_a,산업_p)}</td>
      <td>{d_rate(열병_a,열병_p)}</td><td>{d_rate(합계_a,합계_p)}</td>
    </tr>
    <tr>
      <td class="td-label">증감</td>
      <td>{d_inc(공동_a,공동_p)}</td><td>{d_inc(단독_a,단독_p)}</td><td>{d_inc(소계_a,소계_p)}</td>
      <td>{d_inc(일반_a,일반_p)}</td><td>{d_inc(업무_a,업무_p)}</td><td>{d_inc(산업_a,산업_p)}</td>
      <td>{d_inc(열병_a,열병_p)}</td><td>{d_inc(합계_a,합계_p)}</td>
    </tr>
  </tbody>
</table>
<p style="font-size:12px;color:#555;">※ (괄호)는 누계 기준임.</p>"""
st.markdown(html2, unsafe_allow_html=True)

# ── 산업용 신규 업체 ──
industry_df = pd.DataFrame()
if dev_detail is not None:
    industry_df = dev_detail[dev_detail['용도'].astype(str).str.contains('산업', na=False)].copy()

if not industry_df.empty:
    st.markdown(f'<div class="box-title">🏭 {sel_month}월 산업용 신규 업체 현황</div>',
                unsafe_allow_html=True)
    total_m3 = industry_df['월사용예정량'].sum()
    total_gj = industry_df['월간개발량'].sum()
    ind_rows = ""
    for _, row in industry_df.iterrows():
        공일 = str(row['공급일'])[:10] if pd.notna(row['공급일']) else "-"
        ind_rows += f"""<tr>
          <td>{row['신청명']}</td><td>{row['업종']}</td>
          <td>{fmt(row['월사용예정량'])} ㎥</td>
          <td>{fmt(row['월간개발량'],2)} GJ</td>
          <td>{공일}</td>
          <td style="font-size:11px;text-align:left">{row['주소']}</td></tr>"""
    st.markdown(f"""
    <table>
      <thead><tr>
        <th>업체명</th><th>업종</th><th>월사용예정량</th>
        <th>열량(GJ)</th><th>공급일</th><th>주소</th>
      </tr></thead>
      <tbody>
        {ind_rows}
        <tr style="background:#dce6f5;font-weight:700;">
          <td colspan="2">합 계</td>
          <td>{fmt(total_m3)} ㎥</td><td>{fmt(total_gj,2)} GJ</td>
          <td colspan="2"></td>
        </tr>
      </tbody>
    </table>""", unsafe_allow_html=True)
elif f_dev:
    st.info(f"✅ {sel_month}월 신규 산업용 업체 없음")
else:
    st.warning("③ 신규개발량 상세 파일을 업로드하면 산업용 업체 현황이 표시됩니다.")

# ── PDF 출력 ──
st.markdown("---")
st.markdown("""
<div style='text-align:center;padding:12px;color:#555;font-size:13px;'>
🖨️ <b>PDF 출력</b>: Ctrl+P → 대상: PDF로 저장 → 레이아웃: 가로 → 저장
</div>""", unsafe_allow_html=True)
