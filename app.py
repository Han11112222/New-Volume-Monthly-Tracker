import streamlit as st
import pandas as pd
import io
import urllib.request
import urllib.parse

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
/* 헤더 */
thead th {
    background: #1e3a6b; color: white !important; font-weight: 700;
    border: 1px solid #2d5fa8; padding: 7px 10px;
    text-align: center !important; white-space: nowrap;
    font-size: 13px; vertical-align: middle;
}
thead th.th-sub { background: #2d5fa8 !important; }
/* 데이터 셀 */
td {
    border: 1px solid #ddd; padding: 7px 10px;
    text-align: center; white-space: nowrap; font-size: 13px;
    color: #222 !important; vertical-align: middle;
    background: #ffffff; font-weight: normal;
}
/* 숫자 기본: 검정, 볼드 없음 */
td span, td small { color: #222 !important; font-size: 12px !important; font-weight: normal !important; }
/* 실적만 bold */
td b { color: #222 !important; font-weight: 700 !important; }
/* 달성률 색상만 유지 */
.rate-red { color: #c0392b !important; font-weight: 700 !important; }
.rate-ok  { color: #1a7a1a !important; font-weight: 700 !important; }
/* 구분 라벨 - 파란배경 흰글씨 */
.td-label {
    background: #dce6f5 !important; font-weight: 700;
    color: #1e3a6b !important; text-align: center !important;
}
/* 소분류 라벨 - 연한배경 */
.td-sub-label {
    background: #eef2fa !important; font-weight: 600;
    color: #333 !important; text-align: center !important;
}
/* 짝수행 줄무늬 */
tr:nth-child(even) td { background: #f8f9fc; }
tr:nth-child(odd)  td { background: #ffffff; }
/* 라벨행은 줄무늬 적용 안 함 */
tr td.td-label, tr td.td-sub-label { background: inherit; }
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
    .upload-box, header, [data-testid='stSidebar'] { display: none !important; }
}
</style>
""", unsafe_allow_html=True)

# ── 헬퍼 ──────────────────────────────────────────────
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

def d_inc(a, p):
    try: return fmt(float(a or 0) - float(p or 0))
    except: return "-"

def cell(당월, 누계):
    # 당월: 일반 검정, 누계: small 검정
    return f"<span style='color:#222'>{fmt(당월)}</span><br><small style='color:#555;font-size:11px'>({fmt(누계)})</small>"

def cell_bold(당월, 누계):
    # 실적행: 당월 bold, 누계 small
    return f"<b style='color:#222'>{fmt(당월)}</b><br><small style='color:#555;font-size:11px'>({fmt(누계)})</small>"

def rate_cell(당월_a, 당월_p, 누계_a, 누계_p):
    return f"{rate_html(당월_a,당월_p)}<br><small style='font-size:11px'>({rate_html(누계_a,누계_p)})</small>"

def inc_cell(당월_a, 당월_p, 누계_a, 누계_p):
    return f"<span style='color:#222'>{d_inc(당월_a,당월_p)}</span><br><small style='color:#555;font-size:11px'>({d_inc(누계_a,누계_p)})</small>"

# ── GitHub 자동 로드 (계획 데이터) ──────────────────────
BASE = "https://raw.githubusercontent.com/Han11112222/New-Volume-Monthly-Tracker/main"

@st.cache_data(ttl=3600)
def load_github():
    errors = []
    result = {}
    files = {
        "new_1": {
            "fname": "new_1.(작성용)6월 영업현황 보고(20260626).xlsx",
            "sheets": ["3_1. 개발량 계획",
                       "(회의자료 입력용)공급전 및 공급량 현황"]
        },
        "new_2": {
            "fname": "new_2.(통합)신규개발량(202606)_5월 확정공급량적용_20260626.xlsx",
            "sheets": ["3_1. 개발량 계획", "3_2. 개발량 실적"]
        }
    }
    for key, info in files.items():
        try:
            url = f"{BASE}/{urllib.parse.quote(info['fname'])}"
            req = urllib.request.urlopen(url, timeout=15)
            buf = io.BytesIO(req.read())
            result[key] = {}
            for s in info["sheets"]:
                try:
                    result[key][s] = pd.read_excel(buf, sheet_name=s, header=None)
                    buf.seek(0)
                except: pass
        except Exception as e:
            errors.append(f"{key}: {e}")
    return result, errors

# ── 구글 스프레드시트 총공급량 ──────────────────────────
GSHEET_URL = "https://docs.google.com/spreadsheets/d/13HrIz6OytYDykXeXzXJ02I6XbaKin1YaKBoO2kBd6Bs/export?format=csv&gid=0"

@st.cache_data(ttl=300, show_spinner=False)
def load_supply_gsheet():
    try:
        req = urllib.request.Request(GSHEET_URL,
              headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            raw = resp.read()
        df = pd.read_csv(io.BytesIO(raw), header=0)
        df = df.iloc[:, :6]
        df.columns = ['일자','공급량_MJ','공급량_M3','평균기온','최저기온','최고기온']
        df['일자'] = pd.to_datetime(df['일자'], errors='coerce')
        df = df.dropna(subset=['일자'])
        df['공급량_MJ'] = pd.to_numeric(
            df['공급량_MJ'].astype(str).str.replace(',','').str.replace(' ',''),
            errors='coerce')
        df = df.dropna(subset=['공급량_MJ'])
        df['공급량_GJ'] = df['공급량_MJ'] / 1000
        df['연'] = df['일자'].dt.year
        df['월'] = df['일자'].dt.month
        return df.groupby(['연','월'])['공급량_GJ'].sum().reset_index(), None
    except Exception as e:
        return None, str(e)

def get_gj(mdf, year, month, cum=False):
    if mdf is None: return None
    if cum:
        f = mdf[(mdf['연']==year) & (mdf['월']<=month)]
        return float(f['공급량_GJ'].sum()) if len(f)>0 else None
    r = mdf[(mdf['연']==year) & (mdf['월']==month)]
    return float(r['공급량_GJ'].values[0]) if len(r)>0 else None

# ── 영업일보 로드 ────────────────────────────────────────
@st.cache_data
def load_xlsm(b, fn):
    buf = io.BytesIO(b)
    out = {}
    for s in ['영업일보', '공급전 계획(2026년)', '공급량 계획_MJ(2026년)']:
        try:
            out[s] = pd.read_excel(buf, sheet_name=s, header=None, engine='openpyxl')
            buf.seek(0)
        except: pass
    return out

# ── 신규개발량 파일 로드 (산업용 업체) ──────────────────
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
📁 <b>매월 2개 파일만 업로드</b>하면 자동으로 보고서가 생성됩니다.<br>
&nbsp;&nbsp;① 영업일보_월간_YYYYMM.xlsm &nbsp;|&nbsp; ② 신규개발량_YYYYMM.xlsx (산업용 업체)<br>
&nbsp;&nbsp;📡 계획 데이터는 GitHub에서, 총공급량은 구글시트에서 <b>자동 로드</b>됩니다.
</div>
""", unsafe_allow_html=True)

col_u1, col_u2 = st.columns(2)
with col_u1: f_xlsm = st.file_uploader("① 영업일보 (xlsm)", type=['xlsm','xlsx'])
with col_u2: f_dev  = st.file_uploader("② 신규개발량 (xlsx) — 산업용 업체", type=['xlsx'])

st.markdown("---")
col_s1, col_s2, col_s3, col_s4 = st.columns([1,2,2,1])
with col_s1:
    sel_year  = st.number_input("보고 연도", value=2026, step=1, format="%d")
    sel_month = st.selectbox("보고 월", list(range(1,13)), index=5)

with st.spinner("📡 데이터 자동 로드 중..."):
    gh_data, gh_errors = load_github()
    monthly_df, gs_err = load_supply_gsheet()

auto_actual = get_gj(monthly_df, int(sel_year), sel_month)
auto_cum    = get_gj(monthly_df, int(sel_year), sel_month, cum=True)

c1, c2 = st.columns(2)
with c1:
    if gh_errors: st.warning("⚠️ GitHub: " + " | ".join(gh_errors))
    else: st.success("✅ GitHub 계획 데이터 자동 로드 완료")
with c2:
    if gs_err: st.warning(f"⚠️ 구글시트: {gs_err}")
    elif auto_actual: st.success(f"✅ 구글시트 자동합산: 당월 **{auto_actual:,.0f} GJ** | 누계 **{auto_cum:,.0f} GJ**")
    else: st.info(f"ℹ️ {int(sel_year)}년 {sel_month}월 구글시트 데이터 없음 → 수동 입력")

with col_s2:
    총공_당실 = st.number_input("총공급량 당월 실적 (GJ) 📡자동",
        value=int(round(auto_actual)) if auto_actual else 0,
        step=100, format="%d")
with col_s3:
    총공_누실 = st.number_input("총공급량 누계 실적 (GJ) 📡자동",
        value=int(round(auto_cum)) if auto_cum else 0,
        step=100, format="%d")
with col_s4:
    st.markdown("<br><br>", unsafe_allow_html=True)
    if st.button("✅ 확인/적용", use_container_width=True):
        st.success(f"✅ 당월 **{총공_당실:,}** | 누계 **{총공_누실:,}** GJ")

if not f_xlsm:
    st.info("👆 영업일보 파일을 업로드해주세요.")
    st.stop()

# ════════════════════════════════════════════════════
# 데이터 파싱
# ════════════════════════════════════════════════════
sheets = load_xlsm(f_xlsm.read(), f_xlsm.name); f_xlsm.seek(0)
dev_detail = load_dev_detail(f_dev.read(), f_dev.name) if f_dev else None

# 영업일보 시트 (★ 핵심 실적 데이터)
df_il   = sheets.get('영업일보')               # 모든 실적값
df_plan = sheets.get('공급전 계획(2026년)')    # 신규개발전 상세 계획
df_vol  = sheets.get('공급량 계획_MJ(2026년)') # 공급량 계획

# GitHub (계획)
df_n1p   = gh_data.get("new_1",{}).get("3_1. 개발량 계획")
df_n1rpt = gh_data.get("new_1",{}).get("(회의자료 입력용)공급전 및 공급량 현황")
df_n2p   = gh_data.get("new_2",{}).get("3_1. 개발량 계획")
df_n2r   = gh_data.get("new_2",{}).get("3_2. 개발량 실적")   # 신규개발량 누적 실적

m  = sel_month - 1
mc = m + 4   # 공급전계획: col4=1월
nc = m + 2   # new_1/2: col2=1월
vc = m + 3   # 공급량계획: col3=1월

def s(df, r, c): return safe(df, r, c) if df is not None else None

# ════════════════════════════════════════════════════
# ★ 영업일보 시트에서 실적/계획 직접 추출
# row7=합계, row8=공동, row9=단독, row10=영업용(일반), row11=업무용, row12=산업용, row13=열병합
# col2=전월말, col3=당월계획, col4=당월실적, col5=폐전당월, col9=당년계획, col10=당년실적
# col11=당년폐전누계, col14=누계(①+②-③+④)
# ════════════════════════════════════════════════════
rows_il = {'합계':7, '공동':8, '단독':9, '일반':10, '업무':11, '산업':12, '열병합':13}

def il(row_key, col):
    r = rows_il.get(row_key)
    return s(df_il, r, col) if r is not None else None

# 신규개발전 당월 실적
공동_a  = il('공동',4);  단독_a  = il('단독',4)
소계_a  = (공동_a or 0)+(단독_a or 0)
일반_a  = il('일반',4);  업무_a  = il('업무',4)
산업_a  = il('산업',4);  열병_a  = il('열병합',4)
합계_a  = il('합계',4)
신규_당실 = 합계_a

# 신규개발전 누계 실적 (당년실적)
공동_ca = il('공동',10); 단독_ca = il('단독',10)
소계_ca = (공동_ca or 0)+(단독_ca or 0)
일반_ca = il('일반',10); 업무_ca = il('업무',10)
산업_ca = il('산업',10); 열병_ca = il('열병합',10)
합계_ca = il('합계',10)
신규_누실 = 합계_ca

# 폐전 당월/누계 실적
폐전_당실 = il('합계',5)
폐전_누실 = il('합계',11)

# 순증가
순증_당실 = (float(신규_당실 or 0) - float(폐전_당실 or 0)) if 신규_당실 else None
순증_누실 = (float(신규_누실 or 0) - float(폐전_누실 or 0)) if 신규_누실 else None

# ════════════════════════════════════════════════════
# 계획값 (GitHub new_1 3_1 + 영업일보 공급전계획)
# ════════════════════════════════════════════════════
신규_연간 = s(df_n1p, 15, 14)
신규_당계 = s(df_n1p, 15, nc)
신규_누계 = s(df_n1p, 16, nc)

폐전_연간 = s(df_n1p, 34, 14)
폐전_당계 = s(df_n1p, 34, nc)
폐전_누계 = s(df_n1p, 35, nc)

순증_연간 = (신규_연간 or 0)-(폐전_연간 or 0)
순증_당계 = (신규_당계 or 0)-(폐전_당계 or 0)
순증_누계 = (신규_누계 or 0)-(폐전_누계 or 0)

# 신규개발량 계획 (new_2 3_1, MJ→GJ)
개발량_연간 = (s(df_n2p,103,14) or 0)/1000
개발량_당계 = (s(df_n2p,103,nc) or 0)/1000
개발량_누계 = (s(df_n2p,105,nc) or 0)/1000

# 신규개발량 실적
# 당월: 신규개발량 파일 월간개발량 합산
# 누계: new_2 3_2 row94(누적행), col(nc) MJ÷1000=GJ
개발량_당실 = dev_detail['월간개발량'].sum() if dev_detail is not None else None
개발량_누실_v = s(df_n2r, 94, nc)
개발량_누실 = float(개발량_누실_v)/1000 if 개발량_누실_v else None

# 총공급량 계획
총공_연간 = s(df_n1rpt, 10, 2)   # new_1 회의자료 row10,col2
총공_당계 = (s(df_vol, 5, vc) or 0)/1000
총공_누계 = (s(df_vol, 6, vc) or 0)/1000
총공_당실_v = 총공_당실 if 총공_당실 > 0 else None
총공_누실_v = 총공_누실 if 총공_누실 > 0 else None

# 신규개발전 상세 계획 (new_1 3_1)
공동_p  = s(df_n1p,2,nc);  단독_p  = s(df_n1p,4,nc)
소계_p  = (공동_p or 0)+(단독_p or 0)
일반_p  = s(df_n1p,7,nc);  업무_p  = s(df_n1p,9,nc)
산업_p  = s(df_n1p,11,nc); 열병_p  = s(df_n1p,13,nc); 합계_p  = 신규_당계
공동_cp = s(df_n1p,3,nc);  단독_cp = s(df_n1p,5,nc)
소계_cp = (공동_cp or 0)+(단독_cp or 0)
일반_cp = s(df_n1p,8,nc);  업무_cp = s(df_n1p,10,nc)
산업_cp = s(df_n1p,12,nc); 열병_cp = s(df_n1p,14,nc); 합계_cp = 신규_누계

입력필요 = '<span style="color:#aaa;font-size:11px;">입력필요</span>'

# ════════════════════════════════════════════════════
# 화면 출력
# ════════════════════════════════════════════════════
st.markdown(f"### 📅 **{int(sel_year)}년 {sel_month}월** 영업현황 보고")

# ── 표1: 공급전 및 공급량 현황 ──
st.markdown('<div class="box-title">📋 공급전 및 공급량 현황</div>', unsafe_allow_html=True)

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
      <td>{fmt(신규_연간)}</td>
      <td>{fmt(신규_당계)}</td><td>{fmt(신규_당실)}</td>
      <td>{rate_html(신규_당실,신규_당계)}</td>
      <td>{fmt(신규_누계)}</td><td>{fmt(신규_누실)}</td>
      <td>{rate_html(신규_누실,신규_누계)}</td>
    </tr>
    <tr>
      <td class="td-sub-label">폐전</td>
      <td>{fmt(폐전_연간)}</td>
      <td>{fmt(폐전_당계)}</td><td>{fmt(폐전_당실)}</td>
      <td>{rate_html(폐전_당실,폐전_당계)}</td>
      <td>{fmt(폐전_누계)}</td><td>{fmt(폐전_누실)}</td>
      <td>{rate_html(폐전_누실,폐전_누계)}</td>
    </tr>
    <tr>
      <td class="td-sub-label">순증가</td>
      <td>{fmt(순증_연간)}</td>
      <td>{fmt(순증_당계)}</td><td>{fmt(순증_당실)}</td>
      <td>{rate_html(순증_당실,순증_당계)}</td>
      <td>{fmt(순증_누계)}</td><td>{fmt(순증_누실)}</td>
      <td>{rate_html(순증_누실,순증_누계)}</td>
    </tr>
    <tr>
      <td class="td-label" rowspan="2">공급량<br>(GJ)</td>
      <td class="td-sub-label">신규개발량</td>
      <td>{fmt(개발량_연간)}</td>
      <td>{fmt(개발량_당계)}</td>
      <td>{'<b>'+fmt(개발량_당실)+'</b>' if 개발량_당실 else 입력필요}</td>
      <td>{rate_html(개발량_당실,개발량_당계) if 개발량_당실 else '-'}</td>
      <td>{fmt(개발량_누계)}</td>
      <td>{'<b>'+fmt(개발량_누실)+'</b>' if 개발량_누실 else '-'}</td>
      <td>{rate_html(개발량_누실,개발량_누계) if 개발량_누실 else '-'}</td>
    </tr>
    <tr>
      <td class="td-sub-label">총공급량</td>
      <td>{fmt(총공_연간)}</td>
      <td>{fmt(총공_당계)}</td>
      <td>{'<b>'+fmt(총공_당실_v)+'</b>' if 총공_당실_v else 입력필요}</td>
      <td>{rate_html(총공_당실_v,총공_당계) if 총공_당실_v else '-'}</td>
      <td>{fmt(총공_누계)}</td>
      <td>{'<b>'+fmt(총공_누실_v)+'</b>' if 총공_누실_v else 입력필요}</td>
      <td>{rate_html(총공_누실_v,총공_누계) if 총공_누실_v else '-'}</td>
    </tr>
  </tbody>
</table>"""
st.markdown(html1, unsafe_allow_html=True)

# ── 표2: 신규개발전 상세 ──
st.markdown(f'<div class="box-title">📋 {sel_month}월 신규개발전 상세 현황</div>',
            unsafe_allow_html=True)
st.markdown('<div class="note-right">(단위 : 전)</div>', unsafe_allow_html=True)

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
      <th class="th-sub">공동주택</th><th class="th-sub">단독주택</th><th class="th-sub">소계</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="td-label">계획</td>
      <td>{cell(공동_p,공동_cp)}</td><td>{cell(단독_p,단독_cp)}</td><td>{cell(소계_p,소계_cp)}</td>
      <td>{fmt(일반_p)}</td><td>{fmt(업무_p)}</td>
      <td>{fmt(산업_p)}</td><td>{fmt(열병_p)}</td>
      <td>{cell(합계_p,합계_cp)}</td>
    </tr>
    <tr>
      <td class="td-label">실적</td>
      <td>{cell_bold(공동_a,공동_ca)}</td><td>{cell_bold(단독_a,단독_ca)}</td><td>{cell_bold(소계_a,소계_ca)}</td>
      <td>{cell_bold(일반_a,일반_ca)}</td><td>{cell_bold(업무_a,업무_ca)}</td>
      <td>{cell_bold(산업_a,산업_ca)}</td><td>{cell_bold(열병_a,열병_ca)}</td>
      <td>{cell_bold(합계_a,합계_ca)}</td>
    </tr>
    <tr>
      <td class="td-label">달성률</td>
      <td>{rate_cell(공동_a,공동_p,공동_ca,공동_cp)}</td>
      <td>{rate_cell(단독_a,단독_p,단독_ca,단독_cp)}</td>
      <td>{rate_cell(소계_a,소계_p,소계_ca,소계_cp)}</td>
      <td>{rate_cell(일반_a,일반_p,일반_ca,일반_cp)}</td>
      <td>{rate_cell(업무_a,업무_p,업무_ca,업무_cp)}</td>
      <td>{rate_cell(산업_a,산업_p,산업_ca,산업_cp)}</td>
      <td>{rate_cell(열병_a,열병_p,열병_ca,열병_cp)}</td>
      <td>{rate_cell(합계_a,합계_p,합계_ca,합계_cp)}</td>
    </tr>
    <tr>
      <td class="td-label">증감</td>
      <td>{inc_cell(공동_a,공동_p,공동_ca,공동_cp)}</td>
      <td>{inc_cell(단독_a,단독_p,단독_ca,단독_cp)}</td>
      <td>{inc_cell(소계_a,소계_p,소계_ca,소계_cp)}</td>
      <td>{inc_cell(일반_a,일반_p,일반_ca,일반_cp)}</td>
      <td>{inc_cell(업무_a,업무_p,업무_ca,업무_cp)}</td>
      <td>{inc_cell(산업_a,산업_p,산업_ca,산업_cp)}</td>
      <td>{inc_cell(열병_a,열병_p,열병_ca,열병_cp)}</td>
      <td>{inc_cell(합계_a,합계_p,합계_ca,합계_cp)}</td>
    </tr>
  </tbody>
</table>
<p style="font-size:12px;color:#555;">※ (괄호)는 누계 기준임.</p>"""
st.markdown(html2, unsafe_allow_html=True)

# ── 산업용 신규 업체 ──
if dev_detail is not None:
    industry_df = dev_detail[dev_detail['용도'].astype(str).str.contains('산업', na=False)].copy()
    if not industry_df.empty:
        st.markdown(f'<div class="box-title">🏭 {sel_month}월 산업용 신규 업체 현황</div>',
                    unsafe_allow_html=True)
        ind_rows = "".join([
            f"<tr><td>{row['신청명']}</td><td>{row['업종']}</td>"
            f"<td>{fmt(row['월사용예정량'])} ㎥</td>"
            f"<td>{fmt(row['월간개발량'],2)} GJ</td>"
            f"<td>{str(row['공급일'])[:10] if pd.notna(row['공급일']) else '-'}</td>"
            f"<td style='font-size:11px;text-align:left'>{row['주소']}</td></tr>"
            for _, row in industry_df.iterrows()
        ])
        st.markdown(f"""
        <table>
          <thead><tr>
            <th>업체명</th><th>업종</th><th>월사용예정량</th>
            <th>열량(GJ)</th><th>공급일</th><th>주소</th>
          </tr></thead>
          <tbody>{ind_rows}
            <tr style="background:#dce6f5;font-weight:700;">
              <td colspan="2">합 계</td>
              <td>{fmt(industry_df['월사용예정량'].sum())} ㎥</td>
              <td>{fmt(industry_df['월간개발량'].sum(),2)} GJ</td>
              <td colspan="2"></td>
            </tr>
          </tbody>
        </table>""", unsafe_allow_html=True)
    else:
        st.info(f"✅ {sel_month}월 신규 산업용 업체 없음")
else:
    st.warning("② 신규개발량 파일을 업로드하면 산업용 업체 현황이 표시됩니다.")

st.markdown("---")

# ── 출력 버튼 (JavaScript로 작동) ──
st.markdown("""
<style>
.print-btn-wrap { text-align: center; padding: 20px 0; }
.print-btn {
    background: linear-gradient(135deg, #1e3a6b, #2d5fa8);
    color: white; font-size: 16px; font-weight: 700;
    padding: 14px 60px; border-radius: 8px; cursor: pointer;
    border: none; box-shadow: 0 4px 12px rgba(30,58,107,0.3);
}
.print-btn:hover { opacity: 0.85; }
.print-guide { font-size: 12px; color: #888; margin-top: 10px; }
@media print {
    .stApp > header, .stSidebar, .stButton,
    .stFileUploader, .stNumberInput, .stSelectbox,
    .upload-box, [data-testid="stToolbar"],
    .print-btn-wrap { display: none !important; }
    .report-header { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    table { page-break-inside: avoid; }
}
</style>
<div class="print-btn-wrap">
    <button class="print-btn" onclick="window.print()">🖨️ 보고서 출력 / PDF 저장</button>
    <div class="print-guide">
        버튼 클릭 후 → 대상: <b>PDF로 저장</b> → 레이아웃: <b>가로</b> → 저장
    </div>
</div>
<script>
// Streamlit iframe 내에서 부모 윈도우 출력 트리거
function triggerPrint() {
    window.parent.postMessage({type: "print"}, "*");
}
window.addEventListener("message", function(e) {
    if (e.data && e.data.type === "print") { window.print(); }
});
</script>
""", unsafe_allow_html=True)
