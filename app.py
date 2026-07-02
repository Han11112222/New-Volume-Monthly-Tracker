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
th, td { border: 1px solid #aaa; padding: 6px 10px; text-align: center; white-space: nowrap; }
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

# ── GitHub에서 new_2 자동 로드 ────────────────────────
GITHUB_RAW = "https://raw.githubusercontent.com/Han11112222/New-Volume-Monthly-Tracker/main"
NEW2_FNAME = "new_2.(통합)신규개발량(202606)_5월 확정공급량적용_20260626.xlsx"

@st.cache_data(ttl=3600)
def load_new2_github():
    try:
        url = f"{GITHUB_RAW}/{urllib.parse.quote(NEW2_FNAME)}"
        req = urllib.request.urlopen(url, timeout=15)
        buf = io.BytesIO(req.read())
        out = {}
        for s in ['3_1. 개발량 계획', '3_2. 개발량 실적']:
            try: out[s] = pd.read_excel(buf, sheet_name=s, header=None); buf.seek(0)
            except: pass
        return out, None
    except Exception as e:
        return {}, str(e)

# ── 구글 스프레드시트 공급량 자동 로드 ──────────────────
GSHEET_ID  = '13HrIz6OytYDykXeXzXJ02I6XbaKin1YaKBoO2kBd6Bs'
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}/export?format=csv&gid=0"

@st.cache_data(ttl=600)
def load_supply_gsheet():
    try:
        df = pd.read_csv(GSHEET_URL, header=0)
        cols = ['일자','공급량_MJ','공급량_M3','평균기온','최저기온','최고기온']
        df.columns = cols[:len(df.columns)] + [f'_c{i}' for i in range(max(0,len(df.columns)-6))]
        df['일자'] = pd.to_datetime(df['일자'], errors='coerce')
        df = df.dropna(subset=['일자'])
        df['공급량_MJ'] = pd.to_numeric(
            df['공급량_MJ'].astype(str).str.replace(',',''), errors='coerce')
        df['공급량_GJ'] = df['공급량_MJ'] / 1000
        df['연'] = df['일자'].dt.year
        df['월'] = df['일자'].dt.month
        monthly = df.groupby(['연','월'])['공급량_GJ'].sum().reset_index()
        return monthly, None
    except Exception as e:
        return None, str(e)

def get_monthly_gj(mdf, year, month):
    if mdf is None: return None
    r = mdf[(mdf['연']==year) & (mdf['월']==month)]
    return float(r['공급량_GJ'].values[0]) if len(r) > 0 else None

def get_cum_gj(mdf, year, month):
    if mdf is None: return None
    f = mdf[(mdf['연']==year) & (mdf['월']<=month)]
    return float(f['공급량_GJ'].sum()) if len(f) > 0 else None

# ── 엑셀 파일 로드 ───────────────────────────────────────
@st.cache_data
def load_xlsm(b, fn):
    buf = io.BytesIO(b)
    out = {}
    for s in ['공급전 계획(2026년)', '영업현황']:
        try: out[s] = pd.read_excel(buf, sheet_name=s, header=None, engine='openpyxl'); buf.seek(0)
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
📁 <b>매월 2개 파일만 업로드</b>하면 자동으로 보고서가 생성됩니다.<br>
&nbsp;&nbsp;① 영업일보_월간_YYYYMM.xlsm &nbsp;|&nbsp; ② 신규개발량_YYYYMM.xlsx<br>
&nbsp;&nbsp;📡 계획 데이터(new_2)는 GitHub에서, 총공급량은 구글시트에서 <b>자동 로드</b>됩니다.
</div>
""", unsafe_allow_html=True)

col_u1, col_u2 = st.columns(2)
with col_u1: f_xlsm = st.file_uploader("① 영업일보 (xlsm)", type=['xlsm','xlsx'])
with col_u2: f_dev  = st.file_uploader("② 신규개발량 (xlsx) — 산업용 업체 확인", type=['xlsx'])

# ── 보고 월 선택 ──
st.markdown("---")
col_s1, col_s2, col_s3, col_s4 = st.columns([1, 2, 2, 1])
with col_s1:
    sel_year  = st.number_input("보고 연도", value=2026, step=1, format="%d")
    sel_month = st.selectbox("보고 월", list(range(1,13)), index=5)

# ── 자동 로드 ──
with st.spinner("📡 GitHub & 구글시트에서 데이터 자동 로드 중..."):
    sheets_n2, n2_err   = load_new2_github()
    monthly_df, gs_err  = load_supply_gsheet()

auto_actual = get_monthly_gj(monthly_df, int(sel_year), sel_month)
auto_cum    = get_cum_gj(monthly_df, int(sel_year), sel_month)

# 상태 표시
c1, c2 = st.columns(2)
with c1:
    if n2_err:
        st.warning(f"⚠️ GitHub new_2 로드 실패: {n2_err}")
    else:
        st.success("✅ GitHub new_2 계획 데이터 자동 로드 완료")
with c2:
    if gs_err:
        st.warning(f"⚠️ 구글시트 로드 실패: {gs_err}")
    elif auto_actual:
        st.success(f"✅ 구글시트 자동합산: 당월 **{auto_actual:,.0f} GJ** | 누계 **{auto_cum:,.0f} GJ**")

# ── 총공급량 입력 (자동값 디폴트, 수정 가능) ──
with col_s2:
    총공_당실 = st.number_input(
        "총공급량 당월 실적 (GJ) 📡자동",
        value=int(round(auto_actual)) if auto_actual else 0,
        step=100, format="%d",
        help="구글시트 자동합산값. 필요시 수정 가능.")
with col_s3:
    총공_누실 = st.number_input(
        "총공급량 누계 실적 (GJ) 📡자동",
        value=int(round(auto_cum)) if auto_cum else 0,
        step=100, format="%d",
        help="1월~당월 누계 자동합산값. 필요시 수정 가능.")
with col_s4:
    st.markdown("<br><br>", unsafe_allow_html=True)
    confirmed = st.button("✅ 확인/적용", use_container_width=True)

if confirmed:
    st.success(f"✅ 적용: 당월 **{총공_당실:,} GJ** | 누계 **{총공_누실:,} GJ**")

if not f_xlsm:
    st.info("👆 영업일보 파일을 먼저 업로드해주세요.")
    st.stop()

# ════════════════════════════════════════════════════
# 데이터 파싱
# ════════════════════════════════════════════════════
sheets_xlsm = load_xlsm(f_xlsm.read(), f_xlsm.name); f_xlsm.seek(0)
dev_detail  = load_dev_detail(f_dev.read(), f_dev.name) if f_dev else None

df_plan = sheets_xlsm.get('공급전 계획(2026년)')
df_biz  = sheets_xlsm.get('영업현황')
df_n2p  = sheets_n2.get('3_1. 개발량 계획')   # GitHub에서 로드된 계획
df_n2r  = sheets_n2.get('3_2. 개발량 실적')   # GitHub에서 로드된 실적

m  = sel_month - 1
nc = m + 2   # new_2: col2=1월 ~ col13=12월

def p(df, r, c): return safe(df, r, c) if df is not None else None

# ── 연간계획 (new_2 기준) ──
신규_연간   = p(df_n2p, 15, 14)
폐전_연간   = p(df_n2p, 34, 14)
순증_연간   = (신규_연간 or 0) - (폐전_연간 or 0)
개발량_연간 = p(df_n2p, 103, 14)
if 개발량_연간: 개발량_연간 /= 1000
총공_연간   = 45384427

# ── 당월 계획 (new_2 3_1) ──
신규_당계   = p(df_n2p, 15, nc)
폐전_당계   = p(df_n2p, 34, nc)
순증_당계   = (신규_당계 or 0) - (폐전_당계 or 0)
개발량_당계 = p(df_n2p, 103, nc)
if 개발량_당계: 개발량_당계 /= 1000
공동_p = p(df_n2p, 2,  nc); 단독_p = p(df_n2p, 4,  nc)
소계_p = (공동_p or 0)+(단독_p or 0)
일반_p = p(df_n2p, 7,  nc); 업무_p = p(df_n2p, 9,  nc)
산업_p = p(df_n2p, 11, nc); 열병_p = p(df_n2p, 13, nc)
합계_p = 신규_당계

# ── 누계 계획 (new_2 3_1) ──
신규_누계   = p(df_n2p, 16, nc)
폐전_누계   = p(df_n2p, 35, nc)
순증_누계   = (신규_누계 or 0) - (폐전_누계 or 0)
개발량_누계 = p(df_n2p, 104, nc)
if 개발량_누계: 개발량_누계 /= 1000
공동_cp = p(df_n2p, 3,  nc); 단독_cp = p(df_n2p, 5,  nc)
소계_cp = (공동_cp or 0)+(단독_cp or 0)
일반_cp = p(df_n2p, 8,  nc); 업무_cp = p(df_n2p, 10, nc)
산업_cp = p(df_n2p, 12, nc); 열병_cp = p(df_n2p, 14, nc)
합계_cp = 신규_누계
총공_누계_p = 개발량_누계

# ── 실적 (영업일보 영업현황) ──
cats_biz = ['공동주택','단독주택','소계','일반용','업무용','산업용','열병합','합계']
act = {}
폐전_당실 = 폐전_누실 = 신규_당실 = 순증_당실 = None
if df_biz is not None:
    for i, cat in enumerate(cats_biz):
        act[cat] = safe(df_biz, 25, i+2)
    폐전_당실 = safe(df_biz, 7, 9)
    폐전_누실 = safe(df_biz, 6, 9)
    신규_당실 = act.get('합계')
    순증_당실 = (float(신규_당실 or 0)-float(폐전_당실 or 0)) if 신규_당실 else None

# ── 신규개발량 실적 (new_2 3_2) ──
개발량_당실 = p(df_n2r, 92, nc)
개발량_누실 = p(df_n2r, 94, nc)   # 누적(row94) = 1월~당월 합계
if 개발량_당실: 개발량_당실 /= 1000
if 개발량_누실: 개발량_누실 /= 1000

공동_a=act.get('공동주택'); 단독_a=act.get('단독주택')
소계_a=act.get('소계');    일반_a=act.get('일반용')
업무_a=act.get('업무용');  산업_a=act.get('산업용')
열병_a=act.get('열병합');  합계_a=act.get('합계')

# 누계 실적 (영업현황 row21: 공급전 총누계)
공동_ca=safe(df_biz,21,2) if df_biz is not None else None
단독_ca=safe(df_biz,21,3) if df_biz is not None else None
소계_ca=safe(df_biz,21,4) if df_biz is not None else None
일반_ca=safe(df_biz,21,5) if df_biz is not None else None
업무_ca=safe(df_biz,21,6) if df_biz is not None else None
산업_ca=safe(df_biz,21,7) if df_biz is not None else None
열병_ca=safe(df_biz,21,8) if df_biz is not None else None
합계_ca=safe(df_biz,21,9) if df_biz is not None else None

총공_당실_v = 총공_당실 if 총공_당실 > 0 else None
총공_누실_v = 총공_누실 if 총공_누실 > 0 else None
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
      <td>{rate_html(신규_당실, 신규_당계)}</td>
      <td>{fmt(신규_누계)}</td><td>{fmt(신규_당실)}</td>
      <td>{rate_html(신규_당실, 신규_누계)}</td>
    </tr>
    <tr>
      <td class="td-sub-label">폐전</td>
      <td>{fmt(폐전_연간)}</td>
      <td>{fmt(폐전_당계)}</td><td>{fmt(폐전_당실)}</td>
      <td>{rate_html(폐전_당실, 폐전_당계)}</td>
      <td>{fmt(폐전_누계)}</td><td>{fmt(폐전_누실)}</td>
      <td>{rate_html(폐전_누실, 폐전_누계)}</td>
    </tr>
    <tr>
      <td class="td-sub-label">순증가</td>
      <td>{fmt(순증_연간)}</td>
      <td>{fmt(순증_당계)}</td><td>{fmt(순증_당실)}</td>
      <td>{rate_html(순증_당실, 순증_당계)}</td>
      <td>{fmt(순증_누계)}</td><td>-</td><td>-</td>
    </tr>
    <tr>
      <td class="td-label" rowspan="2">공급량<br>(GJ)</td>
      <td class="td-sub-label">신규개발량</td>
      <td>{fmt(개발량_연간)}</td>
      <td>{fmt(개발량_당계)}</td><td>{fmt(개발량_당실)}</td>
      <td>{rate_html(개발량_당실, 개발량_당계)}</td>
      <td>{fmt(개발량_누계)}</td><td>{fmt(개발량_누실)}</td>
      <td>{rate_html(개발량_누실, 개발량_누계)}</td>
    </tr>
    <tr>
      <td class="td-sub-label">총공급량</td>
      <td>{fmt(총공_연간)}</td>
      <td>{fmt(개발량_당계)}</td>
      <td>{'<b>'+fmt(총공_당실_v)+'</b>' if 총공_당실_v else 입력필요}</td>
      <td>{rate_html(총공_당실_v, 개발량_당계) if 총공_당실_v else '-'}</td>
      <td>{fmt(총공_누계_p)}</td>
      <td>{'<b>'+fmt(총공_누실_v)+'</b>' if 총공_누실_v else 입력필요}</td>
      <td>{rate_html(총공_누실_v, 총공_누계_p) if 총공_누실_v else '-'}</td>
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
      <th class="th-sub">공동주택</th>
      <th class="th-sub">단독주택</th>
      <th class="th-sub">소계</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="td-label">계획</td>
      <td>{fmt(공동_p)}<br><small style='color:#888'>({fmt(공동_cp)})</small></td>
      <td>{fmt(단독_p)}<br><small style='color:#888'>({fmt(단독_cp)})</small></td>
      <td>{fmt(소계_p)}<br><small style='color:#888'>({fmt(소계_cp)})</small></td>
      <td>{fmt(일반_p)}</td><td>{fmt(업무_p)}</td>
      <td>{fmt(산업_p)}</td><td>{fmt(열병_p)}</td>
      <td>{fmt(합계_p)}<br><small style='color:#888'>({fmt(합계_cp)})</small></td>
    </tr>
    <tr>
      <td class="td-label">실적</td>
      <td>{fmt(공동_a)}<br><small style='color:#555'>({fmt(공동_ca)})</small></td>
      <td>{fmt(단독_a)}<br><small style='color:#555'>({fmt(단독_ca)})</small></td>
      <td>{fmt(소계_a)}<br><small style='color:#555'>({fmt(소계_ca)})</small></td>
      <td>{fmt(일반_a)}<br><small style='color:#555'>({fmt(일반_ca)})</small></td>
      <td>{fmt(업무_a)}<br><small style='color:#555'>({fmt(업무_ca)})</small></td>
      <td>{fmt(산업_a)}<br><small style='color:#555'>({fmt(산업_ca)})</small></td>
      <td>{fmt(열병_a)}<br><small style='color:#555'>({fmt(열병_ca)})</small></td>
      <td>{fmt(합계_a)}<br><small style='color:#555'>({fmt(합계_ca)})</small></td>
    </tr>
    <tr>
      <td class="td-label">달성률</td>
      <td>{rate_html(공동_a,공동_p)}<br><small style='color:#888'>({rate_html(공동_ca,공동_cp)})</small></td>
      <td>{rate_html(단독_a,단독_p)}<br><small style='color:#888'>({rate_html(단독_ca,단독_cp)})</small></td>
      <td>{rate_html(소계_a,소계_p)}<br><small style='color:#888'>({rate_html(소계_ca,소계_cp)})</small></td>
      <td>{rate_html(일반_a,일반_p)}<br><small style='color:#888'>({rate_html(일반_ca,일반_cp)})</small></td>
      <td>{rate_html(업무_a,업무_p)}<br><small style='color:#888'>({rate_html(업무_ca,업무_cp)})</small></td>
      <td>{rate_html(산업_a,산업_p)}<br><small style='color:#888'>({rate_html(산업_ca,산업_cp)})</small></td>
      <td>{rate_html(열병_a,열병_p)}<br><small style='color:#888'>({rate_html(열병_ca,열병_cp)})</small></td>
      <td>{rate_html(합계_a,합계_p)}<br><small style='color:#888'>({rate_html(합계_ca,합계_cp)})</small></td>
    </tr>
    <tr>
      <td class="td-label">증감</td>
      <td>{d_inc(공동_a,공동_p)}<br><small style='color:#888'>({d_inc(공동_ca,공동_cp)})</small></td>
      <td>{d_inc(단독_a,단독_p)}<br><small style='color:#888'>({d_inc(단독_ca,단독_cp)})</small></td>
      <td>{d_inc(소계_a,소계_p)}<br><small style='color:#888'>({d_inc(소계_ca,소계_cp)})</small></td>
      <td>{d_inc(일반_a,일반_p)}<br><small style='color:#888'>({d_inc(일반_ca,일반_cp)})</small></td>
      <td>{d_inc(업무_a,업무_p)}<br><small style='color:#888'>({d_inc(업무_ca,업무_cp)})</small></td>
      <td>{d_inc(산업_a,산업_p)}<br><small style='color:#888'>({d_inc(산업_ca,산업_cp)})</small></td>
      <td>{d_inc(열병_a,열병_p)}<br><small style='color:#888'>({d_inc(열병_ca,열병_cp)})</small></td>
      <td>{d_inc(합계_a,합계_p)}<br><small style='color:#888'>({d_inc(합계_ca,합계_cp)})</small></td>
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
            f"<tr><td>{r['신청명']}</td><td>{r['업종']}</td>"
            f"<td>{fmt(r['월사용예정량'])} ㎥</td>"
            f"<td>{fmt(r['월간개발량'],2)} GJ</td>"
            f"<td>{str(r['공급일'])[:10] if pd.notna(r['공급일']) else '-'}</td>"
            f"<td style='font-size:11px;text-align:left'>{r['주소']}</td></tr>"
            for _, r in industry_df.iterrows()
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

# ── PDF 출력 ──
st.markdown("---")
st.markdown("""
<div style='text-align:center;padding:12px;color:#555;font-size:13px;'>
🖨️ <b>PDF 출력</b>: Ctrl+P → 대상: PDF로 저장 → 레이아웃: 가로 → 저장
</div>""", unsafe_allow_html=True)
