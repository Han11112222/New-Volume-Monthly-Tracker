import streamlit as st
import pandas as pd
import io
import urllib.request
import urllib.parse

st.set_page_config(layout="wide", page_title="마케팅본부 월간 보고서")

# ── 스타일 ──────────────────────────────────────────────
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
.upload-box {
    background: #eef4fc; border: 1px solid #2d5fa8;
    border-radius: 6px; padding: 10px 16px; margin-bottom: 14px; font-size: 13px;
}
table { border-collapse: collapse; width: 100%; margin-bottom: 12px; }
th {
    background: #1e3a6b; color: white; font-weight: 700;
    border: 1px solid #2d5fa8; padding: 7px 10px;
    text-align: center; font-size: 13px; vertical-align: middle;
}
th.th-sub { background: #2d5fa8; }
td {
    border: 1px solid #ddd; padding: 7px 10px;
    text-align: center; font-size: 13px;
    vertical-align: middle; background: #ffffff;
}
tr:nth-child(even) td { background: #f8f9fc; }
td.lbl  { background: #dce6f5 !important; font-weight: 700;
           color: #1e3a6b; text-align: center; }
td.slbl { background: #eef2fa !important; font-weight: 600;
           color: #333; text-align: center; }
@media print {
    .no-print { display: none !important; }
    thead th { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
}
</style>
""", unsafe_allow_html=True)

# ── 헬퍼 ──────────────────────────────────────────────
# 숫자: 검정 #222, 13px, normal
NUM  = "style='color:#222;font-size:13px;font-weight:normal'"
SNUM = "style='color:#444;font-size:11px;font-weight:normal'"  # 괄호 누계

def fmt(v, dec=0):
    if v is None or (isinstance(v, float) and pd.isna(v)): return "-"
    try:
        f = float(v)
        return f"{f:,.{dec}f}" if dec > 0 else f"{int(round(f)):,}"
    except: return str(v)

def n(v, dec=0):
    """숫자 셀 값 — 검정, 13px, normal"""
    return f"<span {NUM}>{fmt(v,dec)}</span>"

def rate_html(a, p):
    try:
        val = float(a or 0) / float(p or 1) * 100
        col = "#1a7a1a" if val >= 100 else "#c0392b"
        return f"<span style='color:{col};font-size:13px;font-weight:normal'>{val:.1f}%</span>"
    except: return "<span style='color:#222;font-size:13px'>-</span>"

def safe(df, r, c):
    try:
        v = df.iloc[r, c]
        return None if (isinstance(v, float) and pd.isna(v)) else v
    except: return None

def d_inc(a, p):
    try: return fmt(float(a or 0) - float(p or 0))
    except: return "-"

def cell(당월, 누계):
    return f"<span {NUM}>{fmt(당월)}</span><br><span {SNUM}>({fmt(누계)})</span>"

def rate_cell(da, dp, ca, cp):
    return f"{rate_html(da,dp)}<br><span {SNUM}>({rate_html(ca,cp)})</span>"

def inc_cell(da, dp, ca, cp):
    return (f"<span {NUM}>{d_inc(da,dp)}</span>"
            f"<br><span {SNUM}>({d_inc(ca,cp)})</span>")

입력필요 = "<span style='color:#aaa;font-size:11px'>입력필요</span>"

# ── GitHub 자동 로드 ──────────────────────────────────
BASE = "https://raw.githubusercontent.com/Han11112222/New-Volume-Monthly-Tracker/main"

@st.cache_data(ttl=3600)
def load_github():
    errors, result = [], {}
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
            buf = io.BytesIO(urllib.request.urlopen(url, timeout=15).read())
            result[key] = {}
            for s in info["sheets"]:
                try:
                    result[key][s] = pd.read_excel(buf, sheet_name=s, header=None)
                    buf.seek(0)
                except: pass
        except Exception as e:
            errors.append(f"{key}: {e}")
    return result, errors

# ── 구글시트 총공급량 ──────────────────────────────────
GSHEET_URL = ("https://docs.google.com/spreadsheets/d/"
              "13HrIz6OytYDykXeXzXJ02I6XbaKin1YaKBoO2kBd6Bs/export?format=csv&gid=0")

@st.cache_data(ttl=300, show_spinner=False)
def load_gsheet():
    try:
        req = urllib.request.Request(GSHEET_URL, headers={"User-Agent":"Mozilla/5.0"})
        raw = urllib.request.urlopen(req, timeout=15).read()
        df  = pd.read_csv(io.BytesIO(raw), header=0).iloc[:,:6]
        df.columns = ['일자','공급량_MJ','공급량_M3','평균기온','최저기온','최고기온']
        df['일자'] = pd.to_datetime(df['일자'], errors='coerce')
        df = df.dropna(subset=['일자'])
        df['공급량_MJ'] = pd.to_numeric(
            df['공급량_MJ'].astype(str).str.replace(',','').str.strip(), errors='coerce')
        df['GJ'] = df['공급량_MJ'] / 1000
        df['연'] = df['일자'].dt.year; df['월'] = df['일자'].dt.month
        return df.groupby(['연','월'])['GJ'].sum().reset_index(), None
    except Exception as e:
        return None, str(e)

def get_gj(mdf, yr, mo, cum=False):
    if mdf is None: return None
    f = mdf[(mdf['연']==yr) & (mdf['월']<=(mo if cum else mo)) &
            (mdf['월']>=(1 if cum else mo))]
    return float(f['GJ'].sum()) if len(f)>0 else None

# ── 파일 로드 ──────────────────────────────────────────
@st.cache_data
def load_xlsm(b, fn):
    buf = io.BytesIO(b); out = {}
    for s in ['영업일보','공급전 계획(2026년)','공급량 계획_MJ(2026년)']:
        try: out[s] = pd.read_excel(buf, sheet_name=s, header=None, engine='openpyxl'); buf.seek(0)
        except: pass
    return out

@st.cache_data
def load_dev(b, fn):
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
<div class="upload-box no-print">
📁 <b>매월 2개 파일만 업로드</b>하면 자동으로 보고서가 생성됩니다.<br>
&nbsp;&nbsp;① 영업일보_월간_YYYYMM.xlsm &nbsp;|&nbsp; ② 신규개발량_YYYYMM.xlsx<br>
&nbsp;&nbsp;📡 계획 데이터(GitHub) 및 총공급량(구글시트)은 <b>자동 로드</b>됩니다.
</div>""", unsafe_allow_html=True)

cu1, cu2 = st.columns(2)
with cu1: f_xlsm = st.file_uploader("① 영업일보 (xlsm)", type=['xlsm','xlsx'])
with cu2: f_dev  = st.file_uploader("② 신규개발량 (xlsx)", type=['xlsx'])

st.markdown("---")
cs1,cs2,cs3,cs4 = st.columns([1,2,2,1])
with cs1:
    sel_year  = st.number_input("보고 연도", value=2026, step=1, format="%d")
    sel_month = st.selectbox("보고 월", list(range(1,13)), index=5)

with st.spinner("📡 데이터 자동 로드 중..."):
    gh, gh_err = load_github()
    mdf, gs_err = load_gsheet()

auto_act = get_gj(mdf, int(sel_year), sel_month)
auto_cum = get_gj(mdf, int(sel_year), sel_month, cum=True)

cc1,cc2 = st.columns(2)
with cc1:
    if gh_err: st.warning("⚠️ GitHub: " + " | ".join(gh_err))
    else: st.success("✅ GitHub 계획 데이터 로드 완료")
with cc2:
    if gs_err: st.warning(f"⚠️ 구글시트: {gs_err}")
    elif auto_act: st.success(f"✅ 구글시트: 당월 **{auto_act:,.0f}** | 누계 **{auto_cum:,.0f}** GJ")
    else: st.info("ℹ️ 구글시트 데이터 없음 → 수동 입력")

with cs2:
    총공_당실 = st.number_input("총공급량 당월 실적 (GJ) 📡자동",
        value=int(round(auto_act)) if auto_act else 0, step=100, format="%d")
with cs3:
    총공_누실 = st.number_input("총공급량 누계 실적 (GJ) 📡자동",
        value=int(round(auto_cum)) if auto_cum else 0, step=100, format="%d")
with cs4:
    st.markdown("<br><br>", unsafe_allow_html=True)
    if st.button("✅ 확인/적용", use_container_width=True):
        st.success(f"✅ 당월 **{총공_당실:,}** | 누계 **{총공_누실:,}** GJ")

# PDF 출력 버튼 (Streamlit 네이티브)
st.markdown("<br>", unsafe_allow_html=True)
col_pdf = st.columns([3,1,3])
with col_pdf[1]:
    if st.button("🖨️ PDF 출력", use_container_width=True, type="primary"):
        st.markdown("""
        <script>setTimeout(function(){ window.print(); }, 300);</script>
        """, unsafe_allow_html=True)
        st.info("브라우저 인쇄창 → 대상: PDF로 저장 → 레이아웃: 가로")

if not f_xlsm:
    st.info("👆 영업일보 파일을 업로드해주세요.")
    st.stop()

# ════════════════════════════════════════════════════
# 데이터 파싱
# ════════════════════════════════════════════════════
sheets = load_xlsm(f_xlsm.read(), f_xlsm.name); f_xlsm.seek(0)
dev_df = load_dev(f_dev.read(), f_dev.name) if f_dev else None

df_il   = sheets.get('영업일보')
df_vol  = sheets.get('공급량 계획_MJ(2026년)')
df_n1p  = gh.get("new_1",{}).get("3_1. 개발량 계획")
df_n1rt = gh.get("new_1",{}).get("(회의자료 입력용)공급전 및 공급량 현황")
df_n2p  = gh.get("new_2",{}).get("3_1. 개발량 계획")
df_n2r  = gh.get("new_2",{}).get("3_2. 개발량 실적")

m  = sel_month - 1
nc = m + 2   # new_1/2: col2=1월
vc = m + 3   # 공급량계획: col3=1월

def s(df, r, c): return safe(df, r, c) if df is not None else None

# ── 실적 (영업일보 시트) ──
RI = {'합계':7,'공동':8,'단독':9,'일반':10,'업무':11,'산업':12,'열병합':13}
def il(k, c): return s(df_il, RI.get(k), c)

공동_a=il('공동',4); 단독_a=il('단독',4); 소계_a=(공동_a or 0)+(단독_a or 0)
일반_a=il('일반',4); 업무_a=il('업무',4); 산업_a=il('산업',4)
열병_a=il('열병합',4); 합계_a=il('합계',4); 신규_당실=합계_a

공동_ca=il('공동',10); 단독_ca=il('단독',10); 소계_ca=(공동_ca or 0)+(단독_ca or 0)
일반_ca=il('일반',10); 업무_ca=il('업무',10); 산업_ca=il('산업',10)
열병_ca=il('열병합',10); 합계_ca=il('합계',10); 신규_누실=합계_ca

폐전_당실=il('합계',5); 폐전_누실=il('합계',11)
순증_당실=(float(신규_당실 or 0)-float(폐전_당실 or 0)) if 신규_당실 else None
순증_누실=(float(신규_누실 or 0)-float(폐전_누실 or 0)) if 신규_누실 else None

# ── 계획 (GitHub) ──
신규_연간=s(df_n1p,15,14); 신규_당계=s(df_n1p,15,nc); 신규_누계=s(df_n1p,16,nc)
폐전_연간=s(df_n1p,34,14); 폐전_당계=s(df_n1p,34,nc); 폐전_누계=s(df_n1p,35,nc)
순증_연간=(신규_연간 or 0)-(폐전_연간 or 0)
순증_당계=(신규_당계 or 0)-(폐전_당계 or 0)
순증_누계=(신규_누계 or 0)-(폐전_누계 or 0)

개발량_연간=(s(df_n2p,103,14) or 0)/1000
개발량_당계=(s(df_n2p,103,nc) or 0)/1000
개발량_누계=(s(df_n2p,105,nc) or 0)/1000
개발량_당실=dev_df['월간개발량'].sum() if dev_df is not None else None
_nv=s(df_n2r,94,nc); 개발량_누실=float(_nv)/1000 if _nv else None

총공_연간=s(df_n1rt,10,2)
총공_당계=(s(df_vol,5,vc) or 0)/1000
총공_누계=(s(df_vol,6,vc) or 0)/1000
총공_당실_v=총공_당실 if 총공_당실>0 else None
총공_누실_v=총공_누실 if 총공_누실>0 else None

공동_p=s(df_n1p,2,nc); 단독_p=s(df_n1p,4,nc); 소계_p=(공동_p or 0)+(단독_p or 0)
일반_p=s(df_n1p,7,nc); 업무_p=s(df_n1p,9,nc); 산업_p=s(df_n1p,11,nc)
열병_p=s(df_n1p,13,nc); 합계_p=신규_당계
공동_cp=s(df_n1p,3,nc); 단독_cp=s(df_n1p,5,nc); 소계_cp=(공동_cp or 0)+(단독_cp or 0)
일반_cp=s(df_n1p,8,nc); 업무_cp=s(df_n1p,10,nc); 산업_cp=s(df_n1p,12,nc)
열병_cp=s(df_n1p,14,nc); 합계_cp=신규_누계

# ════════════════════════════════════════════════════
# 표 출력
# ════════════════════════════════════════════════════
st.markdown(f"### 📅 **{int(sel_year)}년 {sel_month}월** 영업현황 보고")

# ── 표1: 공급전 및 공급량 현황 ──
st.markdown('<div class="box-title">📋 공급전 및 공급량 현황</div>', unsafe_allow_html=True)
st.markdown(f"""
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
      <td class="lbl" rowspan="3">공급전<br>(전)</td>
      <td class="slbl">신규개발전</td>
      <td>{n(신규_연간)}</td>
      <td>{n(신규_당계)}</td><td>{n(신규_당실)}</td>
      <td>{rate_html(신규_당실,신규_당계)}</td>
      <td>{n(신규_누계)}</td><td>{n(신규_누실)}</td>
      <td>{rate_html(신규_누실,신규_누계)}</td>
    </tr>
    <tr>
      <td class="slbl">폐전</td>
      <td>{n(폐전_연간)}</td>
      <td>{n(폐전_당계)}</td><td>{n(폐전_당실)}</td>
      <td>{rate_html(폐전_당실,폐전_당계)}</td>
      <td>{n(폐전_누계)}</td><td>{n(폐전_누실)}</td>
      <td>{rate_html(폐전_누실,폐전_누계)}</td>
    </tr>
    <tr>
      <td class="slbl">순증가</td>
      <td>{n(순증_연간)}</td>
      <td>{n(순증_당계)}</td><td>{n(순증_당실)}</td>
      <td>{rate_html(순증_당실,순증_당계)}</td>
      <td>{n(순증_누계)}</td><td>{n(순증_누실)}</td>
      <td>{rate_html(순증_누실,순증_누계)}</td>
    </tr>
    <tr>
      <td class="lbl" rowspan="2">공급량<br>(GJ)</td>
      <td class="slbl">신규개발량</td>
      <td>{n(개발량_연간)}</td>
      <td>{n(개발량_당계)}</td>
      <td>{n(개발량_당실) if 개발량_당실 else 입력필요}</td>
      <td>{rate_html(개발량_당실,개발량_당계) if 개발량_당실 else n(None)}</td>
      <td>{n(개발량_누계)}</td>
      <td>{n(개발량_누실) if 개발량_누실 else n(None)}</td>
      <td>{rate_html(개발량_누실,개발량_누계) if 개발량_누실 else n(None)}</td>
    </tr>
    <tr>
      <td class="slbl">총공급량</td>
      <td>{n(총공_연간)}</td>
      <td>{n(총공_당계)}</td>
      <td>{n(총공_당실_v) if 총공_당실_v else 입력필요}</td>
      <td>{rate_html(총공_당실_v,총공_당계) if 총공_당실_v else n(None)}</td>
      <td>{n(총공_누계)}</td>
      <td>{n(총공_누실_v) if 총공_누실_v else 입력필요}</td>
      <td>{rate_html(총공_누실_v,총공_누계) if 총공_누실_v else n(None)}</td>
    </tr>
  </tbody>
</table>""", unsafe_allow_html=True)

# ── 표2: 신규개발전 상세 현황 ──
st.markdown(f'<div class="box-title">📋 {sel_month}월 신규개발전 상세 현황</div>',
            unsafe_allow_html=True)
st.markdown('<div class="note-right">(단위 : 전)</div>', unsafe_allow_html=True)
st.markdown(f"""
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
      <td class="lbl">계획</td>
      <td>{cell(공동_p,공동_cp)}</td><td>{cell(단독_p,단독_cp)}</td><td>{cell(소계_p,소계_cp)}</td>
      <td>{n(일반_p)}</td><td>{n(업무_p)}</td>
      <td>{n(산업_p)}</td><td>{n(열병_p)}</td><td>{cell(합계_p,합계_cp)}</td>
    </tr>
    <tr>
      <td class="lbl">실적</td>
      <td>{cell(공동_a,공동_ca)}</td><td>{cell(단독_a,단독_ca)}</td><td>{cell(소계_a,소계_ca)}</td>
      <td>{cell(일반_a,일반_ca)}</td><td>{cell(업무_a,업무_ca)}</td>
      <td>{cell(산업_a,산업_ca)}</td><td>{cell(열병_a,열병_ca)}</td><td>{cell(합계_a,합계_ca)}</td>
    </tr>
    <tr>
      <td class="lbl">달성률</td>
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
      <td class="lbl">증감</td>
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
<p style="font-size:12px;color:#555;">※ (괄호)는 누계 기준임.</p>""",
unsafe_allow_html=True)

# ── 산업용 신규 업체 ──
if dev_df is not None:
    ind = dev_df[dev_df['용도'].astype(str).str.contains('산업', na=False)].copy()
    if not ind.empty:
        st.markdown(f'<div class="box-title">🏭 {sel_month}월 산업용 신규 업체 현황</div>',
                    unsafe_allow_html=True)
        rows_html = "".join([
            f"<tr>"
            f"<td>{r['신청명']}</td><td>{r['업종']}</td>"
            f"<td>{fmt(r['월사용예정량'])} ㎥</td>"
            f"<td>{fmt(r['월간개발량'],2)} GJ</td>"
            f"<td>{str(r['공급일'])[:10] if pd.notna(r['공급일']) else '-'}</td>"
            f"<td style='text-align:left;font-size:11px'>{r['주소']}</td>"
            f"</tr>"
            for _, r in ind.iterrows()
        ])
        st.markdown(f"""
        <table>
          <thead><tr>
            <th>업체명</th><th>업종</th><th>월사용예정량</th>
            <th>열량(GJ)</th><th>공급일</th><th>주소</th>
          </tr></thead>
          <tbody>{rows_html}
            <tr style="background:#dce6f5">
              <td colspan="2" style="font-weight:700">합 계</td>
              <td>{fmt(ind['월사용예정량'].sum())} ㎥</td>
              <td>{fmt(ind['월간개발량'].sum(),2)} GJ</td>
              <td colspan="2"></td>
            </tr>
          </tbody>
        </table>""", unsafe_allow_html=True)
    else:
        st.info(f"✅ {sel_month}월 신규 산업용 업체 없음")
else:
    st.warning("② 신규개발량 파일을 업로드하면 산업용 업체 현황이 표시됩니다.")

# ── PDF 출력 안내 ──────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style='text-align:center; padding:16px; background:#f8f9fc; border-radius:8px;'>
  <p style='font-size:15px; font-weight:700; color:#1e3a6b; margin-bottom:8px;'>
    🖨️ 보고서 PDF 저장 방법
  </p>
  <p style='font-size:13px; color:#555; margin:0;'>
    브라우저 메뉴 또는 단축키 <b>Ctrl + P</b> (Mac: ⌘+P)<br>
    → 대상: <b>PDF로 저장</b> &nbsp;→&nbsp; 레이아웃: <b>가로</b> &nbsp;→&nbsp; <b>저장</b>
  </p>
</div>
""", unsafe_allow_html=True)
