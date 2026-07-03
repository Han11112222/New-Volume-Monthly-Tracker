import streamlit as st
import pandas as pd
import io
import urllib.request
import urllib.parse
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              numbers as xl_numbers)
from openpyxl.utils import get_column_letter

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
td.lbl  { background: #dce6f5 !important; font-weight: 700; color: #1e3a6b; }
td.slbl { background: #eef2fa !important; font-weight: 600; color: #333; }
</style>
""", unsafe_allow_html=True)

# ── 헬퍼 ──────────────────────────────────────────────
NUM  = "style='color:#222;font-size:13px;font-weight:normal'"
SNUM = "style='color:#444;font-size:11px;font-weight:normal'"

def fmt(v, dec=0):
    if v is None or (isinstance(v, float) and pd.isna(v)): return "-"
    try:
        f = float(v)
        return f"{f:,.{dec}f}" if dec > 0 else f"{int(round(f)):,}"
    except: return str(v)

def fv(v):
    """순수 숫자값 반환 (엑셀용)"""
    if v is None or (isinstance(v, float) and pd.isna(v)): return None
    try: return float(v)
    except: return None

def n(v, dec=0):
    return f"<span {NUM}>{fmt(v,dec)}</span>"

def rate_html(a, p):
    try:
        val = float(a or 0) / float(p or 1) * 100
        col = "#1a7a1a" if val >= 100 else "#c0392b"
        return f"<span style='color:{col};font-size:13px;font-weight:normal'>{val:.1f}%</span>"
    except: return "<span style='color:#222;font-size:13px'>-</span>"

def rate_val(a, p):
    try: return float(a or 0) / float(p or 1) * 100
    except: return None

def safe(df, r, c):
    try:
        v = df.iloc[r, c]
        return None if (isinstance(v, float) and pd.isna(v)) else v
    except: return None

def d_inc(a, p):
    try: return fmt(float(a or 0) - float(p or 0))
    except: return "-"

def d_inc_v(a, p):
    try: return float(a or 0) - float(p or 0)
    except: return None

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
            "sheets": ["3_1. 개발량 계획", "(회의자료 입력용)공급전 및 공급량 현황"]
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
            for sh in info["sheets"]:
                try:
                    result[key][sh] = pd.read_excel(buf, sheet_name=sh, header=None)
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
    for sh in ['영업일보','공급전 계획(2026년)','공급량 계획_MJ(2026년)']:
        try: out[sh] = pd.read_excel(buf, sheet_name=sh, header=None, engine='openpyxl'); buf.seek(0)
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

# ── 엑셀 다운로드 생성 ────────────────────────────────
def make_excel(yr, mo, data1, data2, ind_df):
    wb = Workbook()
    ws = wb.active
    ws.title = f"{yr}년{mo}월 영업현황"

    # 스타일 정의
    hdr_fill   = PatternFill("solid", fgColor="1E3A6B")
    sub_fill   = PatternFill("solid", fgColor="2D5FA8")
    lbl_fill   = PatternFill("solid", fgColor="DCE6F5")
    slbl_fill  = PatternFill("solid", fgColor="EEF2FA")
    hdr_font   = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=10)
    lbl_font   = Font(name="맑은 고딕", bold=True, color="1E3A6B", size=10)
    slbl_font  = Font(name="맑은 고딕", bold=False, color="333333", size=10)
    num_font   = Font(name="맑은 고딕", size=10)
    red_font   = Font(name="맑은 고딕", size=10, color="C0392B")
    grn_font   = Font(name="맑은 고딕", size=10, color="1A7A1A")
    center     = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin       = Side(style="thin", color="CCCCCC")
    border     = Border(left=thin, right=thin, top=thin, bottom=thin)
    num_fmt    = '#,##0'
    pct_fmt    = '0.0%'

    def hc(ws, row, col, val, fill=None, font=None, fmt=None):
        cell = ws.cell(row=row, column=col, value=val)
        cell.alignment = center
        cell.border    = border
        if fill: cell.fill = fill
        if font: cell.font = font
        else:    cell.font = num_font
        if fmt:  cell.number_format = fmt
        return cell

    def rate_font(a, p):
        try:
            v = float(a or 0)/float(p or 1)
            return grn_font if v >= 1 else red_font
        except: return num_font

    # ── 표1 헤더 ──
    ws.merge_cells('A1:B2'); hc(ws,1,1,f"구 분",hdr_fill,hdr_font)
    ws.merge_cells('C1:C2'); hc(ws,1,3,"연간계획",hdr_fill,hdr_font)
    ws.merge_cells('D1:F1'); hc(ws,1,4,f"{mo}월 (당월)",hdr_fill,hdr_font)
    ws.merge_cells('G1:I1'); hc(ws,1,7,f"{mo}월 (누계)",hdr_fill,hdr_font)
    for c,v in [(4,"계획"),(5,"실적"),(6,"달성률"),(7,"계획"),(8,"실적"),(9,"달성률")]:
        hc(ws,2,c,v,sub_fill,hdr_font)

    # ── 표1 데이터 ──
    rows1 = data1  # list of tuples
    r = 3
    for i, row in enumerate(rows1):
        구분1, 구분2, 연간, 당계, 당실, 누계, 누실 = row
        if 구분1:
            end = r + (2 if 구분1=="공급전\n(전)" else 1)
            ws.merge_cells(start_row=r, start_column=1, end_row=end, end_column=1)
            hc(ws,r,1,구분1,lbl_fill,lbl_font)
        hc(ws,r,2,구분2,slbl_fill,slbl_font)
        hc(ws,r,3,fv(연간),None,num_font,num_fmt)
        hc(ws,r,4,fv(당계),None,num_font,num_fmt)
        hc(ws,r,5,fv(당실),None,num_font,num_fmt)
        rv1 = rate_val(당실, 당계)
        hc(ws,r,6,rv1/100 if rv1 is not None else None,None,rate_font(당실,당계),pct_fmt)
        hc(ws,r,7,fv(누계),None,num_font,num_fmt)
        hc(ws,r,8,fv(누실),None,num_font,num_fmt)
        rv2 = rate_val(누실, 누계)
        hc(ws,r,9,rv2/100 if rv2 is not None else None,None,rate_font(누실,누계),pct_fmt)
        r += 1

    # ── 표2 헤더 (2행 아래) ──
    r += 1
    title_row = r
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=9)
    hc(ws,r,1,f"{mo}월 신규개발전 상세 현황 (단위: 전)",hdr_fill,hdr_font)
    r += 1
    hc(ws,r,1,"구 분",hdr_fill,hdr_font)
    ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=4)
    hc(ws,r,2,"주택용",hdr_fill,hdr_font)
    for c,v in [(5,"일반용"),(6,"업무용"),(7,"산업용"),(8,"열병합"),(9,"합계")]:
        hc(ws,r,c,v,hdr_fill,hdr_font)
    r += 1
    for c,v in [(1,""),(2,"공동주택"),(3,"단독주택"),(4,"소계"),
                (5,""),(6,""),(7,""),(8,""),(9,"")]:
        hc(ws,r,c,v,sub_fill,hdr_font)

    # ── 표2 데이터 ──
    r += 1
    for row in data2:
        구분 = row[0]
        vals = row[1:]
        hc(ws,r,1,구분,lbl_fill,lbl_font)
        for ci, (당월_v, 누계_v) in enumerate(zip(vals[::2], vals[1::2])):
            col = ci + 2
            if 구분 == "달성률":
                rv = rate_val(당월_v, None)  # 이미 비율값
                cell_obj = hc(ws,r,col,
                    당월_v/100 if 당월_v is not None else None,
                    None, rate_font(당월_v, 100), pct_fmt)
            elif 구분 == "증감":
                hc(ws,r,col,fv(당월_v),None,num_font,num_fmt)
            else:
                hc(ws,r,col,fv(당월_v),None,num_font,num_fmt)
        r += 1

    # ── 산업용 업체 (있으면) ──
    if ind_df is not None and not ind_df.empty:
        r += 1
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=6)
        hc(ws,r,1,f"{mo}월 산업용 신규 업체 현황",hdr_fill,hdr_font)
        r += 1
        for c,v in enumerate(["업체명","업종","월사용예정량(㎥)","열량(GJ)","공급일","주소"],1):
            hc(ws,r,c,v,sub_fill,hdr_font)
        r += 1
        for _, row in ind_df.iterrows():
            hc(ws,r,1,row['신청명'],None,num_font)
            hc(ws,r,2,row['업종'],None,num_font)
            hc(ws,r,3,fv(row['월사용예정량']),None,num_font,num_fmt)
            hc(ws,r,4,fv(row['월간개발량']),None,num_font,'#,##0.00')
            공일 = str(row['공급일'])[:10] if pd.notna(row['공급일']) else '-'
            hc(ws,r,5,공일,None,num_font)
            c6 = ws.cell(row=r,column=6,value=row['주소'])
            c6.alignment = Alignment(horizontal="left",vertical="center")
            c6.border = border; c6.font = num_font
            r += 1
        # 합계
        hc(ws,r,1,"합 계",lbl_fill,lbl_font)
        hc(ws,r,2,"",lbl_fill,lbl_font)
        hc(ws,r,3,float(ind_df['월사용예정량'].sum()),lbl_fill,lbl_font,num_fmt)
        hc(ws,r,4,float(ind_df['월간개발량'].sum()),lbl_fill,lbl_font,'#,##0.00')

    # 열 너비
    widths = [10,14,12,12,12,12,12,12,12]
    for i,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ════════════════════════════════════════════════════
# UI
# ════════════════════════════════════════════════════
st.markdown('<div class="report-header">📊 마케팅본부 _ 월간 영업현황 보고서</div>',
            unsafe_allow_html=True)

st.markdown("""
<div class="upload-box">
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

if not f_xlsm:
    st.info("👆 영업일보 파일을 업로드해주세요.")
    st.stop()

# ── 데이터 파싱 ──────────────────────────────────────
sheets = load_xlsm(f_xlsm.read(), f_xlsm.name); f_xlsm.seek(0)
dev_df = load_dev(f_dev.read(), f_dev.name) if f_dev else None

df_il   = sheets.get('영업일보')
df_vol  = sheets.get('공급량 계획_MJ(2026년)')
df_n1p  = gh.get("new_1",{}).get("3_1. 개발량 계획")
df_n1rt = gh.get("new_1",{}).get("(회의자료 입력용)공급전 및 공급량 현황")
df_n2p  = gh.get("new_2",{}).get("3_1. 개발량 계획")
df_n2r  = gh.get("new_2",{}).get("3_2. 개발량 실적")

m  = sel_month - 1
nc = m + 2
vc = m + 3

def s(df, r, c): return safe(df, r, c) if df is not None else None

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

# ── 화면 출력 ──────────────────────────────────────
st.markdown(f"### 📅 **{int(sel_year)}년 {sel_month}월** 영업현황 보고")

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
      <td>{n(신규_연간)}</td><td>{n(신규_당계)}</td><td>{n(신규_당실)}</td>
      <td>{rate_html(신규_당실,신규_당계)}</td>
      <td>{n(신규_누계)}</td><td>{n(신규_누실)}</td>
      <td>{rate_html(신규_누실,신규_누계)}</td>
    </tr>
    <tr>
      <td class="slbl">폐전</td>
      <td>{n(폐전_연간)}</td><td>{n(폐전_당계)}</td><td>{n(폐전_당실)}</td>
      <td>{rate_html(폐전_당실,폐전_당계)}</td>
      <td>{n(폐전_누계)}</td><td>{n(폐전_누실)}</td>
      <td>{rate_html(폐전_누실,폐전_누계)}</td>
    </tr>
    <tr>
      <td class="slbl">순증가</td>
      <td>{n(순증_연간)}</td><td>{n(순증_당계)}</td><td>{n(순증_당실)}</td>
      <td>{rate_html(순증_당실,순증_당계)}</td>
      <td>{n(순증_누계)}</td><td>{n(순증_누실)}</td>
      <td>{rate_html(순증_누실,순증_누계)}</td>
    </tr>
    <tr>
      <td class="lbl" rowspan="2">공급량<br>(GJ)</td>
      <td class="slbl">신규개발량</td>
      <td>{n(개발량_연간)}</td><td>{n(개발량_당계)}</td>
      <td>{n(개발량_당실) if 개발량_당실 else 입력필요}</td>
      <td>{rate_html(개발량_당실,개발량_당계) if 개발량_당실 else n(None)}</td>
      <td>{n(개발량_누계)}</td>
      <td>{n(개발량_누실) if 개발량_누실 else n(None)}</td>
      <td>{rate_html(개발량_누실,개발량_누계) if 개발량_누실 else n(None)}</td>
    </tr>
    <tr>
      <td class="slbl">총공급량</td>
      <td>{n(총공_연간)}</td><td>{n(총공_당계)}</td>
      <td>{n(총공_당실_v) if 총공_당실_v else 입력필요}</td>
      <td>{rate_html(총공_당실_v,총공_당계) if 총공_당실_v else n(None)}</td>
      <td>{n(총공_누계)}</td>
      <td>{n(총공_누실_v) if 총공_누실_v else 입력필요}</td>
      <td>{rate_html(총공_누실_v,총공_누계) if 총공_누실_v else n(None)}</td>
    </tr>
  </tbody>
</table>""", unsafe_allow_html=True)

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
ind = None
if dev_df is not None:
    ind = dev_df[dev_df['용도'].astype(str).str.contains('산업', na=False)].copy()
    if not ind.empty:
        st.markdown(f'<div class="box-title">🏭 {sel_month}월 산업용 신규 업체 현황</div>',
                    unsafe_allow_html=True)
        rows_html = "".join([
            f"<tr><td>{r['신청명']}</td><td>{r['업종']}</td>"
            f"<td>{fmt(r['월사용예정량'])} ㎥</td>"
            f"<td>{fmt(r['월간개발량'],2)} GJ</td>"
            f"<td>{str(r['공급일'])[:10] if pd.notna(r['공급일']) else '-'}</td>"
            f"<td style='text-align:left;font-size:11px'>{r['주소']}</td></tr>"
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
        ind = None
else:
    st.warning("② 신규개발량 파일을 업로드하면 산업용 업체 현황이 표시됩니다.")

# ── 엑셀 다운로드 버튼 (하단) ──────────────────────────
st.markdown("---")

# 엑셀 데이터 준비
data1 = [
    ("공급전\n(전)", "신규개발전", 신규_연간, 신규_당계, 신규_당실, 신규_누계, 신규_누실),
    (None,          "폐전",       폐전_연간, 폐전_당계, 폐전_당실, 폐전_누계, 폐전_누실),
    (None,          "순증가",     순증_연간, 순증_당계, 순증_당실, 순증_누계, 순증_누실),
    ("공급량\n(GJ)", "신규개발량", 개발량_연간, 개발량_당계, 개발량_당실, 개발량_누계, 개발량_누실),
    (None,           "총공급량",  총공_연간, 총공_당계, 총공_당실_v, 총공_누계, 총공_누실_v),
]
data2 = [
    ("계획",  공동_p, 공동_cp, 단독_p, 단독_cp, 소계_p, 소계_cp,
               일반_p, 일반_cp, 업무_p, 업무_cp, 산업_p, 산업_cp, 열병_p, 열병_cp, 합계_p, 합계_cp),
    ("실적",  공동_a, 공동_ca, 단독_a, 단독_ca, 소계_a, 소계_ca,
               일반_a, 일반_ca, 업무_a, 업무_ca, 산업_a, 산업_ca, 열병_a, 열병_ca, 합계_a, 합계_ca),
    ("달성률", rate_val(공동_a,공동_p), rate_val(공동_ca,공동_cp),
               rate_val(단독_a,단독_p), rate_val(단독_ca,단독_cp),
               rate_val(소계_a,소계_p), rate_val(소계_ca,소계_cp),
               rate_val(일반_a,일반_p), rate_val(일반_ca,일반_cp),
               rate_val(업무_a,업무_p), rate_val(업무_ca,업무_cp),
               rate_val(산업_a,산업_p), rate_val(산업_ca,산업_cp),
               rate_val(열병_a,열병_p), rate_val(열병_ca,열병_cp),
               rate_val(합계_a,합계_p), rate_val(합계_ca,합계_cp)),
    ("증감",  d_inc_v(공동_a,공동_p), d_inc_v(공동_ca,공동_cp),
               d_inc_v(단독_a,단독_p), d_inc_v(단독_ca,단독_cp),
               d_inc_v(소계_a,소계_p), d_inc_v(소계_ca,소계_cp),
               d_inc_v(일반_a,일반_p), d_inc_v(일반_ca,일반_cp),
               d_inc_v(업무_a,업무_p), d_inc_v(업무_ca,업무_cp),
               d_inc_v(산업_a,산업_p), d_inc_v(산업_ca,산업_cp),
               d_inc_v(열병_a,열병_p), d_inc_v(열병_ca,열병_cp),
               d_inc_v(합계_a,합계_p), d_inc_v(합계_ca,합계_cp)),
]

xl_bytes = make_excel(int(sel_year), sel_month, data1, data2, ind)
fname = f"{int(sel_year)}년{sel_month}월_영업현황보고.xlsx"

col_dl = st.columns([2,1,2])
with col_dl[1]:
    st.download_button(
        label="📥 엑셀 파일로 저장",
        data=xl_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )
