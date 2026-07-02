import streamlit as st
import pandas as pd
import io
from datetime import datetime

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
.auto-badge {
    background: #27ae60; color: white; font-size: 11px;
    padding: 2px 8px; border-radius: 10px; margin-left: 6px;
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

# ── 구글 스프레드시트에서 공급량 자동 로드 ────────────────
GSHEET_ID = '13HrIz6OytYDykXeXzXJ02I6XbaKin1YaKBoO2kBd6Bs'
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}/export?format=csv&gid=0"

@st.cache_data(ttl=600)
def load_supply_gsheet():
    """구글 스프레드시트에서 일별 공급량 로드 → 월별 합산"""
    try:
        df = pd.read_csv(GSHEET_URL)
        df.columns = ['일자','공급량_MJ','공급량_M3','평균기온','최저기온','최고기온'] + \
                     [f'col{i}' for i in range(len(df.columns)-6)]
        df['일자'] = pd.to_datetime(df['일자'], errors='coerce')
        df = df.dropna(subset=['일자'])
        df['공급량_MJ'] = pd.to_numeric(df['공급량_MJ'].astype(str).str.replace(',',''), errors='coerce')
        df['공급량_GJ'] = df['공급량_MJ'] / 1000
        df['연'] = df['일자'].dt.year
        df['월'] = df['일자'].dt.month
        # 월별 합산
        monthly = df.groupby(['연','월'])['공급량_GJ'].sum().reset_index()
        return df, monthly, None
    except Exception as e:
        return None, None, str(e)

def get_monthly_actual(monthly_df, year, month):
    """특정 연월 공급량 합계(GJ) 반환"""
    if monthly_df is None: return None
    row = monthly_df[(monthly_df['연']==year) & (monthly_df['월']==month)]
    return float(row['공급량_GJ'].values[0]) if len(row) > 0 else None

def get_cum_actual(monthly_df, year, month):
    """1월~해당월 누계 공급량(GJ) 반환"""
    if monthly_df is None: return None
    filtered = monthly_df[(monthly_df['연']==year) & (monthly_df['월']<=month)]
    return float(filtered['공급량_GJ'].sum()) if len(filtered) > 0 else None

# ── 엑셀 파일 로드 ───────────────────────────────────────
@st.cache_data
def load_xlsm(b, fn):
    buf = io.BytesIO(b)
    out = {}
    targets = ['공급전 계획(2026년)', '공급량 계획_MJ(2026년)', '영업현황']
    for s in targets:
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
<div class="upload-box">
📁 <b>매월 2개 파일만 업로드</b>하면 자동으로 보고서가 생성됩니다.<br>
&nbsp;&nbsp;① 영업일보_월간_YYYYMM.xlsm &nbsp;|&nbsp; ② 신규개발량_YYYYMM.xlsx<br>
&nbsp;&nbsp;📡 총공급량 실적은 구글 스프레드시트에서 <b>자동 합산</b>됩니다.
</div>
""", unsafe_allow_html=True)

col_u1, col_u2 = st.columns(2)
with col_u1: f_xlsm = st.file_uploader("① 영업일보 (xlsm)", type=['xlsm','xlsx'])
with col_u2: f_dev  = st.file_uploader("② 신규개발량 (xlsx)", type=['xlsx'])

# ── 보고 월 선택 ──
st.markdown("---")
col_m1, col_m2 = st.columns([1, 3])
with col_m1:
    sel_year  = st.number_input("보고 연도", value=2026, step=1)
    sel_month = st.selectbox("보고 월", list(range(1,13)), index=5)

# ── 구글 스프레드시트 자동 로드 ──
with st.spinner("📡 공급량 데이터 자동 로드 중..."):
    daily_df, monthly_df, gsheet_err = load_supply_gsheet()

auto_actual = get_monthly_actual(monthly_df, int(sel_year), sel_month)
auto_cum    = get_cum_actual(monthly_df, int(sel_year), sel_month)

with col_m2:
    st.markdown("<br>", unsafe_allow_html=True)
    if gsheet_err:
        st.warning(f"⚠️ 구글 스프레드시트 자동 로드 실패: {gsheet_err}\n\n스프레드시트 공유설정을 '링크가 있는 모든 사용자 → 뷰어'로 변경해주세요.")
    elif auto_actual:
        st.success(f"✅ 구글 시트에서 자동 로드: 당월 **{auto_actual:,.0f} GJ** | 누계 **{auto_cum:,.0f} GJ**")
    else:
        st.info(f"ℹ️ {sel_year}년 {sel_month}월 데이터가 구글 시트에 아직 없습니다.")

# ── 수기 입력 (자동값이 디폴트, 수정 가능) ──
st.markdown("**📝 총공급량 실적 확인 / 수정** <span class='auto-badge'>구글시트 자동값</span>",
            unsafe_allow_html=True)

col_s1, col_s2, col_s3 = st.columns([2, 2, 1])
with col_s1:
    total_actual = st.number_input(
        f"총공급량 당월 실적 (GJ)",
        value=int(round(auto_actual)) if auto_actual else 0,
        step=1000, format="%d",
        help="구글 스프레드시트에서 자동 계산된 값. 필요 시 수정 가능.")
with col_s2:
    total_cum = st.number_input(
        f"총공급량 누계 실적 (GJ)",
        value=int(round(auto_cum)) if auto_cum else 0,
        step=1000, format="%d",
        help="1월~당월 누계 자동 합산값. 필요 시 수정 가능.")
with col_s3:
    st.markdown("<br>", unsafe_allow_html=True)
    confirmed = st.button("✅ 확인 / 적용", use_container_width=True)

if confirmed:
    st.success(f"✅ 당월 실적: **{total_actual:,} GJ** | 누계: **{total_cum:,} GJ** 적용됨")

if not f_xlsm:
    st.info("👆 영업일보 파일을 먼저 업로드해주세요.")
    st.stop()

# ════════════════════════════════════════════════════
# 데이터 파싱
# ════════════════════════════════════════════════════
sheets = load_xlsm(f_xlsm.read(), f_xlsm.name); f_xlsm.seek(0)
dev_df = load_dev(f_dev.read(), f_dev.name) if f_dev else None

df_plan = sheets.get('공급전 계획(2026년)')
df_vol  = sheets.get('공급량 계획_MJ(2026년)')
df_biz  = sheets.get('영업현황')

m  = sel_month - 1
mc = m + 4   # 공급전계획: col4=1월

def p(df, r, c): return safe(df, r, c)

# ── 신규개발전 계획 ──
신규_연간 = p(df_plan, 9, 17)
신규_당계 = p(df_plan, 9, mc)
신규_누계 = p(df_plan, 12, mc)
공동_p = p(df_plan, 18, mc); 단독_p = p(df_plan, 19, mc)
소계_p = (공동_p or 0)+(단독_p or 0)
일반_p = p(df_plan, 5, mc);  업무_p = p(df_plan, 6, mc)
산업_p = p(df_plan, 7, mc);  열병_p = p(df_plan, 8, mc)
합계_p = p(df_plan, 9, mc)
공동_cp = p(df_plan, 22, mc); 단독_cp = p(df_plan, 23, mc)
소계_cp = (공동_cp or 0)+(단독_cp or 0)

# ── 공급량 계획 (MJ→GJ) ──
vol_mc = m + 3
총공_당계_p = p(df_vol, 5, vol_mc)
총공_누계_p = p(df_vol, 6, vol_mc)
총공_연간_p = p(df_vol, 5, 16)
if 총공_당계_p: 총공_당계_p /= 1000
if 총공_누계_p: 총공_누계_p /= 1000
if 총공_연간_p: 총공_연간_p /= 1000

# ── 실적 ──
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

공동_a=act.get('공동주택'); 단독_a=act.get('단독주택')
소계_a=act.get('소계');    일반_a=act.get('일반용')
업무_a=act.get('업무용');  산업_a=act.get('산업용')
열병_a=act.get('열병합');  합계_a=act.get('합계')
총공_당실 = total_actual if total_actual > 0 else None
총공_누실 = total_cum    if total_cum    > 0 else None
입력필요  = '<span style="color:#aaa;font-size:11px;">입력필요</span>'

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
      <td>{fmt(신규_누계)}</td><td>-</td><td>-</td>
    </tr>
    <tr>
      <td class="td-sub-label">폐전</td>
      <td>-</td><td>-</td><td>{fmt(폐전_당실)}</td><td>-</td>
      <td>-</td><td>{fmt(폐전_누실)}</td><td>-</td>
    </tr>
    <tr>
      <td class="td-sub-label">순증가</td>
      <td>-</td><td>-</td><td>{fmt(순증_당실)}</td><td>-</td>
      <td>-</td><td>-</td><td>-</td>
    </tr>
    <tr>
      <td class="td-label" rowspan="2">공급량<br>(GJ)</td>
      <td class="td-sub-label">신규개발량</td>
      <td>-</td>
      <td>{fmt(총공_당계_p)}</td><td>-</td><td>-</td>
      <td>{fmt(총공_누계_p)}</td><td>-</td><td>-</td>
    </tr>
    <tr>
      <td class="td-sub-label">총공급량</td>
      <td>{fmt(총공_연간_p)}</td>
      <td>{fmt(총공_당계_p)}</td>
      <td>{'<b>'+fmt(총공_당실)+'</b>' if 총공_당실 else 입력필요}</td>
      <td>{rate_html(총공_당실, 총공_당계_p) if 총공_당실 else '-'}</td>
      <td>{fmt(총공_누계_p)}</td>
      <td>{'<b>'+fmt(총공_누실)+'</b>' if 총공_누실 else 입력필요}</td>
      <td>{rate_html(총공_누실, 총공_누계_p) if 총공_누실 else '-'}</td>
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
      <td>{fmt(공동_p)}</td><td>{fmt(단독_p)}</td><td>{fmt(소계_p)}</td>
      <td>{fmt(일반_p)}</td><td>{fmt(업무_p)}</td>
      <td>{fmt(산업_p)}</td><td>{fmt(열병_p)}</td><td>{fmt(합계_p)}</td>
    </tr>
    <tr>
      <td class="td-label">실적</td>
      <td>{fmt(공동_a)}</td><td>{fmt(단독_a)}</td><td>{fmt(소계_a)}</td>
      <td>{fmt(일반_a)}</td><td>{fmt(업무_a)}</td>
      <td>{fmt(산업_a)}</td><td>{fmt(열병_a)}</td><td>{fmt(합계_a)}</td>
    </tr>
    <tr>
      <td class="td-label">달성률</td>
      <td>{rate_html(공동_a,공동_p)}</td><td>{rate_html(단독_a,단독_p)}</td>
      <td>{rate_html(소계_a,소계_p)}</td>
      <td>{rate_html(일반_a,일반_p)}</td><td>{rate_html(업무_a,업무_p)}</td>
      <td>{rate_html(산업_a,산업_p)}</td><td>{rate_html(열병_a,열병_p)}</td>
      <td>{rate_html(합계_a,합계_p)}</td>
    </tr>
    <tr>
      <td class="td-label">증감</td>
      <td>{d_inc(공동_a,공동_p)}</td><td>{d_inc(단독_a,단독_p)}</td>
      <td>{d_inc(소계_a,소계_p)}</td>
      <td>{d_inc(일반_a,일반_p)}</td><td>{d_inc(업무_a,업무_p)}</td>
      <td>{d_inc(산업_a,산업_p)}</td><td>{d_inc(열병_a,열병_p)}</td>
      <td>{d_inc(합계_a,합계_p)}</td>
    </tr>
  </tbody>
</table>
<p style="font-size:12px;color:#555;">※ (괄호)는 누계 기준임.</p>"""
st.markdown(html2, unsafe_allow_html=True)

# ── 신규개발량 용도별 ──
if dev_df is not None:
    st.markdown(f'<div class="box-title">📦 {sel_month}월 신규개발량 용도별 현황</div>',
                unsafe_allow_html=True)
    pivot = dev_df.groupby('용도').agg(
        건수=('월간개발량','count'),
        월간개발량=('월간개발량','sum'),
        월사용예정량=('월사용예정량','sum')
    ).reset_index()
    p_rows = "".join([
        f"<tr><td>{r['용도']}</td><td>{int(r['건수'])}</td>"
        f"<td>{fmt(r['월사용예정량'])} ㎥</td><td>{fmt(r['월간개발량'],2)} GJ</td></tr>"
        for _, r in pivot.iterrows()
    ])
    st.markdown(f"""
    <table>
      <thead><tr><th>용도</th><th>건수</th><th>월사용예정량(㎥)</th><th>월간개발량(GJ)</th></tr></thead>
      <tbody>{p_rows}
        <tr style="background:#dce6f5;font-weight:700;">
          <td>합 계</td><td>{len(dev_df)}</td>
          <td>{fmt(dev_df['월사용예정량'].sum())} ㎥</td>
          <td>{fmt(dev_df['월간개발량'].sum(),2)} GJ</td>
        </tr>
      </tbody>
    </table>""", unsafe_allow_html=True)

    # 산업용 업체 상세
    industry_df = dev_df[dev_df['용도'].astype(str).str.contains('산업', na=False)].copy()
    if not industry_df.empty:
        st.markdown(f'<div class="box-title">🏭 {sel_month}월 산업용 신규 업체 현황</div>',
                    unsafe_allow_html=True)
        ind_rows = "".join([
            f"<tr><td>{r['신청명']}</td><td>{r['업종']}</td>"
            f"<td>{fmt(r['월사용예정량'])} ㎥</td><td>{fmt(r['월간개발량'],2)} GJ</td>"
            f"<td>{str(r['공급일'])[:10] if pd.notna(r['공급일']) else '-'}</td>"
            f"<td style='font-size:11px;text-align:left'>{r['주소']}</td></tr>"
            for _, r in industry_df.iterrows()
        ])
        st.markdown(f"""
        <table>
          <thead><tr><th>업체명</th><th>업종</th><th>월사용예정량</th>
          <th>열량(GJ)</th><th>공급일</th><th>주소</th></tr></thead>
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
    st.warning("② 신규개발량 파일을 업로드하면 용도별/산업용 현황이 표시됩니다.")

# ── 일별 공급량 참고 (월별 상세 확인용) ──
if daily_df is not None and st.checkbox("📊 일별 공급량 상세 보기 (구글시트 원본)"):
    filtered = daily_df[
        (daily_df['연']==int(sel_year)) & (daily_df['월']==sel_month)
    ][['일자','공급량_MJ','공급량_GJ','평균기온']].copy()
    filtered['일자'] = filtered['일자'].dt.strftime('%Y-%m-%d')
    filtered['공급량_MJ'] = filtered['공급량_MJ'].apply(lambda x: f"{x:,.0f}")
    filtered['공급량_GJ'] = filtered['공급량_GJ'].apply(lambda x: f"{x:,.1f}")
    st.dataframe(filtered, use_container_width=True, hide_index=True)

# ── PDF 출력 ──
st.markdown("---")
st.markdown("""
<div style='text-align:center;padding:12px;color:#555;font-size:13px;'>
🖨️ <b>PDF 출력</b>: Ctrl+P → 대상: PDF로 저장 → 레이아웃: 가로 → 저장
</div>""", unsafe_allow_html=True)
