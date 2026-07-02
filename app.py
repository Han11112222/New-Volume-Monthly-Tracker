import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re

st.set_page_config(layout="wide", page_title="마케팅본부 월간 보고서")

st.markdown("""
<style>
    .report-title {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d6a9f 100%);
        color: white; padding: 18px 28px; border-radius: 8px;
        font-size: 24px; font-weight: 800; margin-bottom: 20px;
    }
    .section-title {
        background-color: #2d6a9f; color: white;
        padding: 6px 16px; border-radius: 4px;
        font-size: 15px; font-weight: 700; margin: 16px 0 10px 0;
        display: inline-block;
    }
    .input-section {
        background: #f8faff; border: 1px solid #c5d8f5;
        border-radius: 8px; padding: 16px; margin-bottom: 16px;
    }
    .metric-card {
        background: #f0f4fa; border-radius: 6px;
        padding: 12px; text-align: center;
        border-left: 4px solid #2d6a9f; margin-bottom: 8px;
    }
    .metric-label { font-size: 12px; color: #666; }
    .metric-value { font-size: 20px; font-weight: 800; color: #1e3a5f; }
    .metric-rate  { font-size: 12px; color: #888; }
    .new-industry {
        background: #fff3cd; border: 1px solid #ffc107;
        border-radius: 6px; padding: 10px 14px; margin: 6px 0;
        font-size: 13px;
    }
    .upload-guide {
        background: #fff8e1; border: 1px solid #ffc107;
        border-radius: 6px; padding: 10px 14px; margin-bottom: 12px; font-size: 13px;
    }
</style>
""", unsafe_allow_html=True)

# ── 헬퍼 ──────────────────────────────────────
def fmt(v, dec=0):
    if v is None or (isinstance(v, float) and pd.isna(v)): return "-"
    try:
        f = float(v)
        return f"{f:,.{dec}f}" if dec > 0 else f"{int(round(f)):,}"
    except: return str(v)

def pct(v):
    if v is None or (isinstance(v, float) and pd.isna(v)): return "-"
    try:
        f = float(v)
        return f"{f*100:.1f}%" if abs(f) <= 5 else f"{f:.1f}%"
    except: return str(v)

def safe(df, r, c):
    try:
        v = df.iloc[r, c]
        return None if pd.isna(v) else v
    except: return None

# ── 구글 스프레드시트 로드 ──────────────────────
SHEET_ID = '16IUVv9kFPWxu6xVEmwXXvlvJEKrZSdjzdsPdhGOO5BY'
GID      = '1975335172'

@st.cache_data(ttl=300)
def load_gsheet():
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
    try:
        df = pd.read_csv(url, header=None)
        return df, None
    except Exception as e:
        return None, str(e)

# ── 영업일보 xlsm 로드 ─────────────────────────
def load_xlsm(file):
    sheets = {}
    names = ['공급전 계획(2026년)', '공급량 계획_MJ(2026년)', '영업현황']
    for s in names:
        try:
            sheets[s] = pd.read_excel(file, sheet_name=s, header=None, engine='openpyxl')
        except: pass
    return sheets

def parse_supply_plan(sheets):
    """공급전 계획 시트 → 월별 계획 추출"""
    df = sheets.get('공급전 계획(2026년)')
    if df is None: return {}
    # row9: 합계 계획, col4~15: 1~12월
    plan = {}
    cat_rows = {
        '가정용': 4, '일반용': 5, '업무용': 6,
        '산업용': 7, '열병합용': 8, '합계': 9
    }
    for cat, ri in cat_rows.items():
        plan[cat] = [safe(df, ri, c) for c in range(4, 16)]
    return plan

def parse_supply_volume_plan(sheets):
    """공급량 계획 시트 → 월별 공급량 계획(GJ) 추출"""
    df = sheets.get('공급량 계획_MJ(2026년)')
    if df is None: return [], []
    # row5: 월 공급량 계획(MJ), col3~14: 1~12월  → GJ = MJ/1000
    plan_mj = [safe(df, 5, c) for c in range(3, 15)]
    plan_gj = [v/1000 if v else None for v in plan_mj]
    return plan_gj

def parse_actual_supply(sheets):
    """영업현황 시트 → 신규개발전 당월 실적"""
    df = sheets.get('영업현황')
    if df is None: return {}
    # row25: 실적, col2~9: 공동, 단독, 계, 영업용, 업무용, 산업용, 열병합, 총계
    cats = ['공동주택', '단독주택', '소계', '일반용', '업무용', '산업용', '열병합', '합계']
    actual = {}
    for i, cat in enumerate(cats):
        actual[cat] = safe(df, 25, i+2)
    # row26: 목표
    plan_row = {}
    for i, cat in enumerate(cats):
        plan_row[cat] = safe(df, 26, i+2)
    return actual, plan_row

# ════════════════════════════════════════════════
# UI 시작
# ════════════════════════════════════════════════
st.markdown('<div class="report-title">📊 마케팅본부 _ 월간 영업현황 보고서</div>',
            unsafe_allow_html=True)

# ── 파일 업로드 ──
st.markdown('<div class="upload-guide">📁 영업일보(xlsm) 파일만 업로드하면 자동으로 보고서가 생성됩니다.</div>',
            unsafe_allow_html=True)

f_xlsm = st.file_uploader("📎 영업일보 파일 업로드 (영업일보_월간_YYYYMM.xlsm)",
                            type=['xlsm', 'xlsx'], key='xlsm')

# ── 수기 입력 영역 ──
st.markdown('<div class="section-title">✏️ 수기 입력 항목</div>', unsafe_allow_html=True)

with st.container():
    st.markdown('<div class="input-section">', unsafe_allow_html=True)
    col_m, col_y = st.columns([1, 1])
    with col_m:
        input_month = st.selectbox("보고 월", list(range(1, 13)), index=5)
    with col_y:
        input_year = st.number_input("보고 연도", value=2026, step=1)

    st.markdown("**📦 총공급량 실적 입력 (고객지원시스템 확인값)**")
    col1, col2 = st.columns(2)
    with col1:
        total_supply_actual = st.number_input(
            f"{input_month}월 총공급량 당월 실적 (GJ)", 
            value=0, step=1000, format="%d")
    with col2:
        total_supply_cum = st.number_input(
            f"{input_month}월 총공급량 누계 실적 (GJ)",
            value=0, step=1000, format="%d")

    st.markdown("**🏭 산업용 신규 업체 (해당 월에 신규 발생 시 입력)**")
    col_a, col_b, col_c = st.columns([2, 1, 1])
    with col_a:
        new_industry_name = st.text_input("업체명", placeholder="예: ㈜대성에너지 신규")
    with col_b:
        new_industry_m3 = st.number_input("월사용예정량 (m³)", value=0, step=100)
    with col_c:
        new_industry_kcal = st.number_input("열량 환산값 (GJ)", value=0.0, step=0.1, format="%.1f")
    st.markdown('</div>', unsafe_allow_html=True)

if not f_xlsm:
    st.info("👆 영업일보 파일을 업로드해주세요.")
    st.stop()

# ── 데이터 로드 ──
sheets = load_xlsm(f_xlsm)
supply_plan = parse_supply_plan(sheets)
vol_plan_gj = parse_supply_volume_plan(sheets)
actual_supply, plan_supply = parse_actual_supply(sheets)

# 구글 스프레드시트 (산업용 업체 목록)
gsheet_df, gsheet_err = load_gsheet()

month_kor = f"{input_year}년 {input_month}월"
months_label = [f"{i}월" for i in range(1, 13)]

st.markdown(f"### 📅 보고 기준: **{month_kor}**")

# ════════════════════════════════════════════════
# 1. 공급전 및 공급량 현황
# ════════════════════════════════════════════════
st.markdown('<div class="section-title">📋 공급전 및 공급량 현황</div>',
            unsafe_allow_html=True)

# 신규개발전 당월 실적/계획
act = actual_supply or {}
pln = plan_supply  or {}

dev_actual  = act.get('합계', 0) or 0
dev_plan    = pln.get('합계', 0) or 0
dev_rate    = dev_actual / dev_plan * 100 if dev_plan else 0

# 누계 계획 (공급전 계획 시트에서)
df_plan = sheets.get('공급전 계획(2026년)')
cum_plan = 0
if df_plan is not None:
    cum_plan_vals = [safe(df_plan, 12, c) for c in range(4, 4+input_month)]
    cum_plan = sum(v for v in cum_plan_vals if v) 

# 메트릭 카드
mc = st.columns(4)
metrics = [
    ("신규개발전 당월 실적", fmt(dev_actual), f"계획 {fmt(dev_plan)} | 달성 {dev_rate:.1f}%", "전"),
    ("신규개발전 누계 계획", fmt(cum_plan), "", "전"),
    ("총공급량 당월 실적", fmt(total_supply_actual), "수기 입력값", "GJ"),
    ("총공급량 누계 실적", fmt(total_supply_cum), "수기 입력값", "GJ"),
]
for i, (label, val, sub, unit) in enumerate(metrics):
    mc[i].markdown(f"""
    <div class="metric-card">
        <div class="metric-label">{label} ({unit})</div>
        <div class="metric-value">{val}</div>
        <div class="metric-rate">{sub}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# 신규개발전 상세 표
st.markdown(f"**{input_month}월 신규개발전 상세 현황** (단위: 전)")
cats = ['공동주택', '단독주택', '소계', '일반용', '업무용', '산업용', '열병합', '합계']
rows_detail = []
for label, d in [('계획', pln), ('실적', act)]:
    row = {'구분': label}
    for c in cats:
        row[c] = fmt(d.get(c))
    rows_detail.append(row)
# 달성률
rate_row = {'구분': '달성률'}
for c in cats:
    try:
        r = float(act.get(c) or 0) / float(pln.get(c) or 1) * 100
        rate_row[c] = f"{r:.1f}%"
    except: rate_row[c] = "-"
rows_detail.append(rate_row)
# 증감
inc_row = {'구분': '증감'}
for c in cats:
    try:
        inc_row[c] = fmt((act.get(c) or 0) - (pln.get(c) or 0))
    except: inc_row[c] = "-"
rows_detail.append(inc_row)

st.dataframe(pd.DataFrame(rows_detail), use_container_width=True, hide_index=True)

# ════════════════════════════════════════════════
# 2. 산업용 신규 업체 현황
# ════════════════════════════════════════════════
st.markdown('<div class="section-title">🏭 산업용 신규 업체 현황</div>',
            unsafe_allow_html=True)

col_gs, col_new = st.columns([2, 1])

with col_gs:
    st.markdown("**구글 스프레드시트 산업용 업체 목록**")
    if gsheet_err:
        st.warning(f"스프레드시트 접근 오류: {gsheet_err}\n\n스프레드시트 공유 설정을 '링크가 있는 모든 사용자'로 변경해주세요.")
    elif gsheet_df is not None:
        # 업체명이 포함된 행 필터링 (산업용 관련)
        try:
            # 헤더 찾기
            header_idx = -1
            for i, row in gsheet_df.iterrows():
                row_str = " ".join([str(v) for v in row if pd.notna(v)])
                if "업체" in row_str or "고객" in row_str or "산업" in row_str:
                    header_idx = i
                    break
            if header_idx >= 0:
                df_show = pd.read_csv(
                    f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}",
                    header=header_idx)
                st.dataframe(df_show.head(20), use_container_width=True)
            else:
                st.dataframe(gsheet_df.head(15), use_container_width=True)
        except:
            st.dataframe(gsheet_df.head(15), use_container_width=True)

with col_new:
    st.markdown("**이번 달 신규 발생 업체**")
    if new_industry_name:
        st.markdown(f"""
        <div class="new-industry">
            🆕 <b>{new_industry_name}</b><br>
            월사용예정량: <b>{fmt(new_industry_m3)} m³</b><br>
            열량 환산: <b>{fmt(new_industry_kcal, 1)} GJ</b>
        </div>""", unsafe_allow_html=True)
    else:
        st.info("이번 달 신규 산업용 업체 없음")

# ════════════════════════════════════════════════
# 3. 월별 그래프
# ════════════════════════════════════════════════
st.markdown('<div class="section-title">📈 월별 계획 vs 실적 그래프</div>',
            unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📦 총공급량", "👥 신규개발전 카테고리별"])

with tab1:
    plan_vals = [v if v else 0 for v in vol_plan_gj]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=months_label, y=plan_vals,
        name='계획(GJ)', marker_color='#2d6a9f', opacity=0.7
    ))
    if total_supply_actual > 0:
        actual_vals = [0]*12
        actual_vals[input_month-1] = total_supply_actual
        fig.add_trace(go.Bar(
            x=[months_label[input_month-1]],
            y=[total_supply_actual],
            name='실적(GJ)', marker_color='#e74c3c'
        ))
    fig.update_layout(
        title=f"{input_year}년 월별 총공급량 계획 vs 실적",
        barmode='group', height=350,
        yaxis_title="공급량 (GJ)",
        plot_bgcolor='white',
        yaxis=dict(gridcolor='#eee'),
        legend=dict(orientation='h', y=1.1)
    )
    st.plotly_chart(fig, use_container_width=True)

with tab2:
    cats_graph = ['가정용', '일반용', '업무용', '산업용']
    colors = ['#2d6a9f', '#27ae60', '#f39c12', '#e74c3c']
    fig2 = make_subplots(rows=2, cols=2, subplot_titles=cats_graph)
    positions = [(1,1),(1,2),(2,1),(2,2)]
    for idx, (cat, pos, color) in enumerate(zip(cats_graph, positions, colors)):
        p_vals = [v if v else 0 for v in (supply_plan.get(cat) or [0]*12)]
        fig2.add_trace(go.Scatter(
            x=months_label, y=p_vals,
            name=f'{cat} 계획',
            line=dict(color=color, dash='dash'),
            showlegend=(idx==0)
        ), row=pos[0], col=pos[1])
        # 실적 (현재 월만)
        a_val = act.get(cat) if cat in ['일반용','업무용','산업용'] else act.get('소계') if cat=='가정용' else None
        if a_val:
            fig2.add_trace(go.Scatter(
                x=[months_label[input_month-1]],
                y=[float(a_val)],
                mode='markers',
                name=f'{cat} 실적',
                marker=dict(color=color, size=10),
                showlegend=(idx==0)
            ), row=pos[0], col=pos[1])
    fig2.update_layout(
        height=480,
        title=f"{input_year}년 신규개발전 카테고리별 계획 추이",
        plot_bgcolor='white'
    )
    st.plotly_chart(fig2, use_container_width=True)

# 달성률 막대
st.markdown(f"**{input_month}월 신규개발전 달성률**")
rate_cats = ['공동주택', '단독주택', '일반용', '업무용', '산업용', '열병합', '합계']
rate_vals = []
for cat in rate_cats:
    try:
        r = float(act.get(cat) or 0) / float(pln.get(cat) or 1) * 100
        rate_vals.append(round(r, 1))
    except: rate_vals.append(0)

fig3 = go.Figure(go.Bar(
    x=rate_cats, y=rate_vals,
    marker_color=['#27ae60' if r >= 100 else '#e74c3c' for r in rate_vals],
    text=[f"{r}%" for r in rate_vals],
    textposition='outside'
))
fig3.add_hline(y=100, line_dash='dash', line_color='#2d6a9f', annotation_text='목표 100%')
fig3.update_layout(
    height=320, yaxis_title="달성률 (%)",
    plot_bgcolor='white',
    yaxis=dict(gridcolor='#eee', range=[0, max(rate_vals + [120]) + 20])
)
st.plotly_chart(fig3, use_container_width=True)

# ════════════════════════════════════════════════
# 4. PDF 출력
# ════════════════════════════════════════════════
st.markdown("---")
st.markdown("""
<div style='text-align:center; padding:16px;'>
    <p style='color:#555; margin-bottom:6px;'>🖨️ PDF로 출력하려면:</p>
    <p style='color:#888; font-size:13px;'>
        브라우저 메뉴 → 인쇄 (Ctrl+P) → 대상: <b>PDF로 저장</b> → 레이아웃: <b>가로</b> → 저장
    </p>
</div>
""", unsafe_allow_html=True)
if st.button("🖨️ PDF 출력 (Ctrl+P)", use_container_width=False):
    st.info("Ctrl+P 를 누르고 → '대상'을 'PDF로 저장' 선택 → 저장!")
