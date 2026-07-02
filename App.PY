import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import glob
import os
import re

st.set_page_config(layout="wide", page_title="마케팅본부 월간 보고서")

st.markdown("""
<style>
    body { font-family: 'Malgun Gothic', sans-serif; }
    .report-title {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d6a9f 100%);
        color: white; padding: 20px 30px; border-radius: 8px;
        font-size: 26px; font-weight: 800; margin-bottom: 24px;
    }
    .section-box {
        border: 2px solid #2d6a9f; border-radius: 8px;
        padding: 16px; margin-bottom: 20px;
    }
    .section-title {
        background-color: #2d6a9f; color: white;
        padding: 6px 16px; border-radius: 4px;
        font-size: 15px; font-weight: 700; margin-bottom: 12px;
        display: inline-block;
    }
    .metric-card {
        background: #f0f4fa; border-radius: 6px;
        padding: 12px 16px; text-align: center;
        border-left: 4px solid #2d6a9f;
    }
    .metric-label { font-size: 12px; color: #666; margin-bottom: 4px; }
    .metric-value { font-size: 20px; font-weight: 800; color: #1e3a5f; }
    .metric-rate { font-size: 13px; color: #888; }
    .table-header {
        background-color: #2d6a9f; color: white;
        font-weight: 700; text-align: center;
    }
    .rate-good { color: #1a7a1a; font-weight: 700; }
    .rate-bad  { color: #c0392b; font-weight: 700; }
    .upload-guide {
        background: #fff8e1; border: 1px solid #ffc107;
        border-radius: 6px; padding: 12px 16px; margin-bottom: 16px;
        font-size: 13px;
    }
</style>
""", unsafe_allow_html=True)

# ── 헬퍼 ──────────────────────────────────────────────
def fmt_num(v, decimals=0):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    try:
        if decimals == 0:
            return f"{int(round(float(v))):,}"
        return f"{float(v):,.{decimals}f}"
    except:
        return str(v)

def fmt_pct(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    try:
        f = float(v)
        if abs(f) > 5:   # 이미 % 단위
            return f"{f:.1f}%"
        return f"{f*100:.1f}%"
    except:
        return str(v)

def rate_color(v):
    try:
        f = float(v)
        val = f * 100 if abs(f) <= 5 else f
        return "rate-good" if val >= 100 else "rate-bad"
    except:
        return ""

def safe(df, row, col):
    try:
        v = df.iloc[row, col]
        return None if pd.isna(v) else v
    except:
        return None

# ── 데이터 읽기 ────────────────────────────────────────
def load_report(file):
    """영업현황 보고 파일 → 주요 지표 추출"""
    xl = pd.read_excel(file, sheet_name='(회의자료 입력용)공급전 및 공급량 현황',
                       header=None)
    # row 6=신규개발전, 7=폐전, 8=순증가, 9=신규개발량, 10=총공급량
    # col: 2=연간계획, 3=당월계획, 4=당월실적, 5=달성률, 6=누계계획, 7=누계실적, 8=누계달성률, 9=익월계획
    rows = {
        '신규개발전': 6, '폐전': 7, '순증가': 8,
        '신규개발량': 9, '총공급량': 10
    }
    cols = {
        '연간계획': 2, '당월계획': 3, '당월실적': 4, '달성률': 5,
        '누계계획': 6, '누계실적': 7, '누계달성률': 8, '익월계획': 9
    }
    data = {}
    for rname, ri in rows.items():
        data[rname] = {cname: safe(xl, ri, ci) for cname, ci in cols.items()}

    # 신규개발전 상세 (당월) — row14~17 / col15~22  (N~W)
    # 행: 계획=6, 실적=7, 달성률=8, 증감=9  /  col: 공동=15, 단독=16, 소계=17, 일반=18, 업무=19, 산업=20, 열병합=21, 합계=22
    detail_rows = {'계획': 6, '실적': 7, '달성률': 8, '증감': 9}
    detail_cols = {'공동주택': 15, '단독주택': 16, '소계': 17,
                   '일반용': 18, '업무용': 19, '산업용': 20, '열병합': 21, '합계': 22}
    detail = {}
    for rn, ri in detail_rows.items():
        detail[rn] = {cn: safe(xl, ri, ci) for cn, ci in detail_cols.items()}

    # 누계기준 신규개발전 상세 — row26~33
    cum_rows = {'계획': 26, '실적': 28, '달성률': 30, '증감': 32}
    cum = {}
    for rn, ri in cum_rows.items():
        cum[rn] = {cn: safe(xl, ri, ci) for cn, ci in detail_cols.items()}

    # 월별 공급량 계획 (row38) / 누계 (row39)
    monthly_plan = [safe(xl, 38, c) for c in range(1, 13)]
    monthly_cum  = [safe(xl, 39, c) for c in range(1, 13)]

    # 보고 월 추출 (파일명에서)
    fname = file.name if hasattr(file, 'name') else str(file)
    m = re.search(r'(\d{6})', fname)
    yyyymm = m.group(1) if m else "202606"
    month = int(yyyymm[4:6])

    return data, detail, cum, monthly_plan, monthly_cum, month, yyyymm


def load_monthly_plan(file):
    """신규개발량 파일 → 공급전 월별 계획/실적"""
    xl = pd.read_excel(file, sheet_name='3_1. 개발량 계획', header=None)
    months = list(range(1, 13))
    categories = {
        '공동주택': 2, '단독주택': 4, '가정용': 6,
        '일반용': 7, '업무용': 9, '산업용': 11, '열병합용': 13
    }
    plan = {}
    for cat, ri in categories.items():
        plan[cat] = [safe(xl, ri, c) for c in range(2, 14)]

    xl2 = pd.read_excel(file, sheet_name='3_2. 개발량 실적', header=None)
    actual = {}
    for cat, ri in categories.items():
        actual[cat] = [safe(xl2, ri, c) for c in range(2, 14)]

    return plan, actual, months


# ── UI ────────────────────────────────────────────────
st.markdown('<div class="report-title">📊 마케팅본부 _ 월간 영업현황 보고서</div>',
            unsafe_allow_html=True)

# 파일 업로드
st.markdown('<div class="upload-guide">📁 매월 초 아래 파일을 업로드하면 자동으로 보고서가 생성됩니다.</div>',
            unsafe_allow_html=True)

col_u1, col_u2 = st.columns(2)
with col_u1:
    f_report = st.file_uploader(
        "① 영업현황 보고 파일 (new_1_...xlsx)",
        type=['xlsx'], key='report')
with col_u2:
    f_devplan = st.file_uploader(
        "② 신규개발량 파일 (new_2_...xlsx)",
        type=['xlsx'], key='devplan')

if not f_report:
    st.info("👆 영업현황 보고 파일을 업로드해주세요.")
    st.stop()

# ── 데이터 로드 ──
data, detail, cum, monthly_plan, monthly_cum, month, yyyymm = load_report(f_report)
year = int(yyyymm[:4])

plan_data, actual_data, months = None, None, None
if f_devplan:
    try:
        plan_data, actual_data, months = load_monthly_plan(f_devplan)
    except:
        pass

month_kor = f"{year}년 {month}월"
st.markdown(f"### 📅 보고 기준: **{month_kor}**")

# ════════════════════════════════════════════════
# 1. 공급전 및 공급량 현황 요약
# ════════════════════════════════════════════════
st.markdown('<div class="section-title">📋 공급전 및 공급량 현황</div>', unsafe_allow_html=True)

# 상단 메트릭 카드
mc = st.columns(5)
items = [
    ("신규개발전", "신규개발전", "전"),
    ("폐전", "폐전", "전"),
    ("순증가", "순증가", "전"),
    ("신규개발량", "신규개발량", "GJ"),
    ("총공급량", "총공급량", "GJ"),
]
for i, (label, key, unit) in enumerate(items):
    d = data.get(key, {})
    실적 = d.get('당월실적')
    달성 = d.get('달성률')
    누계 = d.get('누계실적')
    mc[i].markdown(f"""
    <div class="metric-card">
        <div class="metric-label">{label} ({unit})</div>
        <div class="metric-value">{fmt_num(실적)}</div>
        <div class="metric-rate">당월 달성 {fmt_pct(달성)} | 누계 {fmt_num(누계)}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# 상세 표
def render_summary_table(data_dict, title_prefix=""):
    cols_def = ['구분', '연간계획', f'{month}월 계획', f'{month}월 실적', '달성률', '누계 계획', '누계 실적', '누계 달성률']
    rows_out = []
    for key in ['신규개발전', '폐전', '순증가', '신규개발량', '총공급량']:
        d = data_dict.get(key, {})
        달성 = d.get('달성률')
        누달 = d.get('누계달성률')
        rows_out.append({
            '구분': key,
            '연간계획': fmt_num(d.get('연간계획')),
            f'{month}월 계획': fmt_num(d.get('당월계획')),
            f'{month}월 실적': fmt_num(d.get('당월실적')),
            '달성률': fmt_pct(달성),
            '누계 계획': fmt_num(d.get('누계계획')),
            '누계 실적': fmt_num(d.get('누계실적')),
            '누계 달성률': fmt_pct(누달),
        })
    df_out = pd.DataFrame(rows_out)
    st.dataframe(df_out, use_container_width=True, hide_index=True)

render_summary_table(data)

# ════════════════════════════════════════════════
# 2. 신규개발전 상세 현황 (당월 / 누계)
# ════════════════════════════════════════════════
st.markdown('<div class="section-title">📋 신규개발전 상세 현황</div>', unsafe_allow_html=True)

tab1, tab2 = st.tabs([f"📌 {month}월 (당월)", "📌 누계 기준"])

def render_detail_table(d, month_label):
    cats = ['공동주택', '단독주택', '소계', '일반용', '업무용', '산업용', '열병합', '합계']
    rows = []
    for rn in ['계획', '실적', '달성률', '증감']:
        row = {'구분': rn}
        for cat in cats:
            v = d.get(rn, {}).get(cat)
            if rn == '달성률':
                row[cat] = fmt_pct(v)
            elif rn == '증감':
                try:
                    row[cat] = fmt_num(v) if v is not None else "-"
                except:
                    row[cat] = "-"
            else:
                row[cat] = fmt_num(v)
        rows.append(row)
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

with tab1:
    render_detail_table(detail, f"{month}월")
with tab2:
    render_detail_table(cum, "누계")

# ════════════════════════════════════════════════
# 3. 월별 그래프
# ════════════════════════════════════════════════
st.markdown('<div class="section-title">📈 월별 공급량 계획 vs 실적</div>', unsafe_allow_html=True)

month_labels = [f"{i}월" for i in range(1, 13)]

# 공급량 계획/실적 그래프
fig1 = go.Figure()
plan_vals = [v if v is not None else 0 for v in monthly_plan]
# 실적은 현재 월까지만
actual_vals = [0] * 12
# 영업일보에서 실제 월별 실적 가져오기 (현재는 누계에서 역산)
for i in range(month):
    if i == 0:
        actual_vals[i] = monthly_cum[i] if monthly_cum[i] else 0
    else:
        prev = monthly_cum[i-1] if monthly_cum[i-1] else 0
        curr = monthly_cum[i] if monthly_cum[i] else 0
        actual_vals[i] = curr - prev

fig1.add_trace(go.Bar(
    x=month_labels, y=plan_vals,
    name='계획', marker_color='#2d6a9f', opacity=0.7
))
fig1.add_trace(go.Bar(
    x=month_labels[:month],
    y=actual_vals[:month],
    name='실적', marker_color='#e74c3c', opacity=0.9
))
fig1.update_layout(
    title=f"{year}년 월별 공급량 계획 vs 실적 (GJ)",
    barmode='group', height=380,
    legend=dict(orientation='h', y=1.1),
    yaxis_title="공급량 (GJ)",
    plot_bgcolor='white',
    yaxis=dict(gridcolor='#eee')
)
st.plotly_chart(fig1, use_container_width=True)

# 공급전 계획/실적 그래프 (신규개발량 파일 있을 때)
if plan_data and actual_data:
    st.markdown('<div class="section-title">📈 신규개발전 계획 vs 실적 (카테고리별)</div>',
                unsafe_allow_html=True)

    cats_show = ['가정용', '일반용', '업무용', '산업용']
    colors = ['#2d6a9f', '#e74c3c', '#27ae60', '#f39c12']

    fig2 = make_subplots(rows=2, cols=2,
                         subplot_titles=cats_show,
                         shared_xaxes=False)
    positions = [(1,1),(1,2),(2,1),(2,2)]

    for idx, (cat, pos, color) in enumerate(zip(cats_show, positions, colors)):
        p_vals = [v if v is not None else 0 for v in (plan_data.get(cat) or [0]*12)]
        a_vals = [v if v is not None else 0 for v in (actual_data.get(cat) or [0]*12)]

        fig2.add_trace(go.Scatter(
            x=month_labels, y=p_vals,
            name=f'{cat} 계획', line=dict(color=color, dash='dash'),
            showlegend=(idx == 0)
        ), row=pos[0], col=pos[1])
        fig2.add_trace(go.Scatter(
            x=month_labels[:month], y=a_vals[:month],
            name=f'{cat} 실적', line=dict(color=color),
            showlegend=(idx == 0)
        ), row=pos[0], col=pos[1])

    fig2.update_layout(height=500, title=f"{year}년 신규개발전 카테고리별 계획 vs 실적 (전)",
                       plot_bgcolor='white')
    st.plotly_chart(fig2, use_container_width=True)

    # 당월 달성률 막대
    st.markdown('<div class="section-title">📊 당월 신규개발전 달성률</div>', unsafe_allow_html=True)
    cats_all = ['공동주택', '단독주택', '가정용', '일반용', '업무용', '산업용', '열병합용']
    rate_vals = []
    for cat in cats_all:
        p = (plan_data.get(cat) or [None]*12)[month-1]
        a = (actual_data.get(cat) or [None]*12)[month-1]
        try:
            rate_vals.append(round(float(a)/float(p)*100, 1) if p and float(p) != 0 else 0)
        except:
            rate_vals.append(0)

    bar_colors = ['#27ae60' if r >= 100 else '#e74c3c' for r in rate_vals]
    fig3 = go.Figure(go.Bar(
        x=cats_all, y=rate_vals,
        marker_color=bar_colors,
        text=[f"{r}%" for r in rate_vals],
        textposition='outside'
    ))
    fig3.add_hline(y=100, line_dash='dash', line_color='#2d6a9f', annotation_text='목표 100%')
    fig3.update_layout(
        height=350, yaxis_title="달성률 (%)",
        title=f"{month}월 신규개발전 달성률",
        plot_bgcolor='white',
        yaxis=dict(gridcolor='#eee', range=[0, max(rate_vals + [120])])
    )
    st.plotly_chart(fig3, use_container_width=True)

# ════════════════════════════════════════════════
# 4. PDF 출력 버튼
# ════════════════════════════════════════════════
st.markdown("---")
st.markdown("""
<div style='text-align:center; padding: 20px;'>
    <p style='color:#555; margin-bottom:8px;'>🖨️ 화면을 PDF로 출력하려면:</p>
    <p style='color:#888; font-size:13px;'>브라우저 메뉴 → 인쇄 (Ctrl+P) → 대상: PDF로 저장 → 레이아웃: 가로 → 저장</p>
</div>
""", unsafe_allow_html=True)

col_pdf = st.columns([2, 1, 2])
with col_pdf[1]:
    if st.button("🖨️ PDF 출력", use_container_width=True):
        st.markdown("""
        <script>window.print();</script>
        """, unsafe_allow_html=True)
        st.info("Ctrl+P → PDF로 저장을 선택하세요!")
