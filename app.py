# -*- coding: utf-8 -*-
"""报告智能体 MVP 前端 v2 — 蓝橙配色 / 完整配置模块 / 多标签布局"""
import os
import time
from pathlib import Path
import pandas as pd
import streamlit as st
from docx import Document
import report_agent as ra

# ── Streamlit Cloud 密钥注入 ──────────────────────────────
# 当部署在 Streamlit Community Cloud 时，通过 st.secrets 获取 API Key
# 本地开发时继续读取 mimo_key.txt，无需改动
_cloud_key = st.secrets.get("MIMO_API_KEY", "") if hasattr(st, "secrets") else ""
if _cloud_key and not ra.MIMO_KEY:
    ra.MIMO_KEY = _cloud_key
    try:
        import anthropic as _ac
        ra._client = _ac.Anthropic(api_key=ra.MIMO_KEY, base_url=ra.MIMO_BASE_URL)
    except Exception:
        pass

BASE = Path(__file__).parent
SAMPLE_TRAIN = ra.DATA_DIR / "训练任务数据.xlsx"
SAMPLE_ABN   = ra.DATA_DIR / "异常行为数据.xlsx"
TRAIN_COLS   = ["编号","姓名","中队","训练类别","任务名称","应参训次数","实参训次数","出勤率(%)","考核成绩","是否达标","较上月"]
ABN_COLS     = ["编号","姓名","中队","发生日期","异常类型","行为描述","风险等级","处置措施","是否化解"]
SECTION_META = [
    ("train_overall","训练任务总体完成情况"),("train_cat","分训练类别分析"),
    ("train_squad","分中队对比"),("train_trend","环比趋势"),("train_focus","重点关注人员"),
    ("abn_overall","异常行为总体态势"),("abn_risk","风险等级分布"),
    ("abn_squad","异常分中队分布"),("abn_high","高风险个案"),("abn_disposal","处置与化解情况"),
    ("conclusion","综合研判与结论"),("suggestion","下一步工作建议"),
]
ALL_SQUADS = ["一中队","二中队","三中队"]

# ── 页面配置 ─────────────────────────────────────────────
st.set_page_config(page_title="报告智能体", page_icon="📄", layout="wide",
                   initial_sidebar_state="expanded")

# ── CSS 注入（蓝橙高级感主题）────────────────────────────
st.markdown("""
<style>
/* ── 全局背景 ── */
.stApp { background: #eef2ff; }
.main .block-container { padding: 1.2rem 2.4rem 3rem; max-width: 1100px; }

/* ── 侧边栏 ── */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#0f172a 0%,#1e3a8a 60%,#0f172a 100%) !important;
    border-right: 1px solid #1e3a8a;
}
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stMarkdown p { color:#94a3b8 !important; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color:#e2e8f0 !important; }
section[data-testid="stSidebar"] hr { border-color:#334155 !important; }
section[data-testid="stSidebar"] .stStatusWidget { display:none; }

/* ── 标签页 ── */
.stTabs [data-baseweb="tab-list"] {
    background:#fff; border-radius:14px; padding:5px;
    border:1px solid #dde5f4; gap:4px; margin-bottom:4px;
}
.stTabs [data-baseweb="tab"] {
    border-radius:10px; padding:.45rem 1.6rem;
    font-weight:600; font-size:.9rem; color:#64748b;
    background:transparent; border:none;
    transition:all .2s;
}
.stTabs [aria-selected="true"] {
    background:linear-gradient(135deg,#1e3a8a,#2563eb) !important;
    color:#fff !important; border-radius:10px !important;
}
.stTabs [data-baseweb="tab-highlight"] { display:none !important; }

/* ── 主按钮（橙色）── */
div[data-testid="stBaseButton-primary"] > button,
.stButton > button[kind="primary"] {
    background:linear-gradient(135deg,#f97316,#ea580c) !important;
    color:#fff !important; border:none !important;
    border-radius:14px !important; font-size:1.05rem !important;
    font-weight:700 !important; letter-spacing:.4px;
    padding:.65rem 2rem !important;
    box-shadow:0 4px 20px rgba(249,115,22,.4) !important;
    transition:all .2s !important;
}
div[data-testid="stBaseButton-primary"] > button:hover {
    transform:translateY(-2px);
    box-shadow:0 6px 26px rgba(249,115,22,.5) !important;
}

/* ── 下载按钮（蓝色）── */
div[data-testid="stBaseButton-secondary"] > button,
.stDownloadButton > button {
    background:linear-gradient(135deg,#1e3a8a,#2563eb) !important;
    color:#fff !important; border:none !important;
    border-radius:12px !important; font-weight:600 !important;
    box-shadow:0 4px 16px rgba(37,99,235,.3) !important;
    transition:all .2s !important;
}

/* ── Metric 卡片 ── */
[data-testid="metric-container"] {
    background:#fff; border-radius:14px;
    padding:1rem 1.2rem;
    border-left:4px solid #f97316;
    box-shadow:0 2px 12px rgba(0,0,0,.07);
}
[data-testid="stMetricValue"]  { color:#1e3a8a !important; font-size:2rem !important; font-weight:800 !important; }
[data-testid="stMetricLabel"]  { color:#64748b !important; font-size:.88rem !important; }
[data-testid="stMetricDelta"]  { font-size:.8rem !important; }

/* ── Expander ── */
.streamlit-expanderHeader {
    background:#f0f4ff !important; border-radius:10px !important;
    font-weight:600 !important; color:#1e3a8a !important;
    border:1px solid #dde5f4 !important;
}
/* ── Alert / Info / Success ── */
.stAlert { border-radius:12px !important; }

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    border:2px dashed #93c5fd !important;
    border-radius:14px !important; background:#f0f9ff !important;
}

/* ── Section card helper ── */
.cfg-card {
    background:#fff; border-radius:14px; padding:1.2rem 1.5rem;
    border:1px solid #e2e8f0;
    box-shadow:0 2px 10px rgba(0,0,0,.05); margin-bottom:1rem;
}
.cfg-title {
    font-size:.78rem; font-weight:700; letter-spacing:.8px;
    color:#f97316; text-transform:uppercase; margin-bottom:.7rem;
}
</style>
""", unsafe_allow_html=True)


# ── Session State 初始化 ──────────────────────────────────
def _init():
    defaults = {
        # 基本信息
        "unit":"三监区", "parent":"监狱教育改造科",
        "author":"", "reviewer":"", "approver":"",
        # 时间
        "year":2026, "month":4, "cycle":"月度",
        # 密级
        "classification":"内部资料",
        # 统计配置
        "squads": ALL_SQUADS.copy(),
        "pass_score":60, "pass_att":80, "focus_threshold":70,
        # 章节开关
        "sections":{k:True for k,_ in SECTION_META},
        # AI
        "use_llm":bool(ra.MIMO_KEY), "llm_model":"mimo-v2.5", "llm_style":"规范详实",
        # 运行时
        "use_sample":True, "train_df":None, "abn_df":None, "data_ok":False,
        "result_out":None, "result_stats":None, "result_charts":None,
        "last_gen":"",
    }
    for k,v in defaults.items():
        if k not in st.session_state: st.session_state[k] = v

_init()
S = st.session_state   # shorthand


# ── 侧边栏 ───────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="padding:1.4rem 1rem 1rem;text-align:center;">
      <div style="font-size:3rem;line-height:1;">📄</div>
      <div style="color:#e2e8f0;font-size:1.15rem;font-weight:800;margin:.4rem 0 .2rem;">报告智能体</div>
      <div style="color:#475569;font-size:.78rem;letter-spacing:.5px;">G端安防 · 数据自动成报</div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    llm_ok = bool(ra.MIMO_KEY)
    if llm_ok:
        st.markdown("🟢 &nbsp;**大模型已连接**")
        st.caption(f"Mimo · {ra.MIMO_MODEL}")
    else:
        st.markdown("🔴 &nbsp;**大模型未连接**")
        st.caption("将使用模板模式生成报告")

    if S.last_gen:
        st.divider()
        st.caption(f"⏱ 上次生成：{S.last_gen}")

    if S.result_out and Path(S.result_out).exists():
        st.divider()
        st.download_button("⬇️ 重新下载最新报告",
            data=Path(S.result_out).read_bytes(),
            file_name=Path(S.result_out).name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True)


# ── 顶部 Banner ──────────────────────────────────────────
st.markdown("""
<div style="background:linear-gradient(135deg,#0f172a 0%,#1e3a8a 50%,#1d4ed8 100%);
     padding:1.8rem 2.2rem;border-radius:18px;margin-bottom:1.6rem;
     display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:1rem;">
  <div>
    <h1 style="color:#fff;margin:0;font-size:1.7rem;font-weight:800;letter-spacing:-.3px;">
      📄 &nbsp;数据分析报告智能体
    </h1>
    <p style="color:rgba(255,255,255,.65);margin:.35rem 0 0;font-size:.9rem;">
      公安监狱 · 罪犯训练任务与异常行为数据 → 规范公文报告 · 全程自动化生成
    </p>
  </div>
  <div style="display:flex;gap:.6rem;flex-wrap:wrap;">
    <span style="background:rgba(249,115,22,.18);border:1px solid rgba(249,115,22,.45);
         border-radius:20px;padding:.3rem .9rem;color:#fed7aa;font-size:.78rem;font-weight:600;">
      v2.0
    </span>
    <span style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);
         border-radius:20px;padding:.3rem .9rem;color:#e0e7ff;font-size:.78rem;">
      Mimo 大模型加持
    </span>
  </div>
</div>
""", unsafe_allow_html=True)


# ── 三个主标签 ────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["  📁  上传数据  ", "  ⚙️  报告配置  ", "  🚀  生成 & 下载  "])


# ════════════════════════════════════════════════════════
#  TAB 1 · 上传数据
# ════════════════════════════════════════════════════════
with tab1:
    st.markdown("#### 数据来源")
    S.use_sample = st.toggle("使用内置示例数据（快速体验）", value=S.use_sample)

    if S.use_sample:
        st.info("✅ 已加载示例数据（60名罪犯，180条训练记录，28起异常）。如需使用真实数据，关闭上方开关即可上传。")
        try:
            S.train_df = pd.read_excel(SAMPLE_TRAIN)
            S.abn_df   = pd.read_excel(SAMPLE_ABN)
            S.data_ok  = True
        except Exception as e:
            st.error(f"示例数据加载失败：{e}"); S.data_ok = False
    else:
        c1, c2 = st.columns(2)
        tf = c1.file_uploader("📊 训练任务数据 (.xlsx)", type=["xlsx"], key="tf_up")
        af = c2.file_uploader("⚠️ 异常行为数据 (.xlsx)", type=["xlsx"], key="af_up")

        with st.expander("📥 不清楚表格格式？下载模板"):
            d1, d2 = st.columns(2)
            if SAMPLE_TRAIN.exists():
                d1.download_button("训练任务 模板.xlsx", SAMPLE_TRAIN.read_bytes(), "训练任务数据_模板.xlsx")
            if SAMPLE_ABN.exists():
                d2.download_button("异常行为 模板.xlsx", SAMPLE_ABN.read_bytes(), "异常行为数据_模板.xlsx")

        if tf and af:
            try:
                S.train_df = pd.read_excel(tf); S.abn_df = pd.read_excel(af)
                miss1 = [c for c in TRAIN_COLS if c not in S.train_df.columns]
                miss2 = [c for c in ABN_COLS   if c not in S.abn_df.columns]
                if miss1: st.error(f"训练表缺少列：{miss1}"); S.data_ok=False
                elif miss2: st.error(f"异常表缺少列：{miss2}"); S.data_ok=False
                else: st.success(f"✅ 数据已就绪：训练 {len(S.train_df)} 条 / 异常 {len(S.abn_df)} 条")
                S.data_ok = not miss1 and not miss2
            except Exception as e:
                st.error(f"读取失败：{e}"); S.data_ok = False
        else:
            st.warning("请上传两张数据表，或打开「示例数据」开关。")
            S.data_ok = False

    # 数据预览
    if S.data_ok and S.train_df is not None:
        with st.expander("🔍 数据预览"):
            pc1, pc2 = st.columns(2)
            pc1.caption("训练任务数据（前5行）")
            pc1.dataframe(S.train_df.head(), use_container_width=True)
            pc2.caption("异常行为数据（前5行）")
            pc2.dataframe(S.abn_df.head(), use_container_width=True)

    st.markdown("---")
    st.caption("👉 数据确认后，切换到 **报告配置** 标签页进行个性化设置，再到 **生成 & 下载** 一键出报告。")


# ════════════════════════════════════════════════════════
#  TAB 2 · 报告配置
# ════════════════════════════════════════════════════════
with tab2:
    col_l, col_r = st.columns(2)

    # ── 基本信息 ──
    with col_l:
        st.markdown('<div class="cfg-card"><div class="cfg-title">📋 基本信息</div>', unsafe_allow_html=True)
        S.unit     = st.text_input("报送单位",   S.unit)
        S.parent   = st.text_input("呈报单位",   S.parent)
        c1,c2      = st.columns(2)
        S.author   = c1.text_input("编制人",     S.author)
        S.reviewer = c2.text_input("审核人",     S.reviewer)
        S.approver = st.text_input("批准人（可选）", S.approver)
        st.markdown('</div>', unsafe_allow_html=True)

    # ── 时间与密级 ──
    with col_r:
        st.markdown('<div class="cfg-card"><div class="cfg-title">📅 时间与输出</div>', unsafe_allow_html=True)
        tc1,tc2 = st.columns(2)
        S.year  = tc1.number_input("年份", 2020, 2035, S.year)
        S.month = tc2.selectbox("月份", list(range(1,13)), index=S.month-1)
        S.cycle = st.selectbox("汇报周期", ["日","周","月度","季度","年度"], index=["日","周","月度","季度","年度"].index(S.cycle))
        S.classification = st.selectbox("密级标注", ["内部资料","保密","机密","公开"], index=["内部资料","保密","机密","公开"].index(S.classification))
        st.markdown('</div>', unsafe_allow_html=True)

    # ── 统计配置 ──
    st.markdown('<div class="cfg-card"><div class="cfg-title">📊 统计配置</div>', unsafe_allow_html=True)
    sc1, sc2, sc3 = st.columns(3)
    S.pass_score      = sc1.slider("成绩达标线（分）",    0, 100, S.pass_score,  step=5)
    S.pass_att        = sc2.slider("出勤率达标线（%）",   0, 100, S.pass_att,    step=5)
    S.focus_threshold = sc3.slider("重点关注成绩线（分）",0, 100, S.focus_threshold, step=5)
    # 中队筛选（从已加载数据中读取可用中队）
    avail_sq = sorted(S.train_df["中队"].unique().tolist()) if (S.data_ok and S.train_df is not None) else ALL_SQUADS
    S.squads = st.multiselect("统计范围（中队）", avail_sq, default=[q for q in S.squads if q in avail_sq] or avail_sq)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── AI 设置 ──
    st.markdown('<div class="cfg-card"><div class="cfg-title">🤖 大模型设置</div>', unsafe_allow_html=True)
    ai1, ai2, ai3 = st.columns(3)
    S.use_llm   = ai1.checkbox("启用大模型润色", value=S.use_llm, disabled=not ra.MIMO_KEY)
    S.llm_model = ai2.selectbox("模型", ["mimo-v2.5","mimo-v2-pro"], index=["mimo-v2.5","mimo-v2-pro"].index(S.llm_model) if S.llm_model in ["mimo-v2.5","mimo-v2-pro"] else 0)
    S.llm_style = ai3.selectbox("写作风格", ["规范详实","简洁精炼","严谨正式"], index=["规范详实","简洁精炼","严谨正式"].index(S.llm_style))
    if not ra.MIMO_KEY:
        st.caption("⚠️ 未检测到 Mimo Key，将使用模板模式生成报告。把 Key 写入 mimo_key.txt 后重启即可开启大模型生成。")
    elif S.use_llm:
        st.caption(f"💡 启用后大模型将根据统计数据直接撰写每个章节，每次生成约需 1–3 分钟。")
    st.markdown('</div>', unsafe_allow_html=True)

    # ── 章节开关 ──
    with st.expander("📑 章节精细控制（可关闭不需要的章节）", expanded=False):
        cols = st.columns(3)
        for i, (key, label) in enumerate(SECTION_META):
            S.sections[key] = cols[i % 3].checkbox(label, value=S.sections.get(key, True), key=f"sec_{key}")

    st.markdown("---")
    st.caption("👉 配置完成后，切换到 **生成 & 下载** 标签页一键出报告。")


# ════════════════════════════════════════════════════════
#  TAB 3 · 生成 & 下载
# ════════════════════════════════════════════════════════
with tab3:
    # 配置摘要卡
    period_str = f"{S.year}年{S.month}月"
    st.markdown(f"""
    <div style="background:#fff;border-radius:14px;padding:1rem 1.4rem;
         border-left:4px solid #2563eb;box-shadow:0 2px 10px rgba(0,0,0,.06);margin-bottom:1.2rem;
         display:flex;flex-wrap:wrap;gap:1rem;align-items:center;">
      <div><span style="color:#94a3b8;font-size:.8rem;">单位</span>
           <div style="font-weight:700;color:#1e293b;">{S.unit}</div></div>
      <div style="color:#dde5f4;">|</div>
      <div><span style="color:#94a3b8;font-size:.8rem;">呈报</span>
           <div style="font-weight:700;color:#1e293b;">{S.parent}</div></div>
      <div style="color:#dde5f4;">|</div>
      <div><span style="color:#94a3b8;font-size:.8rem;">期间</span>
           <div style="font-weight:700;color:#1e293b;">{period_str} {S.cycle}</div></div>
      <div style="color:#dde5f4;">|</div>
      <div><span style="color:#94a3b8;font-size:.8rem;">密级</span>
           <div style="font-weight:700;color:{'#c00000' if S.classification in ('保密','机密') else '#1e293b'};">【{S.classification}】</div></div>
      <div style="color:#dde5f4;">|</div>
      <div><span style="color:#94a3b8;font-size:.8rem;">大模型</span>
           <div style="font-weight:700;color:#1e293b;">{'✅ ' + S.llm_model if (S.use_llm and ra.MIMO_KEY) else '⬜ 模板'}</div></div>
    </div>
    """, unsafe_allow_html=True)

    if not S.data_ok:
        st.warning("⚠️ 请先在「上传数据」标签页准备好数据，再来这里生成报告。")
    else:
        if st.button("🚀  一键生成报告", type="primary", use_container_width=True):
            # 同步配置到引擎
            ra.META.update({
                "unit": S.unit, "parent": S.parent,
                "author": S.author, "reviewer": S.reviewer, "approver": S.approver,
                "period": f"{S.year}年{S.month}月", "cycle": S.cycle,
            })
            ra.USE_LLM    = bool(S.use_llm and ra._client is not None)
            ra.MIMO_MODEL = S.llm_model
            ra.LLM_STYLE  = S.llm_style

            t0 = time.time()
            with st.spinner("正在统计分析、生成图表与报告……" +
                            ("大模型正在撰写报告内容，请耐心等待（约1-3分钟）" if ra.USE_LLM else "")):
                try:
                    stats  = ra.analyze(S.train_df, S.abn_df,
                                        pass_score=S.pass_score, pass_att=S.pass_att,
                                        focus_threshold=S.focus_threshold,
                                        squads=S.squads if S.squads else None)
                    charts = ra.make_charts(stats)
                    out    = ra.build_report(stats, charts,
                                             sections=S.sections,
                                             classification=S.classification)
                    S.result_out    = str(out)
                    S.result_stats  = stats
                    S.result_charts = charts
                    S.last_gen      = time.strftime("%Y-%m-%d %H:%M")
                    elapsed = time.time() - t0
                    st.success(f"✅ 报告生成完成！用时 {elapsed:.0f} 秒" +
                               ("（大模型生成）" if ra.USE_LLM else "（模板模式）"))
                except Exception as e:
                    st.error(f"生成失败：{e}"); st.stop()

    # ── 结果展示 ──
    if S.result_out and Path(S.result_out).exists():
        out_path = Path(S.result_out)

        st.download_button("⬇️  下载 Word 报告（.docx）",
            data=out_path.read_bytes(), file_name=out_path.name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True)

        st.markdown("---")

        # 关键指标
        if S.result_stats:
            st_s = S.result_stats
            m1,m2,m3,m4,m5 = st.columns(5)
            m1.metric("纳入罪犯",  st_s["inmate_count"])
            m2.metric("训练记录",  st_s["train_records"])
            m3.metric("达标率",    f"{st_s['pass_rate']}%")
            m4.metric("平均成绩",  st_s["avg_score"])
            m5.metric("异常起数",  st_s["abn_total"])

        # 图表
        if S.result_charts:
            st.markdown("#### 图表预览")
            g1, g2 = st.columns(2)
            g1.image(str(S.result_charts["squad"]), use_container_width=True)
            g2.image(str(S.result_charts["cat"]),   use_container_width=True)
            g3, g4 = st.columns(2)
            g3.image(str(S.result_charts["abn_type"]), use_container_width=True)
            g4.image(str(S.result_charts["abn_risk"]), use_container_width=True)

        # 全文预览
        with st.expander("📖 报告全文预览"):
            for p in Document(S.result_out).paragraphs:
                if p.text.strip():
                    st.write(p.text)
