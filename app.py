# -*- coding: utf-8 -*-
"""报告智能体 · Essay 版前端 v3 — 暖奶油杂志风 · 四页路由"""
import os, time
from pathlib import Path
import pandas as pd
import streamlit as st
from docx import Document
import report_agent as ra

# ── Streamlit Cloud 密钥注入 ──────────────────────────────
_cloud_key = st.secrets.get("MIMO_API_KEY", "") if hasattr(st, "secrets") else ""
if _cloud_key and not ra.MIMO_KEY:
    ra.MIMO_KEY = _cloud_key
    try:
        import anthropic as _ac
        ra._client = _ac.Anthropic(api_key=ra.MIMO_KEY, base_url=ra.MIMO_BASE_URL)
    except Exception:
        pass

BASE         = Path(__file__).parent
SAMPLE_TRAIN = ra.DATA_DIR / "训练任务数据.xlsx"
SAMPLE_ABN   = ra.DATA_DIR / "异常行为数据.xlsx"
TRAIN_COLS   = ["编号","姓名","中队","训练类别","任务名称","应参训次数",
                "实参训次数","出勤率(%)","考核成绩","是否达标","较上月"]
ABN_COLS     = ["编号","姓名","中队","发生日期","异常类型",
                "行为描述","风险等级","处置措施","是否化解"]
SECTION_META = [
    ("train_overall","训练任务总体完成情况"), ("train_cat","分训练类别分析"),
    ("train_squad","分中队对比"),            ("train_trend","环比趋势"),
    ("train_focus","重点关注人员"),          ("abn_overall","异常行为总体态势"),
    ("abn_risk","风险等级分布"),             ("abn_squad","异常分中队分布"),
    ("abn_high","高风险个案"),               ("abn_disposal","处置与化解情况"),
    ("conclusion","综合研判与结论"),         ("suggestion","下一步工作建议"),
]
ALL_SQUADS = ["一中队","二中队","三中队"]

# ── 页面配置（必须第一个 Streamlit 调用）──────────────────
st.set_page_config(page_title="报告智能体", page_icon="📋",
                   layout="wide", initial_sidebar_state="expanded")

# ══════════════════════════════════════════════════════════
#  Essay 主题 CSS
# ══════════════════════════════════════════════════════════
st.markdown("""<style>
/* ── 全局背景 & 字体 ── */
.stApp { background: #FAF9F5 !important; }
.stApp > header { background: transparent !important; box-shadow: none !important; }
.main .block-container { padding: 0 2.4rem 4rem !important; max-width: 1160px; }
body, p, span, div { font-family: -apple-system,"Helvetica Neue",Arial,sans-serif; color: rgba(15,12,8,0.92); }
#MainMenu, footer, .stDeployButton { display: none !important; }

/* ── 侧边栏容器 ── */
section[data-testid="stSidebar"] {
    background: #F8F7F3 !important;
    border-right: 1px solid rgba(15,12,8,0.07) !important;
}
section[data-testid="stSidebar"] > div { padding: 0 !important; }

/* ── 侧边栏所有按钮 → 导航风格 ── */
section[data-testid="stSidebar"] .stButton > button {
    background: transparent !important;
    border: none !important;
    border-radius: 8px !important;
    text-align: left !important;
    padding: 0.46rem 0.85rem !important;
    color: rgba(15,12,8,0.55) !important;
    font-size: 0.86rem !important;
    font-weight: 500 !important;
    box-shadow: none !important;
    width: 100% !important;
    transition: background .15s !important;
}
section[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(15,12,8,0.05) !important;
    color: rgba(15,12,8,0.88) !important;
    transform: none !important;
    box-shadow: none !important;
}
/* 激活状态：primary 类型按钮在侧边栏显示为淡赤陶 */
section[data-testid="stSidebar"] div[data-testid="stBaseButton-primary"] > button {
    background: rgba(217,119,87,0.1) !important;
    color: #B85A35 !important;
    font-weight: 600 !important;
}
/* 侧边栏下载按钮 */
section[data-testid="stSidebar"] .stDownloadButton > button {
    background: #D97757 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    font-size: 0.8rem !important;
    box-shadow: none !important;
    padding: 0.38rem 0.8rem !important;
}
section[data-testid="stSidebar"] .stDownloadButton > button:hover {
    background: #C4623D !important;
    transform: none !important;
}

/* ── 输入控件 ── */
.stTextInput input, .stNumberInput input, .stTextArea textarea {
    background: #fff !important;
    border: 1px solid rgba(15,12,8,0.13) !important;
    border-radius: 8px !important;
    color: rgba(15,12,8,0.92) !important;
    font-size: 0.9rem !important;
}
.stTextInput input:focus, .stNumberInput input:focus {
    border-color: #D97757 !important;
    box-shadow: 0 0 0 3px rgba(217,119,87,0.13) !important;
}
[data-baseweb="select"] > div {
    background: #fff !important;
    border: 1px solid rgba(15,12,8,0.13) !important;
    border-radius: 8px !important;
}

/* ── Toggle ── */
[data-testid="stToggle"] span[aria-checked="true"] { background: #D97757 !important; }

/* ── Slider ── */
[data-baseweb="slider"] [role="slider"] { background: #D97757 !important; border-color: #D97757 !important; }
[data-testid="stSliderTrackFill"] { background: #D97757 !important; }

/* ── 主按钮（赤陶）── */
div[data-testid="stBaseButton-primary"] > button {
    background: #D97757 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 10px !important;
    font-size: 0.93rem !important;
    font-weight: 600 !important;
    padding: 0.58rem 1.6rem !important;
    box-shadow: 0 2px 10px rgba(217,119,87,0.3) !important;
    transition: all .18s !important;
    letter-spacing: .1px !important;
}
div[data-testid="stBaseButton-primary"] > button:hover {
    background: #C4623D !important;
    box-shadow: 0 4px 18px rgba(217,119,87,0.4) !important;
    transform: translateY(-1px) !important;
}

/* ── 次要按钮 ── */
div[data-testid="stBaseButton-secondary"] > button {
    background: transparent !important;
    color: rgba(15,12,8,0.62) !important;
    border: 1px solid rgba(15,12,8,0.16) !important;
    border-radius: 8px !important;
    font-size: 0.88rem !important;
    box-shadow: none !important;
    transition: all .15s !important;
}
div[data-testid="stBaseButton-secondary"] > button:hover {
    border-color: #D97757 !important;
    color: #D97757 !important;
    background: rgba(217,119,87,0.04) !important;
    transform: none !important;
    box-shadow: none !important;
}

/* ── 下载按钮（蓝）── */
.stDownloadButton > button {
    background: #2A78D6 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
    box-shadow: 0 2px 8px rgba(42,120,214,0.25) !important;
}
.stDownloadButton > button:hover { background: #1d64b8 !important; }

/* ── Metric 卡片 ── */
[data-testid="metric-container"] {
    background: #fff !important;
    border: 1px solid rgba(15,12,8,0.08) !important;
    border-radius: 12px !important;
    padding: 0.95rem 1.1rem !important;
}
[data-testid="stMetricValue"] {
    color: rgba(15,12,8,0.9) !important;
    font-size: 1.65rem !important;
    font-weight: 700 !important;
    font-family: Georgia, serif !important;
}
[data-testid="stMetricLabel"] {
    color: rgba(15,12,8,0.42) !important;
    font-size: 0.7rem !important;
    text-transform: uppercase !important;
    letter-spacing: .5px !important;
}

/* ── Expander ── */
.streamlit-expanderHeader {
    background: #F8F7F3 !important;
    border: 1px solid rgba(15,12,8,0.08) !important;
    border-radius: 8px !important;
    color: rgba(15,12,8,0.62) !important;
    font-size: 0.88rem !important;
}

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    border: 2px dashed rgba(217,119,87,0.3) !important;
    border-radius: 12px !important;
    background: rgba(250,249,245,0.5) !important;
}
[data-testid="stFileUploader"]:hover { border-color: rgba(217,119,87,0.55) !important; }

/* ── Alert ── */
[data-testid="stAlertContainer"] { border-radius: 10px !important; }

/* ── Progress & Spinner ── */
.stProgress > div > div { background: #D97757 !important; }
[data-testid="stSpinner"] > div { border-top-color: #D97757 !important; }

/* ══════════════════════════════════════════
   自定义工具类
   ══════════════════════════════════════════ */

/* 上眉标签 */
.eyebrow {
    font-size: .7rem; font-weight: 700; letter-spacing: 1.4px;
    text-transform: uppercase; color: #D97757; margin-bottom: .45rem;
}
/* 主标题 */
.pg-title {
    font-family: Georgia,'Anthropic Serif',serif;
    font-size: 2.2rem; font-weight: 700; color: rgba(15,12,8,0.92);
    line-height: 1.14; letter-spacing: -.4px; margin: 0 0 .7rem;
}
/* 副标题描述 */
.pg-desc { font-size: .94rem; color: rgba(15,12,8,0.5); line-height: 1.65; max-width: 540px; margin: 0; }

/* 区块标签 */
.sec-label {
    font-size: .68rem; font-weight: 700; letter-spacing: .9px;
    text-transform: uppercase; color: rgba(15,12,8,0.36); margin-bottom: .55rem;
}

/* 分割线 */
.hr { border: none; border-top: 1px solid rgba(15,12,8,0.08); margin: 1.4rem 0; }

/* 徽章 */
.chip {
    display: inline-block; padding: .18rem .65rem; border-radius: 20px;
    font-size: .7rem; font-weight: 600; letter-spacing: .2px;
}
.chip-terra { background: rgba(217,119,87,.1); color: #B85A35; border: 1px solid rgba(217,119,87,.22); }
.chip-blue  { background: rgba(42,120,214,.1); color: #1d5fa8; border: 1px solid rgba(42,120,214,.22); }
.chip-green { background: rgba(22,163,74,.1);  color: #15803d; border: 1px solid rgba(22,163,74,.22); }
.chip-grey  { background: rgba(15,12,8,.06);   color: rgba(15,12,8,.55); border: 1px solid rgba(15,12,8,.1); }
.chip-red   { background: rgba(168,78,46,.1);  color: #A84E2E; border: 1px solid rgba(168,78,46,.22); }

/* TL;DR 框 */
.tldr-box {
    background: rgba(217,119,87,.06); border-left: 3px solid #D97757;
    border-radius: 0 10px 10px 0; padding: 1rem 1.4rem; margin: .9rem 0 1.5rem;
}
.tldr-eye  { font-size: .65rem; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; color: #D97757; margin-bottom: .3rem; }
.tldr-text { font-size: .9rem; line-height: 1.72; color: rgba(15,12,8,0.74); }

/* 报告正文区块 */
.rpt-sec-num   { font-family: Georgia,serif; font-size: 1.7rem; font-weight: 700; color: rgba(15,12,8,.07); margin-right: .45rem; display: inline; }
.rpt-sec-title { font-family: Georgia,serif; font-size: 1.05rem; font-weight: 700; color: rgba(15,12,8,.88); display: inline; }
.rpt-sub       { font-weight: 700; font-size: .93rem; color: rgba(15,12,8,.8); font-family: Georgia,serif; margin: 1.1rem 0 .35rem; }
.rpt-body      {
    background: #fff; border: 1px solid rgba(15,12,8,.07);
    border-radius: 12px; padding: 1.15rem 1.5rem; margin: .4rem 0 .9rem;
    font-size: .9rem; line-height: 1.78; color: rgba(15,12,8,.73);
}
.rpt-meta      { font-size: .78rem; color: rgba(15,12,8,.42); padding: .2rem 0 .8rem; }

/* 档案库卡片 */
.lib-card {
    background: #fff; border: 1px solid rgba(15,12,8,.08);
    border-radius: 14px; padding: 1.25rem 1.4rem; margin-bottom: .5rem;
    transition: all .18s;
}
.lib-card:hover {
    border-color: rgba(217,119,87,.3);
    box-shadow: 0 4px 18px rgba(15,12,8,.07);
    transform: translateY(-2px);
}
.lib-bar   { width: 26px; height: 3px; background: #D97757; border-radius: 2px; margin-bottom: .7rem; }
.lib-title { font-family: Georgia,serif; font-size: .9rem; font-weight: 700; color: rgba(15,12,8,.88); margin-bottom: .3rem; line-height: 1.35; }
.lib-meta  { font-size: .72rem; color: rgba(15,12,8,.42); line-height: 1.6; }

/* 状态指示点 */
.dot-green { width:6px;height:6px;background:#16a34a;border-radius:50%;display:inline-block;margin-right:5px;vertical-align:middle; }
.dot-grey  { width:6px;height:6px;background:rgba(15,12,8,.25);border-radius:50%;display:inline-block;margin-right:5px;vertical-align:middle; }

/* 来源行 */
.src-row {
    padding: .7rem 1rem; border-radius: 9px; font-size: .86rem;
    color: rgba(15,12,8,.72); background: #fff;
    border: 1px solid rgba(15,12,8,.07); margin-bottom: .4rem;
    display: flex; gap: .8rem; align-items: center;
}
</style>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
#  Session State 初始化
# ══════════════════════════════════════════════════════════
def _init():
    defaults = {
        "page": "home",
        # 基本信息
        "unit": "三监区", "parent": "监狱教育改造科",
        "author": "", "reviewer": "", "approver": "",
        # 时间
        "year": 2026, "month": 4, "cycle": "月度",
        "classification": "内部资料",
        # 统计参数
        "squads": ALL_SQUADS.copy(),
        "pass_score": 60, "pass_att": 80, "focus_threshold": 70,
        # 章节开关
        "sections": {k: True for k, _ in SECTION_META},
        # AI
        "use_llm": bool(ra.MIMO_KEY), "llm_model": "mimo-v2.5", "llm_style": "规范详实",
        # 数据
        "use_sample": True, "train_df": None, "abn_df": None, "data_ok": False,
        # 结果
        "result_out": None, "result_stats": None, "result_charts": None,
        "last_gen": "", "gen_elapsed": 0, "gen_llm": False,
        # 档案库
        "library": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init()
S = st.session_state

# 自动加载示例数据
if S.use_sample and not S.data_ok:
    try:
        S.train_df = pd.read_excel(SAMPLE_TRAIN)
        S.abn_df   = pd.read_excel(SAMPLE_ABN)
        S.data_ok  = True
    except Exception:
        S.data_ok = False


# ══════════════════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════════════════
with st.sidebar:
    # ── 品牌区 ──
    st.markdown("""
    <div style="padding:1.5rem 1.2rem 1rem;border-bottom:1px solid rgba(15,12,8,.07);">
      <div style="font-family:Georgia,serif;font-size:1rem;font-weight:700;
                  color:rgba(15,12,8,.88);letter-spacing:-.2px;line-height:1.3;">
        数据分析<br>报告智能体
      </div>
      <div style="font-size:.68rem;color:rgba(15,12,8,.36);margin-top:4px;letter-spacing:.4px;">
        G 端安防 · 公安监狱
      </div>
    </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:.55rem'></div>", unsafe_allow_html=True)

    # ── 导航 ──
    NAV = [
        ("home",    "🏠", "首页"),
        ("config",  "⚙️", "报告配置"),
        ("report",  "📋", "报告查看"),
        ("library", "📚", "报告档案"),
    ]
    for pg, icon, lbl in NAV:
        disabled = (pg == "report") and not S.result_out
        btn_type = "primary" if S.page == pg else "secondary"
        if st.button(f"{icon}  {lbl}", key=f"nav_{pg}",
                     type=btn_type, use_container_width=True,
                     disabled=disabled):
            S.page = pg
            st.rerun()

    st.markdown(
        "<div style='height:.75rem;border-top:1px solid rgba(15,12,8,.06);margin-top:.35rem'></div>",
        unsafe_allow_html=True)

    # ── 模型状态 ──
    llm_ok = bool(ra.MIMO_KEY)
    dot_cls = "dot-green" if llm_ok else "dot-grey"
    llm_txt = "大模型已连接 · Mimo" if llm_ok else "模板模式（无 API Key）"
    st.markdown(f"""
    <div style="padding:.4rem .9rem;font-size:.77rem;color:rgba(15,12,8,.52);">
      <span class="{dot_cls}"></span>{llm_txt}
    </div>""", unsafe_allow_html=True)

    if S.last_gen:
        st.markdown(f"""
        <div style="padding:0 .9rem .3rem;font-size:.7rem;color:rgba(15,12,8,.34);">
          上次生成：{S.last_gen}
        </div>""", unsafe_allow_html=True)

    # ── 快速下载 ──
    if S.result_out and Path(S.result_out).exists():
        st.markdown("<div style='padding:0.2rem 0.5rem 0.1rem'>", unsafe_allow_html=True)
        st.download_button("⬇ 下载最新报告",
            data=Path(S.result_out).read_bytes(),
            file_name=Path(S.result_out).name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True, key="sb_dl")
        st.markdown("</div>", unsafe_allow_html=True)

    # ── 报告页目录 ──
    if S.page == "report" and S.result_stats:
        st.markdown("""
        <div style="padding:.9rem .9rem .3rem;font-size:.62rem;font-weight:700;
                    letter-spacing:.9px;color:rgba(15,12,8,.3);text-transform:uppercase;">
          目录
        </div>""", unsafe_allow_html=True)
        TOC = [
            ("报告概述",      ""),
            ("训练总体情况",  "train_overall"),
            ("类别分析",      "train_cat"),
            ("中队对比",      "train_squad"),
            ("环比趋势",      "train_trend"),
            ("重点关注人员",  "train_focus"),
            ("异常总体态势",  "abn_overall"),
            ("风险等级分布",  "abn_risk"),
            ("异常中队分布",  "abn_squad"),
            ("高风险个案",    "abn_high"),
            ("处置与化解",    "abn_disposal"),
            ("综合研判",      "conclusion"),
            ("工作建议",      "suggestion"),
        ]
        for lbl, key in TOC:
            if key and not S.sections.get(key, True):
                continue
            st.markdown(f"""
            <div style="padding:.26rem .9rem .26rem 1.1rem;font-size:.78rem;
                        color:rgba(15,12,8,.48);line-height:1.4;">
              {lbl}
            </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
#  PAGE · HOME  首页
# ══════════════════════════════════════════════════════════
if S.page == "home":

    # Hero 区
    st.markdown("""
    <div style="padding:2.8rem 0 1.2rem;">
      <div class="eyebrow">报告智能体 · 数据驱动公文</div>
      <h1 class="pg-title">将训练与异常数据<br>转化为规范公文报告</h1>
      <p class="pg-desc">
        导入训练任务与异常行为数据，系统自动完成统计分析、图表生成，
        大模型直接撰写符合监狱系统公文规范的 Word 报告。
      </p>
    </div>
    <div style="display:flex;gap:.45rem;flex-wrap:wrap;margin-bottom:1.8rem;">
      <span class="chip chip-terra">📊 自动统计分析</span>
      <span class="chip chip-blue">🤖 大模型直接撰写</span>
      <span class="chip chip-green">📄 Word 公文输出</span>
      <span class="chip chip-grey">📈 可视化图表</span>
    </div>
    <hr class="hr">
    """, unsafe_allow_html=True)

    # 数据来源
    st.markdown('<div class="sec-label">数据来源</div>', unsafe_allow_html=True)
    new_sample = st.toggle(
        "使用内置示例数据（60名罪犯 · 180条训练记录 · 28起异常）",
        value=S.use_sample)
    if new_sample != S.use_sample:
        S.use_sample = new_sample
        S.data_ok    = False
        S.train_df   = None
        S.abn_df     = None

    if S.use_sample:
        if not S.data_ok:
            try:
                S.train_df = pd.read_excel(SAMPLE_TRAIN)
                S.abn_df   = pd.read_excel(SAMPLE_ABN)
                S.data_ok  = True
            except Exception as e:
                st.error(f"示例数据加载失败：{e}")
        if S.data_ok:
            st.markdown("""
            <div style="background:rgba(22,163,74,.06);border:1px solid rgba(22,163,74,.18);
                        border-radius:9px;padding:.65rem 1.1rem;font-size:.85rem;
                        color:rgba(15,12,8,.68);margin:.4rem 0 .8rem;">
              ✓ 示例数据已就绪，可直接生成报告
            </div>""", unsafe_allow_html=True)
    else:
        c1, c2 = st.columns(2)
        tf = c1.file_uploader("📊 训练任务数据 (.xlsx)", type=["xlsx"], key="tf_up")
        af = c2.file_uploader("⚠️ 异常行为数据 (.xlsx)", type=["xlsx"], key="af_up")

        with st.expander("下载数据格式模板"):
            d1, d2 = st.columns(2)
            if SAMPLE_TRAIN.exists():
                d1.download_button("训练任务模板.xlsx",
                    SAMPLE_TRAIN.read_bytes(), "训练任务数据_模板.xlsx", key="tmpl_train")
            if SAMPLE_ABN.exists():
                d2.download_button("异常行为模板.xlsx",
                    SAMPLE_ABN.read_bytes(), "异常行为数据_模板.xlsx", key="tmpl_abn")

        if tf and af:
            try:
                S.train_df = pd.read_excel(tf)
                S.abn_df   = pd.read_excel(af)
                miss1 = [c for c in TRAIN_COLS if c not in S.train_df.columns]
                miss2 = [c for c in ABN_COLS   if c not in S.abn_df.columns]
                if miss1:
                    st.error(f"训练表缺少列：{miss1}"); S.data_ok = False
                elif miss2:
                    st.error(f"异常表缺少列：{miss2}"); S.data_ok = False
                else:
                    st.success(f"✅ 数据已就绪：训练 {len(S.train_df)} 条 / 异常 {len(S.abn_df)} 条")
                S.data_ok = not (miss1 or miss2)
            except Exception as e:
                st.error(f"读取失败：{e}"); S.data_ok = False
        elif not S.data_ok:
            st.caption("请上传两张数据表，或打开示例数据开关。")

    # 数据预览
    if S.data_ok and S.train_df is not None:
        with st.expander("数据预览"):
            p1, p2 = st.columns(2)
            p1.caption("训练任务数据（前5行）")
            p1.dataframe(S.train_df.head(), use_container_width=True)
            p2.caption("异常行为数据（前5行）")
            p2.dataframe(S.abn_df.head(), use_container_width=True)

    st.markdown('<hr class="hr">', unsafe_allow_html=True)

    # CTA 按钮行
    btn_col, cfg_col = st.columns([5, 2])

    with btn_col:
        if st.button("生成报告 →", type="primary",
                     use_container_width=True, disabled=not S.data_ok):
            # 同步配置到引擎
            ra.META.update({
                "unit": S.unit, "parent": S.parent,
                "author": S.author, "reviewer": S.reviewer, "approver": S.approver,
                "period": f"{S.year}年{S.month}月", "cycle": S.cycle,
            })
            ra.USE_LLM    = bool(S.use_llm and ra._client is not None)
            ra.MIMO_MODEL = S.llm_model
            ra.LLM_STYLE  = S.llm_style

            spin_msg = "正在统计分析与生成图表…"
            if ra.USE_LLM:
                spin_msg += "  大模型正在撰写报告内容，约 1–3 分钟，请耐心等待"

            t0 = time.time()
            with st.spinner(spin_msg):
                try:
                    stats  = ra.analyze(
                        S.train_df, S.abn_df,
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
                    S.gen_elapsed   = int(time.time() - t0)
                    S.last_gen      = time.strftime("%Y-%m-%d %H:%M")
                    S.gen_llm       = ra.USE_LLM
                    # 加入档案库
                    S.library.append({
                        "title": f"{S.unit}{S.year}年{S.month}月{S.cycle}报告",
                        "path":  str(out),
                        "time":  S.last_gen,
                        "llm":   ra.USE_LLM,
                        "stats": {
                            "inmate_count": stats["inmate_count"],
                            "pass_rate":    stats["pass_rate"],
                            "abn_total":    stats["abn_total"],
                        },
                    })
                    S.page = "report"
                    st.rerun()
                except Exception as e:
                    st.error(f"生成失败：{e}")

    with cfg_col:
        if st.button("⚙️  报告配置", use_container_width=True):
            S.page = "config"
            st.rerun()

    # 底部提示
    st.markdown("""
    <div style="margin-top:1.5rem;font-size:.78rem;color:rgba(15,12,8,.36);">
      首次使用：先用示例数据体验完整流程，再换入真实业务数据。<br>
      如需启用大模型生成，请将 Mimo API Key 写入 <code>mimo_key.txt</code>。
    </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
#  PAGE · CONFIG  报告配置
# ══════════════════════════════════════════════════════════
elif S.page == "config":

    st.markdown("""
    <div style="padding:2.5rem 0 1.3rem;">
      <div class="eyebrow">报告配置</div>
      <h2 style="font-family:Georgia,serif;font-size:1.7rem;font-weight:700;
                 color:rgba(15,12,8,.92);margin:0;letter-spacing:-.3px;">
        自定义报告参数
      </h2>
    </div>""", unsafe_allow_html=True)

    col_l, col_r = st.columns(2, gap="large")

    # ── 左列 ──
    with col_l:
        st.markdown('<div class="sec-label">基本信息</div>', unsafe_allow_html=True)
        S.unit     = st.text_input("报送单位",      S.unit)
        S.parent   = st.text_input("呈报单位",      S.parent)
        a1, a2     = st.columns(2)
        S.author   = a1.text_input("编制人",        S.author)
        S.reviewer = a2.text_input("审核人",        S.reviewer)
        S.approver = st.text_input("批准人（可选）", S.approver)

        st.markdown("<div style='height:.7rem'></div>", unsafe_allow_html=True)
        st.markdown('<div class="sec-label">大模型设置</div>', unsafe_allow_html=True)
        ml1, ml2 = st.columns(2)
        S.use_llm   = ml1.checkbox("启用大模型", value=S.use_llm,
                                    disabled=not ra.MIMO_KEY)
        S.llm_model = ml2.selectbox("模型版本", ["mimo-v2.5","mimo-v2-pro"],
                                     index=0 if S.llm_model == "mimo-v2.5" else 1)
        S.llm_style = st.selectbox("写作风格", ["规范详实","简洁精炼","严谨正式"],
                                    index=["规范详实","简洁精炼","严谨正式"].index(S.llm_style))
        if not ra.MIMO_KEY:
            st.caption("⚠️ 未检测到 Mimo API Key，将使用模板模式生成报告。"
                       "将 Key 写入 mimo_key.txt 后重启即可启用大模型。")
        elif S.use_llm:
            st.caption("大模型将根据统计数据直接撰写各章节内容，每次约需 1–3 分钟。")

    # ── 右列 ──
    with col_r:
        st.markdown('<div class="sec-label">时间与密级</div>', unsafe_allow_html=True)
        t1, t2  = st.columns(2)
        S.year  = t1.number_input("年份", 2020, 2035, S.year)
        S.month = t2.selectbox("月份", list(range(1, 13)), index=S.month - 1)
        S.cycle = st.selectbox("汇报周期", ["日","周","月度","季度","年度"],
                                index=["日","周","月度","季度","年度"].index(S.cycle))
        S.classification = st.selectbox(
            "密级标注", ["内部资料","保密","机密","公开"],
            index=["内部资料","保密","机密","公开"].index(S.classification))

        st.markdown("<div style='height:.7rem'></div>", unsafe_allow_html=True)
        st.markdown('<div class="sec-label">统计参数</div>', unsafe_allow_html=True)
        S.pass_score      = st.slider("成绩达标线（分）",      0, 100, S.pass_score,      step=5)
        S.pass_att        = st.slider("出勤率达标线（%）",     0, 100, S.pass_att,        step=5)
        S.focus_threshold = st.slider("重点关注成绩线（分）",  0, 100, S.focus_threshold, step=5)
        avail_sq = (sorted(S.train_df["中队"].unique().tolist())
                    if (S.data_ok and S.train_df is not None) else ALL_SQUADS)
        S.squads = st.multiselect("统计范围（中队）", avail_sq,
                                   default=[q for q in S.squads if q in avail_sq] or avail_sq)

    st.markdown('<hr class="hr">', unsafe_allow_html=True)

    with st.expander("精细控制章节（可关闭不需要的章节）", expanded=False):
        cols = st.columns(3)
        for i, (key, lbl) in enumerate(SECTION_META):
            S.sections[key] = cols[i % 3].checkbox(
                lbl, value=S.sections.get(key, True), key=f"sec_{key}")

    st.markdown("<div style='height:.5rem'></div>", unsafe_allow_html=True)

    b1, b2, _ = st.columns([1, 1, 3])
    if b1.button("← 返回首页", type="secondary", use_container_width=True):
        S.page = "home"; st.rerun()
    if b2.button("生成报告 →", type="primary", use_container_width=True,
                 disabled=not S.data_ok):
        S.page = "home"; st.rerun()   # rerun 到首页触发生成逻辑


# ══════════════════════════════════════════════════════════
#  PAGE · REPORT  报告查看
# ══════════════════════════════════════════════════════════
elif S.page == "report":

    if not S.result_out or not Path(S.result_out).exists():
        st.warning("暂无报告，请先在首页生成。")
        if st.button("← 返回首页"):
            S.page = "home"; st.rerun()
        st.stop()

    out_path  = Path(S.result_out)
    llm_label = f"Mimo · {S.llm_model}" if S.gen_llm else "模板模式"
    clf_chip  = "chip-red" if S.classification in ("保密","机密") else "chip-grey"

    # ── 报告标题区 ──
    st.markdown(f"""
    <div style="padding:2.2rem 0 1.2rem;border-bottom:1px solid rgba(15,12,8,.08);">
      <div class="eyebrow">{S.unit} · {S.year}年{S.month}月 {S.cycle}</div>
      <h1 style="font-family:Georgia,serif;font-size:1.8rem;font-weight:700;
                 color:rgba(15,12,8,.92);margin:0 0 .9rem;letter-spacing:-.35px;">
        罪犯训练任务与异常行为数据分析报告
      </h1>
      <div style="display:flex;gap:.5rem;flex-wrap:wrap;align-items:center;">
        <span class="chip {clf_chip}">【{S.classification}】</span>
        <span class="chip chip-blue">{llm_label}</span>
        <span style="font-size:.72rem;color:rgba(15,12,8,.36);">
          生成于 {S.last_gen} · 用时 {S.gen_elapsed}s
        </span>
      </div>
    </div>""", unsafe_allow_html=True)

    # ── 操作行 ──
    dl_col, ng_col, _ = st.columns([2, 1, 3])
    dl_col.download_button(
        "⬇ 下载 Word 报告（.docx）",
        data=out_path.read_bytes(), file_name=out_path.name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True, key="rpt_dl")
    if ng_col.button("生成新报告", use_container_width=True):
        S.page = "home"; st.rerun()

    # ── 关键指标 ──
    if S.result_stats:
        st_s = S.result_stats
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("纳入罪犯",  st_s["inmate_count"])
        m2.metric("训练记录",  st_s["train_records"])
        m3.metric("达标率",    f"{st_s['pass_rate']}%")
        m4.metric("平均成绩",  st_s["avg_score"])
        m5.metric("异常起数",  st_s["abn_total"])

        # ── TL;DR ──
        abn_main = st_s["abn_by_type"].index[0] if len(st_s["abn_by_type"]) else "—"
        h_cnt    = len(st_s["abn_high"])
        wt       = st_s["by_squad"]["达标率"].idxmin()
        st.markdown(f"""
        <div class="tldr-box">
          <div class="tldr-eye">核心摘要</div>
          <div class="tldr-text">
            本期纳入统计 <strong>{st_s['inmate_count']}</strong> 名罪犯，
            训练整体达标率 <strong>{st_s['pass_rate']}%</strong>，
            平均考核成绩 <strong>{st_s['avg_score']}</strong> 分，
            平均出勤率 <strong>{st_s['avg_att']}%</strong>。
            其中 <strong>{wt}</strong> 训练达标率相对最低，需重点督导。
            本期共发生异常行为 <strong>{st_s['abn_total']}</strong> 起，
            以 <strong>{abn_main}</strong> 为主，
            高风险个案 <strong>{h_cnt}</strong> 起需逐一落实处置化解。
          </div>
        </div>""", unsafe_allow_html=True)

    # ── 图表 ──
    if S.result_charts:
        st.markdown('<div class="sec-label" style="margin-top:1.3rem">可视化图表</div>',
                    unsafe_allow_html=True)
        g1, g2 = st.columns(2)
        g1.image(str(S.result_charts["squad"]), use_container_width=True)
        g2.image(str(S.result_charts["cat"]),   use_container_width=True)
        g3, g4 = st.columns([3, 2])
        g3.image(str(S.result_charts["abn_type"]), use_container_width=True)
        g4.image(str(S.result_charts["abn_risk"]), use_container_width=True)

    # ── 报告全文 ──
    st.markdown('<div class="sec-label" style="margin-top:1.6rem">报告全文</div>',
                unsafe_allow_html=True)

    H1_PREFIXES = ("一、","二、","三、","四、","五、","六、","七、","八、","九、","十、")
    H2_PREFIXES = ("（一）","（二）","（三）","（四）","（五）","（六）","（七）","（八）","（九）")
    sec_n = 0

    for para in Document(S.result_out).paragraphs:
        txt = para.text.strip()
        if not txt:
            continue
        sty = para.style.name if para.style else ""
        if any(txt.startswith(p) for p in H1_PREFIXES) or "Heading 1" in sty:
            sec_n += 1
            st.markdown(f"""
            <div style="margin:1.8rem 0 .55rem;display:flex;align-items:baseline;gap:.45rem;">
              <span class="rpt-sec-num">{sec_n:02d}</span>
              <span class="rpt-sec-title">{txt}</span>
            </div>""", unsafe_allow_html=True)
        elif any(txt.startswith(p) for p in H2_PREFIXES):
            st.markdown(f'<div class="rpt-sub">{txt}</div>', unsafe_allow_html=True)
        elif txt.startswith("【") and txt.endswith("】"):
            pass   # 密级标注已在标题区显示，正文跳过
        elif "报送单位：" in txt or "呈报：" in txt:
            st.markdown(f'<div class="rpt-meta">{txt}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="rpt-body">{txt}</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
#  PAGE · LIBRARY  报告档案
# ══════════════════════════════════════════════════════════
elif S.page == "library":

    st.markdown("""
    <div style="padding:2.5rem 0 1.3rem;">
      <div class="eyebrow">报告档案</div>
      <h2 style="font-family:Georgia,serif;font-size:1.7rem;font-weight:700;
                 color:rgba(15,12,8,.92);margin:0;letter-spacing:-.3px;">
        历史生成报告
      </h2>
    </div>""", unsafe_allow_html=True)

    if not S.library:
        st.markdown("""
        <div style="text-align:center;padding:4.5rem 2rem;color:rgba(15,12,8,.3);">
          <div style="font-size:2.5rem;margin-bottom:.7rem;">📚</div>
          <div style="font-family:Georgia,serif;font-size:1.05rem;
                      margin-bottom:.35rem;color:rgba(15,12,8,.42);">
            还没有历史报告
          </div>
          <div style="font-size:.82rem;">
            生成第一份报告后，它会出现在这里
          </div>
        </div>""", unsafe_allow_html=True)
        if st.button("← 去首页生成报告"):
            S.page = "home"; st.rerun()
    else:
        lib = list(reversed(S.library))
        N   = 3
        for i in range(0, len(lib), N):
            batch = lib[i:i + N]
            cols  = st.columns(N)
            for j, item in enumerate(batch):
                with cols[j]:
                    stats    = item.get("stats", {})
                    llm_chip = "🤖 大模型" if item.get("llm") else "📝 模板"
                    st.markdown(f"""
                    <div class="lib-card">
                      <div class="lib-bar"></div>
                      <div class="lib-title">{item['title']}</div>
                      <div class="lib-meta">
                        {item['time']}<br>
                        达标率 {stats.get('pass_rate','—')}% ·
                        异常 {stats.get('abn_total','—')} 起<br>
                        {llm_chip}
                      </div>
                    </div>""", unsafe_allow_html=True)
                    if Path(item["path"]).exists():
                        st.download_button(
                            "⬇ 下载",
                            data=Path(item["path"]).read_bytes(),
                            file_name=Path(item["path"]).name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            key=f"lib_dl_{i}_{j}")
