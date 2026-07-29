from __future__ import annotations

import base64
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"


def _data_uri(path: Path, mime: str = "image/svg+xml") -> str:
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def apply_theme() -> None:
    background = _data_uri(ASSETS / "matching_mapper_background.svg")
    st.markdown(
        f"""
        <style>
        :root {{
            --mm-primary: #4f6df5;
            --mm-primary-2: #7658f5;
            --mm-text: #17213b;
            --mm-muted: #65708a;
            --mm-border: rgba(79, 109, 245, .16);
            --mm-card: rgba(255,255,255,.90);
        }}
        .stApp {{
            background-image: linear-gradient(rgba(247,249,255,.91), rgba(247,249,255,.95)), url('{background}');
            background-size: cover;
            background-attachment: fixed;
            color: var(--mm-text);
        }}
        [data-testid="stHeader"] {{ background: rgba(247,249,255,.72); backdrop-filter: blur(16px); }}
        #MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] {{ display:none !important; }}
        [data-testid="stSidebar"] {{
            background: rgba(255,255,255,.94);
            border-right: 1px solid var(--mm-border);
        }}
        [data-testid="stSidebar"] > div:first-child {{ padding-top: 1rem; }}
        .block-container {{ max-width: 1500px; padding-top: 1.5rem; padding-bottom: 4rem; }}
        h1, h2, h3 {{ color: var(--mm-text); letter-spacing: -.025em; }}
        p, label, .stMarkdown {{ color: var(--mm-text); }}
        .mm-brand {{ display:flex; align-items:center; gap:12px; margin: 2px 0 18px; }}
        .mm-brand img {{ width:44px; height:44px; filter: drop-shadow(0 8px 12px rgba(70,82,180,.18)); }}
        .mm-brand-title {{ font-size:1.05rem; font-weight:800; color:#17213b; }}
        .mm-brand-sub {{ font-size:.76rem; color:#79839b; margin-top:2px; }}
        .mm-kicker {{ text-transform:uppercase; letter-spacing:.15em; color:#4f6df5; font-size:.70rem; font-weight:800; }}
        .mm-hero {{
            display:grid; grid-template-columns:minmax(0,1.25fr) minmax(340px,.75fr); gap:28px;
            border:1px solid rgba(79,109,245,.18); border-radius:28px; padding:34px 38px;
            background:linear-gradient(120deg, rgba(255,255,255,.97), rgba(244,247,255,.87));
            box-shadow:0 20px 55px rgba(49,68,137,.11); overflow:hidden; margin-bottom:22px;
        }}
        .mm-hero h1 {{ font-size:clamp(2.5rem,5vw,4.7rem); line-height:.98; margin:.55rem 0 1rem; background:linear-gradient(90deg,#3f67f4,#7758ef); -webkit-background-clip:text; color:transparent; }}
        .mm-hero p {{ max-width:790px; font-size:1.05rem; line-height:1.75; color:#52607c; }}
        .mm-hero-visual {{ display:flex; align-items:center; justify-content:center; }}
        .mm-hero-visual img {{ width:100%; max-width:560px; }}
        .mm-chip-row {{ display:flex; flex-wrap:wrap; gap:9px; margin-top:20px; }}
        .mm-chip {{ border:1px solid rgba(79,109,245,.22); border-radius:999px; padding:8px 12px; background:rgba(255,255,255,.78); color:#33436d; font-size:.78rem; font-weight:700; }}
        .mm-card {{
            min-height:205px; border:1px solid var(--mm-border); border-radius:20px; padding:22px;
            background:var(--mm-card); box-shadow:0 12px 34px rgba(52,72,139,.08); margin-bottom:16px;
        }}
        .mm-card h3 {{ margin:.55rem 0 .65rem; font-size:1.2rem; }}
        .mm-card p {{ color:#69748e; line-height:1.6; min-height:55px; }}
        .mm-card-foot {{ color:#405bc7; font-weight:750; font-size:.82rem; margin-top:14px; }}
        .mm-section-head {{
            border:1px solid var(--mm-border); border-radius:22px; padding:24px 26px; margin-bottom:18px;
            background:linear-gradient(120deg,rgba(255,255,255,.96),rgba(243,246,255,.88));
            box-shadow:0 12px 34px rgba(52,72,139,.08);
        }}
        .mm-section-head h1 {{ margin:.35rem 0 .55rem; font-size:2.35rem; }}
        .mm-section-head p {{ margin:0; color:#69748e; line-height:1.65; max-width:980px; }}
        .mm-step {{ display:inline-flex; align-items:center; gap:8px; padding:7px 11px; border-radius:999px; background:#eef2ff; color:#405bc7; font-size:.78rem; font-weight:750; margin:0 6px 7px 0; }}
        [data-testid="stFileUploader"] {{ background:rgba(255,255,255,.86); border:1px solid var(--mm-border); border-radius:18px; padding:10px; }}
        [data-testid="stExpander"] {{ background:rgba(255,255,255,.78); border:1px solid var(--mm-border); border-radius:16px; }}
        div.stButton > button, div.stDownloadButton > button {{
            border-radius:12px; min-height:46px; font-weight:750; border:1px solid rgba(79,109,245,.25);
        }}
        div.stButton > button[kind="primary"], div.stDownloadButton > button[kind="primary"] {{
            background:linear-gradient(90deg,#4f6df5,#7658f5); color:white; border:none;
            box-shadow:0 10px 20px rgba(79,109,245,.20);
        }}
        [data-testid="stMetric"] {{ background:rgba(255,255,255,.86); border:1px solid var(--mm-border); border-radius:16px; padding:14px 16px; }}
        .mm-note {{ padding:14px 16px; border-radius:14px; background:rgba(238,242,255,.85); border:1px solid rgba(79,109,245,.14); color:#52607c; line-height:1.55; }}
        .mm-footer {{ margin-top:42px; padding:20px 10px 4px; border-top:1px solid rgba(79,109,245,.14); text-align:center; color:#7a849d; font-size:.78rem; line-height:1.7; }}
        .mm-side-status {{ margin-top:20px; padding:14px; border-radius:15px; border:1px solid var(--mm-border); background:rgba(247,249,255,.9); color:#68738d; font-size:.75rem; line-height:1.6; }}
        .mm-side-rights {{ margin-top:12px; padding:14px; border-radius:15px; border:1px solid var(--mm-border); background:rgba(255,255,255,.9); color:#68738d; font-size:.72rem; line-height:1.6; }}
        code {{ color:#3c56be; }}
        @media (max-width: 980px) {{ .mm-hero {{ grid-template-columns:1fr; padding:28px; }} .mm-hero-visual {{ display:none; }} }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def sidebar_brand() -> None:
    logo = _data_uri(ASSETS / "logo.svg")
    st.sidebar.markdown(
        f"""
        <div class="mm-brand">
          <img src="{logo}" alt="Matching Mapper World logo" />
          <div><div class="mm-brand-title">Matching Mapper World</div><div class="mm-brand-sub">Diamond Matching Studio</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def sidebar_footer() -> None:
    st.sidebar.markdown(
        """
        <div class="mm-side-rights"><strong>Developed by Keyur Bhojak</strong><br/>For diamond matching and stock recommendation analysis.<br/>© 2026 All rights reserved.</div>
        """,
        unsafe_allow_html=True,
    )


def page_footer() -> None:
    st.markdown(
        """
        <div class="mm-footer"><strong>Developed by Keyur Bhojak</strong> for diamond matching and stock recommendation analysis.<br/>© 2026 Matching Mapper World. All rights reserved.</div>
        """,
        unsafe_allow_html=True,
    )


def hero() -> None:
    visual = _data_uri(ASSETS / "hero_visual.svg")
    st.markdown(
        f"""
        <section class="mm-hero">
          <div>
            <div class="mm-kicker">Diamond intelligence • smart matching</div>
            <h1>Matching Mapper World</h1>
            <p>One professional workspace for suggested-stock generation, demand-filter matching, output prioritization, and workbook separation—designed for clear, reliable daily use.</p>
            <div class="mm-chip-row">
              <span class="mm-chip">4 ready-to-use services</span><span class="mm-chip">Exact & tolerance matching</span><span class="mm-chip">Flexible file support</span>
            </div>
          </div>
          <div class="mm-hero-visual"><img src="{visual}" alt="Diamond matching overview" /></div>
        </section>
        """,
        unsafe_allow_html=True,
    )


def section_header(kicker: str, title: str, description: str, steps: list[str]) -> None:
    step_html = "".join(f'<span class="mm-step">{index}. {text}</span>' for index, text in enumerate(steps, 1))
    st.markdown(
        f"""
        <section class="mm-section-head">
          <div class="mm-kicker">{kicker}</div><h1>{title}</h1><p>{description}</p>
          <div style="margin-top:16px">{step_html}</div>
        </section>
        """,
        unsafe_allow_html=True,
    )
