import base64
import html
import json
import os
import re
import tempfile
import unicodedata
from datetime import datetime
from pathlib import Path
import numpy as np
import pandas as pd
import requests
import sqlalchemy as sa
import urllib.parse
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(
    page_title="ItaÃº Asset | Dashboard de Carteiras",
    page_icon="ðŸ“Š",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_DIR = Path(__file__).resolve().parent

BASE_PATH = APP_DIR / "base.csv"

LIBRARY_PATH = APP_DIR / "base.xlsx"

LOGO_PATH = APP_DIR / "logo.png"

SQL_PATH = APP_DIR / "consulta.sql"

RAIOX_SQL_PATH = APP_DIR / "consulta_raiox.sql"

EXTERNAL_RAIOX_SQL_PATH = APP_DIR / "consulta_raiox_externo.sql"

FUND_NAMES_PATH = APP_DIR / "fund_names.json"

COMDINHEIRO_API_URL = "https://www.comdinheiro.com.br/Clientes/API/EndPoint001.ph"

COMDINHEIRO_API_USERNAME = "rfcredit"

COMDINHEIRO_API_PASSWORD = "CreditoIAM300b"

RV_LIBRARY_SHEET = "biblioteca_rv"

MARKET_BASE_PATH = APP_DIR / "base_mercado.xlsx"

EXTERNAL_COMPARISON_FUNDS = {}

def load_market_funds(path, modified_time):
    del modified_time
    market = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    market.columns = [str(c).strip().lower() for c in market.columns]
    required = ["category", "cnpj", "name"]
    missing = [c for c in required if c not in market.columns]
    if missing:
        raise ValueError(f"A base de mercado nÃ£o contÃ©m a(s) coluna(s) obrigatÃ³ria(s): {', '.join(missing)}")
    market = market[required].copy()
    for c in required:
        market[c] = market[c].astype("string").str.strip()
    market = market.dropna(subset=required)
    market = market[(market[required] != "").all(axis=1)].copy()
    market["cnpj_normalizado"] = market["cnpj"].map(normalize_cnpj)
    duplicated = market.loc[market["cnpj_normalizado"].duplicated(False), "cnpj"].unique()
    if len(duplicated):
        raise ValueError(f"CNPJs duplicados em base_mercado.xlsx: {', '.join(duplicated)}")
    market["external_key"] = "EXT_" + market["cnpj_normalizado"]
    return market.sort_values(["category", "name"], kind="stable").reset_index(drop=True)

def refresh_external_comparison_funds():
    if not MARKET_BASE_PATH.exists():
        raise FileNotFoundError(f"Salve {MARKET_BASE_PATH.name} na mesma pasta do dashboard.")
    market = load_market_funds(MARKET_BASE_PATH, MARKET_BASE_PATH.stat().st_mtime_ns)
    EXTERNAL_COMPARISON_FUNDS.clear()
    EXTERNAL_COMPARISON_FUNDS.update({
        r.external_key: {"category": r.category, "cnpj": r.cnpj, "name": r.name}
        for r in market.itertuples(index=False)
    })
    return market

def is_external_comparison_fund(value):
    return str(value) in EXTERNAL_COMPARISON_FUNDS

def comparison_fund_name(value):
    if is_external_comparison_fund(value):
        return EXTERNAL_COMPARISON_FUNDS[str(value)]["name"]
    return get_fund_name(value)

def comparison_fund_label(value):
    if is_external_comparison_fund(value):
        item = EXTERNAL_COMPARISON_FUNDS[str(value)]
        return f"{item['category']} | {item['name']}"
    return f"{get_fund_name(value)} | {value}"

def apply_global_styles():
    st.markdown(
        f"""
        <style>
        /* -------- Fonte / base -------- */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

        html, body, [class*="css"], .stApp {{
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            color: {DARK};
        }}
        .stApp {{
            background: {LIGHT_BG};
        }}

        /* -------- Sidebar escura (identidade Itau) -------- */
        section[data-testid="stSidebar"] {{
            background: {DARK} !important;
            border-right: 3px solid {ORANGE};
        }}
        section[data-testid="stSidebar"] * {{
            color: #FFFFFF !important;
        }}
        section[data-testid="stSidebar"] h1,
        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3 {{
            color: #FFFFFF !important;
            font-weight: 600 !important;
            text-transform: lowercase;
            letter-spacing: 0.01em;
        }}
        section[data-testid="stSidebar"] label {{
            color: #C9CED6 !important;
            font-size: 0.82rem !important;
            font-weight: 500 !important;
            text-transform: lowercase;
        }}
        section[data-testid="stSidebar"] .stSelectbox > div > div,
        section[data-testid="stSidebar"] .stDateInput > div > div,
        section[data-testid="stSidebar"] .stTextInput > div > div {{
            background: {DARK_2} !important;
            border: 1px solid #3A4453 !important;
            border-radius: 4px !important;
        }}
    section[data-testid="stSidebar"] hr {{
        border-color: #3A4453 !important;
        margin: 0.8rem 0 !important;
    }}

    /* --------- Expander SQL sempre legÃ­vel --------- */
    section[data-testid="stSidebar"] details[data-testid="stExpander"] {{
        background: transparent !important;
        border: 1px solid #3A4453 !important;
        border-radius: 6px !important;
        overflow: hidden !important;
    }}
    section[data-testid="stSidebar"] details[data-testid="stExpander"] > summary {{
        background: {DARK_2} !important;
        color: #FFFFFF !important;
        border-radius: 5px !important;
    }}
    section[data-testid="stSidebar"] details[data-testid="stExpander"] > summary:hover,
    section[data-testid="stSidebar"] details[data-testid="stExpander"][open] > summary {{
        background: #354153 !important;
        color: #FFFFFF !important;
    }}
    section[data-testid="stSidebar"] details[data-testid="stExpander"] > summary p,
    section[data-testid="stSidebar"] details[data-testid="stExpander"] > summary span,
    section[data-testid="stSidebar"] details[data-testid="stExpander"] > summary svg {{
        color: #FFFFFF !important;
        fill: #FFFFFF !important;
    }}

    /* --------- Inputs do painel SQL --------- */
    section[data-testid="stSidebar"] input,
    section[data-testid="stSidebar"] textarea {{
        background-color: {DARK_2} !important;
        color: #FFFFFF !important;
        caret-color: {ORANGE} !important;
        -webkit-text-fill-color: #FFFFFF !important;
        border: 1px solid #3A4453 !important;
    }}
    section[data-testid="stSidebar"] input::placeholder,
    section[data-testid="stSidebar"] textarea::placeholder {{
        color: #9FA8B5 !important;
        opacity: 1 !important;
        -webkit-text-fill-color: #9FA8B5 !important;
    }}
    section[data-testid="stSidebar"] input:disabled {{
        background-color: #354153 !important;
        color: #D9DEE5 !important;
        -webkit-text-fill-color: #D9DEE5 !important;
        opacity: 1 !important;
    }}
    section[data-testid="stSidebar"] div[data-baseweb="input"],
    section[data-testid="stSidebar"] div[data-baseweb="textarea"],
    section[data-testid="stSidebar"] div[data-baseweb="base-input"] {{
        background-color: {DARK_2} !important;
        color: #FFFFFF !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stTextArea"] > div,
    section[data-testid="stSidebar"] [data-testid="stTextInput"] > div {{
        background-color: transparent !important;
    }}
    /* Radio da sidebar */
    section[data-testid="stSidebar"] [role="radiogroup"] label {{
        color: #FFFFFF !important;
    }}

    /* --------- Hero / cabecalho --------- */
    .itau-hero {{
        background: {DARK};
        color: #FFFFFF;
        padding: 28px 32px 24px 32px;
        border-radius: 6px;
        margin-bottom: 22px;
        position: relative;
        overflow: hidden;
        box-shadow: 0 2px 12px rgba(28,36,46,0.08);
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 24px;
    }}
    .itau-hero::before {{
        content: '';
        position: absolute;
        top: 0; left: 0;
        width: 6px; height: 100%;
        background: {ORANGE};
    }}
    .itau-hero-content {{
        flex: 1;
        min-width: 0;
    }}
    .itau-hero-logo {{
        flex-shrink: 0;
        display: flex;
        align-items: center;
        justify-content: flex-end;
    }}
    .itau-hero-logo img {{
        max-height: 64px;
        width: auto;
        filter: brightness(0) invert(1);
    }}
    .itau-hero h1 {{
        font-size: 1.9rem !important;
        font-weight: 700 !important;
        color: #FFFFFF !important;
        margin: 0 0 6px 0 !important;
        letter-spacing: -0.02em;
    }}
    .itau-hero .subtitle {{
        color: #C9CED6;
        font-size: 0.92rem;
        font-weight: 400;
        letter-spacing: 0.01em;
    }}
    .itau-hero .brand-tag {{
        display: inline-block;
        color: {ORANGE};
        font-size: 0.75rem;
        font-weight: 600;
        text-transform: lowercase;
        letter-spacing: 0.15em;
        margin-bottom: 8px;
    }}

    /* --------- Section headers com underline laranja --------- */
    .itau-section-title {{
        display: flex;
        align-items: center;
        gap: 14px;
        margin: 8px 0 18px 0;
    }}
    .itau-section-title h3 {{
        font-size: 1.15rem !important;
        font-weight: 600 !important;
        color: {DARK} !important;
        margin: 0 !important;
        text-transform: lowercase;
        letter-spacing: -0.01em;
    }}
    .itau-section-title .accent {{
        flex: 1;
        height: 2px;
        background: {ORANGE};
        opacity: 0.85;
    }}

    /* --------- Metric cards (estilo "R$ 1,3 tri") --------- */
    div[data-testid="stMetric"] {{
        background: #FFFFFF;
        border: 1px solid {BORDER};
        border-left: 4px solid {ORANGE};
        border-radius: 4px;
        padding: 14px 18px 12px 18px;
        box-shadow: 0 1px 3px rgba(28,36,46,0.04);
        transition: box-shadow 0.2s ease;
    }}
    div[data-testid="stMetric"]:hover {{
        box-shadow: 0 4px 12px rgba(28,36,46,0.08);
    }}
    div[data-testid="stMetric"] label {{
        color: {TEXT_MUTED} !important;
        font-size: 0.72rem !important;
        font-weight: 500 !important;
        text-transform: lowercase !important;
        letter-spacing: 0.05em !important;
    }}
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
        color: {ORANGE} !important;
        font-size: 1.75rem !important;
        font-weight: 700 !important;
        letter-spacing: -0.02em;
        line-height: 1.1;
    }}

    /* --------- Tabs (nav Itau) --------- */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 4px;
        background: transparent;
        border-bottom: 1px solid {BORDER};
        padding: 0;
    }}
    .stTabs [data-baseweb="tab"] {{
        background: transparent !important;
        color: {TEXT_MUTED} !important;
        font-weight: 500;
        font-size: 0.92rem;
        text-transform: lowercase;
        letter-spacing: 0.01em;
        padding: 10px 20px !important;
        border-radius: 0 !important;
        border-bottom: 3px solid transparent !important;
        transition: all 0.15s ease;
    }}
    .stTabs [data-baseweb="tab"]:hover {{
        color: {DARK} !important;
        background: rgba(255,98,0,0.04) !important;
    }}
    .stTabs [aria-selected="true"] {{
        color: {DARK} !important;
        border-bottom: 3px solid {ORANGE} !important;
        font-weight: 600 !important;
    }}

    /* --------- Botoes --------- */
    .stDownloadButton button, .stButton button {{
        background: {ORANGE} !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 4px !important;
        font-weight: 600 !important;
        text-transform: lowercase !important;
        letter-spacing: 0.02em !important;
        padding: 10px 22px !important;
        transition: background 0.15s ease !important;
    }}
    .stDownloadButton button:hover, .stButton button:hover {{
        background: #E55700 !important;
        color: #FFFFFF !important;
    }}
    .stDownloadButton button[kind="secondary"] {{
        background: #FFFFFF !important;
        color: {DARK} !important;
        border: 1.5px solid {DARK} !important;
    }}

    /* --------- Seletor horizontal das classes --------- */
    /* O key do widget gera uma classe prÃ³pria no Streamlit. Isso evita
       depender apenas do data-testid, que varia entre versÃµes. */
    .st-key-active_dashboard_section,
    .stkey-active_dashboard_section,
    div[data-testid="stSegmentedControl"] {{
        width: 100%;
        margin: 4px 0 18px 0;
    }}
    .st-key-active_dashboard_section button,
    .stkey-active_dashboard_section button,
    div[data-testid="stSegmentedControl"] button {{
        min-height: 40px !important;
        padding: 8px 16px !important;
        background: #FFFFFF !important;
        color: {DARK} !important;
        border-color: {BORDER} !important;
        font-family: 'Inter', sans-serif !important;
        font-size: 0.83rem !important;
        font-weight: 500 !important;
        box-shadow: none !important;
        transition: background-color 0.15s ease, color 0.15s ease !important;
    }}

    .st-key-active_dashboard_section button:hover,
    .stkey-active_dashboard_section button:hover,
    div[data-testid="stSegmentedControl"] button:hover {{
        background: #F1F2F4 !important;
        color: {DARK} !important;
        border-color: #C9CDD3 !important;
    }}

    /* Na versÃ£o atual do Streamlit, o segmento escolhido Ã© um botÃ£o primary.
       Mantemos seletores adicionais para compatibilidade entre versÃµes. */
    .st-key-active_dashboard_section button[kind="primary"],
    .st-key-active_dashboard_section button[data-testid="stBaseButton-primary"],
    .st-key-active_dashboard_section button[aria-pressed="true"],
    .st-key-active_dashboard_section button[data-selected="true"],
    .stkey-active_dashboard_section button[kind="primary"],
    .stkey-active_dashboard_section button[data-testid="stBaseButton-primary"],
    .stkey-active_dashboard_section button[aria-pressed="true"],
    .stkey-active_dashboard_section button[data-selected="true"],
    div[data-testid="stSegmentedControl"] button[kind="primary"],
    div[data-testid="stSegmentedControl"] button[data-testid="stBaseButton-primary"],
    div[data-testid="stSegmentedControl"] button[aria-pressed="true"],
    div[data-testid="stSegmentedControl"] button[data-selected="true"] {{
        background: {DARK_2} !important;
        color: #FF6200 !important;
        border-color: {DARK_2} !important;
        box-shadow: inset 0 -3px 0 #FF6200 !important;
        font-weight: 700 !important;
    }}

    .st-key-active_dashboard_section button[kind="primary"] *,
    .st-key-active_dashboard_section button[data-testid="stBaseButton-primary"] *,
    .st-key-active_dashboard_section button[aria-pressed="true"] *,
    .st-key-active_dashboard_section button[data-selected="true"] *,
    .stkey-active_dashboard_section button[kind="primary"] *,
    .stkey-active_dashboard_section button[data-testid="stBaseButton-primary"] *,
    .stkey-active_dashboard_section button[aria-pressed="true"] *,
    .stkey-active_dashboard_section button[data-selected="true"] *,
    div[data-testid="stSegmentedControl"] button[kind="primary"] *,
    div[data-testid="stSegmentedControl"] button[data-testid="stBaseButton-primary"] *,
    div[data-testid="stSegmentedControl"] button[aria-pressed="true"] *,
    div[data-testid="stSegmentedControl"] button[data-selected="true"] * {{
        color: #FF6200 !important;
    }}

    /* --------- Radios inline (corpo principal) --------- */
    div[role="radiogroup"] > label {{
        background: #FFFFFF;
        border: 1px solid {BORDER};
        border-radius: 4px;
        padding: 6px 12px;
        margin-right: 6px !important;
        font-size: 0.85rem;
        transition: all 0.15s ease;
        color: {DARK} !important;
    }}
    div[role="radiogroup"] > label p,
    div[role="radiogroup"] > label span {{
        color: {DARK} !important;
    }}

    /* --------- Radios da sidebar (fundo escuro, texto branco) --------- */
    section[data-testid="stSidebar"] div[role="radiogroup"] > label {{
        background: {DARK_2} !important;
        border: 1px solid #3A4453 !important;
        color: #FFFFFF !important;
    }}
    section[data-testid="stSidebar"] div[role="radiogroup"] > label p,
    section[data-testid="stSidebar"] div[role="radiogroup"] > label span,
    section[data-testid="stSidebar"] div[role="radiogroup"] > label div {{
        color: #FFFFFF !important;
    }}
    /* Bolinha selecionada em laranja na sidebar */
    section[data-testid="stSidebar"] div[role="radiogroup"] input:checked + div,
    section[data-testid="stSidebar"] div[role="radiogroup"] [data-checked="true"] {{
        background-color: {ORANGE} !important;
        border-color: {ORANGE} !important;
    }}

    /* --------- Titulos gerais e caption --------- */
    h1, h2, h3, h4 {{
        color: {DARK};
        font-weight: 600;
        letter-spacing: -0.01em;
    }}
    .stCaption, .caption {{
        color: {TEXT_MUTED};
    }}

    /* --------- Divisor --------- */
    hr {{
        border-color: {BORDER};
    }}

    /* --------- Info/warning boxes --------- */
    div[data-testid="stAlert"] {{
        border-radius: 4px;
        border-left: 4px solid {ORANGE};
    }}

    /* --------- Tabelas HTML embutidas --------- */
    table.itau-table {{
        width: 100%;
        border-collapse: collapse;
        font-family: 'Inter', sans-serif;
        font-size: 0.87rem;
        margin: 8px 0;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

DARK = "#1C242E"
DARK_2 = "#2A3340"
ORANGE = "#FF6200"
ORANGE_SOFT = "#FF8A3D"
LIGHT_BG = "#F5F6F8"
LIGHT_BLUE = "#E9F4FA"
BORDER = "#E4E6EA"
TEXT_MUTED = "#5B6472"

COLORS = [
    "#FF6200",  # laranja Itau
    "#1C242E",  # navy
    "#1A9DC4",  # azul
    "#F45880",  # rosa
    "#A850CA",  # roxo
    "#10B981",  # verde
    "#F59E0B",  # ambar
    "#6B7280",  # cinza
    "#7988D1",
    "#CC0000",
    "#35ADD0",
    "#7C3AED",
    "#84CC16",
    "#D97706",
    "#64748B",
]

REQUIRED = {"date", "codigo_IAM_fundo", "codigo_IAM_ativo", "perc_pl"}

FUND_NAME_MAP = {
    "HIGHYEMM53556": "HIGHYEMM53556",
    "ITAU_DI_348": "ITAU_DI_348",
    "DIFIIRF52888": "DIFIIRF52888",
    "FLEXPREVDI1928": "FLEXPREVDI1928",
    "SPECIAL_DI_192": "SPECIAL_DI_192",
    "FIDELILCMM771": "FIDELILCMM771",
    "INTRAGITAU675": "PRIVILEGE",
    "WEALTHMAS53004": "WEALTHMAS53004",
    "DIFERENRF505": "DIFERENRF505",
    "ACTFIXRF51044": "ACTFIXRF51044",
    "ACTIVEMM52230": "ACTIVEMM52230",
    "ITAURFIQ53566": "ITAURFIQ53566",
    "ACTIVFIX53867": "ACTIVFIX53867",
    "ADFIM54217": "ADFIM54217",
    "TOP_MIX_HY_200": "TOP_MIX_HY_200",
    "ACTFIXFIM456": "ACTFIXFIM456",
    "DEBINFRARF494": "DEBINFRARF494",
    "PRECIFIM55518": "PRECIFIM55518",
    "SINFONIA718": "SINFONIA718",
    "DEBINFRA55846": "DEBINFRA55846",
    "ACTINFRF55874": "ACTINFRF55874",
    "INFRARF55751": "INFRARF55751",
    "INFRAS56712": "INFRAS56712",
    "FIM57631": "FIM57631",
    "INVEGRADERF1625": "INVEGRADERF1625",
    "FIM57240": "FIM57240",
    "FLEXPF159079": "FLEXPF159079",
    "XPSFIM57442": "XPSFIM57442",
}

RV_LIBRARY_FIELDS = [
    "setor",
    "grupo_industria_gics",
    "subsetor",
    "issuer",
    "pais",
]

RV_SOURCE_COLUMNS = {
    "setor": ["setor", "setor_dashboard", "setor_iam", "gics_sector_name"],
    "grupo_industria_gics": ["grupo_industria_gics"],
    "subsetor": ["subsetor", "subsetor_iam"],
    "issuer": ["issuer", "emissor"],
    "pais": ["pais", "pais_dashboard", "tipo_pais_exposicao"],
}

BRAZIL = {"BRAZIL", "BRASIL", "BR", "BRA"}

NORTH_AMERICA = {"UNITED STATES", "UNITED STATES OF AMERICA", "USA", "US", "ESTADOS UNIDOS", "CANADA", "CA", "CAN"}

LATAM = {"ARGENTINA", "CHILE", "COLOMBIA", "MEXICO", "PERU", "URUGUAY", "PARAGUAY", "BOLIVIA", "ECUADOR", "VENEZUELA", "PANAMA", "COSTA RICA", "GUATEMALA", "HONDURAS", "EL SALVADOR", "NICARAGUA", "BELIZE", "DOMINICAN REPUBLIC", "REPUBLICA DOMINICANA", "JAMAICA", "TRINIDAD AND TOBAGO", "PUERTO RICO", "BAHAMAS", "BARBADOS", "CAYMAN ISLANDS", "CAYMAN"}

EUROPE = {"UNITED KINGDOM", "UK", "GREAT BRITAIN", "REINO UNIDO", "GERMANY", "ALEMANHA", "FRANCE", "FRANCA", "ITALY", "ITALIA", "SPAIN", "ESPANHA", "PORTUGAL", "NETHERLANDS", "HOLANDA", "BELGIUM", "BELGICA", "SWITZERLAND", "SUICA", "AUSTRIA", "IRELAND", "IRLANDA", "DENMARK", "DINAMARCA", "SWEDEN", "SUECIA", "NORWAY", "NORUEGA", "FINLAND", "FINLANDIA", "POLAND", "POLONIA", "GREECE", "GRECIA", "LUXEMBOURG", "LUXEMBURGO", "TURKEY", "TURKIYE", "TURQUIA", "RUSSIA", "CZECH REPUBLIC", "CZECHIA", "HUNGARY", "HUNGRIA", "ROMANIA", "ROMENIA"}

ASIA = {"CHINA", "HONG KONG", "JAPAN", "JAPAO", "SOUTH KOREA", "COREIA DO SUL", "TAIWAN", "INDIA", "INDONESIA", "SINGAPORE", "SINGAPURA", "MALAYSIA", "MALASIA", "THAILAND", "TAILANDIA", "VIETNAM", "PHILIPPINES", "FILIPINAS", "PAKISTAN", "BANGLADESH", "ISRAEL", "SAUDI ARABIA", "ARABIA SAUDITA", "UNITED ARAB EMIRATES", "EMIRADOS ARABES UNIDOS", "QATAR", "KUWAIT"}

def clean_text(series):
    return (
        series.astype("string")
        .str.strip()
        .replace({"": pd.NA, "None": pd.NA, "nan": pd.NA})
    )

def first_existing(df, names, default="Nao classificado"):
    result = pd.Series(pd.NA, index=df.index, dtype="object")
    for name in names:
        if name in df.columns:
            result = result.fillna(clean_text(df[name]))
    return result.fillna(default)

def normalized_key(value):
    if value is None or pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", text).strip().upper()

UNCLASSIFIED_KEYS = {
    "", "NAO CLASSIFICADO", "NAO CLASSIFICADA", "N/A", "N A", "NA",
    "NAN", "NONE", "NULL", "<NA>", "-",
}

def is_unclassified(value):
    return normalized_key(value) in UNCLASSIFIED_KEYS

def load_custom_fund_names():
    """Carrega nomes cadastrados pelo usuÃ¡rio sem alterar o cÃ³digo-fonte."""
    if not FUND_NAMES_PATH.exists():
        return {}
    try:
        content = json.loads(FUND_NAMES_PATH.read_text(encoding="utf-8"))
        return {str(code).strip(): str(name).strip() for code, name in content.items()}
    except Exception:
        return {}

CUSTOM_FUND_NAME_MAP = load_custom_fund_names()

def get_fund_name(fund_code):
    """Retorna nome cadastrado; usa o cÃ³digo IAM quando nÃ£o houver cadastro."""
    code = str(fund_code).strip()
    return CUSTOM_FUND_NAME_MAP.get(code, FUND_NAME_MAP.get(code, code))

def _rv_reference_columns(data):
    """ObtÃ©m ticker, risco de mercado e o pacote cadastral usado pela biblioteca."""
    references = {
        "ticker": first_existing(
            data,
            ["ticker", "ticker_dashboard", "ticker_ativo", "ticker_bolsa", "security_name", "codigo_IAM_ativo"],
            pd.NA,
        ),
        "risco_mercado": first_existing(
            data,
            ["fator_risco", "fator_risco_dashboard", "fator_risco_mercado"],
            pd.NA,
        ),
    }
    for field, source_columns in RV_SOURCE_COLUMNS.items():
        references[field] = first_existing(data, source_columns, pd.NA)
    return references

def update_rv_library_and_sectors(data, library=None):
    """
    MantÃ©m a biblioteca cadastral de Renda VariÃ¡vel por ticker e preenche, na
    base, apenas atributos vazios ou nÃ£o classificados encontrados na biblioteca.

    Valores preenchidos manualmente na biblioteca tÃªm prioridade. Campos vazios
    da biblioteca sÃ£o complementados pelo histÃ³rico vÃ¡lido da prÃ³pria base.
    """
    result = data.copy()
    refs = _rv_reference_columns(result)
    ticker_key = refs["ticker"].map(normalized_key)
    rv_mask = refs["risco_mercado"].map(normalized_key).eq("RENDA VARIAVEL")
    library_columns = ["ticker", *RV_LIBRARY_FIELDS]

    if library is None:
        library = pd.DataFrame(columns=library_columns)
    else:
        library = library.copy()
        library.columns = [str(column).strip() for column in library.columns]
        for column in library_columns:
            if column not in library.columns:
                library[column] = pd.NA
        library = library[library_columns]

    library["_ticker_key"] = library["ticker"].map(normalized_key)
    library = (
        library.loc[library["_ticker_key"].ne("")]
        .drop_duplicates("_ticker_key", keep="last")
        .reset_index(drop=True)
    )

    # Cria na biblioteca todos os tickers de Renda VariÃ¡vel presentes na base.
    ticker_candidates = pd.DataFrame({
        "ticker": refs["ticker"],
        "_ticker_key": ticker_key,
        "_date": pd.to_datetime(result.get("date"), errors="coerce"),
    }).loc[rv_mask & ticker_key.ne("")]
    if not ticker_candidates.empty:
        ticker_candidates = (
            ticker_candidates.sort_values("_date", ascending=False, kind="stable")
            .drop_duplicates("_ticker_key", keep="first")
        )
        missing_tickers = ticker_candidates.loc[
            ~ticker_candidates["_ticker_key"].isin(library["_ticker_key"])
        ].copy()
        for field in RV_LIBRARY_FIELDS:
            missing_tickers[field] = pd.NA
        library = pd.concat(
            [library, missing_tickers[library_columns + ["_ticker_key"]]],
            ignore_index=True,
        )

    # Completa cada atributo da biblioteca de forma independente. Assim, um
    # ticker pode aprender paÃ­s de uma linha e subsetor de outra, sem sobrescrever
    # ajustes vÃ¡lidos ou manuais jÃ¡ existentes na biblioteca.
    row_dates = (
        pd.to_datetime(result["date"], errors="coerce")
        if "date" in result.columns
        else pd.Series(pd.NaT, index=result.index)
    )
    for field in RV_LIBRARY_FIELDS:
        library[field] = library[field].astype("object")
        current = refs[field]
        history = pd.DataFrame({
            "_ticker_key": ticker_key,
            "_value": current,
            "_date": row_dates,
        }).loc[rv_mask & ticker_key.ne("") & ~current.map(is_unclassified)]
        if history.empty:
            continue
        latest = (
            history.sort_values("_date", ascending=False, kind="stable")
            .drop_duplicates("_ticker_key", keep="first")
            .set_index("_ticker_key")["_value"]
        )
        missing = library[field].map(is_unclassified)
        library.loc[missing, field] = library.loc[missing, "_ticker_key"].map(latest)

    # Aplica a biblioteca na base somente onde o atributo estÃ¡ vazio ou nÃ£o
    # classificado, registrando cada correÃ§Ã£o realizada.
    corrections = []
    library_lookup = library.set_index("_ticker_key")
    for field in RV_LIBRARY_FIELDS:
        current = refs[field]
        fill_values = ticker_key.map(library_lookup[field])
        apply_mask = (
            rv_mask
            & ticker_key.ne("")
            & current.map(is_unclassified)
            & ~fill_values.map(is_unclassified)
        )
        if not apply_mask.any():
            continue
        if field in result.columns:
            result[field] = result[field].astype("object")
        else:
            result[field] = current.astype("object")
        result.loc[apply_mask, field] = fill_values[apply_mask]
        corrections.append(pd.DataFrame({
            "ticker": refs["ticker"][apply_mask],
            "campo": field,
            "valor_anterior": current[apply_mask],
            "valor_novo": fill_values[apply_mask],
        }))

    corrections = (
        pd.concat(corrections, ignore_index=True)
        if corrections
        else pd.DataFrame(columns=["ticker", "campo", "valor_anterior", "valor_novo"])
    )

    library_output = (
        library[library_columns]
        .sort_values("ticker", key=lambda values: values.astype("string").str.upper(), kind="stable")
        .reset_index(drop=True)
    )
    return result, library_output, corrections

def load_rv_library(path):
    """LÃª a biblioteca cadastral; retorna a estrutura completa quando a aba nÃ£o existe"""
    try:
        return pd.read_excel(path, sheet_name=RV_LIBRARY_SHEET, engine="openpyxl")
    except (FileNotFoundError, ValueError):
        return pd.DataFrame(columns=["ticker", *RV_LIBRARY_FIELDS])

def load_csv(path, modified_time):
    del modified_time
    df = pd.read_csv(
        path,
        sep=";",
        decimal=",",
        encoding="utf-8-sig",
        low_memory=False,
    )

    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    missing = REQUIRED - set(df.columns)

    if missing:
        raise ValueError(
            f"Colunas obrigatÃ³rias ausentes: {sorted(missing)}"
        )

    # Datas do CSV estÃ£o no padrÃ£o brasileiro.
    df["date"] = pd.to_datetime(
        df["date"],
        format="%d/%m/%Y",
        errors="coerce",
    ).dt.normalize()

    df["codigo_IAM_fundo"] = (
        clean_text(df["codigo_IAM_fundo"])
        .str.upper()
    )

    df["codigo_IAM_ativo"] = clean_text(
        df["codigo_IAM_ativo"]
    )

    # O read_csv jÃ¡ utiliza decimal=",", mas mantemos a
    # conversÃ£o numÃ©rica para validar os campos.
    df["perc_pl"] = pd.to_numeric(
        df["perc_pl"],
        errors="coerce",
    )

    if "eq_ano" not in df.columns:
        df["eq_ano"] = np.nan

    df["eq_ano"] = pd.to_numeric(
        df["eq_ano"],
        errors="coerce",
    )

    if "quantidade" not in df.columns:
        df["quantidade"] = np.nan

    df["quantidade"] = pd.to_numeric(
        df["quantidade"],
        errors="coerce",
    )

    if "tipo_fechamento" not in df.columns:
        df["tipo_fechamento"] = pd.NA

    df["tipo_fechamento"] = clean_text(
        df["tipo_fechamento"]
    )

    if "vencimento" in df.columns:
        df["vencimento"] = pd.to_datetime(
            df["vencimento"],
            dayfirst=True,
            errors="coerce",
        ).dt.normalize()

    rows_before = len(df)

    df = df.dropna(
        subset=[
            "date",
            "codigo_IAM_fundo",
            "codigo_IAM_ativo",
            "perc_pl",
        ]
    ).copy()

    rows_removed = rows_before - len(df)

    if df.empty:
        raise ValueError(
            "A base ficou vazia apÃ³s validar date, codigo_IAM_fundo, "
            "codigo_IAM_ativo e perc_pl. Verifique o formato do CSV."
        )

    print(
        f"[INFO] CSV carregado: {len(df):,} linhas vÃ¡lidas; "
        f"{rows_removed:,} linhas removidas."
    )

    library = load_rv_library(LIBRARY_PATH)

    df, _, _ = update_rv_library_and_sectors(
        df,
        library,
    )

    return df

def load_raiox_default_probability_history(
    server,
    driver,
    username,
    password,
    fund_code,
    start_date,
    end_date,
):
    """Retorna a soma diÃ¡ria da PD ponderada por perc_pl para o fundo."""
    query = """
    SELECT
        P.[date],
        P.codigo_IAM_fundo,
        SUM(
            CAST(P.perc_pl AS float) * CAST(R.PD AS float)
        ) AS probabilidade_default
    FROM VIEW_FUNDOS_POSICAO AS P
    INNER JOIN dbsx362.dbo.CLUSTER_RATING AS R
      ON LTRIM(RTRIM(P.rating_itau)) = LTRIM(RTRIM(R.rating_iam))
    WHERE P.codigo_IAM_fundo = ?
      AND P.[date] >= ?
      AND P.[date] <= ?
      AND NULLIF(LTRIM(RTRIM(P.isin)), '') IS NOT NULL
      AND NULLIF(LTRIM(RTRIM(P.rating_itau)), '') IS NOT NULL
      AND P.perc_pl IS NOT NULL
      AND R.PD IS NOT NULL
    GROUP BY
        P.[date],
        P.codigo_IAM_fundo
    ORDER BY
        P.[date];
    """
    params = [
        str(fund_code).strip().upper(),
        pd.Timestamp(start_date).date(),
        pd.Timestamp(end_date).date(),
    ]
    engine = conn(username=username, password=password)
    try:
        result = pd.read_sql(query, engine, params=tuple(params))
    finally:
        engine.dispose()

    if result.empty:
        return result
    result.columns = [str(column).strip() for column in result.columns]
    result["date"] = pd.to_datetime(result["date"], errors="coerce").dt.normalize()
    result["probabilidade_default"] = pd.to_numeric(
        result["probabilidade_default"], errors="coerce"
    )
    return result.dropna(subset=["date", "probabilidade_default"]).copy()

def raiox_default_probability_history_chart(data, title):
    """GrÃ¡fico em linha da soma diÃ¡ria da PD ponderada."""
    if data.empty:
        return None
    fig = px.line(
        data.sort_values(["Fundo", "date"], kind="stable"),
        x="date",
        y="probabilidade_default",
        color="Fundo",
        markers=True,
        title=title,
        color_discrete_sequence=[ORANGE, DARK],
        labels={
            "date": "Data",
            "probabilidade_default": "Probabilidade de Default",
        },
    )
    apply_layout(fig, title)
    fig.update_layout(
        height=500,
        hovermode="x unified",
        margin=dict(l=10, r=20, t=70, b=35),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
        ),
    )
    fig.update_yaxes(tickformat=".3%", title="Probabilidade de Default")
    fig.update_xaxes(title="", tickformat="%d/%m/%Y")
    fig.update_traces(
        hovertemplate=(
            "Data: %{x|%d/%m/%Y}<br>"
            "Probabilidade de Default: %{y:.4%}"
            "<extra>%{fullData.name}</extra>"
        )
    )
    return fig

def load_raiox_default_probability(
    server,
    driver,
    username,
    password,
    fund_code,
    reference_date,
):
    """Retorna a soma da PD para o fundo e a data-base selecionados."""
    query = """
    SELECT
        SUM(CAST(P.perc_pl AS float) * CAST(R.PD AS float)) AS probabilidade_default
    FROM VIEW_FUNDOS_POSICAO AS P
    INNER JOIN dbsx362.dbo.CLUSTER_RATING AS R
      ON LTRIM(RTRIM(P.rating_itau)) = LTRIM(RTRIM(R.rating_iam))
    WHERE P.codigo_IAM_fundo = ?
      AND P.[date] = ?
    """
    params = [
        str(fund_code).strip().upper(),
        pd.Timestamp(reference_date).date(),
    ]
    engine = conn(username=username, password=password)
    try:
        result = pd.read_sql(query, engine, params=tuple(params))
    finally:
        engine.dispose()

    if result.empty or "probabilidade_default" not in result.columns:
        return np.nan
    value = pd.to_numeric(result["probabilidade_default"], errors="coerce").iloc[0]
    return float(value) if pd.notna(value) else np.nan

def load_credit_raiox(
    sql_path,
    sql_modified_time,
    server,
    driver,
    username,
    password,
    fund_code,
    reference_date,
):
    """
    Executa consulta_raiox.sql e retorna a base analÃ­tica
    utilizada pela aba Raio-X
    """
    del sql_modified_time

    sql_file = Path(sql_path)
    if not sql_file.exists():
        raise FileNotFoundError(
            f"Arquivo SQL nÃ£o encontrado: {sql_file}"
        )

    query = sql_file.read_text(
        encoding="utf-8"
    )

    # A query usada pelo Python deve possuir:
    # DECLARE @DataRef date = ?;
    # DECLARE @CodigoFundo varchar(100) = ?;
    #
    # No SSMS, esses parÃ¢metros devem ser substituÃ­dos
    # por valores fixos para teste.

    params = [
        pd.Timestamp(reference_date).date(),
        str(fund_code).strip().upper(),
    ]

    engine = conn(username=username, password=password)
    try:
        result = pd.read_sql(query, engine, params=tuple(params))
    finally:
        engine.dispose()

    if result.empty:
        return result

    result.columns = [
        str(column).strip()
        for column in result.columns
    ]

    if "date" in result.columns:
        result["date"] = pd.to_datetime(
            result["date"],
            errors="coerce",
        ).dt.normalize()

    numeric_columns = [
        "perc_pl",
        "perc_nav",
        "modified_duration",
        "spread_cdi_ativo",
        "spread_ipca_ativo",
        "spread_ntnb_ativo",
        "spread_equivalente_ativo",
        "perc_pl_credito",
        "perc_pl_credito_sem_cotas",
        "perc_pl_credito_tem_spread",
        "perc_pl_considerado",
        "spread_equivalente",
        "spread_equivalente_credito",
        "spread_ntnb_equivalente",
        "spread_ntnb_equivalente_credito",
        "spread_cdi",
        "spread_cdi_credito",
        "premio",
        "premio_credito",
        "moddur_credito",
        "moddur_fundo",
        "patrimonio",
        "perc_debin",
        "taxa_administracao_total",
        "taxa_performance",
        "yield_imab_referencia",
        "carrego_liquido",
        "carrego_liquido_b_equivalente",
        "carrego_liquido_imab",
    ]

    for column in numeric_columns:
        if column in result.columns:
            result[column] = pd.to_numeric(
                result[column],
                errors="coerce",
            )

    text_columns = [
        "instrumento",
        "setor",
        "grupo_economico",
        "indexador",
        "rating_externo",
        "benchmark",
        "benchmark_ajustado",
    ]

    for column in text_columns:
        if column in result.columns:
            result[column] = clean_text(
                result[column]
            )

    return result

def load_external_comparison_raiox(external_key, reference_date, username, password):
    import time
    diagnostics = {"steps": [], "tables": {}, "raw_json": None}
    def step(name, status, detail):
        diagnostics["steps"].append({"etapa": name, "status": status, "detalhe": str(detail)})
        st.session_state["external_raiox_diagnostics"] = diagnostics

    item = EXTERNAL_COMPARISON_FUNDS[external_key]
    reference_date = pd.Timestamp(reference_date).normalize()
    query_path = build_comdinheiro_query_path(item["cnpj"], reference_date)
    step("Cadastro", "SUCESSO", f"{item['name']} | CNPJ {item['cnpj']}")
    step("Data-base", "SUCESSO", f"{reference_date:%d/%m/%Y}")
    step("Query path", "SUCESSO", query_path)
    def request_portfolio(request_date):
        url = COMDINHEIRO_API_URL
        querystring = {
            "code": "import_data"
        }
        cnpj = normalize_cnpj(item["cnpj"])
        data_api = pd.Timestamp(request_date).strftime("%Y%m%d")
        request_path = (
            f"CarteiraFundo006-{cnpj}-{data_api}-"
            "acrescentar_cnpj_fundo_isin-1-default-5-2-"
            "desc_ativo-final_valor+percent"
        )
        payload_api = {
            "username": COMDINHEIRO_API_USERNAME,
            "password": COMDINHEIRO_API_PASSWORD,
            "URL": request_path,
            "format": "json3"
        }
        response = requests.post(
            url,
            params=querystring,
            data=payload_api,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=120,
        )
        response.raise_for_status()
        return response.json()

    started = time.perf_counter()
    try:
        payload = request_portfolio(reference_date)
    except Exception as exc:
        step("API Comdinheiro", "ERRO", exc)
        raise
    diagnostics["raw_json"] = payload
    step("API Comdinheiro", "SUCESSO", f"Resposta recebida em {time.perf_counter() - started:.1f}s")

    try:
        portfolio = parse_comdinheiro_portfolio(payload, item["cnpj"], reference_date)
    except Exception as exc:
        step("Leitura da carteira", "ERRO", exc)
        raise
    step("Leitura da carteira", "SUCESSO", f"{len(portfolio)} linhas retornadas")

    # Apenas as linhas finais da hierarquia (sem filhos) evitam dupla contagem.
    leaves = portfolio.loc[portfolio["linha_falha"]].copy()

    try:
        sql_data = load_isin_credit_information(
            server="SQNPRC028",
            driver="ODBC Driver 17 for SQL Server",
            username=username,
            password=password,
            isins=leaves["isin"].tolist(),
            reference_date=reference_date,
        )
    except Exception as exc:
        step("InformaÃ§Ãµes internas por ISIN", "ERRO", exc)
        raise
    step("InformaÃ§Ãµes internas por ISIN", "SUCESSO", f"{len(sql_data)} ISINs encontrados")

    enriched = enrich_comdinheiro_portfolio(leaves, sql_data)
    diagnostics["tables"]["Carteira Comdinheiro"] = portfolio
    diagnostics["tables"]["Carteira enriquecida"] = enriched
    step(
        "AssociaÃ§Ã£o",
        "SUCESSO",
        enriched["status_associacao"].value_counts().to_dict(),
    )

    def text_or_default(column, default="NÃ£o classificado"):
        if column in enriched.columns:
            return clean_text(enriched[column]).fillna(default)
        return pd.Series(default, index=enriched.index, dtype="object")

    instrument = text_or_default("tipo", pd.NA)
    if "descricao_comdinheiro" in enriched.columns:
        instrument = instrument.fillna(clean_text(enriched["descricao_comdinheiro"]))
    group = text_or_default("ativo_sql", pd.NA)
    if "ativo_comdinheiro" in enriched.columns:
        group = group.fillna(clean_text(enriched["ativo_comdinheiro"]))

    top_level = portfolio.loc[portfolio["nivel_hierarquia"].eq(0)]
    patrimonio = (
        top_level["posicao_final"].sum(min_count=1)
        if "posicao_final" in top_level.columns
        else np.nan
    )

    secondary = pd.DataFrame({
        "date": reference_date,
        "codigo_IAM_fundo": external_key,
        "isin": enriched["isin"],
        "codigo_IAM_ativo": enriched.get("codigo_IAM_ativo"),
        "instrumento": instrument.fillna("NÃ£o classificado"),
        "grupo_economico": group.fillna("NÃ£o classificado"),
        "indexador": text_or_default("indexador"),
        "rating_externo": text_or_default("rating_itau"),
        "setor": text_or_default("setor"),
        "perc_pl": enriched["perc_pl"],
        "perc_nav": enriched["perc_pl"],
        "PD": enriched["PD"],
        "contribuicao_PD": enriched["contribuicao_PD"],
        "status_associacao": enriched["status_associacao"],
        "patrimonio": patrimonio,
        "benchmark": "CDI",
    })
    secondary_pd = enriched["contribuicao_PD"].sum(min_count=1)
    secondary_pd = float(secondary_pd) if pd.notna(secondary_pd) else np.nan
    step("PD ponderada", "SUCESSO", fmt_pct(secondary_pd, 4))
    return secondary, secondary_pd

def build_raiox_distribution(
    data,
    category,
    value_column="perc_pl",
    denominator=None,
    top_n=None,
    others_position="last",
):
    """
    Agrupa a carteira por categoria. Com denominator=None o percentual Ã© o
    prÃ³prio valor (perc_pl jÃ¡ estÃ¡ em fraÃ§Ã£o do PL); com denominator="total"
    os valores sÃ£o normalizados pela soma; com um nÃºmero, divide por ele.
    """
    columns = [category, "valor", "percentual", "ordem"]
    if (
        data is None
        or data.empty
        or category not in data.columns
        or value_column not in data.columns
    ):
        return pd.DataFrame(columns=columns)

    base = data[[category, value_column]].copy()
    base[category] = clean_text(base[category]).fillna("NÃ£o classificado")
    base[value_column] = pd.to_numeric(base[value_column], errors="coerce").fillna(0.0)

    grouped = (
        base.groupby(category, as_index=False, dropna=False)[value_column]
        .sum()
        .rename(columns={value_column: "valor"})
    )
    grouped[category] = grouped[category].astype(str)

    if denominator is None:
        grouped["percentual"] = grouped["valor"]
    else:
        total = grouped["valor"].sum() if denominator == "total" else float(denominator)
        grouped["percentual"] = grouped["valor"] / total if total else np.nan

    grouped["_ranking"] = grouped["percentual"].abs()
    grouped = grouped.sort_values(
        "_ranking", ascending=False, kind="stable"
    ).reset_index(drop=True)

    if top_n is not None and len(grouped) > top_n:
        top = grouped.head(top_n).copy()
        rest = grouped.iloc[top_n:]
        others = pd.DataFrame({
            category: [f"Outros ({len(rest)})"],
            "valor": [rest["valor"].sum()],
            "percentual": [rest["percentual"].sum()],
            "_ranking": [rest["_ranking"].sum()],
        })
    else:
        top = grouped
        others = grouped.iloc[0:0].copy()

    grouped = pd.concat(
        [
            top,
            others,
        ],
        ignore_index=True,
    )

    if others_position == "last":
        grouped["_is_others"] = (
            grouped[category]
            .astype(str)
            .str.startswith("Outros")
            .astype(int)
        )

        grouped = grouped.sort_values(
            [
                "_is_others",
                "_ranking",
            ],
            ascending=[
                True,
                False,
            ],
            kind="stable",
        )

    elif others_position == "first":
        grouped["_is_others"] = np.where(
            grouped[category]
            .astype(str)
            .str.startswith("Outros"),
            0,
            1,
        )

        grouped = grouped.sort_values(
            [
                "_is_others",
                "_ranking",
            ],
            ascending=[
                True,
                False,
            ],
            kind="stable",
        )

    grouped["ordem"] = np.arange(
        1,
        len(grouped) + 1,
    )

    return grouped.drop(
        columns=[
            "_ranking",
            "_is_others",
        ],
        errors="ignore",
    ).reset_index(drop=True)

def apply_layout(fig, title=None):
    fig.update_layout(
        template="plotly_white",
        hovermode="x unified",
        legend_title_text="",
        margin=dict(l=10, r=20, t=60, b=35),

        paper_bgcolor="white",
        plot_bgcolor="white",

        font=dict(
            color=DARK,
            family="Inter, sans-serif",
            size=11,
        ),

        title=dict(
            text=title if title else fig.layout.title.text,
            font=dict(
                color=DARK,
                size=13,
                family="Inter, sans-serif"
            ),
            x=0.01,
            xanchor="left",
        ) if (title or fig.layout.title.text) else None,

        legend=dict(
            font=dict(
                color=DARK,
                size=11
            ),
            bgcolor="rgba(255,255,255,0)"
        )
    )

    fig.update_xaxes(
        color=DARK,
        tickfont=dict(
            color=DARK,
            size=11,
        ),
        title_font=dict(
            color=DARK,
            size=13,
        ),
        gridcolor="#F1F2F3",
        linecolor="#D1D1D0",
    )

    fig.update_yaxes(
        color=DARK,
        tickfont=dict(
            color=DARK,
            size=11,
        ),
        title_font=dict(
            color=DARK,
            size=13,
        ),
        gridcolor="#F1F2F3",
        linecolor="#D1D1D0",
    )

    return fig

def raiox_horizontal_bar(
    data,
    category,
    title,
    color=ORANGE,
    height=420,
):
    if data.empty:
        return None

    chart_data = data.copy()

    chart_data["texto"] = chart_data[
        "percentual"
    ].map(
        lambda value: (
            "-"
            if pd.isna(value)
            else f"{value:.1%}".replace(".", ",")
        )
    )

    # Plotly horizontal mostra o Ãºltimo item no topo.
    chart_data = chart_data.sort_values(
        "percentual",
        ascending=False,
        kind="stable",
    )

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=chart_data["percentual"],
            y=chart_data[category],
            orientation="h",
            marker=dict(
                color=color,
            ),
            text=chart_data["texto"],
            textposition="outside",
            cliponaxis=False,
            customdata=np.column_stack(
                [
                    chart_data["valor"],
                ]
            ),
            hovertemplate=(
                "%{y}<br>"
                "Percentual: %{x:.2%}<br>"
                "ExposiÃ§Ã£o: %{customdata[0]:.2%}"
                "<extra></extra>"
            )
        )
    )

    apply_layout(
        fig,
        title,
    )

    fig.update_layout(
        height=height,
        showlegend=False,
        hovermode="closest",
        margin=dict(
            l=10,
            r=80,
            t=60,
            b=35,
        )
    )

    fig.update_xaxes(
        tickformat=".0%",
        title="",
        showgrid=False,
        zeroline=False,
        visible=False,
    )

    fig.update_yaxes(
        title="",
        showgrid=False,
        autorange="reversed",
    )

    return fig

def raiox_pie_chart(
    data,
    category,
    title,
    height=410,
):
    if data.empty:
        return None

    chart_data = data.copy()

    # Evita fatias sem valor ou negativas no grÃ¡fico de pizza.
    chart_data = chart_data.loc[
        chart_data["percentual"].gt(0)
    ].copy()

    if chart_data.empty:
        return None

    fig = go.Figure()

    fig.add_trace(
        go.Pie(
            labels=chart_data[category],
            values=chart_data["percentual"],
            sort=False,
            direction="clockwise",
            marker=dict(
                colors=COLORS,
                line=dict(
                    color="#FFFFFF",
                    width=1.5,
                ),
            ),
            texttemplate=(
                "%{label}<br>%{percent:.1%}"
            ),
            textposition="auto",
            hovertemplate=(
                "%{label}<br>"
                "Percentual: %{value:.2%}"
                "<extra></extra>"
            ),
        )
    )

    apply_layout(
        fig,
        title,
    )

    fig.update_layout(
        height=height,
        showlegend=True,
        hovermode="closest",
        margin=dict(
            l=10,
            r=20,
            t=60,
            b=35,
        ),
    )

    return fig

def format_raiox_money(value):
    if pd.isna(value):
        return "-"
    # MantÃ©m o padrÃ£o visual da macro em R$ milhÃµes.
    return (
        f"{value / 1_000_000:,.1f}"
        .replace(".", "x")
        .replace(",", ".")
        .replace("x", ",")
    )

def format_raiox_duration(value):
    if pd.isna(value):
        return "-"
    return (
        f"{value:.1f} anos"
        .replace(".", ",")
    )

def format_raiox_spread(
    value,
    benchmark="CDI",
):
    if pd.isna(value):
        return "-"
    signal = "+" if value >= 0 else "-"
    formatted = (
        f"{abs(value):.2f}"
        .replace(".", ",")
    )
    benchmark_text = (
        str(benchmark).strip()
        if pd.notna(benchmark)
        else "CDI"
    )
    return (
        f"{benchmark_text} {signal} {formatted}%"
    )

def simulate_sql_visao_geral(df):
    """
    Simula a execuÃ§Ã£o da Etapa 3 (CTE VisaoGeral) utilizando o DataFrame da Carteira (mock).
    Em produÃ§Ã£o, isso seria uma chamada pd.read_sql() contra a query montada no banco.
    """
    if df.empty:
        return {}
    
    qtd_ativos = df["codigo_IAM_ativo"].nunique() if "codigo_IAM_ativo" in df else 0
    pl_total = df["perc_pl"].fillna(0).sum() if "perc_pl" in df else 0
    
    # % CrÃ©dito
    if "tipo_ativo" in df and "perc_pl" in df:
        perc_credito = df.loc[df["tipo_ativo"] == "Credito", "perc_pl"].fillna(0).sum()
    else:
        perc_credito = 0
        
    # % Infraestrutura
    if "setor_dashboard" in df and "perc_pl" in df:
        perc_infra = df.loc[df["setor_dashboard"] == "Infraestrutura", "perc_pl"].fillna(0).sum()
    else:
        perc_infra = 0
        
    # Duration CrÃ©ditos
    if "tipo_ativo" in df and "perc_pl" in df and "duration" in df:
        cred_df = df[df["tipo_ativo"] == "Credito"]
        s_peso = cred_df["perc_pl"].fillna(0).sum()
        duration_creditos = (cred_df["perc_pl"].fillna(0) * cred_df["duration"].fillna(0)).sum() / s_peso if s_peso != 0 else 0
    else:
        duration_creditos = 0
        
    # Duration Fundo
    if "perc_pl" in df and "duration" in df:
        s_peso_total = df["perc_pl"].fillna(0).sum()
        duration_fundo = (df["perc_pl"].fillna(0) * df["duration"].fillna(0)).sum() / s_peso_total if s_peso_total != 0 else 0
    else:
        duration_fundo = 0

    return {
        "qtd_ativos": qtd_ativos,
        "pl_total": pl_total,
        "perc_credito": perc_credito,
        "perc_infra": perc_infra,
        "duration_creditos": duration_creditos,
        "duration_fundo": duration_fundo,
        
        # Mantendo compatibilidade com mÃ©tricas que ainda nÃ£o estÃ£o no CTE:
        "benchmark": "CDI",
        "carrego_credito": 0.0,
        "carrego_fundo": 0.0,
        "probabilidade_default": 0.0
    }

def render_raiox_general_view(raiox):
    # SIMULAÃ‡ÃƒO: Executa a "query" (agregaÃ§Ã£o) da VisÃ£o Geral baseada na carteira bruta
    agg = simulate_sql_visao_geral(raiox)
    
    benchmark = agg.get("benchmark", "CDI")
    patrimonio = agg.get("pl_total", 0)
    perc_credito = agg.get("perc_credito", 0)
    perc_infra = agg.get("perc_infra", 0)
    carrego_credito = agg.get("carrego_credito", 0)
    carrego_fundo = agg.get("carrego_fundo", 0)
    duration_credito = agg.get("duration_creditos", 0)
    duration_fundo = agg.get("duration_fundo", 0)
    probabilidade_default = agg.get("probabilidade_default", 0)

    rows = [
        (
            "PL (R$ mn)",
            format_raiox_money(patrimonio),
        ),
        (
            "% CrÃ©dito / % Infra",
            (
                f"{fmt_pct(perc_credito)} / "
                f"{fmt_pct(perc_infra)}"
            ),
        ),
        (
            "Carrego dos CrÃ©ditos",
            format_raiox_spread(
                carrego_credito,
                benchmark,
            ),
        ),
        (
            "Carrego do Fundo",
            format_raiox_spread(
                carrego_fundo,
                benchmark,
            ),
        ),
        (
            "Duration dos CrÃ©ditos",
            format_raiox_duration(
                duration_credito
            ),
        ),
        (
            "Duration do Fundo",
            format_raiox_duration(
                duration_fundo
            ),
        ),
        (
            "Probabilidade de Default (PD)",
            fmt_pct(probabilidade_default, 2),
        )
    ]

    rows_html = "".join(
        (
            '<div class="rx-general-row">'
            f'<div class="rx-general-label">{html.escape(label)}</div>'
            f'<div class="rx-general-value">{html.escape(value)}</div>'
            '</div>'
        )
        for label, value in rows
    )

    st.markdown(
        f"""
        <div class="rx-general-card">
            <div class="rx-general-title">VisÃ£o Geral</div>
            {rows_html}
        </div>
        """,
        unsafe_allow_html=True,
    )

def build_raiox_excel(
    raiox,
    instrument_data,
    group_data,
    indexor_data,
    rating_data,
    sector_data,
    instrument_fig,
    group_fig,
    indexor_fig,
    rating_fig,
    sector_fig,
):
    wb = Workbook()
    ws_base = wb.active
    ws_base.title = "Base_Raiox"

    for row in (
        pd.DataFrame(raiox)
        .fillna("")
        .itertuples(index=False)
    ):
        pass

    ws_base.append(list(raiox.columns))

    for row in (
        raiox.fillna("")
        .values
        .tolist()
    ):
        ws_base.append(row)

    sheets = {
        "Instrumento": (
            instrument_data,
            instrument_fig,
        ),
        "Grupo_Economico": (
            group_data,
            group_fig,
        ),
        "Indexador": (
            indexor_data,
            indexor_fig,
        ),
        "Rating": (
            rating_data,
            rating_fig,
        ),
        "Setor": (
            sector_data,
            sector_fig,
        )
    }

    for sheet_name, (
        data,
        figure,
    ) in sheets.items():
        ws = wb.create_sheet(
            title=sheet_name
        )

        if not data.empty:
            ws.append(
                list(data.columns)
            )

            for row in (
                data.fillna("")
                .values
                .tolist()
            ):
                ws.append(row)

        try:
            image_bytes = figure.to_image(
                format="png",
                width=1400,
                height=800,
                scale=2,
            )

            image_stream = BytesIO(
                image_bytes
            )

            image = XLImage(
                image_stream
            )

            image.width = 980
            image.height = 500

            ws.add_image(
                image,
                "H2",
            )

        except Exception:
            pass

    output = BytesIO()

    wb.save(output)

    output.seek(0)

    return output

def section_title(text):
    st.markdown(
        f'<div class="itau-section-title"><h3>{html.escape(text)}</h3>'
        f'<div class="itau-section-title-accent"></div></div>',
        unsafe_allow_html=True,
    )

def show(fig, section, title, include_html=True):
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})
    report_figures.append((section, title, fig, "include_html" if include_html else "exclude_html"))

def build_raiox_comparison(first_data, second_data, category, first_label, second_label):
    """Une as categorias dos dois fundos e calcula a diferenÃ§a Fundo 1 - Fundo 2."""
    first = first_data[[category, "valor", "percentual"]].rename(
        columns={"valor": f"valor_{first_label}", "percentual": first_label}
    )
    second = second_data[[category, "valor", "percentual"]].rename(
        columns={"valor": f"valor_{second_label}", "percentual": second_label}
    )
    comparison = first.merge(second, on=category, how="outer").fillna(0.0)
    comparison["Diferenca"] = comparison[first_label] - comparison[second_label]
    comparison["RelevÃ¢ncia combinada"] = (
        comparison[first_label].abs() + comparison[second_label].abs()
    )
    return comparison.sort_values(
        "RelevÃ¢ncia combinada", ascending=False, kind="stable"
    ).reset_index(drop=True)

def raiox_comparison_chart(
    first_data,
    second_data,
    category,
    title,
    first_label,
    second_label,
    orientation="h",
    category_order=None,
    top=None,
    height=460,
):
    comparison = build_raiox_comparison(
        first_data, second_data, category, first_label, second_label
    )
    if category_order is not None:
        comparison["_order"] = comparison[category].map(category_order).fillna(99)
        comparison = comparison.sort_values("_order", kind="stable")
    elif top is not None:
        comparison = comparison.head(top)
    if comparison.empty:
        return None, comparison

    fig = go.Figure()
    for label, color in ((first_label, ORANGE), (second_label, DARK)):
        other_label = second_label if label == first_label else first_label
        common = dict(
            name=label,
            marker_color=color,
            text=comparison[label].map(
                lambda value: f"{value:.1%}".replace(".", ",")
            ),
            textposition="outside",
            cliponaxis=False,
            customdata=np.column_stack(
                [comparison[other_label], comparison["Diferenca"]]
            ),
            hovertemplate=(
                f"{label}: %{{x:.2%}}<br>"
                f"{other_label}: %{{customdata[0]:.2%}}<br>"
                "DiferenÃ§a Fundo 1 - Fundo 2: %{customdata[1]:+.2%}"
                "<extra></extra>"
            ),
        )
        if orientation == "h":
            fig.add_trace(
                go.Bar(
                    x=comparison[label],
                    y=comparison[category],
                    orientation="h",
                    **common,
                )
            )
        else:
            fig.add_trace(
                go.Bar(
                    x=comparison[category],
                    y=comparison[label],
                    orientation="v",
                    **common,
                )
            )

    apply_layout(fig, title)
    fig.update_layout(
        barmode="group",
        height=height,
        hovermode="closest",
        margin=dict(
            l=10,
            r=85 if orientation == "h" else 20,
            t=75,
            b=110 if orientation == "v" else 30,
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.03,
            xanchor="left",
            x=0,
        ),
    )
    if orientation == "h":
        fig.update_xaxes(tickformat=".0%", title="")
        fig.update_yaxes(title="", autorange="reversed")
    else:
        fig.update_yaxes(tickformat=".0%", title="")
        fig.update_xaxes(title="", tickangle=-35)
    return fig, comparison

def build_raiox_group_top(data, top_n=10):
    """Retorna TOP N, linha Outros e composiÃ§Ã£o detalhada de Outros."""
    full = build_raiox_distribution(
        data,
        category="grupo_economico",
        value_column="perc_pl",
        denominator=None,
        top_n=None,
        others_position="last",
    )
    full = full.loc[full["percentual"].abs().gt(1e-12)].copy()
    full = full.sort_values(
        "percentual",
        ascending=False,
        kind="stable",
    ).reset_index(drop=True)

    top = full.head(top_n).copy()
    others_detail = full.iloc[top_n:].copy()
    if not others_detail.empty:
        others_row = pd.DataFrame({
            "grupo_economico": [f"Outros ({len(others_detail)} sun)"],
            "valor": [others_detail["valor"].sum()],
            "percentual": [others_detail["percentual"].sum()],
            "ordem": [top_n + 1],
        })
        top = pd.concat([top, others_row], ignore_index=True)

    top["ordem"] = np.arange(1, len(top) + 1)
    return top, others_detail

def raiox_group_bar_chart(
    data,
    others_detail,
    title,
    color,
    height,
    axis_max=None,
):
    """GrÃ¡fico TOP 30 com Outros fixado no rodapÃ© e hover detalhado."""
    if data.empty:
        return None

    chart_data = data.copy()
    regular = chart_data.loc[
        ~chart_data["grupo_economico"].astype(str).str.startswith("Outros")
    ].sort_values("percentual", ascending=False, kind="stable")
    others = chart_data.loc[
        chart_data["grupo_economico"].astype(str).str.startswith("Outros")
    ]
    chart_data = pd.concat([regular, others], ignore_index=True)

    others_lines = "<br>".join(
        f"{html.escape(str(row.grupo_economico))}: {row.percentual:.2%}"
        for row in others_detail.itertuples(index=False)
    )
    hover_text = []
    for row in chart_data.itertuples(index=False):
        if str(row.grupo_economico).startswith("Outros"):
            hover_text.append(
                f"<b>{html.escape(str(row.grupo_economico))}</b><br>"
                f"Percentual total: {row.percentual:.2%}<br>"
                f"<b>Grupos incluÃ­dos:</b><br>{others_lines}"
            )
        else:
            hover_text.append(
                f"<b>{html.escape(str(row.grupo_economico))}</b><br>"
                f"Percentual: {row.percentual:.2%}"
            )

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=chart_data["percentual"],
            y=chart_data["grupo_economico"],
            orientation="h",
            marker=dict(color=color),
            width=0.66,
            text=chart_data["percentual"].map(
                lambda value: f"{value:.1%}".replace(".", ",")
            ),
            textposition="outside",
            cliponaxis=False,
            customdata=np.array(hover_text, dtype=object).reshape(-1, 1),
            hovertemplate="%{customdata[0]}<extra></extra>",
        )
    )

    apply_layout(fig, title)
    fig.update_layout(
        height=height,
        showlegend=False,
        hovermode="closest",
        bargap=0.20,
        margin=dict(l=10, r=75, t=70, b=25),
    )

    fig.update_xaxes(
        range=[0, axis_max * 1.20] if axis_max and axis_max > 0 else None,
        visible=False,
        title="",
        showgrid=False,
        zeroline=False,
    )

    fig.update_yaxes(
        title="",
        showgrid=False,
        autorange="reversed",
        categoryorder="array",
        categoryarray=chart_data["grupo_economico"].tolist(),
        tickfont=dict(size=10),
    )

    return fig

def raiox_group_overlap_chart(
    first_full,
    second_full,
    first_label,
    second_label,
    height,
    top_n=15,
):
    """Ranking vertical conectado dos TOP 15 independentes dos dois fundos."""
    first_top = first_full.head(top_n).copy().reset_index(drop=True)
    second_top = second_full.head(top_n).copy().reset_index(drop=True)
    if first_top.empty and second_top.empty:
        return None, pd.DataFrame()

    first_top["rank_first"] = np.arange(1, len(first_top) + 1)
    second_top["rank_second"] = np.arange(1, len(second_top) + 1)
    first_top["x"] = np.linspace(0.04, 0.96, max(len(first_top), 1))
    second_top["x"] = np.linspace(0.04, 0.96, max(len(second_top), 1))

    first_lookup = first_top.set_index("grupo_economico")
    second_lookup = second_top.set_index("grupo_economico")
    shared = list(set(first_lookup.index) & set(second_lookup.index))

    fig = go.Figure()
    for group in shared:
        first_row = first_lookup.loc[group]
        second_row = second_lookup.loc[group]
        delta = first_row["percentual"] - second_row["percentual"]
        fig.add_trace(
            go.Scatter(
                x=[first_row["x"], second_row["x"]],
                y=[1, 0],
                mode="lines",
                line=dict(color="rgba(255,98,0,0.43)", width=2.2),
                customdata=[
                    [
                        group,
                        first_row["percentual"],
                        second_row["percentual"],
                        delta,
                    ]
                ]
                * 2,
                hovertemplate=(
                    "<b>%{customdata[0]}</b><br>"
                    f"{first_label}: %{{customdata[1]:.2%}}<br>"
                    f"{second_label}: %{{customdata[2]:.2%}}<br>"
                    "Delta: %{customdata[3]:+.2%}<extra></extra>"
                ),
                showlegend=False,
            )
        )

    for source, y_value, label, rank_col, other_lookup, color in [
        (first_top, 1, first_label, "rank_first", second_lookup, ORANGE),
        (second_top, 0, second_label, "rank_second", first_lookup, DARK),
    ]:
        custom = []
        node_colors = []
        for row in source.itertuples(index=False):
            group = row.grupo_economico
            own_value = row.percentual
            if group in other_lookup.index:
                other_value = float(other_lookup.loc[group, "percentual"])
                other_rank_col = "rank_second" if rank_col == "rank_first" else "rank_first"
                other_rank = int(other_lookup.loc[group, other_rank_col])
                overlap = "Compartilhado"
                node_colors.append(color)
            else:
                other_value = 0.0
                other_rank = 0
                overlap = "Exclusivo do TOP 15"
                node_colors.append("#A7A0B5")
            custom.append([group, own_value, int(getattr(row, rank_col)), other_value, other_rank, overlap])

        fig.add_trace(
            go.Scatter(
                x=source["x"],
                y=[y_value] * len(source),
                mode="markers+text",
                marker=dict(size=17, color=node_colors, line=dict(color="#FFFFFF", width=1.2)),
                text=[str(index) for index in source[rank_col]],
                textposition="middle center",
                textfont=dict(color="#FFFFFF", size=9),
                customdata=np.array(custom, dtype=object),
                hovertemplate=(
                    "<b>%{customdata[0]}</b><br>"
                    f"Ranking: %{{customdata[2]}} | PosiÃ§Ã£o %{{customdata[1]:.2%}}<br>"
                    "Outro fundo: %{customdata[3]:.2%} | PosiÃ§Ã£o %{customdata[4]}<br>"
                    "%{customdata[5]}<extra></extra>"
                ),
                showlegend=False,
            )
        )

    for row in first_top.itertuples(index=False):
        fig.add_annotation(
            x=row.x,
            y=1.08,
            text=f"{html.escape(str(row.grupo_economico))}<br><b>({row.percentual:.1%})</b>",
            showarrow=False,
            textangle=-55,
            xanchor="left",
            yanchor="bottom",
            font=dict(size=9, color=DARK),
            align="left",
        )
    for row in second_top.itertuples(index=False):
        fig.add_annotation(
            x=row.x,
            y=-0.16,
            text=f"{html.escape(str(row.grupo_economico))}<br><b>({row.percentual:.1%})</b>",
            showarrow=False,
            textangle=-55,
            xanchor="left",
            yanchor="bottom",
            font=dict(size=9, color=DARK),
            align="left",
        )

    fig.add_annotation(x=0, y=1.23, xanchor="left", showarrow=False, text=f"<b>{html.escape(first_label)}</b>", font=dict(color=ORANGE, size=12))
    fig.add_annotation(x=0, y=-0.23, xanchor="left", showarrow=False, text=f"<b>{html.escape(second_label)}</b>", font=dict(color=DARK, size=12))
    apply_layout(fig, "Overlap TOP 15")
    fig.update_layout(
        height=height,
        hovermode="closest",
        showlegend=False,
        margin=dict(l=10, r=10, t=135, b=135),
        xaxis=dict(range=[-0.03, 1.03], visible=False, fixedrange=True),
        yaxis=dict(range=[-0.30, 1.30], visible=False, fixedrange=True),
    )

    overlap_table = first_top[["grupo_economico", "percentual", "rank_first"]].merge(
        second_top[["grupo_economico", "percentual", "rank_second"]],
        on="grupo_economico",
        suffixes=(f" {first_label}", f" {second_label}"),
    )
    return fig, overlap_table


# ---------------------------------------------------------------------------
# FunÃ§Ãµes auxiliares do Raio-X
# ---------------------------------------------------------------------------

def first_valid_value(data, columns, default=np.nan):
    """Retorna o primeiro valor nÃ£o vazio encontrado nas colunas informadas."""
    if data is None or getattr(data, "empty", True):
        return default
    for column in columns:
        if column not in data.columns:
            continue
        values = data[column].dropna()
        if values.dtype == object or str(values.dtype).startswith("string"):
            values = values[values.astype(str).str.strip().ne("")]
        if not values.empty:
            return values.iloc[0]
    return default


def fmt_pct(value, decimals=1):
    if value is None or pd.isna(value):
        return "-"
    return f"{float(value):.{decimals}%}".replace(".", ",")


def period_label(value, dates=None, frequency="Mensal"):
    del dates
    ts = pd.Timestamp(value)
    if frequency == "Mensal":
        return f"{ts:%m/%Y} ({ts:%d/%m/%Y})"
    return f"{ts:%d/%m/%Y}"


def _raiox_signature(fund_code, second_fund_code, raiox_date, second_raiox_date):
    return (
        str(fund_code),
        str(second_fund_code),
        str(pd.Timestamp(raiox_date).date()) if raiox_date is not None else "",
        str(pd.Timestamp(second_raiox_date).date()) if second_raiox_date is not None else "",
    )


def _set_raiox_page(value):
    st.session_state["raiox_page"] = int(value)


def _render_raiox_navigation(page, pages, key_suffix=""):
    total = len(pages)
    prev_col, info_col, next_col = st.columns([1, 4, 1])
    prev_col.button(
        "â—€ anterior",
        key=f"raiox_prev_{key_suffix}",
        disabled=page <= 0,
        use_container_width=True,
        on_click=_set_raiox_page,
        args=(max(page - 1, 0),),
    )
    info_col.markdown(
        "<div style='text-align:center;padding-top:8px;'>"
        f"Etapa <b>{page + 1}</b> de <b>{total}</b> Â· {html.escape(str(pages[page]))}"
        "</div>",
        unsafe_allow_html=True,
    )
    next_col.button(
        "prÃ³xima â–¶",
        key=f"raiox_next_{key_suffix}",
        disabled=page >= total - 1,
        use_container_width=True,
        on_click=_set_raiox_page,
        args=(min(page + 1, total - 1),),
    )


def raiox_page_distribution(data, category, value_column):
    """DistribuiÃ§Ã£o completa (sem agrupar em Outros), ordenada pelo percentual."""
    distribution = build_raiox_distribution(
        data,
        category=category,
        value_column=value_column,
        denominator=None,
        top_n=None,
        others_position="last",
    )
    if distribution.empty:
        return distribution
    distribution = distribution.loc[distribution["percentual"].abs().gt(1e-12)]
    return distribution.sort_values(
        "percentual", ascending=False, kind="stable"
    ).reset_index(drop=True)


def raiox_instrument_delta_chart(comparison, category, title, first_label, second_label):
    """Barras horizontais com a diferenÃ§a Fundo 1 - Fundo 2 por categoria."""
    if comparison is None or comparison.empty:
        return None
    chart_data = comparison.sort_values("Diferenca", ascending=False, kind="stable")
    colors = np.where(chart_data["Diferenca"] >= 0, ORANGE, DARK)
    fig = go.Figure(
        go.Bar(
            x=chart_data["Diferenca"],
            y=chart_data[category],
            orientation="h",
            marker_color=colors,
            text=chart_data["Diferenca"].map(lambda value: f"{value:+.1%}".replace(".", ",")),
            textposition="outside",
            cliponaxis=False,
            customdata=np.column_stack([chart_data[first_label], chart_data[second_label]]),
            hovertemplate=(
                "%{y}<br>"
                f"{html.escape(first_label)}: %{{customdata[0]:.2%}}<br>"
                f"{html.escape(second_label)}: %{{customdata[1]:.2%}}<br>"
                "DiferenÃ§a: %{x:+.2%}<extra></extra>"
            ),
        )
    )
    apply_layout(fig, title)
    fig.update_layout(
        height=max(380, 34 * len(chart_data) + 120),
        showlegend=False,
        hovermode="closest",
        margin=dict(l=10, r=85, t=70, b=30),
    )
    fig.update_xaxes(tickformat=".0%", title="", zeroline=True, zerolinecolor="#9AA1AC")
    fig.update_yaxes(title="", autorange="reversed")
    return fig


def render_others_detail(others_detail, label):
    if others_detail is None or others_detail.empty:
        return
    with st.expander(f"ComposiÃ§Ã£o de Outros - {label}"):
        table = others_detail[["grupo_economico", "percentual"]].copy()
        table["percentual"] = table["percentual"].map(lambda value: fmt_pct(value, 2))
        st.dataframe(table, use_container_width=True, hide_index=True)


def render_external_raiox_diagnostics():
    diagnostics = st.session_state.get("external_raiox_diagnostics")
    if not diagnostics:
        return
    with st.expander("Diagnóstico da carteira externa (Comdinheiro)"):
        st.dataframe(pd.DataFrame(diagnostics.get("steps", [])), use_container_width=True, hide_index=True)
        for name, table in diagnostics.get("tables", {}).items():
            st.markdown(f"*{name}*")
            st.dataframe(table, use_container_width=True, hide_index=True)



def _show_equivalent_query(title, fund, second_fund, comparison_mode):
    f1 = str(fund).strip().upper()
    f2 = str(second_fund).strip().upper() if second_fund else ""
    
    try:
        with open("consulta_raiox.sql", "r", encoding="utf-8") as f:
            base_sql = f.read()
    except FileNotFoundError:
        base_sql = "-- Arquivo de query não encontrado."
        
    if comparison_mode and f2:
        base_sql = base_sql.replace("'MOCK_FUNDO_PRINCIPAL'", f"'{f1}'").replace("'EXT_00000000000191'", f"'{f2}'")
    else:
        base_sql = base_sql.replace("'MOCK_FUNDO_PRINCIPAL',", f"'{f1}'").replace("'EXT_00000000000191'", "")
        bs_lines = base_sql.splitlines()
        clean_lines = [line for line in bs_lines if line.strip() != f"'{f1}'"]
        base_sql = chr(10).join(clean_lines)

    cte_map = {
        "Visão geral": "VisaoGeral",
        "Visão geral dos dois fundos": "VisaoGeral",
        "Instrumentos": "Instrumentos",
        "Setores": "Setores",
        "Ratings": "Ratings",
        "Indexadores": "Indexadores",
        "Grupos econômicos do fundo principal": "GruposEconomicos",
        "Grupos econômicos do fundo comparado": "GruposEconomicos",
        "Delta do instrumentos": "DeltaInstrumentos",
        "Overlap dos grupos": "OverlapGrupos"
    }
    
    target_cte = cte_map.get(title, "Carteira")
    
    if title != "Exportação e auditoria":
        final_sql = base_sql.split("-- O Python fará o SELECT final")[0]
        final_sql += chr(10) + "-- ============================================================" + chr(10) + "-- ETAPA FINAL: SELECT DA ABA ATUAL" + chr(10) + "-- ============================================================" + chr(10) + "SELECT * FROM " + target_cte + ";"
    else:
        final_sql = base_sql.split("-- O Python fará o SELECT final")[0]
        final_sql += chr(10) + "-- ============================================================" + chr(10) + "-- ETAPA FINAL: EXPORTAÇÃO" + chr(10) + "-- ============================================================" + chr(10) + "SELECT * FROM Carteira;"
        
    with st.expander(f"📋 Copiar Query SQL Equivalente - {title}"):
        st.caption("Esta consulta demonstra o código SQL exato que o Backend rodaria para trazer os dados apenas desta etapa.")
        st.code(final_sql, language="sql")


def render_credit_raiox():
    section_title("raio-x por etapas")
    username = st.session_state.get("sql_username", "")
    password = st.session_state.get("sql_password", "")
    if not username or not password:
        st.info("Informe usuÃ¡rio e senha SQL na barra lateral.")
        return

    comparison_mode = fund_comparison_mode and second_fund is not None
    if comparison_mode:
        first_col, second_col = st.columns(2)
        raiox_date = first_col.selectbox(
            "data-base do fundo principal", dates_desc,
            format_func=lambda value: period_label(value, dates, frequency),
            key="raiox_reference_date",
        )
        if is_external_comparison_fund(second_fund):
            external_start = pd.Timestamp(start_date).to_period("M").to_timestamp(how="end").normalize()
            external_limit = min(pd.Timestamp(end_date), pd.Timestamp("2026-06-01"))
            external_end = pd.Timestamp(external_limit).to_period("M").to_timestamp(how="end").normalize()
            external_dates = list(pd.date_range(external_start, external_end, freq="ME"))[::-1]
            if not external_dates:
                external_dates = [external_end]
            second_raiox_date = second_col.selectbox(
                "data-base do fundo comparado", external_dates,
                format_func=lambda value: pd.Timestamp(value).strftime("%d/%m/%Y"),
                key="raiox_second_reference_date_external",
            )
        else:
            second_raiox_date = second_col.selectbox(
                "data-base do fundo comparado", dates_desc,
                index=dates_desc.index(raiox_date),
                format_func=lambda value: period_label(value, dates, frequency),
                key="raiox_second_reference_date",
            )
    else:
        raiox_date = st.selectbox(
            "data-base de raio-x", dates_desc,
            format_func=lambda value: period_label(value, dates, frequency),
            key="raiox_reference_date",
        )
        second_raiox_date = None

    signature = _raiox_signature(fund, second_fund, raiox_date, second_raiox_date)
    run_requested = st.session_state.pop("raiox_run_requested", False)
    if run_requested:
        try:
            with st.spinner("Carregando somente as tabelas-base das carteiras..."):
                primary = load_credit_raiox(
                    sql_path=str(RAIOX_SQL_PATH),
                    sql_modified_time=RAIOX_SQL_PATH.stat().st_mtime_ns,
                    server="SQNPRC028", driver="ODBC Driver 17 for SQL Server",
                    username=username, password=password, fund_code=fund,
                    reference_date=raiox_date,
                )
                primary_pd = load_raiox_default_probability(
                    server="SQNPRC028", driver="ODBC Driver 17 for SQL Server",
                    username=username, password=password, fund_code=fund,
                    reference_date=raiox_date,
                )
                if not primary.empty:
                    primary["probabilidade_default"] = primary_pd

                secondary = None
                if comparison_mode:
                    if is_external_comparison_fund(second_fund):
                        secondary, secondary_pd = load_external_comparison_raiox(
                            second_fund, second_raiox_date, username, password
                        )
                    else:
                        secondary = load_credit_raiox(
                            sql_path=str(RAIOX_SQL_PATH),
                            sql_modified_time=RAIOX_SQL_PATH.stat().st_mtime_ns,
                            server="SQNPRC028", driver="ODBC Driver 17 for SQL Server",
                            username=username, password=password, fund_code=second_fund,
                            reference_date=second_raiox_date,
                        )
                        secondary_pd = load_raiox_default_probability(
                            server="SQNPRC028", driver="ODBC Driver 17 for SQL Server",
                            username=username, password=password,
                            fund_code=second_fund, reference_date=second_raiox_date,
                        )
                    if secondary is not None and not secondary.empty:
                        secondary["probabilidade_default"] = secondary_pd

                st.session_state["raiox_base_primary"] = primary
                st.session_state["raiox_base_secondary"] = secondary
                st.session_state["raiox_loaded_signature"] = signature
                st.session_state["raiox_page"] = 0
                st.session_state.pop("raiox_pd_history", None)
        except Exception as exc:
            st.error(f"Falha ao carregar as carteiras-base: {exc}")
            return

    primary = st.session_state.get("raiox_base_primary")
    secondary = st.session_state.get("raiox_base_secondary")
    loaded_signature = st.session_state.get("raiox_loaded_signature")
    if primary is None:
        st.info("Configure os parÃ¢metros e clique em *Rodar Raio-X* na barra lateral.")
        return
    if loaded_signature != signature:
        st.warning("Os parÃ¢metros foram alterados. Clique em *Rodar Raio-X* para atualizar as bases.")
        return
    if primary.empty:
        st.warning("A carteira principal nÃ£o retornou dados.")
        return

    pages = [
        "VisÃ£o geral dos dois fundos", "Instrumentos", "Delta do instrumentos",
        "Grupos econÃ´micos do fundo principal", "Grupos econÃ´micos do fundo comparado",
        "Overlap dos grupos", "Indexadores", "Ratings", "Setores",
        "EvoluÃ§Ã£o da PD", "ExportaÃ§Ã£o e auditoria",
    ] if comparison_mode else [
        "VisÃ£o geral", "Instrumentos", "Grupos econÃ´micos do fundo principal", "Indexadores",
        "Ratings", "Setores", "EvoluÃ§Ã£o da PD", "ExportaÃ§Ã£o e auditoria",
    ]
    current_max_page = min(int(st.session_state.get("raiox_page", 0)), len(pages) - 1)
    first_label = get_fund_name(fund)
    second_label = comparison_fund_name(second_fund) if comparison_mode else ""

    for current_step in range(current_max_page + 1):
        title = pages[current_step]
        st.subheader(title)
        _show_equivalent_query(title, fund, second_fund, comparison_mode)

        if title in ("VisÃ£o geral", "VisÃ£o geral dos dois fundos"):
            if comparison_mode:
                left, right = st.columns(2)
                with left:
                    st.markdown(f"*{first_label}*")
                    render_raiox_general_view(primary)
                with right:
                    st.markdown(f"*{second_label}*")
                    render_raiox_general_view(secondary)
            else:
                render_raiox_general_view(primary)

        elif title == "Instrumentos":
            first_data = raiox_page_distribution(primary, "instrumento", "perc_nav")
            if comparison_mode:
                second_data = raiox_page_distribution(secondary, "instrumento", "perc_nav")
                comparison = build_raiox_comparison(
                    first_data, second_data, "instrumento", first_label, second_label
                )
                fig, _ = raiox_comparison_chart(
                    first_data, second_data, "instrumento", "Instrumentos",
                    first_label, second_label
                )
            else:
                fig = raiox_horizontal_bar(first_data, "instrumento", "Instrumentos")
            if fig is not None:
                st.plotly_chart(fig, use_container_width=True)

        elif title == "Delta do instrumentos":
            first_data = raiox_page_distribution(primary, "instrumento", "perc_nav")
            second_data = raiox_page_distribution(secondary, "instrumento", "perc_nav")
            comparison = build_raiox_comparison(
                first_data, second_data, "instrumento", first_label, second_label
            )
            fig = raiox_instrument_delta_chart(
                comparison, "instrumento", "Delta do instrumentos",
                first_label, second_label,
            )
            if fig is not None:
                st.plotly_chart(fig, use_container_width=True)

        elif title == "Grupos econÃ´micos do fundo principal":
            data, others = build_raiox_group_top(primary, top_n=30)
            fig = raiox_group_bar_chart(data, others, first_label, ORANGE, 760)
            if fig is not None:
                st.plotly_chart(fig, use_container_width=True)
                render_others_detail(others, first_label)

        elif title == "Grupos econÃ´micos do fundo comparado":
            data, others = build_raiox_group_top(secondary, top_n=30)
            fig = raiox_group_bar_chart(data, others, second_label, DARK, 760)
            if fig is not None:
                st.plotly_chart(fig, use_container_width=True)
                render_others_detail(others, second_label)

        elif title == "Overlap dos grupos":
            first_data = raiox_page_distribution(primary, "grupo_economico", "perc_pl")
            second_data = raiox_page_distribution(secondary, "grupo_economico", "perc_pl")
            fig, overlap = raiox_group_overlap_chart(
                first_data, second_data, first_label, second_label, 728
            )
            if fig is not None:
                st.plotly_chart(fig, use_container_width=True)
                st.dataframe(overlap, use_container_width=True, hide_index=True)

        elif title in ("Indexadores", "Ratings", "Setores"):
            category = {"Indexadores": "indexador", "Ratings": "rating_externo", "Setores": "setor"}[title]
            first_data = raiox_page_distribution(primary, category, "perc_pl")
            if comparison_mode:
                second_data = raiox_page_distribution(secondary, category, "perc_pl")
                fig, _ = raiox_comparison_chart(
                    first_data, second_data, category, title,
                    first_label, second_label
                )
            else:
                fig = raiox_pie_chart(first_data, category, title)
            if fig is not None:
                st.plotly_chart(fig, use_container_width=True)

        elif title == "EvoluÃ§Ã£o da PD":
            history_key = (fund, second_fund, str(start_date), str(end_date))
            cached = st.session_state.get("raiox_pd_history")
            if cached is None or cached["key"] != history_key:
                with st.spinner("Consultando histÃ³rico de PD somente nesta etapa..."):
                    frames = []
                    first_history = load_raiox_default_probability_history(
                        server="SQNPRC028", driver="ODBC Driver 17 for SQL Server",
                        username=username, password=password, fund_code=fund,
                        start_date=start_date, end_date=end_date,
                    )
                    if not first_history.empty:
                        first_history["Fundo"] = first_label
                        frames.append(first_history)
                    if comparison_mode and not is_external_comparison_fund(second_fund):
                        second_history = load_raiox_default_probability_history(
                            server="SQNPRC028", driver="ODBC Driver 17 for SQL Server",
                            username=username, password=password, fund_code=second_fund,
                            start_date=start_date, end_date=end_date,
                        )
                        if not second_history.empty:
                            second_history["Fundo"] = second_label
                            frames.append(second_history)
                    history = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
                    st.session_state["raiox_pd_history"] = {"key": history_key, "data": history}
            history = st.session_state["raiox_pd_history"]["data"]
            fig = raiox_default_probability_history_chart(history, "EvoluÃ§Ã£o da Probabilidade de Default")
            if fig is not None:
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("NÃ£o hÃ¡ histÃ³rico de PD disponÃ­vel para esta seleÃ§Ã£o.")

        elif title == "ExportaÃ§Ã£o e auditoria":
            st.caption("A exportaÃ§Ã£o Ã© construÃ­da somente ao abrir esta etapa.")
            st.markdown("#### Carteira principal")
            st.dataframe(primary, use_container_width=True, hide_index=True)
            st.download_button(
                "Baixar carteira principal em CSV",
                primary.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"raiox_{fund}.csv", mime="text/csv",
            )
            if comparison_mode and secondary is not None:
                st.markdown("#### Carteira comparada")
                st.dataframe(secondary, use_container_width=True, hide_index=True)
                st.download_button(
                    "Baixar carteira comparada em CSV",
                    secondary.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"raiox_{second_fund}.csv", mime="text/csv",
                )
                if is_external_comparison_fund(second_fund):
                    render_external_raiox_diagnostics()

            st.divider()
    if current_max_page < len(pages) - 1:
        if st.button(f"AvanÃ§ar para: {pages[current_max_page + 1]}", type="primary", use_container_width=True):
            st.session_state["raiox_page"] = current_max_page + 1
            st.rerun()

def normalize_cnpj(cnpj):
    normalized = re.sub(r"\D", "", str(cnpj or ""))
    if len(normalized) != 14:
        raise ValueError("O CNPJ deve possuir exatamente 14 dÃ­gitos.")
    return normalized

def build_comdinheiro_query_path(cnpj, reference_date):
    return (
        "CarteiraFundo006-"
        f"{normalize_cnpj(cnpj)}-"
        f"{pd.Timestamp(reference_date).strftime('%Y%m%d')}-"
        "acrescentar_cnpj_fundo_isin-1-default-5-2-"
        "desc_ativo-final_valor+percent"
    )

def parse_brazilian_number(series):
    normalized = (
        series.astype("string")
        .str.strip()
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .replace({"": pd.NA, "None": pd.NA, "nan": pd.NA})
    )
    return pd.to_numeric(normalized, errors="coerce")

def parse_comdinheiro_portfolio(payload, cnpj, reference_date):
    if not isinstance(payload, dict):
        raise TypeError("A resposta da API nÃ£o possui formato de objeto JSON.")
    tab2 = payload.get("tables", {}).get("tab2")
    if not isinstance(tab2, dict):
        raise KeyError("A resposta nÃ£o contÃ©m tables.tab2.")
    result = pd.DataFrame.from_dict(
        tab2,
        orient="index"
    )

    if "lin0" not in result.index:
        raise KeyError("A resposta nÃ£o contÃ©m o cabeÃ§alho tables.tab2.lin0.")
    result.columns = result.loc["lin0"].astype(str).str.strip().tolist()
    result = result.drop(index="lin0").reset_index(drop=True)

    if result.empty:
        raise ValueError("A carteira retornada pelo Comdinheiro estÃ¡ vazia.")
    missing_api = {"ISIN", "X2", "Indice"} - set(result.columns)
    if missing_api:
        raise KeyError(f"A carteira retornada nÃ£o contÃ©m: {', '.join(sorted(missing_api))}")
    result = result.rename(columns={
        "Indice": "indice",
        "Ativo": "ativo_comdinheiro",
        "DescriÃ§Ã£o": "descricao_comdinheiro",
        "Compras (R$)": "compras_valor",
        "Vendas (R$)": "vendas_valor",
        "Final (R$)": "posicao_final",
        "X1": "percentual_original",
        "X2": "percentual_2_original",
        "cnpj_fundo": "cnpj_fundo_carteira",
        "ISIN": "isin",
    })

    for column in ["compras_valor", "vendas_valor", "posicao_final"]:
        if column in result.columns:
            result[column] = parse_brazilian_number(result[column])

    result["percentual_2_original"] = pd.to_numeric(
        result["percentual_2_original"]
        .astype("string")
        .str.strip()
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False),
        errors="coerce"
    )
    result["perc_pl"] = result["percentual_2_original"] / 100.0
    result["isin"] = (
        result["isin"].astype("string").str.strip().str.upper()
        .replace({"": pd.NA, "None": pd.NA, "NAN": pd.NA})
    )
    result["cnpj_consultado"] = normalize_cnpj(cnpj)
    result["data_carteira"] = pd.Timestamp(reference_date).normalize()
    result["capturado_em"] = pd.Timestamp.now()
    result["indice_normalizado"] = (
        result["indice"].astype("string").str.strip().str.strip("*")
    )
    indices = result["indice_normalizado"].dropna().tolist()
    parent_indices = {
        parent for parent in indices
        if any(child != parent and child.startswith(parent + ".") for child in indices)
    }
    result["possui_filho"] = result["indice_normalizado"].isin(parent_indices)
    result["nivel_hierarquia"] = result["indice_normalizado"].str.count(r"\.").fillna(0).astype(int)
    result["linha_falha"] = ~result["possui_filho"]

    # Ajusta o peso dos nÃ­veis explodidos. A API informa o percentual de um
    # filho em relacao ao seu pai, e nÃ£o diretamente em relacao ao PL do fundo.
    # Assim, o peso efetivo de cada linha Ã© o produto dos percentuais ao longo
    # de toda a cadeia hierÃ¡rquica.
    pesos_por_indice = {}
    ordem_hierarquia = result.sort_values(
        ["nivel_hierarquia"],
        kind="stable",
    ).index

    for idx in ordem_hierarquia:
        indice = result.at[idx, "indice_normalizado"]
        percentual = result.at[idx, "perc_pl"]
        if pd.isna(percentual) or pd.isna(indice):
            peso_efetivo = np.nan
        else:
            indice_str = str(indice)
            if "." not in indice_str:
                peso_efetivo = float(percentual)
            else:
                indice_pai = indice_str.rsplit(".", 1)[0]
                peso_pai = pesos_por_indice.get(indice_pai, float(percentual))
                peso_efetivo = float(peso_pai) * float(percentual)

        pesos_por_indice[str(indice)] = peso_efetivo
        result.at[idx, "perc_pl"] = peso_efetivo

    return result

def load_isin_credit_information(
    server,
    driver,
    username,
    password,
    isins,
    reference_date,
):
    """ObtÃ©m a observaÃ§Ã£o interna mais recente, atÃ© a data-base, para cada ISIN."""
    normalized_isins = sorted({
        str(value).strip().upper()
        for value in isins
        if pd.notna(value) and str(value).strip()
    })
    if not normalized_isins:
        return pd.DataFrame(columns=[
            "data_informacao_sql", "isin", "codigo_IAM_ativo", "rating_itau",
            "ativo_sql", "tipo", "PD", "cluster", "quantidade_registros_isin",
        ])

    placeholders = ", ".join("?" for _ in normalized_isins)
    query = f"""
    WITH ISIN_BASE AS
    (
        SELECT
            P.[date] AS data_informacao_sql,
            UPPER(LTRIM(RTRIM(P.isin))) AS isin,
            P.codigo_IAM_ativo,
            P.rating_itau,
            P.ativo AS ativo_sql,
            P.tipo,
            R.PD,
            R.[cluster],
            COUNT(*) OVER (
                PARTITION BY UPPER(LTRIM(RTRIM(P.isin)))
            ) AS quantidade_registros_isin,
            ROW_NUMBER() OVER (
                PARTITION BY UPPER(LTRIM(RTRIM(P.isin)))
                ORDER BY
                    P.[date] DESC,
                    CASE WHEN R.PD IS NOT NULL THEN 0 ELSE 1 END,
                    CASE WHEN NULLIF(LTRIM(RTRIM(P.rating_itau)), '') IS NOT NULL THEN 0 ELSE 1 END,
                    P.codigo_IAM_ativo
            ) AS rn
        FROM VIEW_FUNDOS_POSICAO AS P
        LEFT JOIN dbsx362.dbo.CLUSTER_RATING AS R
          ON LTRIM(RTRIM(P.rating_itau)) = LTRIM(RTRIM(R.rating_iam))
        WHERE P.[date] <= ?
          AND UPPER(LTRIM(RTRIM(P.isin))) IN ({placeholders})
          AND NULLIF(LTRIM(RTRIM(P.isin)), '') IS NOT NULL
    )
    SELECT
        data_informacao_sql,
        isin,
        codigo_IAM_ativo,
        rating_itau,
        ativo_sql,
        tipo,
        PD,
        cluster,
        quantidade_registros_isin
    FROM ISIN_BASE
    WHERE rn = 1
    """
    params = [pd.Timestamp(reference_date).date(), *normalized_isins]
    engine = conn(username=username, password=password)
    try:
        result = pd.read_sql(query, engine, params=tuple(params))
    finally:
        engine.dispose()

    if result.empty:
        return result
    result.columns = [str(column).strip() for column in result.columns]
    result["isin"] = result["isin"].astype("string").str.strip().str.upper()
    result["quantidade_registros_isin"] = pd.to_numeric(
        result["quantidade_registros_isin"], errors="coerce"
    ).fillna(0).astype(int)
    return result

def enrich_comdinheiro_portfolio(portfolio, sql_data):
    result = portfolio.copy()
    if sql_data is None or sql_data.empty:
        for column in [
            "data_informacao_sql", "codigo_IAM_ativo", "rating_itau", "ativo_sql",
            "tipo", "PD", "cluster", "quantidade_registros_isin"
        ]:
            result[column] = pd.NA
    else:
        result = result.merge(sql_data, on="isin", how="left", validate="m:1")

    result["PD"] = pd.to_numeric(result.get("PD"), errors="coerce")
    result["perc_pl"] = pd.to_numeric(result["perc_pl"], errors="coerce")
    result["contribuicao_PD"] = result["perc_pl"] * result["PD"]

    result["status_associacao"] = np.select(
        [
            result["isin"].isna(),
            result["codigo_IAM_ativo"].isna(),
            result["rating_itau"].astype("string").str.strip().isin(["", "N A"]),
            result["PD"].isna(),
        ],
        ["SEM_ISIN", "ISIN_NAO_ENCONTRADO", "SEM_RATING", "SEM_PD"],
        default="ASSOCIADO",
    )
    return result

def conn(database="DBSX337", username="", password=""):
    """Cria o engine SQL Server no padrÃ£o Ãºnico do dashboard."""
    if not username or not password:
        raise ValueError("Informe usuÃ¡rio e senha para autenticaÃ§Ã£o SQL.")

    params = urllib.parse.quote_plus(
        "Driver={ODBC Driver 17 for SQL Server};"
        "Server=SQNPRC028;"
        f"Database={database};"
        f"Uid={username};"
        f"Pwd={password};"
    )

    engine = sa.create_engine(
        f"mssql+pyodbc:///?odbc_connect={params}",
        fast_executemany=True,
    )
    return engine

def main():
    global fund, second_fund, fund_comparison_mode, start_date, end_date
    global dates, dates_desc, frequency, report_figures, report_tables
    global report_overview_blocks, comparison_df, df, raw, funds, fund_name

    apply_global_styles()
    st.title("Raio-X de Carteiras")
    st.caption("AnÃ¡lise mensal da composiÃ§Ã£o de crÃ©dito, com comparaÃ§Ã£o entre fundos internos e fundos de mercado.")

    usar_mock = False
    if not BASE_PATH.exists() or not RAIOX_SQL_PATH.exists() or not EXTERNAL_RAIOX_SQL_PATH.exists():
        st.warning("âš ï¸ Modo de design ativo: Arquivos faltantes detectados. O dashboard estÃ¡ rodando com dados **MOCK** fictÃ­cios para facilitar os ajustes de design.")
        usar_mock = True

    if usar_mock:
        # Cria arquivos SQL vazios temporÃ¡rios se nÃ£o existirem para evitar FileNotFoundError no read_text()
        for path in [RAIOX_SQL_PATH, EXTERNAL_RAIOX_SQL_PATH]:
            if not path.exists():
                path.write_text("-- mock sql")

        # Mock de dados brutos (base.csv) simulando estrutura de crÃ©dito
        mock_dates = pd.date_range("2023-01-01", "2023-12-31", freq="MS")
        mock_data = []
        for d in mock_dates:
            for f in ["MOCK_FUNDO_PRINCIPAL", "MOCK_FUNDO_SECUNDARIO"]:
                for i in range(15):
                    mock_data.append({
                        "date": d,
                        "codigo_IAM_fundo": f,
                        "codigo_IAM_ativo": f"MOCK_ATIVO_{i}",
                        "perc_pl": np.random.uniform(0.01, 0.08),
                        "tipo_ativo": np.random.choice(["DebÃªntures", "CDB", "LF", "FIDC"]),
                        "setor_dashboard": np.random.choice(["Energia", "Saneamento", "Varejo", "Bancos", "Infraestrutura"]),
                        "rating_itau": np.random.choice(["AAA", "AA", "A", "BBB"]),
                        "PD": np.random.uniform(0.01, 0.10),
                        "ativo_sql": "S",
                        "tipo": "Credito Privado",
                        "cluster": np.random.choice(["High Grade", "High Yield"]),
                        "indexador": np.random.choice(["CDI", "IPCA", "PRE"]),
                        "duration": np.random.uniform(1, 5),
                        "emissor": f"Emissor Mock {i % 5}",
                        "cnpj_emissor": f"00.000.000/0001-0{i % 5}",
                        "quantidade_registros_isin": 1,
                        "isin": f"BRMOCKBND00{i}"
                    })
        raw = pd.DataFrame(mock_data)
        
        # Mock de fundos externos
        EXTERNAL_COMPARISON_FUNDS.clear()
        EXTERNAL_COMPARISON_FUNDS.update({
            "EXT_00000000000191": {"category": "Mock Categoria", "cnpj": "00.000.000/0001-91", "name": "MOCK FUNDO CONCORRENTE"}
        })

        # Previne erros de conexÃ£o com banco de dados
        def mock_conn(*args, **kwargs):
            class DummyEngine:
                def dispose(self): pass
            return DummyEngine()
        globals()['conn'] = mock_conn

        # Previne falha no pd.read_sql
        def mock_read_sql(*args, **kwargs):
            return raw.copy()
        pd.read_sql = mock_read_sql

        # Previne falha na API da ComDinheiro (evita requests.post)
        def mock_load_external_comparison_raiox(*args, **kwargs):
            return raw.copy(), 0.05
        globals()['load_external_comparison_raiox'] = mock_load_external_comparison_raiox

    else:
        raw = load_csv(str(BASE_PATH), BASE_PATH.stat().st_mtime_ns)
        try:
            refresh_external_comparison_funds()
        except Exception as exc:
            st.error(f"Falha ao carregar base_mercado.xlsx: {exc}")
            st.stop()

    funds = sorted(
        raw["codigo_IAM_fundo"].dropna().astype(str).str.strip().str.upper().unique().tolist()
    )
    if not funds and not usar_mock:
        st.error("Nenhum fundo vÃ¡lido foi encontrado no base.csv.")
        st.stop()

    with st.sidebar:
        st.markdown("*Raio-X*")
        st.markdown("*Credenciais SQL*")
        st.text_input("usuÃ¡rio SQL", key="sql_username")
        st.text_input("senha SQL", type="password", key="sql_password")
        st.divider()
        fund = st.selectbox(
            "fundo principal", funds,
            format_func=lambda code: f"{get_fund_name(code)} | {code}",
        )
        fund_name = get_fund_name(fund)
        fund_comparison_mode = st.toggle("comparar dois fundos", value=True)
        second_fund = None
        if fund_comparison_mode:
            options = [code for code in funds if code != fund] + list(EXTERNAL_COMPARISON_FUNDS)
            second_fund = st.selectbox(
                "fundo comparado", options,
                format_func=comparison_fund_label,
                key="sidebar_second_fund",
            )

        available = pd.to_datetime(
            raw.loc[
                raw["codigo_IAM_fundo"].astype(str).str.strip().str.upper().eq(fund),
                "date",
            ],
            errors="coerce",
        ).dropna()
        # Sempre o primeiro dia do mÃªs para cada mÃªs com dado disponÃ­vel
        monthly_dates = sorted(list({
            pd.Timestamp(value).replace(day=1)
            for value in available
        }))
        if not monthly_dates:
            st.error("O fundo principal nÃ£o possui datas mensais vÃ¡lidas.")
            st.stop()
        start_date = st.selectbox(
            "mÃªs inicial", monthly_dates,
            format_func=lambda value: pd.Timestamp(value).strftime("%m/%Y"),
        )
        end_options = [value for value in monthly_dates if value >= start_date]
        end_date = st.selectbox(
            "mÃªs final", end_options,
            index=len(end_options) - 1,
            format_func=lambda value: pd.Timestamp(value).strftime("%m/%Y"),
        )

        with st.sidebar:
            st.divider()
            if st.button(
                "Rodar Raio-X",
                type="primary",
                use_container_width=True,
                key="run_raiox_sidebar",
            ):
                st.session_state["raiox_run_requested"] = True
                st.session_state["raiox_page"] = 0

    frequency = "Mensal"
    dates = [pd.Timestamp(value) for value in monthly_dates if pd.Timestamp(start_date) <= pd.Timestamp(value) <= pd.Timestamp(end_date)]
    dates_desc = dates[::-1]
    df = pd.DataFrame()
    comparison_df = pd.DataFrame()
    report_figures = []
    report_tables = []
    report_overview_blocks = []
    render_credit_raiox()

if __name__ == "__main__":
    main()
