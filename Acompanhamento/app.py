"""
Relatório de Acompanhamento — Streamlit 100% Nativo + Plotly
============================================================
Interface construída 100% com componentes nativos do Streamlit
(st.header, st.metric, st.dataframe com Pandas Styler, st.columns, st.divider),
mantendo a identidade visual e os gráficos interativos Plotly.
"""

import sys, os, warnings
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ── Garantir que o diretório do script esteja no path ──
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)
os.chdir(_SCRIPT_DIR)

from acompanhamento import (
    ReadData, ReadExcel, MotorDados,
    CORES_PRATELEIRA, _CORES_CANAL, _ORDEM_CANAL,
    _CORES_PREV, _ORDEM_PREV,
    _eh_rf, _classificar_canal, _classificar_prev,
    _obter_pl_dados, fmt_num, fmt_pct,
    calcular_metricas_fics, calcular_capacity_bloco,
    _calcular_series_receita_total,
    _calcular_series_pfee_taxa_efetiva,
    get_cell_colors,
)

warnings.filterwarnings("ignore")

# ==========================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Relatório de Acompanhamento — Itaú Asset",
    page_icon="📊",
    layout = "centered",
    initial_sidebar_state="collapsed",
)

# ── Customização de tema CSS para componentes nativos do Streamlit ──
st.markdown("""
<style>
    /* Esconde elementos padrão do Streamlit */
    section[data-testid="stSidebar"] { display: none; }
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    .block-container { padding-top: 1rem; padding-bottom: 2rem; max-width: 1200px; }

    /* Estilo dos headers nativos */
    h1, h2, h3 { font-family: Arial, Helvetica, sans-serif !important; }
    
    /* Estilização para métricas nativas do Streamlit */
    div[data-testid="stMetric"] {
        background-color: #f8f9fa;
        border: 1px solid #dee2e6;
        padding: 10px 14px;
        border-radius: 6px;
        text-align: center;
    }
    div[data-testid="stMetricLabel"] { font-size: 12px !important; color: #495057 !important; font-weight: bold; }
    div[data-testid="stMetricValue"] { font-size: 20px !important; color: #212529 !important; font-weight: bold; }

    /* Customização do st.dataframe / st.table */
    .stDataFrame { font-family: Arial, sans-serif !important; }
</style>
""", unsafe_allow_html=True)


# ==========================================
# CONSTANTES
# ==========================================
MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun',
            7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}


# ==========================================
# PLOTLY — LAYOUT PADRÃO
# ==========================================
def _layout(fig, title="", ylabel="", show_legend=True, height=350, pct_y=False, pct_fmt=".0%"):
    fig.update_layout(
        title=dict(text=f"<b>{title}</b>", font=dict(size=12, color="#111", family="Arial"), x=0.5, xanchor="center"),
        yaxis_title=ylabel,
        plot_bgcolor="#ffffff",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Arial", size=10, color="#333"),
        margin=dict(l=45, r=15, t=35, b=35),
        height=height,
        showlegend=show_legend,
        legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5, font=dict(size=9)),
        xaxis=dict(showgrid=False, zeroline=False, showline=True, linecolor="#cccccc", tickformat="%b-%y", tickfont=dict(size=8)),
        yaxis=dict(showgrid=True, gridcolor="#e0e0e0", zeroline=False, showline=False, tickfont=dict(size=9)),
        hovermode="x unified",
    )
    if pct_y:
        fig.update_yaxes(tickformat=pct_fmt)
    return fig


# ==========================================
# FUNÇÕES DE GRÁFICOS PLOTLY
# ==========================================

def plotly_retorno(motor, cnpj, nome_fundo):
    if cnpj not in motor.df_cotas.columns: return go.Figure()
    cotas = motor.df_cotas[cnpj].dropna()
    if cotas.empty or len(cotas) < 2: return go.Figure()
    ret_acum = (cotas / cotas.iloc[0]) - 1
    dias = np.arange(len(ret_acum))
    cdi_acum = (1 + (1.13 ** (1/252) - 1)) ** dias - 1
    cdi_series = pd.Series(cdi_acum, index=ret_acum.index)

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=ret_acum.index, y=ret_acum.values, mode='lines',
                             name=nome_fundo, line=dict(color='#932c3d', width=2.5)))
    fig.add_trace(go.Scatter(x=cdi_series.index, y=cdi_series.values, mode='lines',
                             name='CDI', line=dict(color='#505568', width=2.5)))
    _layout(fig, title=f"Retorno — {nome_fundo}", pct_y=True)
    return fig


def plotly_sharpe_rolling(motor, dicionario_prateleira, window=252):
    cnpjs = [c for c in dicionario_prateleira.keys() if c in motor.df_cotas.columns]
    df_cotas = motor.df_cotas[cnpjs].loc[:motor.data_ref]
    ret_diario = df_cotas.pct_change()
    vol_rolling = ret_diario.rolling(window=window).std(ddof=1) * np.sqrt(252)
    ret_rolling = (df_cotas / df_cotas.shift(window)) - 1
    ret_anualizado = (1 + ret_rolling) ** (252 / window) - 1
    sharpe_rolling = (ret_anualizado / vol_rolling).dropna(how='all')

    fig = go.Figure()
    for i, c in enumerate(cnpjs):
        s = sharpe_rolling[c].dropna()
        if s.empty: continue
        fig.add_trace(go.Scatter(x=s.index, y=s.values, mode='lines',
                                 name=dicionario_prateleira[c],
                                 line=dict(color=CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)], width=1.5)))
    _layout(fig, title="Evolução do Sharpe Líquido (12m)")
    fig.update_yaxes(tickformat=".1f")
    return fig


def plotly_vol_rolling(motor, dicionario_prateleira, window=63, titulo_janela="3m"):
    cnpjs = [c for c in dicionario_prateleira.keys() if c in motor.df_cotas.columns]
    df_cotas = motor.df_cotas[cnpjs].loc[:motor.data_ref]
    ret_diario = df_cotas.pct_change()
    vol_rolling = (ret_diario.rolling(window=window).std(ddof=1) * np.sqrt(252)).dropna(how='all')

    fig = go.Figure()
    for i, c in enumerate(cnpjs):
        s = vol_rolling[c].dropna()
        if s.empty: continue
        fig.add_trace(go.Scatter(x=s.index, y=s.values, mode='lines',
                                 name=dicionario_prateleira[c],
                                 line=dict(color=CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)], width=1.5)))
    _layout(fig, title=f"Evolução da Volatilidade ({titulo_janela})", pct_y=True, pct_fmt=".1%")
    return fig


def plotly_drawdown(motor, cnpj, nome_fundo):
    if cnpj not in motor.df_cotas.columns: return go.Figure()
    cotas = motor.df_cotas[cnpj].dropna()
    if cotas.empty or len(cotas) < 2: return go.Figure()
    picos = cotas.expanding(min_periods=1).max()
    dd = (cotas / picos) - 1
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=dd.index, y=dd.values, mode='lines', fill='tozeroy',
                             fillcolor='rgba(181,37,31,0.35)', line=dict(color='#b5251f', width=0.8), name='Drawdown'))
    _layout(fig, title=f"Drawdown — {nome_fundo}", show_legend=False, height=300, pct_y=True, pct_fmt=".2%")
    return fig


def plotly_representatividade(motor, dicionario_fics):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    pl_fics = motor.df_pl[cnpjs].rename(columns=dicionario_fics)
    rep = pl_fics.div(pl_fics.sum(axis=1), axis=0)
    fig = go.Figure()
    for i, col in enumerate(rep.columns):
        fig.add_trace(go.Scatter(x=rep.index, y=rep[col], mode='lines', stackgroup='one', name=col,
                                 line=dict(width=0.5, color=CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)]),
                                 fillcolor=CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)]))
    _layout(fig, title="Representatividade — FICs", pct_y=True)
    fig.update_yaxes(range=[0, 1])
    return fig


def plotly_pl_fics_linhas(motor, dicionario_fics):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    pl_fics = (motor.df_pl[cnpjs] / 1000).rename(columns=dicionario_fics)
    fig = go.Figure()
    for i, col in enumerate(pl_fics.columns):
        fig.add_trace(go.Scatter(x=pl_fics.index, y=pl_fics[col], mode='lines', name=col,
                                 line=dict(color=CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)], width=1.5)))
    _layout(fig, title="PL FICs — Evolução", ylabel="R$ Bilhões")
    return fig


def plotly_pl_fics_area(motor, dicionario_fics):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    pl_fics = (motor.df_pl[cnpjs] / 1000).rename(columns=dicionario_fics)
    fig = go.Figure()
    for i, col in enumerate(pl_fics.columns):
        fig.add_trace(go.Scatter(x=pl_fics.index, y=pl_fics[col], mode='lines', stackgroup='one', name=col,
                                 line=dict(width=0.5, color=CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)]),
                                 fillcolor=CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)]))
    _layout(fig, title="PL FICs — Composição", ylabel="R$ Bilhões")
    return fig


def plotly_pl_crescimento(motor, dicionario_masters, dicionario_fics):
    cnpjs_m = [c for c in dicionario_masters.keys() if c in motor.df_pl.columns]
    cnpjs_f = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    pl_m = motor.df_pl[cnpjs_m].sum(axis=1) / 1000 if cnpjs_m else pd.Series(dtype=float)
    pl_f = motor.df_pl[cnpjs_f].sum(axis=1) / 1000 if cnpjs_f else pd.Series(dtype=float)
    ratio = (pl_f / pl_m).replace([np.inf, -np.inf], np.nan)

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Scatter(x=pl_m.index, y=pl_m.values, mode='lines', fill='tozeroy',
                             name='PL Masters', line=dict(color='#e74c3c', width=0.5),
                             fillcolor='rgba(231,76,60,0.4)'), secondary_y=False)
    fig.add_trace(go.Scatter(x=pl_f.index, y=pl_f.values, mode='lines', fill='tozeroy',
                             name='PL FICs', line=dict(color='#c0392b', width=0.5),
                             fillcolor='rgba(192,57,43,0.35)'), secondary_y=False)
    fig.add_trace(go.Scatter(x=ratio.index, y=ratio.values, mode='lines',
                             name='FICs/Masters', line=dict(color='#2c3e50', width=2)), secondary_y=True)
    fig.update_layout(
        title=dict(text="<b>PL — Crescimento</b>", font=dict(size=12, color="#111"), x=0.5),
        plot_bgcolor="#fff", paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=45, r=45, t=35, b=35), height=350,
        legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5, font=dict(size=9)),
        hovermode="x unified",
        xaxis=dict(showgrid=False, tickformat="%b-%y", tickfont=dict(size=8)),
        yaxis=dict(title="R$ Bilhões", showgrid=True, gridcolor="#e0e0e0", tickfont=dict(size=9)),
    )
    fig.update_yaxes(title_text="FICs/Masters", tickformat=".0%", showgrid=False, secondary_y=True)
    return fig


def plotly_pl_total_ajustado(motor, dicionario_masters, dict_vol_target):
    cnpjs = [c for c in dicionario_masters.keys() if c in motor.df_pl.columns]
    vol_target = pd.Series([dict_vol_target.get(c, 0.06) for c in cnpjs], index=cnpjs)
    pl_total = motor.df_pl[cnpjs].sum(axis=1) / 1000
    pl_ajustado = motor.df_pl[cnpjs].multiply(vol_target, axis=1).sum(axis=1) / 1000
    pl_total = pl_total[pl_total > 0]; pl_ajustado = pl_ajustado[pl_ajustado > 0]

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=pl_total.index, y=pl_total.values, mode='lines',
                             name='PL Total', line=dict(color='#c0392b', width=2)))
    fig.add_trace(go.Scatter(x=pl_ajustado.index, y=pl_ajustado.values, mode='lines',
                             name='PL Ajustado', line=dict(color='#2c3e50', width=2)))
    _layout(fig, title="PL Total vs PL Ajustado — Masters", ylabel="R$ Bilhões")
    return fig


def plotly_distribuicao_canal(motor, dicionario_fics):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    df_pl = motor.df_pl[cnpjs].loc[:motor.data_ref].fillna(0)
    canal_map = {c: _classificar_canal(dicionario_fics[c]) for c in cnpjs}
    df_canal = pd.DataFrame({cat: df_pl[[c for c in cnpjs if canal_map[c] == cat]].sum(axis=1) for cat in _ORDEM_CANAL})
    df_pct = df_canal.div(df_canal.sum(axis=1).replace(0, np.nan), axis=0).fillna(0)
    fig = go.Figure()
    for cat in _ORDEM_CANAL:
        fig.add_trace(go.Scatter(x=df_pct.index, y=df_pct[cat], mode='lines', stackgroup='one', name=cat,
                                 line=dict(width=0.5, color=_CORES_CANAL[cat]), fillcolor=_CORES_CANAL[cat]))
    _layout(fig, title="Distribuição por Canal — FICs", pct_y=True)
    fig.update_yaxes(range=[0, 1])
    return fig


def plotly_pl_ajustado_estrategia(motor, dicionario_fics, dicionario_masters, dict_vol_target):
    cnpjs_f = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    cnpjs_m = [c for c in dicionario_masters.keys() if c in motor.df_pl.columns]
    def _pl_aj(cnpjs):
        vt = pd.Series([dict_vol_target.get(c, 0.06) for c in cnpjs], index=cnpjs)
        return (motor.df_pl[cnpjs].multiply(vt, axis=1) / 0.06).sum(axis=1) / 1e6
    pl_total = motor.df_pl[cnpjs_f + cnpjs_m].sum(axis=1) / 1e6
    pl_aj_tot = _pl_aj(cnpjs_f + cnpjs_m)
    pl_aj_rf = _pl_aj([c for c in cnpjs_f if _eh_rf(dicionario_fics[c])])
    pl_aj_mm = _pl_aj([c for c in cnpjs_f if not _eh_rf(dicionario_fics[c])])

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=pl_total.index, y=pl_total.values, mode='lines', fill='tozeroy',
                             name='PL Total', line=dict(color='#e0e0e0', width=0.5), fillcolor='rgba(224,224,224,0.5)'))
    fig.add_trace(go.Scatter(x=pl_aj_tot.index, y=pl_aj_tot.values, mode='lines', name='PL Aj Total', line=dict(color='#2c3e50', width=2)))
    fig.add_trace(go.Scatter(x=pl_aj_rf.index, y=pl_aj_rf.values, mode='lines', name='PL Aj RF', line=dict(color='#e74c3c', width=2)))
    fig.add_trace(go.Scatter(x=pl_aj_mm.index, y=pl_aj_mm.values, mode='lines', name='PL Aj MM', line=dict(color='#c0392b', width=2)))
    _layout(fig, title="Evolução do PL Ajustado por Estratégia", ylabel="R$ Bilhões")
    return fig


def plotly_distribuicao_prev(motor, dicionario_fics):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    df_pl = motor.df_pl[cnpjs].loc[:motor.data_ref].fillna(0)
    prev_map = {c: _classificar_prev(dicionario_fics[c]) for c in cnpjs}
    df_prev = pd.DataFrame({cat: df_pl[[c for c in cnpjs if prev_map[c] == cat]].sum(axis=1) for cat in _ORDEM_PREV})
    df_pct = df_prev.div(df_prev.sum(axis=1).replace(0, np.nan), axis=0).fillna(0)
    fig = go.Figure()
    for cat in _ORDEM_PREV:
        fig.add_trace(go.Scatter(x=df_pct.index, y=df_pct[cat], mode='lines', stackgroup='one', name=cat,
                                 line=dict(width=0.5, color=_CORES_PREV[cat]), fillcolor=_CORES_PREV[cat]))
    _layout(fig, title="Distribuição Previdência — FICs", pct_y=True)
    fig.update_yaxes(range=[0, 1])
    return fig


def plotly_taxa_media_fics(motor, dicionario_fics, dict_taxa_adm):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    df_pl = motor.df_pl[cnpjs].loc[:motor.data_ref].fillna(0)
    taxas = pd.Series([dict_taxa_adm.get(c, 0.015) for c in cnpjs], index=cnpjs)
    taxa_media = (df_pl.multiply(taxas, axis=1).sum(axis=1) / df_pl.sum(axis=1).replace(0, np.nan)).fillna(0)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=taxa_media.index, y=taxa_media.values, mode='lines',
                             name='Taxa Média', line=dict(color='#c0392b', width=2)))
    _layout(fig, title="Taxa Média FICs", show_legend=False, pct_y=True, pct_fmt=".2%")
    return fig


def plotly_dist_receita_adm(motor, dicionario_fics, dict_taxa_adm):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    df_pl = motor.df_pl[cnpjs].loc[:motor.data_ref].fillna(0)
    taxas = pd.Series([dict_taxa_adm.get(c, 0.015) for c in cnpjs], index=cnpjs)
    df_rec = df_pl.multiply(taxas, axis=1)
    cnpjs_rf = [c for c in cnpjs if _eh_rf(dicionario_fics[c])]
    cnpjs_mm = [c for c in cnpjs if not _eh_rf(dicionario_fics[c])]
    df_cat = pd.DataFrame({
        'RF': df_rec[cnpjs_rf].sum(axis=1) if cnpjs_rf else pd.Series(0, index=df_pl.index),
        'MM': df_rec[cnpjs_mm].sum(axis=1) if cnpjs_mm else pd.Series(0, index=df_pl.index)})
    df_pct = df_cat.div(df_cat.sum(axis=1).replace(0, np.nan), axis=0).fillna(0)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_pct.index, y=df_pct['RF'], mode='lines', stackgroup='one', name='RF',
                             line=dict(width=0.5, color='#e67e22'), fillcolor='#e67e22'))
    fig.add_trace(go.Scatter(x=df_pct.index, y=df_pct['MM'], mode='lines', stackgroup='one', name='MM',
                             line=dict(width=0.5, color='#c0392b'), fillcolor='#c0392b'))
    _layout(fig, title="Distribuição Receita Adm — Estratégia", pct_y=True)
    fig.update_yaxes(range=[0, 1])
    return fig


def plotly_receita_adm_estrategia(motor, dicionario_fics, dict_taxa_adm):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    df_pl = motor.df_pl[cnpjs].loc[:motor.data_ref].fillna(0)
    taxas = pd.Series([dict_taxa_adm.get(c, 0.015) for c in cnpjs], index=cnpjs)
    df_rec = df_pl.multiply(taxas, axis=1) / 1e6
    cnpjs_rf = [c for c in cnpjs if _eh_rf(dicionario_fics[c])]
    cnpjs_mm = [c for c in cnpjs if not _eh_rf(dicionario_fics[c])]
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_rec.index, y=df_rec[cnpjs_rf].sum(axis=1) if cnpjs_rf else pd.Series(0, index=df_pl.index),
                             mode='lines', name='RF', line=dict(color='#e67e22', width=2)))
    fig.add_trace(go.Scatter(x=df_rec.index, y=df_rec[cnpjs_mm].sum(axis=1) if cnpjs_mm else pd.Series(0, index=df_pl.index),
                             mode='lines', name='MM', line=dict(color='#c0392b', width=2)))
    _layout(fig, title="Receita de Adm por Estratégia — FICs", ylabel="R$ Milhões")
    return fig


def _plotly_basico(df, y_cols, labels, colors, title, ylabel="", chart_type='plot'):
    fig = go.Figure()
    if chart_type in ('stackplot', 'stackplot_custom'):
        for col, label, color in zip(y_cols, labels, colors):
            fig.add_trace(go.Scatter(x=df.index, y=df[col], mode='lines', stackgroup='one', name=label,
                                     line=dict(width=0.5, color=color), fillcolor=color))
        _layout(fig, title=title, ylabel=ylabel)
        if chart_type == 'stackplot':
            fig.update_yaxes(tickformat=".0%", range=[0, 1])
    else:
        for col, label, color in zip(y_cols, labels, colors):
            fig.add_trace(go.Scatter(x=df.index, y=df[col], mode='lines', name=label, line=dict(color=color, width=2)))
        _layout(fig, title=title, ylabel=ylabel)
    return fig


def plotly_rec_total_componentes(df_rec):
    return _plotly_basico(df_rec, ['rec_adm_total','rec_pfee_total'], ['Rec Adm Total','Rec PFee Total'], ['#bcdff0','#1a2439'], 'Receita Total (/ano)', 'R$ Milhões', 'stackplot_custom')

def plotly_rec_total_estrategia(df_rec):
    return _plotly_basico(df_rec, ['rec_tot_rf','rec_tot_mm'], ['RF','MM'], ['#e67e22','#c0392b'], 'Receita Total por Estratégia (/ano)', 'R$ Milhões')

def plotly_rec_adm_estrategia_total(df_rec):
    return _plotly_basico(df_rec, ['rec_adm_rf','rec_adm_mm'], ['RF','MM'], ['#e67e22','#c0392b'], 'Receita Adm por Estratégia (/ano)', 'R$ Milhões')

def plotly_dist_rec_total_componentes(df_rec):
    tot = (df_rec['rec_adm_total'] + df_rec['rec_pfee_total']).replace(0, np.nan)
    df_pct = pd.DataFrame({'adm': (df_rec['rec_adm_total']/tot).fillna(0), 'pfee': (df_rec['rec_pfee_total']/tot).fillna(0)})
    return _plotly_basico(df_pct, ['adm','pfee'], ['Rec Adm Total','Rec PFee Total'], ['#bcdff0','#1a2439'], 'Dist. Receita Total (/ano)', '', 'stackplot')

def plotly_dist_rec_total_estrategia(df_rec):
    tot = (df_rec['rec_tot_rf'] + df_rec['rec_tot_mm']).replace(0, np.nan)
    df_pct = pd.DataFrame({'rf': (df_rec['rec_tot_rf']/tot).fillna(0), 'mm': (df_rec['rec_tot_mm']/tot).fillna(0)})
    return _plotly_basico(df_pct, ['rf','mm'], ['RF','MM'], ['#e67e22','#c0392b'], 'Dist. Receita Total Estratégia (/ano)', '', 'stackplot')

def plotly_dist_rec_adm_estrategia(df_rec):
    tot = (df_rec['rec_adm_rf'] + df_rec['rec_adm_mm']).replace(0, np.nan)
    df_pct = pd.DataFrame({'rf': (df_rec['rec_adm_rf']/tot).fillna(0), 'mm': (df_rec['rec_adm_mm']/tot).fillna(0)})
    return _plotly_basico(df_pct, ['rf','mm'], ['RF','MM'], ['#e67e22','#c0392b'], 'Dist. Receita Adm Estratégia (/ano)', '', 'stackplot')

def plotly_rec_pfee_estrategia(df_pfee):
    return _plotly_basico(df_pfee, ['rec_pfee_rf','rec_pfee_mm'], ['RF','MM'], ['#e67e22','#c0392b'], 'Receita PFee por Estratégia (/ano)', 'R$ Milhões')

def plotly_taxa_efetiva_total(df_pfee):
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_pfee.index, y=df_pfee['taxa_efetiva_total'], mode='lines',
                             name='Taxa Efetiva', line=dict(color='#111111', width=2)))
    _layout(fig, title="Receita Total / PL Ajustado = Taxa Efetiva", show_legend=False, pct_y=True, pct_fmt=".2%")
    return fig

def plotly_dist_pfee_estrategia(df_pfee):
    tot = (df_pfee['rec_pfee_rf'] + df_pfee['rec_pfee_mm']).replace(0, np.nan)
    df_pct = pd.DataFrame({'rf': (df_pfee['rec_pfee_rf']/tot).fillna(0), 'mm': (df_pfee['rec_pfee_mm']/tot).fillna(0)})
    return _plotly_basico(df_pct, ['rf','mm'], ['RF','MM'], ['#e67e22','#c0392b'], 'Dist. Receita PFee Estratégia (/ano)', '', 'stackplot')

def plotly_dist_taxa_efetiva_estrategia(df_pfee):
    tot = (df_pfee['rec_tot_rf'] + df_pfee['rec_tot_mm']).replace(0, np.nan)
    df_pct = pd.DataFrame({'rf': (df_pfee['rec_tot_rf']/tot).fillna(0), 'mm': (df_pfee['rec_tot_mm']/tot).fillna(0)})
    return _plotly_basico(df_pct, ['rf','mm'], ['RF','MM'], ['#e67e22','#c0392b'], 'Receita Total / PL Aj por Estratégia', '', 'stackplot')


# ==========================================
# GERADORES DE PANDAS STYLER (100% NATIVOS)
# ==========================================

def get_styled_retornos(motor, dicionario_prateleira):
    cnpjs = list(dicionario_prateleira.keys())
    df_ret = motor.calcular_retornos(cnpjs).rename(index=dicionario_prateleira)
    df_ret.loc['CDI'] = 0.0

    df_num = df_ret.iloc[:-1]
    col_mins = df_num.min()
    col_maxs = df_num.max()

    def style_cells(df_in):
        styles = pd.DataFrame('', index=df_in.index, columns=df_in.columns)
        for idx, row in df_in.iloc[:-1].iterrows():
            for col_name, val in row.items():
                bg, tx = get_cell_colors(val, col_mins[col_name], col_maxs[col_name])
                styles.loc[idx, col_name] = f'background-color: {bg}; color: {tx}; text-align: center;'
        for col_name in df_in.columns:
            styles.loc['CDI', col_name] = 'background-color: #e9ecef; color: #000; font-weight: bold; text-align: center;'
        return styles

    def fmt_pct_val(v):
        if pd.isna(v): return "-"
        return "{:.2f}%".format(v * 100).replace('.', ',')

    styler = df_ret.style.apply(style_cells, axis=None).format(fmt_pct_val)
    return styler


def get_styled_captacao(motor, dicionario_prateleira):
    cnpjs = list(dicionario_prateleira.keys())
    df_cap = motor.calcular_captacao_janelas(cnpjs).rename(index=dicionario_prateleira)
    df_cap.loc['Soma'] = df_cap.sum(numeric_only=True)

    df_num = df_cap.iloc[:-1]
    col_mins = df_num.min()
    col_maxs = df_num.max()

    def style_cells(df_in):
        styles = pd.DataFrame('', index=df_in.index, columns=df_in.columns)
        for idx, row in df_in.iloc[:-1].iterrows():
            for col_name, val in row.items():
                bg, tx = get_cell_colors(val, col_mins[col_name], col_maxs[col_name])
                styles.loc[idx, col_name] = f'background-color: {bg}; color: {tx}; text-align: center;'
        for col_name in df_in.columns:
            styles.loc['Soma', col_name] = 'background-color: #e9ecef; color: #000; font-weight: bold; text-align: center;'
        return styles

    def fmt_num_val(v):
        if pd.isna(v): return "-"
        return "{:,.0f}".format(v).replace(',', '.')

    styler = df_cap.style.apply(style_cells, axis=None).format(fmt_num_val)
    return styler


def get_styled_volatilidade(motor, dicionario_prateleira):
    cnpjs = list(dicionario_prateleira.keys())
    df_vol = motor.calcular_volatilidade(cnpjs).rename(index=dicionario_prateleira)

    def fmt_pct_val(v):
        if pd.isna(v): return "-"
        return "{:.2f}%".format(v * 100).replace('.', ',')

    styler = df_vol.style.format(fmt_pct_val).set_properties(**{'text-align': 'center'})
    return styler


def get_styled_pl(df_pl_aj, data_ref_str):
    df_formatted = df_pl_aj.copy()

    def style_table(df_in):
        styles = pd.DataFrame('', index=df_in.index, columns=df_in.columns)
        styles[data_ref_str] = 'text-align: center;'
        styles['Vol Target'] = 'text-align: center;'
        styles['PL ajustado (R$MM)'] = 'text-align: center;'
        return styles

    def fmt_val(v):
        if pd.isna(v): return "-"
        if isinstance(v, float) and v < 1.0 and v > 0:
            return "{:.0%}".format(v)
        return "{:,.0f}".format(v).replace(',', '.')

    styler = df_formatted.style.apply(style_table, axis=None).format(fmt_val)
    return styler


# ==========================================
# CARREGAMENTO DE DADOS
# ==========================================
@st.cache_data(show_spinner="Carregando dados...")
def carregar_dados(nome_aba, data_b3):
    excel = ReadExcel(r'base_acompanhamento.xlsm', aba=nome_aba)
    lista_masters, lista_gd, lista_fics, dicionario_prateleira, dicionario_masters, dicionario_fics, dict_vol_target = excel.extrair_listas()
    dict_taxa_adm = {t[0]: t[1] for t in lista_fics}
    cnpjs_m = [t[0] for t in lista_masters]
    cnpjs_g = [t[0] for t in lista_gd]
    cnpjs_f = [t[0] for t in lista_fics]
    dados = ReadData(r'dados_planilha.csv', cnpjs_m, cnpjs_g, cnpjs_f)
    motor = MotorDados(dados, data_b3, lista_masters, lista_fics, lista_gd)
    return motor, lista_masters, lista_gd, lista_fics, dicionario_prateleira, dicionario_masters, dicionario_fics, dict_vol_target, dict_taxa_adm


# ==========================================
# FORMULÁRIO DE CONFIGURAÇÃO (NO TOPO, NATIVO)
# ==========================================
cfg1, cfg2, cfg3 = st.columns([2, 2, 6])
with cfg1:
    nome_aba = st.text_input("Aba do Excel", value="Optimus", placeholder="Aba (ex: Optimus)")
with cfg2:
    data_b3 = st.date_input("Data de referência", value=pd.to_datetime("2026-06-30"))
with cfg3:
    st.empty()

data_b3_str = data_b3.strftime("%Y-%m-%d")

# ── Carregar dados ──
motor, lista_masters, lista_gd, lista_fics, dicionario_prateleira, dicionario_masters, dicionario_fics, dict_vol_target, dict_taxa_adm = carregar_dados(nome_aba, data_b3_str)
data_ref_str = motor.data_ref.strftime('%d/%m/%Y')
cnpjs_prat = list(dicionario_prateleira.keys())
nomes_prat = list(dicionario_prateleira.values())

f_n = lambda v: "{:,.0f}".format(v).replace(',', '.') if pd.notnull(v) else "-"
f_p = lambda v: "{:.0%}".format(v) if pd.notnull(v) else "-"


# ==========================================
# PAGE 1: Resumo & Retornos (100% Nativo Streamlit)
# ==========================================
st.markdown("---")
st.title(f"📊 {nome_aba.upper()}")
st.caption(f"**itaú Asset** | Data de referência: {data_ref_str}")

# KPI Metrics - Masters, FICs, GD
st.subheader("Métricas de Resumo (R$MM)")
col_m, col_f, col_g = st.columns([3, 3, 1])

rec_val, rec_var = motor.calcular_receita(motor.cnpjs_master_tup, 'master')
capt_val, capt_var = motor.calcular_capliq(motor.cnpjs_master, 12)
pl_val, pl_var = _obter_pl_dados(motor, motor.cnpjs_master)

with col_m:
    st.markdown("**Masters**")
    m1, m2, m3 = st.columns(3)
    m1.metric("Receita (adm + PFee)", f_n(rec_val), delta=f_p(rec_var))
    m2.metric("Captação (12m)", f_n(capt_val), delta=f_p(capt_var))
    m3.metric("PL", f_n(pl_val), delta=f_p(pl_var))

rec_f_val, rec_f_var = motor.calcular_receita(motor.cnpjs_fic_tup, 'fic')
capt_f_val, capt_f_var = motor.calcular_capliq(motor.cnpjs_fic, 12)
pl_f_val, pl_f_var = _obter_pl_dados(motor, motor.cnpjs_fic)

with col_f:
    st.markdown("**FICs**")
    f1, f2, f3 = st.columns(3)
    f1.metric("Receita (adm)", f_n(rec_f_val), delta=f_p(rec_f_var))
    f2.metric("Captação (12m)", f_n(capt_f_val), delta=f_p(capt_f_var))
    f3.metric("PL", f_n(pl_f_val), delta=f_p(pl_f_var))

gd_capt, gd_var = motor.calcular_capliq(motor.cnpjs_gd, 12)
with col_g:
    st.markdown("**GD**")
    st.metric("Captação (12m)", f_n(gd_capt), delta=f_p(gd_var))

# Retorno Tabela
st.subheader(f"Retorno Família {nome_aba}")
st.table(get_styled_retornos(motor, dicionario_prateleira))

# Retorno Gráficos
if len(cnpjs_prat) >= 2:
    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(plotly_retorno(motor, cnpjs_prat[0], nomes_prat[0]), width='stretch', key="ret1")
    with c2:
        st.plotly_chart(plotly_retorno(motor, cnpjs_prat[1], nomes_prat[1]), width='stretch', key="ret2")
elif len(cnpjs_prat) >= 1:
    st.plotly_chart(plotly_retorno(motor, cnpjs_prat[0], nomes_prat[0]), width='stretch', key="ret1")

st.divider()


# ==========================================
# PAGE 2: Captação (100% Nativo)
# ==========================================
st.subheader(f"Captação Família {nome_aba} (R$MM)")
st.caption(f"**itaú Asset** | {data_ref_str}")
st.table(get_styled_captacao(motor, dicionario_prateleira))

st.divider()


# ==========================================
# PAGE 3: Volatilidade, Sharpe & Drawdown (100% Nativo)
# ==========================================
st.subheader("Volatilidade e Sharpe - FICs")
st.caption(f"**itaú Asset** | {data_ref_str}")

c3a, c3b = st.columns([1.1, 1.3])
with c3a:
    st.table(get_styled_volatilidade(motor, dicionario_prateleira))
with c3b:
    st.plotly_chart(plotly_sharpe_rolling(motor, dicionario_prateleira, window=252), width='stretch', key="sharpe")

st.plotly_chart(plotly_vol_rolling(motor, dicionario_prateleira, window=63, titulo_janela="3m"), width='stretch', key="vol3m")

c3c, c3d, c3e = st.columns([1, 1, 1])
with c3c:
    st.plotly_chart(plotly_vol_rolling(motor, dicionario_prateleira, window=252, titulo_janela="12m"), width='stretch', key="vol12m")
with c3d:
    if len(cnpjs_prat) >= 1:
        st.plotly_chart(plotly_drawdown(motor, cnpjs_prat[0], nomes_prat[0]), width='stretch', key="dd1")
with c3e:
    if len(cnpjs_prat) >= 2:
        st.plotly_chart(plotly_drawdown(motor, cnpjs_prat[1], nomes_prat[1]), width='stretch', key="dd2")

st.divider()


# ==========================================
# PAGE 4: PL Masters (100% Nativo)
# ==========================================
st.subheader("PL Masters")
df_masters = motor.calcular_pl_ajustado(list(dicionario_masters.keys()),
                                         dict_vol_target=dict_vol_target,
                                         nome_coluna_pl=data_ref_str).rename(index=dicionario_masters)

c4a, c4b = st.columns([1.5, 1])
with c4a:
    st.table(get_styled_pl(df_masters, data_ref_str))
    tot1, tot2 = st.columns(2)
    tot1.metric("PL Total (R$MM)", f_n(df_masters[data_ref_str].sum()))
    tot2.metric("PL Ajustado Total (R$MM)", f_n(df_masters['PL ajustado (R$MM)'].sum()))

with c4b:
    st.plotly_chart(plotly_pl_crescimento(motor, dicionario_masters, dicionario_fics), width='stretch', key="plcresc")
    st.plotly_chart(plotly_pl_total_ajustado(motor, dicionario_masters, dict_vol_target), width='stretch', key="plaj")

st.divider()


# ==========================================
# PAGE 5: PL FICs (1/2) (100% Nativo)
# ==========================================
st.subheader("PL FICs (Parte 1)")
df_fics = motor.calcular_pl_ajustado(list(dicionario_fics.keys()),
                                      nome_coluna_pl=data_ref_str).rename(index=dicionario_fics)
df_fics_p1 = df_fics.iloc[:26]

c5a, c5b = st.columns([1.5, 1])
with c5a:
    st.table(get_styled_pl(df_fics_p1, data_ref_str))
with c5b:
    st.plotly_chart(plotly_pl_fics_linhas(motor, dicionario_fics), width='stretch', key="ficlin")
    st.plotly_chart(plotly_pl_fics_area(motor, dicionario_fics), width='stretch', key="ficarea")

st.divider()


# ==========================================
# PAGE 6: PL FICs (2/2) + Métricas + Capacity (100% Nativo)
# ==========================================
st.subheader("PL FICs (Parte 2)")
df_fics_p2 = df_fics.iloc[26:]
metricas = calcular_metricas_fics(motor, dicionario_fics, dicionario_masters, dict_vol_target)
capacity, remanescente, pct_cap = calcular_capacity_bloco(motor, dicionario_fics, dicionario_masters, dict_vol_target)

c6a, c6b = st.columns([1.5, 1])
with c6a:
    if not df_fics_p2.empty:
        st.table(get_styled_pl(df_fics_p2, data_ref_str))
    tot_f1, tot_f2 = st.columns(2)
    tot_f1.metric("PL Total FICs (R$MM)", f_n(df_fics[data_ref_str].sum()))
    tot_f2.metric("PL Ajustado Total FICs (R$MM)", f_n(df_fics['PL ajustado (R$MM)'].sum()))

with c6b:
    st.plotly_chart(plotly_representatividade(motor, dicionario_fics), width='stretch', key="repr")
    
    st.markdown("**Métricas FICs / Total**")
    m_col1, m_col2 = st.columns(2)
    m_col1.metric("FICs/Total", metricas['fics_total'])
    m_col2.metric("FICs/Total Ajustado", metricas['fics_aj_total'])

    st.markdown("**Métricas Por Estratégia**")
    e1, e2, e3, e4 = st.columns(4)
    e1.metric("PL RF", metricas['pl_rf'], delta=metricas['pct_rf'])
    e2.metric("PL RF Aj", metricas['pl_aj_rf'], delta=metricas['pct_aj_rf'])
    e3.metric("PL MM", metricas['pl_mm'], delta=metricas['pct_mm'])
    e4.metric("PL MM Aj", metricas['pl_aj_mm'], delta=metricas['pct_aj_mm'])

    st.markdown("**Capacity**")
    cap1, cap2, cap3 = st.columns(3)
    cap1.metric("Capacity (R$MM)", capacity)
    cap2.metric("Remanescente (R$MM)", remanescente)
    cap3.metric("% Capacity usado", pct_cap)

st.divider()


# ==========================================
# PAGE 7: Distribuição Canal, Estratégia & Previdência (100% Nativo)
# ==========================================
st.subheader("Distribuição por Canal & Estratégia")

c7a, c7b = st.columns([1.15, 1])
with c7a:
    st.plotly_chart(plotly_distribuicao_canal(motor, dicionario_fics), width='stretch', key="canal")
with c7b:
    st.plotly_chart(plotly_pl_ajustado_estrategia(motor, dicionario_fics, dicionario_masters, dict_vol_target), width='stretch', key="plajest")
    st.plotly_chart(plotly_distribuicao_prev(motor, dicionario_fics), width='stretch', key="prev")

st.divider()


# ==========================================
# PAGE 8: Receita FICs (100% Nativo)
# ==========================================
st.subheader("Receita FICs")
c8a, c8b, c8c = st.columns(3)
with c8a:
    st.plotly_chart(plotly_taxa_media_fics(motor, dicionario_fics, dict_taxa_adm), width='stretch', key="txm")
with c8b:
    st.plotly_chart(plotly_dist_receita_adm(motor, dicionario_fics, dict_taxa_adm), width='stretch', key="dra")
with c8c:
    st.plotly_chart(plotly_receita_adm_estrategia(motor, dicionario_fics, dict_taxa_adm), width='stretch', key="rae")

st.divider()


# ==========================================
# PAGE 9: Receita Total (100% Nativo)
# ==========================================
st.subheader("Receita Total")
df_rec = _calcular_series_receita_total(motor, lista_masters, lista_fics, dicionario_masters, dicionario_fics)

c9a, c9b, c9c = st.columns(3)
with c9a:
    st.plotly_chart(plotly_rec_total_componentes(df_rec), width='stretch', key="rtc")
with c9b:
    st.plotly_chart(plotly_rec_total_estrategia(df_rec), width='stretch', key="rte")
with c9c:
    st.plotly_chart(plotly_rec_adm_estrategia_total(df_rec), width='stretch', key="rae2")

c9d, c9e, c9f = st.columns(3)
with c9d:
    st.plotly_chart(plotly_dist_rec_total_componentes(df_rec), width='stretch', key="drtc")
with c9e:
    st.plotly_chart(plotly_dist_rec_total_estrategia(df_rec), width='stretch', key="drte")
with c9f:
    st.plotly_chart(plotly_dist_rec_adm_estrategia(df_rec), width='stretch', key="drae")

st.divider()


# ==========================================
# PAGE 10: PFee & Taxa Efetiva (100% Nativo)
# ==========================================
st.subheader("Receita de PFee & Taxa Efetiva")
df_pfee = _calcular_series_pfee_taxa_efetiva(motor, lista_masters, lista_fics, dicionario_masters, dicionario_fics, dict_vol_target)

c10a, c10b = st.columns(2)
with c10a:
    st.plotly_chart(plotly_rec_pfee_estrategia(df_pfee), width='stretch', key="rpfe")
with c10b:
    st.plotly_chart(plotly_taxa_efetiva_total(df_pfee), width='stretch', key="tet")

c10c, c10d = st.columns(2)
with c10c:
    st.plotly_chart(plotly_dist_pfee_estrategia(df_pfee), width='stretch', key="dpfe")
with c10d:
    st.plotly_chart(plotly_dist_taxa_efetiva_estrategia(df_pfee), width='stretch', key="dtee")
