
import pandas as pd
import numpy as np

def fmt_num(v):
    return "{:,.0f}".format(v).replace(',', '.') if pd.notnull(v) else "-"

def fmt_pct(v):
    return "{:.1%}".format(v).replace('.', ',') if pd.notnull(v) else "-"

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import io
import base64

def gerar_grafico_retorno(motor, cnpj, nome_fundo):
    if cnpj not in motor.df_cotas.columns:
        return ""

    cotas = motor.df_cotas[cnpj].dropna()
    if cotas.empty or len(cotas) < 2:
        return ""

    ret_acum = (cotas / cotas.iloc[0]) - 1
    
    dias = np.arange(len(ret_acum))
    cdi_diario = (1.13) ** (1/252) - 1
    cdi_acum = (1 + cdi_diario) ** dias - 1
    cdi_series = pd.Series(cdi_acum, index=ret_acum.index)

    fig, ax = plt.subplots(figsize=(6.5, 3.8))
    
    ax.plot(ret_acum.index, ret_acum.values, color='#932c3d', linewidth=2.5, label=nome_fundo)
    ax.plot(cdi_series.index, cdi_series.values, color='#505568', linewidth=2.5, label='CDI')
    
    ax.set_title(f"Retorno - {nome_fundo}", fontsize=12, fontweight='bold', color='#111111', pad=15)
    
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_color('#cccccc')
    
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f"{int(round(y*100))}%"))
    
    max_val = max(ret_acum.max(), cdi_series.max())
    max_lim = np.ceil(max_val * 20) / 20
    if max_lim < max_val + 0.01:
        max_lim += 0.05
    ax.set_ylim(0, max_lim)
    ax.set_yticks(np.arange(0, max_lim + 0.01, 0.05))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        dt = mdates.num2date(x)
        mes = MESES_PT[dt.month]
        ano = str(dt.year)[-2:]
        return f"{mes}-{ano}"
        
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    
    ax.tick_params(axis='x', rotation=90, labelsize=9, color='#cccccc')
    ax.tick_params(axis='y', length=0, labelsize=9)
    
    ax.legend(fontsize=9, loc='upper center', bbox_to_anchor=(0.5, -0.25), ncol=2, frameon=False)
    
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)

    img_str = base64.b64encode(buffer.getvalue()).decode('utf-8')
    return f"<img src='data:image/png;base64,{img_str}' style='max-width:100%; height:auto;' />"




CORES_PRATELEIRA = ['#c0392b', '#e8825a', '#7b1e1e', '#3b6fb5', '#a11a1a', '#f4a6b8',
                    '#2ecc71', '#9b59b6', '#f39c12', '#1abc9c', '#e74c3c', '#3498db']


def _grafico_multilinhas(series_dict, titulo, fmt_y='{:.0%}', figsize=(5.5, 2.8), show_legend=True, legend_loc='best', legend_bbox=None):
    fig, ax = plt.subplots(figsize=figsize)
    for i, (nome, serie) in enumerate(series_dict.items()):
        cor = CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)]
        ax.plot(serie.index, serie.values, color=cor, linewidth=1.5, label=nome)
    ax.set_title(titulo, fontsize=11, fontweight='bold', color='#111111')
    
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_color('#cccccc')
    
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: fmt_y.format(y)))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        dt = mdates.num2date(x)
        mes = MESES_PT[dt.month]
        ano = str(dt.year)[-2:]
        return f"{mes}-{ano}"
        
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=90, labelsize=8, color='#cccccc')
    ax.tick_params(axis='y', length=0, labelsize=8)
    
    if show_legend:
        if legend_bbox:
            ax.legend(fontsize=8, loc=legend_loc, bbox_to_anchor=legend_bbox, ncol=2, frameon=False)
        else:
            ax.legend(fontsize=8, loc=legend_loc, ncol=2, frameon=False)
            
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    img_str = base64.b64encode(buffer.getvalue()).decode('utf-8')
    return f"<img src='data:image/png;base64,{img_str}' style='max-width:100%; height:auto;' />"




def gerar_grafico_sharpe_rolling(motor, dicionario_prateleira, window=252, show_legend=True, legend_loc='best', legend_bbox=None, figsize=(6, 2.8)):
    cnpjs = [c for c in dicionario_prateleira.keys() if c in motor.df_cotas.columns]
    df_cotas = motor.df_cotas[cnpjs].loc[:motor.data_ref]
    ret_diario = df_cotas.pct_change()
    vol_rolling = ret_diario.rolling(window=window).std(ddof=1) * np.sqrt(252)
    ret_rolling = (df_cotas / df_cotas.shift(window)) - 1
    ret_anualizado = (1 + ret_rolling) ** (252 / window) - 1
    sharpe_rolling = ret_anualizado / vol_rolling
    sharpe_rolling = sharpe_rolling.dropna(how='all')
    series = {dicionario_prateleira[c]: sharpe_rolling[c].dropna() for c in cnpjs if not sharpe_rolling[c].dropna().empty}
    return _grafico_multilinhas(series, "Evolução do Sharpe Líquido (12m)", fmt_y='{:.1f}', figsize=figsize, show_legend=show_legend, legend_loc=legend_loc, legend_bbox=legend_bbox)


def gerar_grafico_vol_rolling(motor, dicionario_prateleira, window=63, titulo_janela="3m", show_legend=True, legend_loc='best', legend_bbox=None, figsize=(5.5, 2.8)):
    cnpjs = [c for c in dicionario_prateleira.keys() if c in motor.df_cotas.columns]
    df_cotas = motor.df_cotas[cnpjs].loc[:motor.data_ref]
    ret_diario = df_cotas.pct_change()
    vol_rolling = ret_diario.rolling(window=window).std(ddof=1) * np.sqrt(252)
    vol_rolling = vol_rolling.dropna(how='all')
    series = {dicionario_prateleira[c]: vol_rolling[c].dropna() for c in cnpjs if not vol_rolling[c].dropna().empty}
    return _grafico_multilinhas(series, f"Evolução da Volatilidade ({titulo_janela})", show_legend=show_legend, legend_loc=legend_loc, legend_bbox=legend_bbox, figsize=figsize)


def gerar_grafico_drawdown(motor, cnpj, nome_fundo):
    if cnpj not in motor.df_cotas.columns:
        return ""
    cotas = motor.df_cotas[cnpj].dropna()
    if cotas.empty or len(cotas) < 2:
        return ""
    picos = cotas.expanding(min_periods=1).max()
    dd = (cotas / picos) - 1
    fig, ax = plt.subplots(figsize=(4, 2.6))
    ax.fill_between(dd.index, dd.values, 0, color='#b5251f', alpha=0.7)
    ax.plot(dd.index, dd.values, color='#b5251f', linewidth=0.8)
    ax.set_title(f"Drawdown - {nome_fundo}", fontsize=9, fontweight='bold', color='#333')
    ax.grid(True, linestyle='--', alpha=0.4)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.2%}'.format(y)))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=4))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b-%y'))
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, fontsize=7)
    plt.setp(ax.yaxis.get_majorticklabels(), fontsize=7)
    fig.tight_layout()
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True)
    plt.close(fig)
    img_str = base64.b64encode(buffer.getvalue()).decode('utf-8')
    return f"<img src='data:image/png;base64,{img_str}' style='max-width:100%; height:auto;' />"






def _eh_rf(nome):
    nome_up = nome.upper()
    return any(k in nome_up for k in ['RENDA FIXA', ' RF ', 'RF LP', 'INFRA', 'RF\x00'])




def calcular_metricas_fics(motor, dicionario_fics, dicionario_masters, dict_vol_target):
    cnpjs_f = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    cnpjs_m = [c for c in dicionario_masters.keys() if c in motor.df_pl.columns]
    idx_ref = motor.df_pl.index[motor.df_pl.index <= motor.data_ref].max()

    pl_fics = motor.df_pl[cnpjs_f].loc[idx_ref].sum()
    pl_masters = motor.df_pl[cnpjs_m].loc[idx_ref].sum()
    pl_total_est = pl_fics + pl_masters

    vol_f = pd.Series([dict_vol_target.get(c, 0.06) for c in cnpjs_f], index=cnpjs_f)
    pl_aj_fics = (motor.df_pl[cnpjs_f].loc[idx_ref] * vol_f / 0.06).sum()
    vol_m = pd.Series([dict_vol_target.get(c, 0.06) for c in cnpjs_m], index=cnpjs_m)
    pl_aj_masters = (motor.df_pl[cnpjs_m].loc[idx_ref] * vol_m / 0.06).sum()
    pl_aj_total = pl_aj_fics + pl_aj_masters

    pl_rf = sum(motor.df_pl[c].loc[idx_ref] for c in cnpjs_f if _eh_rf(dicionario_fics[c]))
    pl_mm = sum(motor.df_pl[c].loc[idx_ref] for c in cnpjs_f if not _eh_rf(dicionario_fics[c]))
    pl_aj_rf = sum(motor.df_pl[c].loc[idx_ref] * dict_vol_target.get(c, 0.06) / 0.06
                   for c in cnpjs_f if _eh_rf(dicionario_fics[c]))
    pl_aj_mm = sum(motor.df_pl[c].loc[idx_ref] * dict_vol_target.get(c, 0.06) / 0.06
                   for c in cnpjs_f if not _eh_rf(dicionario_fics[c]))

    fics_total = pl_fics / pl_total_est if pl_total_est else 0
    fics_aj_total = pl_aj_fics / pl_aj_total if pl_aj_total else 0
    pct_rf = pl_rf / (pl_rf + pl_mm) if (pl_rf + pl_mm) else 0
    pct_mm = pl_mm / (pl_rf + pl_mm) if (pl_rf + pl_mm) else 0
    pct_aj_rf = pl_aj_rf / (pl_aj_rf + pl_aj_mm) if (pl_aj_rf + pl_aj_mm) else 0
    pct_aj_mm = pl_aj_mm / (pl_aj_rf + pl_aj_mm) if (pl_aj_rf + pl_aj_mm) else 0

    pl_fics_total_fmt = fmt_num(pl_fics / 1e6)
    pl_aj_fics_total_fmt = fmt_num(pl_aj_fics / 1e6)

    return {
        'pl_fics_total': pl_fics_total_fmt,
        'pl_aj_fics_total': pl_aj_fics_total_fmt,
        'fics_total': fmt_pct(fics_total),
        'fics_aj_total': fmt_pct(fics_aj_total),
        'pl_rf': fmt_num(pl_rf / 1e6), 'pct_rf': fmt_pct(pct_rf),
        'pl_mm': fmt_num(pl_mm / 1e6), 'pct_mm': fmt_pct(pct_mm),
        'pl_aj_rf': fmt_num(pl_aj_rf / 1e6), 'pct_aj_rf': fmt_pct(pct_aj_rf),
        'pl_aj_mm': fmt_num(pl_aj_mm / 1e6), 'pct_aj_mm': fmt_pct(pct_aj_mm),
    }


def calcular_capacity_bloco(motor, dicionario_fics, dicionario_masters, dict_vol_target):
    cnpjs_f = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    idx_ref = motor.df_pl.index[motor.df_pl.index <= motor.data_ref].max()
    vol_f = pd.Series([dict_vol_target.get(c, 0.06) for c in cnpjs_f], index=cnpjs_f)
    pl_aj_total = (motor.df_pl[cnpjs_f].loc[idx_ref] * vol_f / 0.06).sum() / 1e6

    cnpjs_m = [c for c in dicionario_masters.keys() if c in motor.df_fluxo.columns]
    capacity = motor.df_fluxo[cnpjs_m].sum(axis=1).cumsum().iloc[-1] / 1e6 if cnpjs_m else 0.0
    remanescente = capacity - pl_aj_total
    pct_usado = pl_aj_total / capacity if capacity else 0
    return fmt_num(capacity), fmt_num(remanescente), fmt_pct(pct_usado)


def gerar_grafico_representatividade(motor, dicionario_fics):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    pl_fics = motor.df_pl[cnpjs]
    pl_fics = pl_fics.rename(columns=dicionario_fics)
    pl_total = pl_fics.sum(axis=1)
    rep = pl_fics.div(pl_total, axis=0)

    fig, ax = plt.subplots(figsize=(6, 3.5))
    cores = [CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)] for i in range(len(rep.columns))]
    ax.stackplot(rep.index, rep.T, labels=rep.columns, colors=cores, alpha=0.9)
    
    ax.set_ylabel('R$ Bilhões', fontsize=8) # The photo shows 100%, 90% etc. We will format it as %
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=90, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(False)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=3, frameon=False)
    fig.tight_layout()

    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    img_str = base64.b64encode(buffer.getvalue()).decode('utf-8')
    return f"<img src='data:image/png;base64,{img_str}' style='max-width:100%; height:auto;' />"


def gerar_grafico_pl_fics_linhas(motor, dicionario_fics):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    pl_fics = motor.df_pl[cnpjs] / 1000
    pl_fics = pl_fics.rename(columns=dicionario_fics)

    fig, ax = plt.subplots(figsize=(6, 2.8))
    for i, col in enumerate(pl_fics.columns):
        cor = CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)]
        ax.plot(pl_fics.index, pl_fics[col], color=cor, linewidth=1.5, label=col)
    
    ax.set_ylabel('R$ Bilhões', fontsize=8)
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=90, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.4), ncol=3, frameon=False)
    fig.tight_layout()

    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    img_str = base64.b64encode(buffer.getvalue()).decode('utf-8')
    return f"<img src='data:image/png;base64,{img_str}' style='max-width:100%; height:auto;' />"


def gerar_grafico_pl_fics_area(motor, dicionario_fics):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    pl_fics = motor.df_pl[cnpjs] / 1000
    pl_fics = pl_fics.rename(columns=dicionario_fics)

    fig, ax = plt.subplots(figsize=(6, 2.8))
    cores = [CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)] for i in range(len(pl_fics.columns))]
    ax.stackplot(pl_fics.index, pl_fics.T, labels=pl_fics.columns, colors=cores, alpha=0.9)
    
    ax.set_ylabel('R$ Bilhões', fontsize=8)
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=90, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.4), ncol=3, frameon=False)
    fig.tight_layout()

    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    img_str = base64.b64encode(buffer.getvalue()).decode('utf-8')
    return f"<img src='data:image/png;base64,{img_str}' style='max-width:100%; height:auto;' />"




def gerar_grafico_pl_crescimento(motor, dicionario_masters, dicionario_fics):
    cnpjs_m = [c for c in dicionario_masters.keys() if c in motor.df_pl.columns]
    cnpjs_f = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]

    pl_m = motor.df_pl[cnpjs_m].sum(axis=1) / 1000 if cnpjs_m else pd.Series(dtype=float)
    pl_f = motor.df_pl[cnpjs_f].sum(axis=1) / 1000 if cnpjs_f else pd.Series(dtype=float)
    ratio = (pl_f / pl_m).replace([np.inf, -np.inf], np.nan)

    fig, ax1 = plt.subplots(figsize=(6, 3))
    ax1.fill_between(pl_m.index, pl_m.values, alpha=0.9, color='#e74c3c', label='PL Masters')
    ax1.fill_between(pl_f.index, pl_f.values, alpha=0.8, color='#c0392b', label='PL FICs')
    ax1.set_ylabel('R$ Bilhões', fontsize=8)
    ax1.tick_params(axis='y', labelsize=8)
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax1.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax1.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax1.tick_params(axis='x', rotation=90, labelsize=8)

    ax2 = ax1.twinx()
    ax2.plot(ratio.index, ratio.values, color='#2c3e50', linewidth=2.0, label='FICs/Masters')
    ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    ax2.tick_params(axis='y', labelsize=8)

    ax1.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax1.xaxis.grid(False)
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)
    ax1.spines['left'].set_visible(False)
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    ax2.spines['left'].set_visible(False)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=8, loc='lower center', bbox_to_anchor=(0.5, -0.4), ncol=3, frameon=False)
    fig.tight_layout()

    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    img_str = base64.b64encode(buffer.getvalue()).decode('utf-8')
    return f"<img src='data:image/png;base64,{img_str}' style='max-width:100%; height:auto;' />"


def gerar_grafico_pl_total_ajustado(motor, dicionario_masters, dict_vol_target):
    cnpjs = [c for c in dicionario_masters.keys() if c in motor.df_pl.columns]
    vol_target = pd.Series([dict_vol_target.get(c, 0.06) for c in cnpjs], index=cnpjs)

    pl_total = motor.df_pl[cnpjs].sum(axis=1) / 1000
    pl_ajustado = motor.df_pl[cnpjs].multiply(vol_target, axis=1).sum(axis=1) / 1000
    pl_total = pl_total[pl_total > 0]
    pl_ajustado = pl_ajustado[pl_ajustado > 0]

    fig, ax = plt.subplots(figsize=(6, 3))
    ax.plot(pl_total.index, pl_total.values, color='#c0392b', linewidth=2.0, label='PL Total')
    ax.plot(pl_ajustado.index, pl_ajustado.values, color='#2c3e50', linewidth=2.0, label='PL Ajustado')
    ax.set_ylabel('R$ Bilhões', fontsize=8)
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=90, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=8, loc='lower center', bbox_to_anchor=(0.5, -0.4), ncol=2, frameon=False)
    fig.tight_layout()

    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    img_str = base64.b64encode(buffer.getvalue()).decode('utf-8')
    return f"<img src='data:image/png;base64,{img_str}' style='max-width:100%; height:auto;' />"


def _classificar_canal(nome):
    n = nome.upper()
    if any(k in n for k in ['HEDGE PLUS', 'MACRO OPP', 'MACRO OPPORTUNITIES', 'HEDGE ULTRA']):
        return 'Hedge Plus + Macro Opps'
    if 'PRIVATE' in n:
        return 'Private'
    if any(k in n for k in ['DISTRIBUIDOR', 'DISTRIBUID', 'WHG', 'ONZE DISTRIBUID']):
        return 'Distribuidores'
    if any(k in n for k in ['INSTITUCIONAL', 'INST ']):
        return 'Institucionais'
    if 'INFRA' in n:
        return 'Infra'
    if any(k in n for k in ['FOF', 'FEEDER', 'FIFCIC', 'FIF CIC']):
        return 'FOF'
    return 'Varejo'


def _classificar_prev(nome):
    n = nome.upper()
    if not any(k in n for k in ['FLEXPREV', 'PREV', 'PREVIDENCIA', 'PREVIDÊNCIA']):
        return 'Não-Prev'
    if _eh_rf(nome):
        return 'Prev RF'
    return 'Prev MM'


_CORES_CANAL = {
    'Hedge Plus + Macro Opps': '#e8b878',
    'Varejo':                  '#e05a2b',
    'Private':                 '#e893a8',
    'Institucionais':          '#8b1e2e',
    'FOF':                     '#1c2b4a',
    'Infra':                   '#2f7a86',
    'Distribuidores':          '#a9d3e8',
}
_ORDEM_CANAL = ['Hedge Plus + Macro Opps', 'Varejo', 'Private', 'Institucionais', 'FOF', 'Infra', 'Distribuidores']
_ORDEM_PREV  = ['Prev RF', 'Prev MM', 'Não-Prev']
_CORES_PREV  = {'Prev RF': '#8b1e2e', 'Prev MM': '#c0392b', 'Não-Prev': '#b0b0b0'}


def gerar_grafico_distribuicao_canal(motor, dicionario_fics):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    df_pl = motor.df_pl[cnpjs].loc[:motor.data_ref].fillna(0)
    canal_map = {c: _classificar_canal(dicionario_fics[c]) for c in cnpjs}
    df_canal = pd.DataFrame({cat: df_pl[[c for c in cnpjs if canal_map[c] == cat]].sum(axis=1)
                             for cat in _ORDEM_CANAL})
    total = df_canal.sum(axis=1).replace(0, np.nan)
    df_pct = df_canal.div(total, axis=0).fillna(0)
    fig, ax = plt.subplots(figsize=(6, 4))
    cores = [_CORES_CANAL[cat] for cat in _ORDEM_CANAL]
    ax.stackplot(df_pct.index, [df_pct[cat].values for cat in _ORDEM_CANAL],
                 labels=_ORDEM_CANAL, colors=cores, alpha=0.95)
    ax.set_title('Distribuição por Canal - FICs Janeiro', fontsize=10, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    ax.set_ylim(0, 1)
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(False)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=8, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def gerar_grafico_pl_ajustado_estrategia(motor, dicionario_fics, dicionario_masters, dict_vol_target):
    cnpjs_f = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    cnpjs_m = [c for c in dicionario_masters.keys() if c in motor.df_pl.columns]
    def _pl_aj(cnpjs, _dicio=None):
        vt = pd.Series([dict_vol_target.get(c, 0.06) for c in cnpjs], index=cnpjs)
        return (motor.df_pl[cnpjs].multiply(vt, axis=1) / 0.06).sum(axis=1) / 1e6
    cnpjs_f_rf = [c for c in cnpjs_f if _eh_rf(dicionario_fics[c])]
    cnpjs_f_mm = [c for c in cnpjs_f if not _eh_rf(dicionario_fics[c])]
    pl_total   = motor.df_pl[cnpjs_f + cnpjs_m].sum(axis=1) / 1e6
    pl_aj_tot  = _pl_aj(cnpjs_f + cnpjs_m)
    pl_aj_rf   = _pl_aj(cnpjs_f_rf)
    pl_aj_mm   = _pl_aj(cnpjs_f_mm)
    fig, ax = plt.subplots(figsize=(6, 3))
    ax.fill_between(pl_total.index, pl_total.values, color='#e0e0e0', alpha=0.9, label='PL Total')
    ax.plot(pl_aj_tot.index, pl_aj_tot.values, color='#2c3e50', linewidth=2.0, label='PL Ajustado Total')
    ax.plot(pl_aj_rf.index,  pl_aj_rf.values,  color='#e74c3c', linewidth=2.0, label='PL Ajustado RF')
    ax.plot(pl_aj_mm.index,  pl_aj_mm.values,  color='#c0392b', linewidth=2.0, label='PL Ajustado MM')
    ax.set_title('Evolução do PL Ajustado por Estratégia', fontsize=10, fontweight='bold', color='#111111')
    ax.set_ylabel('R$ Bilhões', fontsize=8)
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=0, labelsize=8) # Looks like horizontal or slight in photo, let's keep it straight if space permits or 0. Wait, photo says set-23 dez-23 mar-24. 0 rotation.
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=8, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=4, frameon=False)
    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def gerar_grafico_distribuicao_prev(motor, dicionario_fics):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    df_pl = motor.df_pl[cnpjs].loc[:motor.data_ref].fillna(0)
    prev_map = {c: _classificar_prev(dicionario_fics[c]) for c in cnpjs}
    df_prev = pd.DataFrame({cat: df_pl[[c for c in cnpjs if prev_map[c] == cat]].sum(axis=1)
                            for cat in _ORDEM_PREV})
    total = df_prev.sum(axis=1).replace(0, np.nan)
    df_pct = df_prev.div(total, axis=0).fillna(0)
    fig, ax = plt.subplots(figsize=(6, 3))
    cores = [_CORES_PREV[cat] for cat in _ORDEM_PREV]
    ax.stackplot(df_pct.index, [df_pct[cat].values for cat in _ORDEM_PREV],
                 labels=_ORDEM_PREV, colors=cores, alpha=0.9)
    ax.set_title('Distribuição Estratégias Previdência - FICs Janeiro', fontsize=10, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    ax.set_ylim(0, 1)
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(False)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=8, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=3, frameon=False)
    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def gerar_grafico_taxa_media_fics(motor, dicionario_fics, dict_taxa_adm):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    df_pl = motor.df_pl[cnpjs].loc[:motor.data_ref].fillna(0)
    taxas = pd.Series([dict_taxa_adm.get(c, 0.015) for c in cnpjs], index=cnpjs)
    
    pl_total = df_pl.sum(axis=1).replace(0, np.nan)
    receita_total = df_pl.multiply(taxas, axis=1).sum(axis=1)
    taxa_media = (receita_total / pl_total).fillna(0)
    
    fig, ax = plt.subplots(figsize=(6, 3.5))
    ax.plot(taxa_media.index, taxa_media.values, color='#c0392b', linewidth=2.0)
    ax.set_title('Taxa Média FICs Janeiro', fontsize=10, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f"{y*100:.2f}%".replace('.', ',')))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def gerar_grafico_distribuicao_receita_adm(motor, dicionario_fics, dict_taxa_adm):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    df_pl = motor.df_pl[cnpjs].loc[:motor.data_ref].fillna(0)
    taxas = pd.Series([dict_taxa_adm.get(c, 0.015) for c in cnpjs], index=cnpjs)
    
    df_rec = df_pl.multiply(taxas, axis=1)
    
    cnpjs_rf = [c for c in cnpjs if _eh_rf(dicionario_fics[c])]
    cnpjs_mm = [c for c in cnpjs if not _eh_rf(dicionario_fics[c])]
    
    rec_rf = df_rec[cnpjs_rf].sum(axis=1) if cnpjs_rf else pd.Series(0, index=df_pl.index)
    rec_mm = df_rec[cnpjs_mm].sum(axis=1) if cnpjs_mm else pd.Series(0, index=df_pl.index)
    
    df_cat = pd.DataFrame({'RF': rec_rf, 'MM': rec_mm})
    total = df_cat.sum(axis=1).replace(0, np.nan)
    df_pct = df_cat.div(total, axis=0).fillna(0)
    
    fig, ax = plt.subplots(figsize=(6, 3.5))
    ax.stackplot(df_pct.index, [df_pct['RF'].values, df_pct['MM'].values],
                 labels=['RF', 'MM'], colors=['#e67e22', '#c0392b'], alpha=0.95)
    ax.set_title('Distribuição de Receita de Adm por Estratégia - FICs', fontsize=10, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    ax.set_ylim(0, 1)
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(False)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    
    ax.legend(fontsize=8, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def gerar_grafico_receita_adm_estrategia(motor, dicionario_fics, dict_taxa_adm):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    df_pl = motor.df_pl[cnpjs].loc[:motor.data_ref].fillna(0)
    taxas = pd.Series([dict_taxa_adm.get(c, 0.015) for c in cnpjs], index=cnpjs)
    
    df_rec = df_pl.multiply(taxas, axis=1) / 1e6
    
    cnpjs_rf = [c for c in cnpjs if _eh_rf(dicionario_fics[c])]
    cnpjs_mm = [c for c in cnpjs if not _eh_rf(dicionario_fics[c])]
    
    rec_rf = df_rec[cnpjs_rf].sum(axis=1) if cnpjs_rf else pd.Series(0, index=df_pl.index)
    rec_mm = df_rec[cnpjs_mm].sum(axis=1) if cnpjs_mm else pd.Series(0, index=df_pl.index)
    
    fig, ax = plt.subplots(figsize=(6, 3.5))
    ax.plot(rec_rf.index, rec_rf.values, color='#e67e22', linewidth=2.0, label='RF')
    ax.plot(rec_mm.index, rec_mm.values, color='#c0392b', linewidth=2.0, label='MM')
    ax.set_title('Receita de Adm por Estratégia - FICs', fontsize=10, fontweight='bold', color='#111111')
    ax.set_ylabel('R$ Milhões', fontsize=8)
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=8, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def _calcular_series_receita_total(motor, lista_masters, lista_fics, dicionario_masters, dicionario_fics):
    todos_fundos = []
    for cnpj, t_adm, t_pfee in lista_masters:
        nome = dicionario_masters.get(cnpj, '')
        todos_fundos.append((cnpj, nome, t_adm, t_pfee))
    for cnpj, t_adm, t_pfee in lista_fics:
        nome = dicionario_fics.get(cnpj, '')
        todos_fundos.append((cnpj, nome, t_adm, t_pfee))

    dates = motor.df_pl.loc[:motor.data_ref].index

    rec_adm_total = pd.Series(0.0, index=dates)
    rec_pfee_total = pd.Series(0.0, index=dates)
    rec_tot_rf = pd.Series(0.0, index=dates)
    rec_tot_mm = pd.Series(0.0, index=dates)
    rec_adm_rf = pd.Series(0.0, index=dates)
    rec_adm_mm = pd.Series(0.0, index=dates)

    df_pl = motor.df_pl.loc[:motor.data_ref].fillna(0)
    df_cotas = motor.df_cotas.loc[:motor.data_ref].ffill()

    for cnpj, nome, t_adm, t_pfee in todos_fundos:
        if cnpj not in df_pl.columns or cnpj not in df_cotas.columns:
            continue
        pl = df_pl[cnpj]
        cota = df_cotas[cnpj]

        rec_adm = (pl * t_adm) / 1e6

        cota_12m = cota.shift(252)
        ret_12m = ((cota / cota_12m) - 1).clip(lower=0).fillna(0)
        rec_pfee = (pl * ret_12m * t_pfee) / 1e6

        rec_tot = rec_adm + rec_pfee

        rec_adm_total += rec_adm
        rec_pfee_total += rec_pfee

        if _eh_rf(nome):
            rec_tot_rf += rec_tot
            rec_adm_rf += rec_adm
        else:
            rec_tot_mm += rec_tot
            rec_adm_mm += rec_adm

    return pd.DataFrame({
        'rec_adm_total': rec_adm_total,
        'rec_pfee_total': rec_pfee_total,
        'rec_tot_rf': rec_tot_rf,
        'rec_tot_mm': rec_tot_mm,
        'rec_adm_rf': rec_adm_rf,
        'rec_adm_mm': rec_adm_mm,
    })


def gerar_grafico_rec_total_componentes(df_rec):
    fig, ax = plt.subplots(figsize=(4.5, 3.5))
    ax.stackplot(df_rec.index, [df_rec['rec_adm_total'].values, df_rec['rec_pfee_total'].values],
                 labels=['Receita de Adm Total', 'Receita de PFee Total'],
                 colors=['#bcdff0', '#1a2439'], alpha=0.95)
    ax.set_title('Receita Total (/ano)', fontsize=9, fontweight='bold', color='#111111')
    ax.set_ylabel('R$ Milhões', fontsize=8)
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=7)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def gerar_grafico_rec_total_estrategia(df_rec):
    fig, ax = plt.subplots(figsize=(4.5, 3.5))
    ax.plot(df_rec.index, df_rec['rec_tot_rf'].values, color='#e67e22', linewidth=2.0, label='RF')
    ax.plot(df_rec.index, df_rec['rec_tot_mm'].values, color='#c0392b', linewidth=2.0, label='MM')
    ax.set_title('Receita Total por Estratégia (/ano)', fontsize=9, fontweight='bold', color='#111111')
    ax.set_ylabel('R$ Milhões', fontsize=8)
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=7)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def gerar_grafico_rec_adm_estrategia_total(df_rec):
    fig, ax = plt.subplots(figsize=(4.5, 3.5))
    ax.plot(df_rec.index, df_rec['rec_adm_rf'].values, color='#e67e22', linewidth=2.0, label='RF')
    ax.plot(df_rec.index, df_rec['rec_adm_mm'].values, color='#c0392b', linewidth=2.0, label='MM')
    ax.set_title('Receita de Adm por Estratégia (/ano)', fontsize=9, fontweight='bold', color='#111111')
    ax.set_ylabel('R$ Milhões', fontsize=8)
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=7)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def gerar_grafico_dist_rec_total_componentes(df_rec):
    tot = (df_rec['rec_adm_total'] + df_rec['rec_pfee_total']).replace(0, np.nan)
    pct_adm = (df_rec['rec_adm_total'] / tot).fillna(0)
    pct_pfee = (df_rec['rec_pfee_total'] / tot).fillna(0)

    fig, ax = plt.subplots(figsize=(4.5, 3.5))
    ax.stackplot(df_rec.index, [pct_adm.values, pct_pfee.values],
                 labels=['Receita de Adm Total', 'Receita de PFee Total'],
                 colors=['#bcdff0', '#1a2439'], alpha=0.95)
    ax.set_title('Distribuição de Receita Total (/ano)', fontsize=9, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    ax.set_ylim(0, 1)
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=7)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(False)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def gerar_grafico_dist_rec_total_estrategia(df_rec):
    tot = (df_rec['rec_tot_rf'] + df_rec['rec_tot_mm']).replace(0, np.nan)
    pct_mm = (df_rec['rec_tot_mm'] / tot).fillna(0)
    pct_rf = (df_rec['rec_tot_rf'] / tot).fillna(0)

    fig, ax = plt.subplots(figsize=(4.5, 3.5))
    ax.stackplot(df_rec.index, [pct_rf.values, pct_mm.values],
                 labels=['RF', 'MM'], colors=['#e67e22', '#c0392b'], alpha=0.95)
    ax.set_title('Distribuição de Receita Total por Estratégia (/ano)', fontsize=9, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    ax.set_ylim(0, 1)
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=7)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(False)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def gerar_grafico_dist_rec_adm_estrategia(df_rec):
    tot = (df_rec['rec_adm_rf'] + df_rec['rec_adm_mm']).replace(0, np.nan)
    pct_mm = (df_rec['rec_adm_mm'] / tot).fillna(0)
    pct_rf = (df_rec['rec_adm_rf'] / tot).fillna(0)

    fig, ax = plt.subplots(figsize=(4.5, 3.5))
    ax.stackplot(df_rec.index, [pct_rf.values, pct_mm.values],
                 labels=['RF', 'MM'], colors=['#e67e22', '#c0392b'], alpha=0.95)
    ax.set_title('Distribuição de Receita de Adm por Estratégia (/ano)', fontsize=9, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    ax.set_ylim(0, 1)
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=7)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(False)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def _calcular_series_pfee_taxa_efetiva(motor, lista_masters, lista_fics, dicionario_masters, dicionario_fics, dict_vol_target):
    df_rec = _calcular_series_receita_total(motor, lista_masters, lista_fics, dicionario_masters, dicionario_fics)

    dates = df_rec.index
    df_pl = motor.df_pl.loc[dates].fillna(0)
    df_cotas = motor.df_cotas.loc[dates].ffill()

    rec_pfee_rf = pd.Series(0.0, index=dates)
    rec_pfee_mm = pd.Series(0.0, index=dates)

    cnpjs_f = [c for c in dicionario_fics.keys() if c in df_pl.columns]
    cnpjs_m = [c for c in dicionario_masters.keys() if c in df_pl.columns]

    for cnpj, t_adm, t_pfee in (lista_masters + lista_fics):
        if cnpj not in df_pl.columns or cnpj not in df_cotas.columns or t_pfee == 0:
            continue
        nome = dicionario_masters.get(cnpj, dicionario_fics.get(cnpj, ''))
        pl = df_pl[cnpj]
        cota = df_cotas[cnpj]
        cota_12m = cota.shift(252)
        ret_12m = ((cota / cota_12m) - 1).clip(lower=0).fillna(0)
        pfee = (pl * ret_12m * t_pfee) / 1e6

        if _eh_rf(nome):
            rec_pfee_rf += pfee
        else:
            rec_pfee_mm += pfee

    vol_f = pd.Series([dict_vol_target.get(c, 0.06) for c in cnpjs_f], index=cnpjs_f)
    vol_m = pd.Series([dict_vol_target.get(c, 0.06) for c in cnpjs_m], index=cnpjs_m)

    pl_aj_f = (df_pl[cnpjs_f].multiply(vol_f, axis=1) / 0.06).sum(axis=1) / 1e6 if cnpjs_f else pd.Series(0, index=dates)
    pl_aj_m = (df_pl[cnpjs_m].multiply(vol_m, axis=1) / 0.06).sum(axis=1) / 1e6 if cnpjs_m else pd.Series(0, index=dates)
    pl_aj_tot = (pl_aj_f + pl_aj_m).replace(0, np.nan)

    rec_tot_total = df_rec['rec_adm_total'] + df_rec['rec_pfee_total']
    taxa_efetiva_total = (rec_tot_total / pl_aj_tot).fillna(0)

    df_rec['rec_pfee_rf'] = rec_pfee_rf
    df_rec['rec_pfee_mm'] = rec_pfee_mm
    df_rec['taxa_efetiva_total'] = taxa_efetiva_total

    return df_rec


def gerar_grafico_rec_pfee_estrategia(df_pfee):
    fig, ax = plt.subplots(figsize=(5.5, 3.5))
    ax.plot(df_pfee.index, df_pfee['rec_pfee_rf'].values, color='#e67e22', linewidth=2.0, label='RF')
    ax.plot(df_pfee.index, df_pfee['rec_pfee_mm'].values, color='#c0392b', linewidth=2.0, label='MM')
    ax.set_title('Receita de PFee por Estratégia (/ano)', fontsize=9, fontweight='bold', color='#111111')
    ax.set_ylabel('R$ Milhões', fontsize=8)
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=7)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def gerar_grafico_taxa_efetiva_total(df_pfee):
    fig, ax = plt.subplots(figsize=(5.5, 3.5))
    ax.plot(df_pfee.index, df_pfee['taxa_efetiva_total'].values, color='#111111', linewidth=2.0)
    ax.set_title('Receita Total (/ano) / PL Ajustado = taxa efetiva', fontsize=9, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f"{y*100:.2f}%".replace('.', ',')))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=7)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def gerar_grafico_dist_pfee_estrategia(df_pfee):
    tot = (df_pfee['rec_pfee_rf'] + df_pfee['rec_pfee_mm']).replace(0, np.nan)
    pct_mm = (df_pfee['rec_pfee_mm'] / tot).fillna(0)
    pct_rf = (df_pfee['rec_pfee_rf'] / tot).fillna(0)

    fig, ax = plt.subplots(figsize=(5.5, 3.5))
    ax.stackplot(df_pfee.index, [pct_rf.values, pct_mm.values],
                 labels=['RF', 'MM'], colors=['#e67e22', '#c0392b'], alpha=0.95)
    ax.set_title('Distribuição de Receita de PFee por Estratégia (/ano)', fontsize=9, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    ax.set_ylim(0, 1)
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=7)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(False)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


def gerar_grafico_dist_taxa_efetiva_estrategia(df_pfee):
    tot = (df_pfee['rec_tot_rf'] + df_pfee['rec_tot_mm']).replace(0, np.nan)
    pct_mm = (df_pfee['rec_tot_mm'] / tot).fillna(0)
    pct_rf = (df_pfee['rec_tot_rf'] / tot).fillna(0)

    fig, ax = plt.subplots(figsize=(5.5, 3.5))
    ax.stackplot(df_pfee.index, [pct_rf.values, pct_mm.values],
                 labels=['RF', 'MM'], colors=['#e67e22', '#c0392b'], alpha=0.95)
    ax.set_title('Receita Total (/ano) / PL Ajustado por Estratégia', fontsize=9, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    ax.set_ylim(0, 1)
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        import matplotlib.dates as mdates
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_pt))
    ax.tick_params(axis='x', rotation=45, labelsize=7)
    ax.tick_params(axis='y', labelsize=8)
    
    ax.yaxis.grid(False)
    ax.xaxis.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"




if __name__ == "__main__":
    main()





