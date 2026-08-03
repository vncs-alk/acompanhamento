import pandas as pd
import numpy as np

def _gerar_html_bloco_resumo(titulo, colunas, valores, deltas):
    html = f"<table style='border-collapse: collapse; text-align: center; font-family: Arial, sans-serif; font-size: 13px; margin: 0 auto; min-width: 250px;'>"
    html += f"<thead><tr><th colspan='{len(colunas)}' style='border-bottom: 2px solid #000; padding: 4px 10px; font-size: 15px;'>{titulo}</th></tr><tr>"
    
    for col in colunas:
        html += f"<th style='padding: 8px 10px; font-weight: bold; white-space: normal; vertical-align: bottom;'>{col}</th>"
        
    html += "</tr></thead><tbody><tr style='font-size: 14px;'>"
    
    for val in valores:
        html += f"<td style='padding: 4px 10px;'>{val}</td>"
        
    html += "</tr><tr>"
    
    for _ in colunas:
        html += "<td style='padding: 2px 10px; font-weight: bold; font-size: 13px;'>&Delta; 12m</td>"
        
    html += "</tr><tr style='font-size: 14px;'>"
    
    for delta in deltas:
        html += f"<td style='padding: 4px 10px; padding-bottom: 12px;'>{delta}</td>"
        
    html += "</tr></tbody></table>"
    return html

def _obter_pl_dados(motor, cnpjs):
    idx_ref = motor.df_pl.index[motor.df_pl.index <= motor.data_ref].max()
    data_12m = motor.data_ref - pd.DateOffset(months=12)
    idx_12m = motor.df_pl.index[motor.df_pl.index <= data_12m].max()
    
    pl_atual = motor.df_pl[cnpjs].loc[idx_ref].sum() if pd.notnull(idx_ref) and set(cnpjs).issubset(motor.df_pl.columns) else 0
    pl_12m = motor.df_pl[cnpjs].loc[idx_12m].sum() if pd.notnull(idx_12m) and set(cnpjs).issubset(motor.df_pl.columns) else 0
    
    var_pl = (pl_atual / pl_12m) - 1 if pl_12m != 0 else np.nan
    return pl_atual, var_pl


def gerar_tabela_master(motor):
    rec_val, rec_var = motor.calcular_receita(motor.cnpjs_master_tup, 'master')
    capt_val, capt_var = motor.calcular_capliq(motor.cnpjs_master, 12)
    pl_val, pl_var = _obter_pl_dados(motor, motor.cnpjs_master)
    
    def fmt_num(v): return "{:,.0f}".format(v).replace(',', '.') if pd.notnull(v) else "-"
    def fmt_pct(v): return "{:.0%}".format(v) if pd.notnull(v) else "-"
    
    valores = [fmt_num(rec_val), fmt_num(capt_val), fmt_num(pl_val)]
    deltas = [fmt_pct(rec_var), fmt_pct(capt_var), fmt_pct(pl_var)]
    
    return _gerar_html_bloco_resumo(
        "Masters (R$MM)",
        ['Receita (adm +<br>PFee)', 'Captação<br>(12m)', 'PL'],
        valores,
        deltas
    )


def gerar_tabela_fic(motor):
    rec_val, rec_var = motor.calcular_receita(motor.cnpjs_fic_tup, 'fic')
    capt_val, capt_var = motor.calcular_capliq(motor.cnpjs_fic, 12)
    pl_val, pl_var = _obter_pl_dados(motor, motor.cnpjs_fic)
    
    def fmt_num(v): return "{:,.0f}".format(v).replace(',', '.') if pd.notnull(v) else "-"
    def fmt_pct(v): return "{:.0%}".format(v) if pd.notnull(v) else "-"
    
    valores = [fmt_num(rec_val), fmt_num(capt_val), fmt_num(pl_val)]
    deltas = [fmt_pct(rec_var), fmt_pct(capt_var), fmt_pct(pl_var)]
    
    return _gerar_html_bloco_resumo(
        "FICs (R$MM)",
        ['Receita (adm)', 'Captação<br>(12m)', 'PL'],
        valores,
        deltas
    )


def gerar_tabela_gd(motor):
    capt_val, capt_var = motor.calcular_capliq(motor.cnpjs_gd, 12)
    
    def fmt_num(v): return "{:,.0f}".format(v).replace(',', '.') if pd.notnull(v) else "-"
    def fmt_pct(v): return "{:.0%}".format(v) if pd.notnull(v) else "-"
    
    valores = [fmt_num(capt_val)]
    deltas = [fmt_pct(capt_var)]
    
    return _gerar_html_bloco_resumo(
        "GD (R$MM)",
        ['Captação<br>(12m)'],
        valores,
        deltas
    )


import matplotlib.cm as cm
import numpy as np

def _rgb_to_hex(rgb):
    return '#%02x%02x%02x' % (int(rgb[0]*255), int(rgb[1]*255), int(rgb[2]*255))

def get_cell_colors(val, min_val, max_val, cmap='RdYlGn'):
    if pd.isna(val) or val == "-":
        return '#ffffff', '#000000'
    
    if max_val == min_val:
        norm_val = 0.5
    else:
        norm_val = (val - min_val) / (max_val - min_val)
        
    rgba = cm.get_cmap(cmap)(norm_val)
    bg_color = _rgb_to_hex(rgba)
    
    luminance = 0.299*rgba[0] + 0.587*rgba[1] + 0.114*rgba[2]
    text_color = '#ffffff' if luminance < 0.5 else '#000000'
    
    return bg_color, text_color

def gerar_tabela_retornos(motor, dicionario_prateleira):
    cnpjs = list(dicionario_prateleira.keys())
    df_retornos = motor.calcular_retornos(cnpjs)
    df_retornos = df_retornos.rename(index=dicionario_prateleira)
    
    # CDI fixo (0) por enquanto, a ser ajustado quando vierem os dados reais
    df_retornos.loc['CDI'] = 0.0
    
    # Calculando min e max por coluna (ignorando a ultima linha do CDI e NaNs)
    df_numeric = df_retornos.iloc[:-1]
    col_mins = df_numeric.min()
    col_maxs = df_numeric.max()
    
    def fmt_br_pct(v):
        if pd.isna(v): return "-"
        return "{:.2f}%".format(v * 100).replace('.', ',')

    html = "<table class='tabela-retornos' style='border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 13px; margin-bottom: 40px;'>"
    
    # Cabeçalho da tabela
    html += "<thead><tr>"
    html += "<th style='background-color: #212529; color: white; text-align: left; padding: 8px; border: 1px solid #fff;'>Principais FICs Prateleira</th>"
    
    for col in df_retornos.columns:
        html += f"<th style='background-color: #212529; color: white; text-align: center; padding: 8px; border: 1px solid #fff;'>{col}</th>"
    html += "</tr></thead>"
    
    html += "<tbody>"
    
    # Linhas dos fundos
    for idx, row in df_retornos.iloc[:-1].iterrows():
        html += "<tr>"
        html += f"<th style='text-align: left; padding: 8px; background-color: #ffffff; border: 1px solid #fff; font-weight: bold;'>{idx}</th>"
        for col_name, val in row.items():
            min_v = col_mins[col_name]
            max_v = col_maxs[col_name]
            bg_col, txt_col = get_cell_colors(val, min_v, max_v)
            val_str = fmt_br_pct(val)
            html += f"<td style='text-align: center; padding: 8px; border: 1px solid #fff; background-color: {bg_col}; color: {txt_col};'>{val_str}</td>"
        html += "</tr>"
        
    # Linha do CDI
    cdi_row = df_retornos.iloc[-1]
    html += "<tr>"
    html += f"<th style='text-align: left; padding: 8px; background-color: #e9ecef; border: 1px solid #fff; font-weight: bold; color: #000;'>{cdi_row.name}</th>"
    for val in cdi_row:
        val_str = fmt_br_pct(val)
        html += f"<td style='text-align: center; padding: 8px; border: 1px solid #fff; background-color: #e9ecef; color: #000; font-weight: bold;'>{val_str}</td>"
    html += "</tr>"
    
    html += "</tbody></table>"
    
    return html


def gerar_tabela_captacao(motor, dicionario_prateleira):
    cnpjs = list(dicionario_prateleira.keys())
    df_cap = motor.calcular_captacao_janelas(cnpjs)
    df_cap = df_cap.rename(index=dicionario_prateleira)
    
    df_cap.loc['Soma'] = df_cap.sum(numeric_only=True)
    
    # Calculando min e max por coluna (ignorando a ultima linha Soma e NaNs)
    df_numeric = df_cap.iloc[:-1]
    col_mins = df_numeric.min()
    col_maxs = df_numeric.max()
    
    def formata_br(x):
        if pd.isna(x):
            return "-"
        return "{:,.0f}".format(x).replace(',', '.')

    html = "<table class='tabela-captacao' style='border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 13px; margin-bottom: 40px;'>"
    
    # Cabeçalho da tabela
    html += "<thead><tr>"
    html += "<th style='background-color: #212529; color: white; text-align: left; padding: 8px; border: 1px solid #fff;'>Principais FICs Prateleira</th>"
    
    for col in df_cap.columns:
        html += f"<th style='background-color: #212529; color: white; text-align: center; padding: 8px; border: 1px solid #fff;'>{col}</th>"
    html += "</tr></thead>"
    
    html += "<tbody>"
    
    # Linhas dos fundos
    for idx, row in df_cap.iloc[:-1].iterrows():
        html += "<tr>"
        html += f"<th style='text-align: left; padding: 8px; background-color: #ffffff; border: 1px solid #fff; font-weight: bold;'>{idx}</th>"
        for col_name, val in row.items():
            min_v = col_mins[col_name]
            max_v = col_maxs[col_name]
            bg_col, txt_col = get_cell_colors(val, min_v, max_v)
            val_str = formata_br(val)
            html += f"<td style='text-align: center; padding: 8px; border: 1px solid #fff; background-color: {bg_col}; color: {txt_col};'>{val_str}</td>"
        html += "</tr>"
        
    # Linha da Soma
    soma_row = df_cap.iloc[-1]
    html += "<tr>"
    html += f"<th style='text-align: left; padding: 8px; background-color: #e9ecef; border: 1px solid #fff; font-weight: bold; color: #000;'>{soma_row.name}</th>"
    for val in soma_row:
        val_str = formata_br(val)
        html += f"<td style='text-align: center; padding: 8px; border: 1px solid #fff; background-color: #e9ecef; color: #000; font-weight: bold;'>{val_str}</td>"
    html += "</tr>"
    
    html += "</tbody></table>"
    
    return html

def gerar_tabela_volatilidade(motor, dicionario_prateleira):
    cnpjs = list(dicionario_prateleira.keys())
    df_vol = motor.calcular_volatilidade(cnpjs)
    df_vol = df_vol.rename(index=dicionario_prateleira)
    
    def formata_pct(x):
        if pd.isna(x): return "-"
        return "{:.2f}%".format(x * 100).replace('.', ',')

    html = "<table class='tabela-volatilidade' style='border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 13px;'>"
    html += "<thead><tr>"
    html += "<th style='background-color: #212529; color: white; text-align: left; padding: 6px; border: 1px solid #fff;'>Fundos</th>"
    for col in df_vol.columns:
        html += f"<th style='background-color: #212529; color: white; text-align: center; padding: 6px; border: 1px solid #fff;'>{col}</th>"
    html += "</tr></thead><tbody>"
    
    for idx, row in df_vol.iterrows():
        html += "<tr>"
        html += f"<th style='text-align: left; padding: 6px; border-bottom: 1px solid #eee; background-color: #ffffff; color: #333;'>{idx}</th>"
        for val in row:
            html += f"<td style='text-align: center; padding: 6px; border-bottom: 1px solid #eee; background-color: #ffffff; color: #333;'>{formata_pct(val)}</td>"
        html += "</tr>"
        
    html += "</tbody></table>"
    return html



