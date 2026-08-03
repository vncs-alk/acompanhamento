import sys
import os
import pandas as pd
from bs4 import BeautifulSoup

sys.path.append(r"c:\Users\vncsa\OneDrive\Área de Trabalho\ItauAsset\Andamento\Acompanhamento")
from src.data.engine import MotorDados
from src.data.extractors import ReadExcel, ReadData
from src.components.tables import (
    gerar_tabela_master, 
    gerar_tabela_fic, gerar_tabela_gd, gerar_tabela_retornos,
    gerar_tabela_captacao, gerar_tabela_volatilidade
)

from src.components.charts import (
    gerar_grafico_retorno, gerar_grafico_sharpe_rolling, gerar_grafico_vol_rolling,
    gerar_grafico_drawdown, gerar_grafico_pl_crescimento, gerar_grafico_pl_total_ajustado,
    gerar_grafico_pl_fics_linhas, gerar_grafico_pl_fics_area,
    calcular_metricas_fics, calcular_capacity_bloco, gerar_grafico_representatividade,
    gerar_grafico_distribuicao_canal, gerar_grafico_pl_ajustado_estrategia, gerar_grafico_distribuicao_prev,
    gerar_grafico_taxa_media_fics, gerar_grafico_distribuicao_receita_adm, gerar_grafico_receita_adm_estrategia,
    _calcular_series_receita_total, gerar_grafico_rec_total_componentes, gerar_grafico_rec_total_estrategia,
    gerar_grafico_rec_adm_estrategia_total, gerar_grafico_dist_rec_total_componentes, gerar_grafico_dist_rec_total_estrategia,
    gerar_grafico_dist_rec_adm_estrategia, _calcular_series_pfee_taxa_efetiva, gerar_grafico_rec_pfee_estrategia,
    gerar_grafico_taxa_efetiva_total, gerar_grafico_dist_pfee_estrategia,
    gerar_grafico_dist_taxa_efetiva_estrategia
)

def build(nome_aba_arg="Optimus", data_b3="2026-06-30"):
    excel = ReadExcel(r'base_acompanhamento.xlsm', aba=nome_aba_arg)
    excel.nome_estrategia
    lista_masters, lista_gd, lista_fics, dicionario_prateleira, dicionario_masters, dicionario_fics, dict_vol_target = excel.extrair_listas()
    dict_taxa_adm = {t[0]: t[1] for t in lista_fics}

    data_referencia = data_b3
    cnpjs_m = [t[0] for t in lista_masters]
    cnpjs_g = [t[0] for t in lista_gd]
    cnpjs_f = [t[0] for t in lista_fics]
    
    dados = ReadData(r'dados_planilha.csv', cnpjs_m, cnpjs_g, cnpjs_f)
    motor = MotorDados(dados, data_referencia, lista_masters, lista_fics, lista_gd)
    data_ref_str = motor.data_ref.strftime('%d/%m/%Y')
    
    cnpjs_prat = list(dicionario_prateleira.keys())
    nomes_prat = list(dicionario_prateleira.values())

    modelos_dir = r"c:\Users\vncsa\OneDrive\Área de Trabalho\ItauAsset\Andamento\Acompanhamento\modelos"

    paginas_html = []
    
    def get_soup(i):
        with open(os.path.join(modelos_dir, f'relatorio_pagina{i}.html'), 'r', encoding='utf-8') as f:
            return BeautifulSoup(f, 'html.parser')

    def replace_table(soup, idx, html_str, extract_rows=False):
        tables = soup.find_all('table')
        if idx < len(tables):
            if extract_rows:
                s2 = BeautifulSoup(html_str, 'html.parser')
                trs = s2.find_all('tr')
                tbody = tables[idx].find('tbody')
                if tbody:
                    tbody.clear()
                    for tr in trs:
                        tbody.append(tr)
            else:
                s2 = BeautifulSoup(html_str, 'html.parser')
                new_table = s2.find('table')
                style_tag = s2.find('style')
                if new_table:
                    old_classes = tables[idx].get('class', [])
                    if old_classes:
                        new_table['class'] = old_classes
                    
                    if style_tag:
                        tables[idx].insert_before(style_tag)
                        
                    tables[idx].replace_with(new_table)

    def replace_chart(soup, idx, html_str):
        boxes = soup.find_all('div', class_='chart-box')
        if idx < len(boxes):
            boxes[idx].clear()
            boxes[idx].append(BeautifulSoup(html_str, 'html.parser'))

    # PAGE 1
    s1 = get_soup(1)
    replace_table(s1, 0, gerar_tabela_master(motor))
    replace_table(s1, 1, gerar_tabela_fic(motor))
    replace_table(s1, 2, gerar_tabela_gd(motor))
    replace_table(s1, 3, gerar_tabela_retornos(motor, dicionario_prateleira))
    replace_chart(s1, 0, gerar_grafico_retorno(motor, cnpjs_prat[0], nomes_prat[0]) if len(cnpjs_prat) >= 1 else "")
    replace_chart(s1, 1, gerar_grafico_retorno(motor, cnpjs_prat[1], nomes_prat[1]) if len(cnpjs_prat) >= 2 else "")
    paginas_html.append(s1)

    # PAGE 2
    s2 = get_soup(2)
    replace_table(s2, 0, gerar_tabela_captacao(motor, dicionario_prateleira))
    paginas_html.append(s2)

    # PAGE 3
    s3 = get_soup(3)
    replace_table(s3, 0, gerar_tabela_volatilidade(motor, dicionario_prateleira))
    replace_chart(s3, 0, gerar_grafico_sharpe_rolling(motor, dicionario_prateleira, window=252, show_legend=False))
    replace_chart(s3, 1, gerar_grafico_vol_rolling(motor, dicionario_prateleira, window=63, titulo_janela="3m", show_legend=True, legend_loc='center left', legend_bbox=(1.05, 0.5)))
    replace_chart(s3, 2, gerar_grafico_vol_rolling(motor, dicionario_prateleira, window=252, titulo_janela="12m", show_legend=False))
    replace_chart(s3, 3, gerar_grafico_drawdown(motor, cnpjs_prat[0], nomes_prat[0]) if len(cnpjs_prat) >= 1 else "")
    replace_chart(s3, 4, gerar_grafico_drawdown(motor, cnpjs_prat[1], nomes_prat[1]) if len(cnpjs_prat) >= 2 else "")
    paginas_html.append(s3)

    # PAGE 4
    s4 = get_soup(4)
    df_masters = motor.calcular_pl_ajustado(list(dicionario_masters.keys()), dict_vol_target=dict_vol_target, nome_coluna_pl=data_ref_str).rename(index=dicionario_masters)
    # Tabela 1: Corpo (sem o total)
    def formata_br(x): return "{:,.0f}".format(x).replace(',', '.') if pd.notnull(x) else "-"
    def formata_pct(x): return "{:.0%}".format(x) if pd.notnull(x) else "-"
    
    t1_rows = ""
    for i, (idx, row) in enumerate(df_masters.iterrows()):
        bg = "#fafafa" if i % 2 == 0 else "#ffffff"
        td_left = f"text-align: left; padding: 6px 8px; border-bottom: 1px solid #ddd; background-color: {bg}; color: #333; font-size: 12px;"
        td_center = f"text-align: center; padding: 6px 8px; border-bottom: 1px solid #ddd; background-color: {bg}; color: #333; font-size: 12px;"
        t1_rows += f"<tr><td style='{td_left}'>{idx}</td><td style='{td_center}'>{formata_br(row[data_ref_str])}</td><td style='{td_center}'>{formata_pct(row['Vol Target'])}</td><td style='{td_center}'>{formata_br(row['PL ajustado (R$MM)'])}</td></tr>"
    replace_table(s4, 0, t1_rows, extract_rows=True)
    
    total_pl = df_masters[data_ref_str].sum()
    total_pl_aj = df_masters['PL ajustado (R$MM)'].sum()
    td_tot = "padding: 8px; text-align: center; font-weight: bold; font-size: 14px; border: 1px solid #333; background-color: #ffffff; color: #111;"
    t2_rows = f"<tr><td style='{td_tot}'>{formata_br(total_pl)}</td><td style='{td_tot}'>{formata_br(total_pl_aj)}</td></tr>"
    replace_table(s4, 1, t2_rows, extract_rows=True)

    replace_chart(s4, 0, gerar_grafico_pl_crescimento(motor, dicionario_masters, dicionario_fics))
    replace_chart(s4, 1, gerar_grafico_pl_total_ajustado(motor, dicionario_masters, dict_vol_target))
    paginas_html.append(s4)

    # PAGE 5
    s5 = get_soup(5)
    # Tabela de FICs (parte 1: 0 a 26)
    df_fics = motor.calcular_pl_ajustado(list(dicionario_fics.keys()), nome_coluna_pl=data_ref_str).rename(index=dicionario_fics)
    t1_rows = ""
    for i, (idx, row) in enumerate(df_fics.iloc[0:26].iterrows()):
        bg = "#fafafa" if i % 2 == 0 else "#ffffff"
        td_left = f"text-align: left; padding: 6px 8px; border-bottom: 1px solid #ddd; background-color: {bg}; color: #333; font-size: 12px;"
        td_center = f"text-align: center; padding: 6px 8px; border-bottom: 1px solid #ddd; background-color: {bg}; color: #333; font-size: 12px;"
        t1_rows += f"<tr><td style='{td_left}'>{idx}</td><td style='{td_center}'>{formata_br(row[data_ref_str])}</td><td style='{td_center}'>{formata_pct(row['Vol Target'])}</td><td style='{td_center}'>{formata_br(row['PL ajustado (R$MM)'])}</td></tr>"
    replace_table(s5, 0, t1_rows, extract_rows=True)
    replace_chart(s5, 0, gerar_grafico_pl_fics_linhas(motor, dicionario_fics))
    replace_chart(s5, 1, gerar_grafico_pl_fics_area(motor, dicionario_fics))
    paginas_html.append(s5)

    # PAGE 6
    s6 = get_soup(6)
    t1_rows = ""
    for i, (idx, row) in enumerate(df_fics.iloc[26:].iterrows()):
        bg = "#fafafa" if i % 2 == 0 else "#ffffff"
        td_left = f"text-align: left; padding: 6px 8px; border-bottom: 1px solid #ddd; background-color: {bg}; color: #333; font-size: 12px;"
        td_center = f"text-align: center; padding: 6px 8px; border-bottom: 1px solid #ddd; background-color: {bg}; color: #333; font-size: 12px;"
        t1_rows += f"<tr><td style='{td_left}'>{idx}</td><td style='{td_center}'>{formata_br(row[data_ref_str])}</td><td style='{td_center}'>{formata_pct(row['Vol Target'])}</td><td style='{td_center}'>{formata_br(row['PL ajustado (R$MM)'])}</td></tr>"
    replace_table(s6, 0, t1_rows, extract_rows=True)
    
    total_pl = df_fics[data_ref_str].sum()
    total_pl_aj = df_fics['PL ajustado (R$MM)'].sum()
    td_tot = "padding: 8px; text-align: center; font-weight: bold; font-size: 14px; border: 1px solid #333; background-color: #ffffff; color: #111;"
    t2_rows = f"<tr><td style='{td_tot}'>{formata_br(total_pl)}</td><td style='{td_tot}'>{formata_br(total_pl_aj)}</td></tr>"
    replace_table(s6, 1, t2_rows, extract_rows=True)

    metricas = calcular_metricas_fics(motor, dicionario_fics, dicionario_masters, dict_vol_target)
    th_s = "background-color: #333; color: white; padding: 6px; text-align: center; border: 1px solid #444;"
    td_s = "padding: 6px; text-align: center; border: 1px solid #ddd; background-color: #fff; font-weight: bold;"
    t3_rows = f"<tr><th style='{th_s}'>FICs/Total</th><th style='{th_s}'>FICs/Total Ajustado</th></tr><tr><td style='{td_s}'>{metricas['fics_total']}</td><td style='{td_s}'>{metricas['fics_aj_total']}</td></tr>"
    replace_table(s6, 2, t3_rows, extract_rows=True)

    t4_rows = f"<tr><th colspan='2' style='{th_s}'>PL RF (R$MM)</th><th colspan='2' style='{th_s}'>PL RF Ajustado (R$MM)</th></tr>"
    t4_rows += f"<tr><td style='{td_s}'>{metricas['pl_rf']}</td><td style='{td_s}'>{metricas['pct_rf']}</td><td style='{td_s}'>{metricas['pl_aj_rf']}</td><td style='{td_s}'>{metricas['pct_aj_rf']}</td></tr>"
    t4_rows += f"<tr><th colspan='2' style='{th_s}'>PL MM (R$MM)</th><th colspan='2' style='{th_s}'>PL MM Ajustado (R$MM)</th></tr>"
    t4_rows += f"<tr><td style='{td_s}'>{metricas['pl_mm']}</td><td style='{td_s}'>{metricas['pct_mm']}</td><td style='{td_s}'>{metricas['pl_aj_mm']}</td><td style='{td_s}'>{metricas['pct_aj_mm']}</td></tr>"
    replace_table(s6, 3, t4_rows, extract_rows=True)

    capacity, remanescente, pct_cap = calcular_capacity_bloco(motor, dicionario_fics, dicionario_masters, dict_vol_target)
    th_l = "background-color: #555; color: white; padding: 6px; text-align: left; border: 1px solid #444;"
    t5_rows = f"<tr><td style='{th_l}'>Capacity (R$MM)</td><td style='{td_s}'>{capacity}</td></tr>"
    t5_rows += f"<tr><td style='{th_l}'>Remanescente (R$MM)</td><td style='{td_s}'>{remanescente}</td></tr>"
    t5_rows += f"<tr><td style='{th_l}'>% Capacity usado</td><td style='{td_s}'>{pct_cap}</td></tr>"
    replace_table(s6, 4, t5_rows, extract_rows=True)

    replace_chart(s6, 0, gerar_grafico_representatividade(motor, dicionario_fics))
    paginas_html.append(s6)

    # PAGE 7
    s7 = get_soup(7)
    replace_chart(s7, 0, gerar_grafico_distribuicao_canal(motor, dicionario_fics))
    replace_chart(s7, 1, gerar_grafico_pl_ajustado_estrategia(motor, dicionario_fics, dicionario_masters, dict_vol_target))
    replace_chart(s7, 2, gerar_grafico_distribuicao_prev(motor, dicionario_fics))
    replace_chart(s7, 1, gerar_grafico_pl_ajustado_estrategia(motor, dicionario_fics, dicionario_masters, dict_vol_target))
    replace_chart(s7, 2, gerar_grafico_distribuicao_prev(motor, dicionario_fics))
    paginas_html.append(s7)

    # PAGE 8
    s8 = get_soup(8)
    replace_chart(s8, 0, gerar_grafico_taxa_media_fics(motor, dicionario_fics, dict_taxa_adm))
    replace_chart(s8, 1, gerar_grafico_distribuicao_receita_adm(motor, dicionario_fics, dict_taxa_adm))
    replace_chart(s8, 2, gerar_grafico_receita_adm_estrategia(motor, dicionario_fics, dict_taxa_adm))
    paginas_html.append(s8)

    # PAGE 9
    s9 = get_soup(9)
    df_rec = _calcular_series_receita_total(motor, lista_masters, lista_fics, dicionario_masters, dicionario_fics)
    replace_chart(s9, 0, gerar_grafico_rec_total_componentes(df_rec))
    replace_chart(s9, 1, gerar_grafico_rec_total_estrategia(df_rec))
    replace_chart(s9, 2, gerar_grafico_rec_adm_estrategia_total(df_rec))
    replace_chart(s9, 3, gerar_grafico_dist_rec_total_componentes(df_rec))
    replace_chart(s9, 4, gerar_grafico_dist_rec_total_estrategia(df_rec))
    replace_chart(s9, 5, gerar_grafico_dist_rec_adm_estrategia(df_rec))
    paginas_html.append(s9)

    # PAGE 10
    s10 = get_soup(10)
    df_pfee = _calcular_series_pfee_taxa_efetiva(motor, lista_masters, lista_fics, dicionario_masters, dicionario_fics, dict_vol_target)
    replace_chart(s10, 0, gerar_grafico_rec_pfee_estrategia(df_pfee))
    replace_chart(s10, 1, gerar_grafico_taxa_efetiva_total(df_pfee))
    replace_chart(s10, 2, gerar_grafico_dist_pfee_estrategia(df_pfee))
    replace_chart(s10, 3, gerar_grafico_dist_taxa_efetiva_estrategia(df_pfee))
    paginas_html.append(s10)

    # CONSOLIDAR EM UM UNICO HTML usando unified_style.css
    with open('unified_style.css', 'r', encoding='utf-8') as f:
        unified_style = f.read()
        
    master_html = f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Relatório Completo Acompanhamento</title>
{unified_style}
</head>
<body style="background: #e9ecef; margin: 0; padding: 20px; font-family: Arial, sans-serif;">
'''
    for idx, p in enumerate(paginas_html):
        # We assume p.body has inner elements, we wrap them in a page-section div
        inner_content = p.body.decode_contents()
        master_html += f'<!-- ===== PÁGINA {idx+1} ===== -->\n'
        master_html += f'<div class="page-section" id="pagina-{idx+1}">\n'
        master_html += f'{inner_content}\n'
        master_html += '</div>\n\n'
    
    master_html += '</body></html>'

    with open(r'c:\Users\vncsa\OneDrive\Área de Trabalho\ItauAsset\Andamento\Acompanhamento\Relatorio_Final.html', 'w', encoding='utf-8') as f:
        f.write(master_html)

if __name__ == "__main__":
    build()
