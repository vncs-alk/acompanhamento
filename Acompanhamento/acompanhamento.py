import sys
import os
import io
import base64
import shutil
import tempfile
import warnings

import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
import openpyxl

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.cm as cm

# ==========================================
# CONSTANTES E FUNÇÕES AUXILIARES
# ==========================================
def fmt_num(v):
    return "{:,.0f}".format(v).replace(',', '.') if pd.notnull(v) else "-"

def fmt_pct(v):
    return "{:.1%}".format(v).replace('.', ',') if pd.notnull(v) else "-"

CORES_PRATELEIRA = ['#c0392b', '#e8825a', '#7b1e1e', '#3b6fb5', '#a11a1a', '#f4a6b8',
                    '#2ecc71', '#9b59b6', '#f39c12', '#1abc9c', '#e74c3c', '#3498db']

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

def _eh_rf(nome):
    nome_up = nome.upper()
    return any(k in nome_up for k in ['RENDA FIXA', ' RF ', 'RF LP', 'INFRA', 'RF\x00'])

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


# ==========================================
# CLASSES DE DADOS E MOTOR
# ==========================================
class ReadData():
    def __init__(self, caminho_csv, cnpjs_master, cnpjs_gd, cnpjs_fics):
        self.caminho_csv = caminho_csv
        self.cnpjs_master = cnpjs_master
        self.cnpjs_gd = cnpjs_gd
        self.cnpjs_fics = cnpjs_fics
        
        self.cnpjs_todos = self.cnpjs_master + self.cnpjs_gd + self.cnpjs_fics
        self.carregar_e_filtrar()

    def carregar_e_filtrar(self):
        df = pd.read_csv(self.caminho_csv, index_col=0)
        df['DT_COMPTC'] = pd.to_datetime(df['DT_COMPTC'])
        
        if self.cnpjs_todos:
            df = df[df['CNPJ_FUNDO_CLASSE'].isin(self.cnpjs_todos)]
            
        df['CAPTC_LIQ'] = (df['CAPTC_DIA'] - df['RESG_DIA']) / 1000000
        df['PL_MM'] = df['VL_PATRIM_LIQ'] / 1000000
        
        self.dados = df


class ReadExcel():
    def __init__(self, caminho_excel="base_acompanhamento.xlsm", aba="Optimus"):
        with tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False) as tmp:
            shutil.copyfile(caminho_excel, tmp.name)
            caminho_leitura = tmp.name
        
        self.dados = pd.read_excel(caminho_leitura, sheet_name=aba, header=None)
        self.nome_estrategia = aba
        
        masters = self.dados[self.dados[0] == "Masters"].index[0]
        gd = self.dados[self.dados[0] == "GD"].index[0]
        fics = self.dados[self.dados[0] == "Fics"].index[0]

        self.infos_masters = self.dados.iloc[masters + 1:gd].dropna(how='all').dropna(axis=1, how='all')
        if len(self.infos_masters) > 0: self.infos_masters.columns = self.dados.iloc[masters].dropna()

        self.infos_gd = self.dados.iloc[gd + 1:fics].dropna(how='all').dropna(axis=1, how='all')
        if len(self.infos_gd) > 0: self.infos_gd.columns = self.dados.iloc[gd].dropna()

        self.infos_fics = self.dados.iloc[fics + 1:].dropna(how='all').dropna(axis=1, how='all')
        if len(self.infos_fics) > 0: self.infos_fics.columns = self.dados.iloc[fics].dropna()

        self.cnpjs_prateleira_azul = []
        wb = openpyxl.load_workbook(caminho_leitura, read_only=True, data_only=True)
        if 'Aba_Auxiliar' in wb.sheetnames:
            ws_aux = wb['Aba_Auxiliar']
            for row in ws_aux.iter_rows(min_col=1, max_col=1, values_only=True):
                if row[0]:
                    self.cnpjs_prateleira_azul.append(str(row[0]).strip())
        wb.close()

    def extrair_listas(self):
        dicionario_geral = {}
        dict_vol_target = {}

        def processar_df(df_fundo):
            lista = []
            dic = {}
            if df_fundo is None or df_fundo.empty: return lista, dic
            
            for idx, row in df_fundo.iterrows():
                if pd.isna(row.iloc[0]): continue
                cnpj = str(row.iloc[0]).strip()
                nome = str(row.iloc[1]).strip() if len(row) > 1 and pd.notnull(row.iloc[1]) else ""
                vol_target = float(row.iloc[2]) if len(row) > 2 and pd.notnull(row.iloc[2]) else 0.0
                taxa_adm = float(row.iloc[3]) if len(row) > 3 and pd.notnull(row.iloc[3]) else 0.02
                taxa_perf = float(row.iloc[4]) if len(row) > 4 and pd.notnull(row.iloc[4]) else 0.2
                
                lista.append((cnpj, taxa_adm, taxa_perf))
                dic[cnpj] = nome
                dict_vol_target[cnpj] = vol_target
                dicionario_geral[cnpj] = nome
            return lista, dic

        lista_masters, dicionario_masters = processar_df(self.infos_masters)
        lista_gd, _ = processar_df(self.infos_gd)
        lista_fics, dicionario_fics = processar_df(self.infos_fics)
        
        if self.cnpjs_prateleira_azul:
            dicionario_prateleira = {cnpj: dicionario_geral.get(cnpj, cnpj) for cnpj in self.cnpjs_prateleira_azul}
        else:
            dicionario_prateleira = dicionario_geral
            
        return lista_masters, lista_gd, lista_fics, dicionario_prateleira, dicionario_masters, dicionario_fics, dict_vol_target


class MotorDados:
    def __init__(self, obj_dados, data_referencia, cnpjs_master, cnpjs_fic, cnpjs_gd):
        self.df_base = obj_dados.dados
        self.data_ref = pd.to_datetime(data_referencia)
        
        self.cnpjs_master_tup = cnpjs_master
        self.cnpjs_fic_tup = cnpjs_fic
        self.cnpjs_gd_tup = cnpjs_gd
        
        self.cnpjs_master = [t[0] for t in cnpjs_master]
        self.cnpjs_fic = [t[0] for t in cnpjs_fic]
        self.cnpjs_gd = [t[0] for t in cnpjs_gd]
        
        self._pivotar_dados()

    def _obter_cota_retroativa(self, cnpjs, meses):
        data_alvo = self.data_ref - pd.DateOffset(months=meses)
        idx = self.df_cotas.index[self.df_cotas.index <= data_alvo].max()
        if pd.notnull(idx) and set(cnpjs).issubset(self.df_cotas.columns):
            return self.df_cotas[cnpjs].loc[idx]
        return pd.Series(np.nan, index=cnpjs)

    def calcular_retornos(self, cnpjs):
        idx_ref = self.data_ref
        
        idx_dia = self.df_cotas.index[self.df_cotas.index < idx_ref].max()
        
        mes_ant = idx_ref.replace(day=1) - pd.Timedelta(days=1)
        idx_mes_ant = self.df_cotas.index[self.df_cotas.index <= mes_ant].max()
        
        idx_3m = self.df_cotas.index[self.df_cotas.index <= idx_ref - pd.DateOffset(months=3)].max()
        idx_6m = self.df_cotas.index[self.df_cotas.index <= idx_ref - pd.DateOffset(months=6)].max()
        idx_12m = self.df_cotas.index[self.df_cotas.index <= idx_ref - pd.DateOffset(months=12)].max()
        idx_24m = self.df_cotas.index[self.df_cotas.index <= idx_ref - pd.DateOffset(months=24)].max()
        
        ano_ant = pd.Timestamp(year=idx_ref.year - 1, month=12, day=31)
        idx_ano_ant = self.df_cotas.index[self.df_cotas.index <= ano_ant].max()

        def _get_cota(idx):
            return self.df_cotas[cnpjs].loc[idx] if pd.notnull(idx) and set(cnpjs).issubset(self.df_cotas.columns) else pd.Series(np.nan, index=cnpjs)

        cota_atual = _get_cota(idx_ref)
        cota_inicio = self.df_cotas[cnpjs].apply(lambda x: x.dropna().iloc[0] if not x.dropna().empty else np.nan)

        df_resumo = pd.DataFrame({
            'Dia': (cota_atual / _get_cota(idx_dia)) - 1,
            'Mês atual': (cota_atual / _get_cota(idx_mes_ant)) - 1,
            '3m': (cota_atual / _get_cota(idx_3m)) - 1,
            '6m': (cota_atual / _get_cota(idx_6m)) - 1,
            'Ano': (cota_atual / _get_cota(idx_ano_ant)) - 1,
            '12m': (cota_atual / _get_cota(idx_12m)) - 1,
            '24m': (cota_atual / _get_cota(idx_24m)) - 1,
            'Início': (cota_atual / cota_inicio) - 1
        })
        
        return df_resumo

    def calcular_receita(self, lista_tuplas, tipo):
        idx_ref = self.df_pl.index[self.df_pl.index <= self.data_ref].max()
        data_12m = self.data_ref - pd.DateOffset(months=12)
        idx_12m = self.df_pl.index[self.df_pl.index <= data_12m].max()
        data_24m = self.data_ref - pd.DateOffset(months=24)
        idx_24m = self.df_pl.index[self.df_pl.index <= data_24m].max()

        receita_atual_total = 0
        receita_ant_total = 0

        for cnpj, taxa_adm, taxa_pfee in lista_tuplas:
            if cnpj in self.df_pl.columns and cnpj in self.df_cotas.columns:
                pl_atual = self.df_pl[cnpj].loc[idx_ref] if pd.notnull(idx_ref) else 0
                pl_12m = self.df_pl[cnpj].loc[idx_12m] if pd.notnull(idx_12m) else 0
                
                cota_atual = self.df_cotas[cnpj].loc[idx_ref] if pd.notnull(idx_ref) else np.nan
                cota_12m = self.df_cotas[cnpj].loc[idx_12m] if pd.notnull(idx_12m) else np.nan
                cota_24m = self.df_cotas[cnpj].loc[idx_24m] if pd.notnull(idx_24m) else np.nan

                retorno_atual = (cota_atual / cota_12m) - 1 if pd.notnull(cota_atual) and pd.notnull(cota_12m) else 0
                retorno_ant = (cota_12m / cota_24m) - 1 if pd.notnull(cota_12m) and pd.notnull(cota_24m) else 0

                rec_adm_atual = pl_atual * taxa_adm
                rec_adm_ant = pl_12m * taxa_adm
                
                rec_pfee_atual = pl_atual * max(0, retorno_atual) * taxa_pfee
                rec_pfee_ant = pl_12m * max(0, retorno_ant) * taxa_pfee

                if tipo == 'master':
                    receita_atual_total += (rec_adm_atual + rec_pfee_atual)
                    receita_ant_total += (rec_adm_ant + rec_pfee_ant)
                elif tipo == 'fic':
                    receita_atual_total += rec_adm_atual
                    receita_ant_total += rec_adm_ant

        var_receita = (receita_atual_total / receita_ant_total) - 1 if receita_ant_total != 0 else np.nan
        return receita_atual_total, var_receita

    def calcular_capliq(self, cnpjs, meses):
        data_inicio = self.data_ref - pd.DateOffset(months=meses)
        data_ant_inicio = data_inicio - pd.DateOffset(months=meses)
        data_ant_fim = data_inicio - pd.Timedelta(days=1)
        
        df_f_local = self.df_fluxo[cnpjs] if set(cnpjs).issubset(self.df_fluxo.columns) else pd.DataFrame()
        if df_f_local.empty:
            return 0, np.nan
            
        capt_atual = df_f_local.loc[data_inicio:self.data_ref].sum().sum()
        capt_ant = df_f_local.loc[data_ant_inicio:data_ant_fim].sum().sum()
        
        var_capt = (capt_atual / capt_ant) - 1 if capt_ant != 0 else np.nan
        return capt_atual, var_capt

    def calcular_captacao_janelas(self, cnpjs):
        idx_ref = self.data_ref
        
        janelas = {
            'Dia': idx_ref,
            'Mês atual': pd.Timestamp(year=idx_ref.year, month=idx_ref.month, day=1),
            '3m': idx_ref - pd.DateOffset(months=3) + pd.Timedelta(days=1),
            '6m': idx_ref - pd.DateOffset(months=6) + pd.Timedelta(days=1),
            'Ano': pd.Timestamp(year=idx_ref.year, month=1, day=1),
            '12m': idx_ref - pd.DateOffset(months=12) + pd.Timedelta(days=1),
            '24m': idx_ref - pd.DateOffset(months=24) + pd.Timedelta(days=1)
        }
        
        primeiras_datas = self.df_cotas[cnpjs].apply(lambda x: x.dropna().index[0] if not x.dropna().empty else pd.NaT)
        
        resultados = {}
        for nome_janela, dt_inicio in janelas.items():
            serie_janela = pd.Series(index=cnpjs, dtype=float)
            for cnpj in cnpjs:
                if pd.isna(primeiras_datas[cnpj]) or primeiras_datas[cnpj] > dt_inicio:
                    serie_janela[cnpj] = np.nan
                else:
                    serie_janela[cnpj] = self.df_fluxo.loc[dt_inicio:idx_ref, cnpj].sum()
            resultados[nome_janela] = serie_janela
            
        serie_inicio = pd.Series(index=cnpjs, dtype=float)
        for cnpj in cnpjs:
            if not pd.isna(primeiras_datas[cnpj]):
                serie_inicio[cnpj] = self.df_fluxo.loc[:idx_ref, cnpj].sum()
            else:
                serie_inicio[cnpj] = np.nan
        resultados['Início'] = serie_inicio
        
        return pd.DataFrame(resultados)

    def calcular_volatilidade(self, cnpjs):
        idx_ref = self.data_ref
        
        start_mes = pd.Timestamp(year=idx_ref.year, month=idx_ref.month, day=1)
        start_6m = idx_ref - pd.DateOffset(months=6) + pd.Timedelta(days=1)
        start_12m = idx_ref - pd.DateOffset(months=12) + pd.Timedelta(days=1)
        
        df_ret_diario = self.df_cotas[cnpjs].pct_change()
        
        fator_anualizacao = np.sqrt(252)
        
        vol_mes = df_ret_diario.loc[start_mes:idx_ref].std(ddof=1) * fator_anualizacao
        vol_6m = df_ret_diario.loc[start_6m:idx_ref].std(ddof=1) * fator_anualizacao
        vol_12m = df_ret_diario.loc[start_12m:idx_ref].std(ddof=1) * fator_anualizacao
        
        return pd.DataFrame({
            'Mês': vol_mes,
            '6m': vol_6m,
            '12m': vol_12m
        })

    def calcular_sharpe(self, cnpjs, risk_free_rate=0.10):
        df_vol = self.calcular_volatilidade(cnpjs)
        
        idx_ref = self.data_ref
        start_mes = pd.Timestamp(year=idx_ref.year, month=idx_ref.month, day=1)
        start_6m = idx_ref - pd.DateOffset(months=6) + pd.Timedelta(days=1)
        start_12m = idx_ref - pd.DateOffset(months=12) + pd.Timedelta(days=1)
        
        def _retorno_anualizado(start_date):
            if start_date not in self.df_cotas.index:
                start_date = self.df_cotas.index[self.df_cotas.index >= start_date].min()
            cotas = self.df_cotas[cnpjs].loc[start_date:idx_ref]
            dias_uteis = len(cotas)
            if dias_uteis < 2: return pd.Series(0.0, index=cnpjs)
            retorno_acum = (cotas.iloc[-1] / cotas.iloc[0]) - 1
            return (1 + retorno_acum) ** (252 / dias_uteis) - 1
            
        ret_mes = _retorno_anualizado(start_mes)
        ret_6m = _retorno_anualizado(start_6m)
        ret_12m = _retorno_anualizado(start_12m)
        
        sharpe_mes = (ret_mes - risk_free_rate) / df_vol['Mês']
        sharpe_6m = (ret_6m - risk_free_rate) / df_vol['6m']
        sharpe_12m = (ret_12m - risk_free_rate) / df_vol['12m']
        
        return pd.DataFrame({'Mês': sharpe_mes, '6m': sharpe_6m, '12m': sharpe_12m})

    def calcular_drawdown(self, cnpjs):
        cotas = self.df_cotas[cnpjs].loc[:self.data_ref]
        picos = cotas.expanding(min_periods=1).max()
        drawdowns = (cotas / picos) - 1
        return drawdowns

    def calcular_receita_detalhada(self, lista_tuplas):
        idx_ref = self.df_pl.index[self.df_pl.index <= self.data_ref].max()
        data_12m = self.data_ref - pd.DateOffset(months=12)
        idx_12m = self.df_pl.index[self.df_pl.index <= data_12m].max()
        
        resultados = []
        for cnpj, taxa_adm, taxa_pfee in lista_tuplas:
            if cnpj in self.df_pl.columns and cnpj in self.df_cotas.columns:
                pl_atual = self.df_pl[cnpj].loc[idx_ref] if pd.notnull(idx_ref) else 0.0
                
                cota_atual = self.df_cotas[cnpj].loc[idx_ref] if pd.notnull(idx_ref) else np.nan
                cota_12m = self.df_cotas[cnpj].loc[idx_12m] if pd.notnull(idx_12m) else np.nan
                retorno_12m = (cota_atual / cota_12m) - 1 if pd.notnull(cota_atual) and pd.notnull(cota_12m) else 0.0
                
                rec_adm_ano = pl_atual * taxa_adm
                rec_pfee_ano = pl_atual * max(0, retorno_12m) * taxa_pfee
                rec_total_ano = rec_adm_ano + rec_pfee_ano
                taxa_efetiva = (rec_total_ano / pl_atual) if pl_atual > 0 else 0.0
                
                resultados.append({
                    'CNPJ': cnpj,
                    'Receita Adm / ano': rec_adm_ano,
                    'Receita PFee / ano': rec_pfee_ano,
                    'Receita Total / ano': rec_total_ano,
                    'Taxa Efetiva': taxa_efetiva
                })
        
        return pd.DataFrame(resultados).set_index('CNPJ')

    def calcular_pl_ajustado(self, cnpjs, dict_vol_target=None, nome_coluna_pl=None):
        if nome_coluna_pl is None:
            nome_coluna_pl = self.data_ref.strftime('%d/%m/%Y')
        
        idx_ref = self.df_pl.index[self.df_pl.index <= self.data_ref].max()
        
        cnpjs_validos = [c for c in cnpjs if c in self.df_pl.columns]
        if pd.notnull(idx_ref) and cnpjs_validos:
            pl_atual = self.df_pl[cnpjs_validos].loc[idx_ref].reindex(cnpjs, fill_value=0.0)
        else:
            pl_atual = pd.Series(0.0, index=cnpjs)

        if dict_vol_target:
            vol_target = pd.Series([dict_vol_target.get(c, 0.06) for c in cnpjs], index=cnpjs)
        else:
            cnpjs_cotas = [c for c in cnpjs if c in self.df_cotas.columns]
            if cnpjs_cotas:
                df_vol = self.calcular_volatilidade(cnpjs_cotas)
                vol_target = df_vol[['Mês', '6m', '12m']].mean(axis=1).reindex(cnpjs, fill_value=0.06)
            else:
                vol_target = pd.Series(0.06, index=cnpjs)

        pl_ajustado = pl_atual * vol_target

        df_resultado = pd.DataFrame({
            nome_coluna_pl: pl_atual,
            'Vol Target': vol_target,
            'PL ajustado (R$MM)': pl_ajustado
        })

        return df_resultado

    def _pivotar_dados(self):
        self.df_fluxo = self.df_base.pivot_table(index='DT_COMPTC', columns='CNPJ_FUNDO_CLASSE', values='CAPTC_LIQ', aggfunc='sum').fillna(0).sort_index()
        self.df_cotas = self.df_base.pivot_table(index='DT_COMPTC', columns='CNPJ_FUNDO_CLASSE', values='VL_QUOTA', aggfunc='first').ffill().sort_index()
        self.df_pl = self.df_base.pivot_table(index='DT_COMPTC', columns='CNPJ_FUNDO_CLASSE', values='PL_MM', aggfunc='first').ffill().sort_index()

        if self.data_ref not in self.df_cotas.index:
            self.data_ref = self.df_cotas.index[self.df_cotas.index <= self.data_ref].max()

        self.df_cotas = self.df_cotas.loc[:self.data_ref]
        self.df_fluxo = self.df_fluxo.loc[:self.data_ref]
        self.df_pl = self.df_pl.loc[:self.data_ref]


# ==========================================
# FUNÇÕES DE TABELAS (HTML)
# ==========================================
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
    def f_num(v): return "{:,.0f}".format(v).replace(',', '.') if pd.notnull(v) else "-"
    def f_pct(v): return "{:.0%}".format(v) if pd.notnull(v) else "-"
    valores = [f_num(rec_val), f_num(capt_val), f_num(pl_val)]
    deltas = [f_pct(rec_var), f_pct(capt_var), f_pct(pl_var)]
    return _gerar_html_bloco_resumo("Masters (R$MM)", ['Receita (adm +<br>PFee)', 'Captação<br>(12m)', 'PL'], valores, deltas)

def gerar_tabela_fic(motor):
    rec_val, rec_var = motor.calcular_receita(motor.cnpjs_fic_tup, 'fic')
    capt_val, capt_var = motor.calcular_capliq(motor.cnpjs_fic, 12)
    pl_val, pl_var = _obter_pl_dados(motor, motor.cnpjs_fic)
    def f_num(v): return "{:,.0f}".format(v).replace(',', '.') if pd.notnull(v) else "-"
    def f_pct(v): return "{:.0%}".format(v) if pd.notnull(v) else "-"
    valores = [f_num(rec_val), f_num(capt_val), f_num(pl_val)]
    deltas = [f_pct(rec_var), f_pct(capt_var), f_pct(pl_var)]
    return _gerar_html_bloco_resumo("FICs (R$MM)", ['Receita (adm)', 'Captação<br>(12m)', 'PL'], valores, deltas)

def gerar_tabela_gd(motor):
    capt_val, capt_var = motor.calcular_capliq(motor.cnpjs_gd, 12)
    def f_num(v): return "{:,.0f}".format(v).replace(',', '.') if pd.notnull(v) else "-"
    def f_pct(v): return "{:.0%}".format(v) if pd.notnull(v) else "-"
    valores = [f_num(capt_val)]
    deltas = [f_pct(capt_var)]
    return _gerar_html_bloco_resumo("GD (R$MM)", ['Captação<br>(12m)'], valores, deltas)

def _rgb_to_hex(rgb):
    return '#%02x%02x%02x' % (int(rgb[0]*255), int(rgb[1]*255), int(rgb[2]*255))

def get_cell_colors(val, min_val, max_val, cmap='RdYlGn'):
    if pd.isna(val) or val == "-": return '#ffffff', '#000000'
    norm_val = 0.5 if max_val == min_val else (val - min_val) / (max_val - min_val)
    rgba = cm.get_cmap(cmap)(norm_val)
    bg_color = _rgb_to_hex(rgba)
    luminance = 0.299*rgba[0] + 0.587*rgba[1] + 0.114*rgba[2]
    text_color = '#ffffff' if luminance < 0.5 else '#000000'
    return bg_color, text_color

def gerar_tabela_retornos(motor, dicionario_prateleira):
    cnpjs = list(dicionario_prateleira.keys())
    df_retornos = motor.calcular_retornos(cnpjs)
    df_retornos = df_retornos.rename(index=dicionario_prateleira)
    df_retornos.loc['CDI'] = 0.0
    
    df_numeric = df_retornos.iloc[:-1]
    col_mins = df_numeric.min()
    col_maxs = df_numeric.max()
    
    def fmt_br_pct(v):
        if pd.isna(v): return "-"
        return "{:.2f}%".format(v * 100).replace('.', ',')

    html = "<table class='tabela-retornos' style='border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 13px; margin-bottom: 40px;'>"
    html += "<thead><tr><th style='background-color: #212529; color: white; text-align: left; padding: 8px; border: 1px solid #fff;'>Principais FICs Prateleira</th>"
    
    for col in df_retornos.columns:
        html += f"<th style='background-color: #212529; color: white; text-align: center; padding: 8px; border: 1px solid #fff;'>{col}</th>"
    html += "</tr></thead><tbody>"
    
    for idx, row in df_retornos.iloc[:-1].iterrows():
        html += "<tr>"
        html += f"<th style='text-align: left; padding: 8px; background-color: #ffffff; border: 1px solid #fff; font-weight: bold;'>{idx}</th>"
        for col_name, val in row.items():
            min_v = col_mins[col_name]
            max_v = col_maxs[col_name]
            bg_col, txt_col = get_cell_colors(val, min_v, max_v)
            html += f"<td style='text-align: center; padding: 8px; border: 1px solid #fff; background-color: {bg_col}; color: {txt_col};'>{fmt_br_pct(val)}</td>"
        html += "</tr>"
        
    cdi_row = df_retornos.iloc[-1]
    html += "<tr>"
    html += f"<th style='text-align: left; padding: 8px; background-color: #e9ecef; border: 1px solid #fff; font-weight: bold; color: #000;'>{cdi_row.name}</th>"
    for val in cdi_row:
        html += f"<td style='text-align: center; padding: 8px; border: 1px solid #fff; background-color: #e9ecef; color: #000; font-weight: bold;'>{fmt_br_pct(val)}</td>"
    html += "</tr></tbody></table>"
    return html

def gerar_tabela_captacao(motor, dicionario_prateleira):
    cnpjs = list(dicionario_prateleira.keys())
    df_cap = motor.calcular_captacao_janelas(cnpjs)
    df_cap = df_cap.rename(index=dicionario_prateleira)
    df_cap.loc['Soma'] = df_cap.sum(numeric_only=True)
    
    df_numeric = df_cap.iloc[:-1]
    col_mins = df_numeric.min()
    col_maxs = df_numeric.max()
    
    def formata_br(x): return "-" if pd.isna(x) else "{:,.0f}".format(x).replace(',', '.')

    html = "<table class='tabela-captacao' style='border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 13px; margin-bottom: 40px;'>"
    html += "<thead><tr><th style='background-color: #212529; color: white; text-align: left; padding: 8px; border: 1px solid #fff;'>Principais FICs Prateleira</th>"
    
    for col in df_cap.columns:
        html += f"<th style='background-color: #212529; color: white; text-align: center; padding: 8px; border: 1px solid #fff;'>{col}</th>"
    html += "</tr></thead><tbody>"
    
    for idx, row in df_cap.iloc[:-1].iterrows():
        html += "<tr>"
        html += f"<th style='text-align: left; padding: 8px; background-color: #ffffff; border: 1px solid #fff; font-weight: bold;'>{idx}</th>"
        for col_name, val in row.items():
            min_v = col_mins[col_name]
            max_v = col_maxs[col_name]
            bg_col, txt_col = get_cell_colors(val, min_v, max_v)
            html += f"<td style='text-align: center; padding: 8px; border: 1px solid #fff; background-color: {bg_col}; color: {txt_col};'>{formata_br(val)}</td>"
        html += "</tr>"
        
    soma_row = df_cap.iloc[-1]
    html += "<tr>"
    html += f"<th style='text-align: left; padding: 8px; background-color: #e9ecef; border: 1px solid #fff; font-weight: bold; color: #000;'>{soma_row.name}</th>"
    for val in soma_row:
        html += f"<td style='text-align: center; padding: 8px; border: 1px solid #fff; background-color: #e9ecef; color: #000; font-weight: bold;'>{formata_br(val)}</td>"
    html += "</tr></tbody></table>"
    return html

def gerar_tabela_volatilidade(motor, dicionario_prateleira):
    cnpjs = list(dicionario_prateleira.keys())
    df_vol = motor.calcular_volatilidade(cnpjs)
    df_vol = df_vol.rename(index=dicionario_prateleira)
    
    def formata_pct(x): return "-" if pd.isna(x) else "{:.2f}%".format(x * 100).replace('.', ',')

    html = "<table class='tabela-volatilidade' style='border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 13px;'>"
    html += "<thead><tr><th style='background-color: #212529; color: white; text-align: left; padding: 6px; border: 1px solid #fff;'>Fundos</th>"
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


# ==========================================
# FUNÇÕES DE GRÁFICOS
# ==========================================
def gerar_grafico_retorno(motor, cnpj, nome_fundo):
    if cnpj not in motor.df_cotas.columns: return ""
    cotas = motor.df_cotas[cnpj].dropna()
    if cotas.empty or len(cotas) < 2: return ""

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
    for s in ['top', 'right', 'left']: ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_color('#cccccc')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f"{int(round(y*100))}%"))
    
    max_val = max(ret_acum.max(), cdi_series.max())
    max_lim = np.ceil(max_val * 20) / 20
    if max_lim < max_val + 0.01: max_lim += 0.05
    ax.set_ylim(0, max_lim)
    ax.set_yticks(np.arange(0, max_lim + 0.01, 0.05))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
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


def _grafico_multilinhas(series_dict, titulo, fmt_y='{:.0%}', figsize=(5.5, 2.8), show_legend=True, legend_loc='best', legend_bbox=None):
    fig, ax = plt.subplots(figsize=figsize)
    for i, (nome, serie) in enumerate(series_dict.items()):
        cor = CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)]
        ax.plot(serie.index, serie.values, color=cor, linewidth=1.5, label=nome)
    ax.set_title(titulo, fontsize=11, fontweight='bold', color='#111111')
    
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    for s in ['top', 'right', 'left']: ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_color('#cccccc')
    
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: fmt_y.format(y)))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    def format_date_pt(x, pos=None):
        dt = mdates.num2date(x)
        return f"{MESES_PT[dt.month]}-{str(dt.year)[-2:]}"
        
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
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"

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
    if cnpj not in motor.df_cotas.columns: return ""
    cotas = motor.df_cotas[cnpj].dropna()
    if cotas.empty or len(cotas) < 2: return ""
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
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"


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
    pl_aj_rf = sum(motor.df_pl[c].loc[idx_ref] * dict_vol_target.get(c, 0.06) / 0.06 for c in cnpjs_f if _eh_rf(dicionario_fics[c]))
    pl_aj_mm = sum(motor.df_pl[c].loc[idx_ref] * dict_vol_target.get(c, 0.06) / 0.06 for c in cnpjs_f if not _eh_rf(dicionario_fics[c]))

    fics_total = pl_fics / pl_total_est if pl_total_est else 0
    fics_aj_total = pl_aj_fics / pl_aj_total if pl_aj_total else 0
    pct_rf = pl_rf / (pl_rf + pl_mm) if (pl_rf + pl_mm) else 0
    pct_mm = pl_mm / (pl_rf + pl_mm) if (pl_rf + pl_mm) else 0
    pct_aj_rf = pl_aj_rf / (pl_aj_rf + pl_aj_mm) if (pl_aj_rf + pl_aj_mm) else 0
    pct_aj_mm = pl_aj_mm / (pl_aj_rf + pl_aj_mm) if (pl_aj_rf + pl_aj_mm) else 0

    return {
        'pl_fics_total': fmt_num(pl_fics / 1e6),
        'pl_aj_fics_total': fmt_num(pl_aj_fics / 1e6),
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
    pl_fics = motor.df_pl[cnpjs].rename(columns=dicionario_fics)
    rep = pl_fics.div(pl_fics.sum(axis=1), axis=0)

    fig, ax = plt.subplots(figsize=(6, 3.5))
    cores = [CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)] for i in range(len(rep.columns))]
    ax.stackplot(rep.index, rep.T, labels=rep.columns, colors=cores, alpha=0.9)
    ax.set_ylabel('R$ Bilhões', fontsize=8) 
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{MESES_PT[mdates.num2date(x).month]}-{str(mdates.num2date(x).year)[-2:]}"))
    ax.tick_params(axis='x', rotation=90, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    ax.yaxis.grid(False); ax.xaxis.grid(False)
    for s in ['top', 'right', 'left']: ax.spines[s].set_visible(False)
    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=3, frameon=False)
    fig.tight_layout()

    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"

def gerar_grafico_pl_fics_linhas(motor, dicionario_fics):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    pl_fics = (motor.df_pl[cnpjs] / 1000).rename(columns=dicionario_fics)

    fig, ax = plt.subplots(figsize=(6, 2.8))
    for i, col in enumerate(pl_fics.columns):
        ax.plot(pl_fics.index, pl_fics[col], color=CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)], linewidth=1.5, label=col)
    
    ax.set_ylabel('R$ Bilhões', fontsize=8)
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{MESES_PT[mdates.num2date(x).month]}-{str(mdates.num2date(x).year)[-2:]}"))
    ax.tick_params(axis='x', rotation=90, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    for s in ['top', 'right', 'left']: ax.spines[s].set_visible(False)
    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.4), ncol=3, frameon=False)
    fig.tight_layout()

    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"

def gerar_grafico_pl_fics_area(motor, dicionario_fics):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    pl_fics = (motor.df_pl[cnpjs] / 1000).rename(columns=dicionario_fics)

    fig, ax = plt.subplots(figsize=(6, 2.8))
    cores = [CORES_PRATELEIRA[i % len(CORES_PRATELEIRA)] for i in range(len(pl_fics.columns))]
    ax.stackplot(pl_fics.index, pl_fics.T, labels=pl_fics.columns, colors=cores, alpha=0.9)
    ax.set_ylabel('R$ Bilhões', fontsize=8)
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{MESES_PT[mdates.num2date(x).month]}-{str(mdates.num2date(x).year)[-2:]}"))
    ax.tick_params(axis='x', rotation=90, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    for s in ['top', 'right', 'left']: ax.spines[s].set_visible(False)
    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.4), ncol=3, frameon=False)
    fig.tight_layout()

    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"

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
    ax1.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax1.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{MESES_PT[mdates.num2date(x).month]}-{str(mdates.num2date(x).year)[-2:]}"))
    ax1.tick_params(axis='x', rotation=90, labelsize=8)

    ax2 = ax1.twinx()
    ax2.plot(ratio.index, ratio.values, color='#2c3e50', linewidth=2.0, label='FICs/Masters')
    ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    ax2.tick_params(axis='y', labelsize=8)

    ax1.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax1.xaxis.grid(False)
    for s in ['top', 'right', 'left']: 
        ax1.spines[s].set_visible(False)
        ax2.spines[s].set_visible(False)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=8, loc='lower center', bbox_to_anchor=(0.5, -0.4), ncol=3, frameon=False)
    fig.tight_layout()

    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"

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
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{MESES_PT[mdates.num2date(x).month]}-{str(mdates.num2date(x).year)[-2:]}"))
    ax.tick_params(axis='x', rotation=90, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    for s in ['top', 'right', 'left']: ax.spines[s].set_visible(False)
    ax.legend(fontsize=8, loc='lower center', bbox_to_anchor=(0.5, -0.4), ncol=2, frameon=False)
    fig.tight_layout()

    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"

def gerar_grafico_distribuicao_canal(motor, dicionario_fics):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    df_pl = motor.df_pl[cnpjs].loc[:motor.data_ref].fillna(0)
    canal_map = {c: _classificar_canal(dicionario_fics[c]) for c in cnpjs}
    df_canal = pd.DataFrame({cat: df_pl[[c for c in cnpjs if canal_map[c] == cat]].sum(axis=1) for cat in _ORDEM_CANAL})
    df_pct = df_canal.div(df_canal.sum(axis=1).replace(0, np.nan), axis=0).fillna(0)
    
    fig, ax = plt.subplots(figsize=(6, 4))
    cores = [_CORES_CANAL[cat] for cat in _ORDEM_CANAL]
    ax.stackplot(df_pct.index, [df_pct[cat].values for cat in _ORDEM_CANAL], labels=_ORDEM_CANAL, colors=cores, alpha=0.95)
    ax.set_title('Distribuição por Canal - FICs Janeiro', fontsize=10, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    ax.set_ylim(0, 1)
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{MESES_PT[mdates.num2date(x).month]}-{str(mdates.num2date(x).year)[-2:]}"))
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    ax.yaxis.grid(False); ax.xaxis.grid(False)
    for s in ['top', 'right', 'left']: ax.spines[s].set_visible(False)
    ax.legend(fontsize=8, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()

    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"

def gerar_grafico_pl_ajustado_estrategia(motor, dicionario_fics, dicionario_masters, dict_vol_target):
    cnpjs_f = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    cnpjs_m = [c for c in dicionario_masters.keys() if c in motor.df_pl.columns]
    def _pl_aj(cnpjs):
        vt = pd.Series([dict_vol_target.get(c, 0.06) for c in cnpjs], index=cnpjs)
        return (motor.df_pl[cnpjs].multiply(vt, axis=1) / 0.06).sum(axis=1) / 1e6
    
    pl_total = motor.df_pl[cnpjs_f + cnpjs_m].sum(axis=1) / 1e6
    pl_aj_tot = _pl_aj(cnpjs_f + cnpjs_m)
    pl_aj_rf = _pl_aj([c for c in cnpjs_f if _eh_rf(dicionario_fics[c])])
    pl_aj_mm = _pl_aj([c for c in cnpjs_f if not _eh_rf(dicionario_fics[c])])
    
    fig, ax = plt.subplots(figsize=(6, 3))
    ax.fill_between(pl_total.index, pl_total.values, color='#e0e0e0', alpha=0.9, label='PL Total')
    ax.plot(pl_aj_tot.index, pl_aj_tot.values, color='#2c3e50', linewidth=2.0, label='PL Ajustado Total')
    ax.plot(pl_aj_rf.index, pl_aj_rf.values, color='#e74c3c', linewidth=2.0, label='PL Ajustado RF')
    ax.plot(pl_aj_mm.index, pl_aj_mm.values, color='#c0392b', linewidth=2.0, label='PL Ajustado MM')
    ax.set_title('Evolução do PL Ajustado por Estratégia', fontsize=10, fontweight='bold', color='#111111')
    ax.set_ylabel('R$ Bilhões', fontsize=8)
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{MESES_PT[mdates.num2date(x).month]}-{str(mdates.num2date(x).year)[-2:]}"))
    ax.tick_params(axis='x', rotation=0, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    for s in ['top', 'right', 'left']: ax.spines[s].set_visible(False)
    ax.legend(fontsize=8, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=4, frameon=False)
    fig.tight_layout()

    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"

def gerar_grafico_distribuicao_prev(motor, dicionario_fics):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    df_pl = motor.df_pl[cnpjs].loc[:motor.data_ref].fillna(0)
    prev_map = {c: _classificar_prev(dicionario_fics[c]) for c in cnpjs}
    df_prev = pd.DataFrame({cat: df_pl[[c for c in cnpjs if prev_map[c] == cat]].sum(axis=1) for cat in _ORDEM_PREV})
    df_pct = df_prev.div(df_prev.sum(axis=1).replace(0, np.nan), axis=0).fillna(0)
    
    fig, ax = plt.subplots(figsize=(6, 3))
    cores = [_CORES_PREV[cat] for cat in _ORDEM_PREV]
    ax.stackplot(df_pct.index, [df_pct[cat].values for cat in _ORDEM_PREV], labels=_ORDEM_PREV, colors=cores, alpha=0.9)
    ax.set_title('Distribuição Estratégias Previdência - FICs Janeiro', fontsize=10, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    ax.set_ylim(0, 1)
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{MESES_PT[mdates.num2date(x).month]}-{str(mdates.num2date(x).year)[-2:]}"))
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    ax.yaxis.grid(False); ax.xaxis.grid(False)
    for s in ['top', 'right', 'left']: ax.spines[s].set_visible(False)
    ax.legend(fontsize=8, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=3, frameon=False)
    fig.tight_layout()

    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"

def gerar_grafico_taxa_media_fics(motor, dicionario_fics, dict_taxa_adm):
    cnpjs = [c for c in dicionario_fics.keys() if c in motor.df_pl.columns]
    df_pl = motor.df_pl[cnpjs].loc[:motor.data_ref].fillna(0)
    taxas = pd.Series([dict_taxa_adm.get(c, 0.015) for c in cnpjs], index=cnpjs)
    taxa_media = (df_pl.multiply(taxas, axis=1).sum(axis=1) / df_pl.sum(axis=1).replace(0, np.nan)).fillna(0)
    
    fig, ax = plt.subplots(figsize=(6, 3.5))
    ax.plot(taxa_media.index, taxa_media.values, color='#c0392b', linewidth=2.0)
    ax.set_title('Taxa Média FICs Janeiro', fontsize=10, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f"{y*100:.2f}%".replace('.', ',')))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{MESES_PT[mdates.num2date(x).month]}-{str(mdates.num2date(x).year)[-2:]}"))
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    for s in ['top', 'right', 'left']: ax.spines[s].set_visible(False)
    fig.tight_layout()

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
    
    df_cat = pd.DataFrame({
        'RF': df_rec[cnpjs_rf].sum(axis=1) if cnpjs_rf else pd.Series(0, index=df_pl.index),
        'MM': df_rec[cnpjs_mm].sum(axis=1) if cnpjs_mm else pd.Series(0, index=df_pl.index)
    })
    df_pct = df_cat.div(df_cat.sum(axis=1).replace(0, np.nan), axis=0).fillna(0)
    
    fig, ax = plt.subplots(figsize=(6, 3.5))
    ax.stackplot(df_pct.index, [df_pct['RF'].values, df_pct['MM'].values], labels=['RF', 'MM'], colors=['#e67e22', '#c0392b'], alpha=0.95)
    ax.set_title('Distribuição de Receita de Adm por Estratégia - FICs', fontsize=10, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
    ax.set_ylim(0, 1)
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{MESES_PT[mdates.num2date(x).month]}-{str(mdates.num2date(x).year)[-2:]}"))
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    ax.yaxis.grid(False); ax.xaxis.grid(False)
    for s in ['top', 'right', 'left']: ax.spines[s].set_visible(False)
    ax.legend(fontsize=8, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()

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
    
    fig, ax = plt.subplots(figsize=(6, 3.5))
    ax.plot(df_rec.index, df_rec[cnpjs_rf].sum(axis=1) if cnpjs_rf else pd.Series(0, index=df_pl.index), color='#e67e22', linewidth=2.0, label='RF')
    ax.plot(df_rec.index, df_rec[cnpjs_mm].sum(axis=1) if cnpjs_mm else pd.Series(0, index=df_pl.index), color='#c0392b', linewidth=2.0, label='MM')
    ax.set_title('Receita de Adm por Estratégia - FICs', fontsize=10, fontweight='bold', color='#111111')
    ax.set_ylabel('R$ Milhões', fontsize=8)
    
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{MESES_PT[mdates.num2date(x).month]}-{str(mdates.num2date(x).year)[-2:]}"))
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)
    ax.xaxis.grid(False)
    for s in ['top', 'right', 'left']: ax.spines[s].set_visible(False)
    ax.legend(fontsize=8, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()

    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"

def _calcular_series_receita_total(motor, lista_masters, lista_fics, dicionario_masters, dicionario_fics):
    todos_fundos = []
    for cnpj, t_adm, t_pfee in lista_masters:
        todos_fundos.append((cnpj, dicionario_masters.get(cnpj, ''), t_adm, t_pfee))
    for cnpj, t_adm, t_pfee in lista_fics:
        todos_fundos.append((cnpj, dicionario_fics.get(cnpj, ''), t_adm, t_pfee))

    dates = motor.df_pl.loc[:motor.data_ref].index
    r = {k: pd.Series(0.0, index=dates) for k in ['rec_adm_total', 'rec_pfee_total', 'rec_tot_rf', 'rec_tot_mm', 'rec_adm_rf', 'rec_adm_mm']}
    
    df_pl = motor.df_pl.loc[:motor.data_ref].fillna(0)
    df_cotas = motor.df_cotas.loc[:motor.data_ref].ffill()

    for cnpj, nome, t_adm, t_pfee in todos_fundos:
        if cnpj not in df_pl.columns or cnpj not in df_cotas.columns: continue
        pl = df_pl[cnpj]
        cota = df_cotas[cnpj]
        rec_adm = (pl * t_adm) / 1e6
        rec_pfee = (pl * ((cota / cota.shift(252)) - 1).clip(lower=0).fillna(0) * t_pfee) / 1e6
        rec_tot = rec_adm + rec_pfee

        r['rec_adm_total'] += rec_adm
        r['rec_pfee_total'] += rec_pfee
        if _eh_rf(nome):
            r['rec_tot_rf'] += rec_tot
            r['rec_adm_rf'] += rec_adm
        else:
            r['rec_tot_mm'] += rec_tot
            r['rec_adm_mm'] += rec_adm
    return pd.DataFrame(r)

def _gerar_grafico_basico(df, x_col, y_cols, labels, colors, title, ylabel, chart_type='plot'):
    fig, ax = plt.subplots(figsize=(4.5, 3.5))
    if chart_type == 'stackplot':
        ax.stackplot(df.index, [df[c].values for c in y_cols], labels=labels, colors=colors, alpha=0.95)
        ax.set_ylim(0, 1)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
        ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
        ax.yaxis.grid(False)
    else:
        for c, l, col in zip(y_cols, labels, colors): ax.plot(df.index, df[c].values, color=col, linewidth=2.0, label=l)
        ax.set_ylabel(ylabel, fontsize=8)
        ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0)

    ax.set_title(title, fontsize=9, fontweight='bold', color='#111111')
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{MESES_PT[mdates.num2date(x).month]}-{str(mdates.num2date(x).year)[-2:]}"))
    ax.tick_params(axis='x', rotation=45, labelsize=7)
    ax.tick_params(axis='y', labelsize=8)
    ax.xaxis.grid(False)
    for s in ['top', 'right', 'left']: ax.spines[s].set_visible(False)
    ax.legend(fontsize=7, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=2, frameon=False)
    fig.tight_layout()
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"

def gerar_grafico_rec_total_componentes(df_rec):
    return _gerar_grafico_basico(df_rec, None, ['rec_adm_total', 'rec_pfee_total'], ['Receita de Adm Total', 'Receita de PFee Total'], ['#bcdff0', '#1a2439'], 'Receita Total (/ano)', 'R$ Milhões', 'stackplot_custom')

def gerar_grafico_rec_total_estrategia(df_rec):
    return _gerar_grafico_basico(df_rec, None, ['rec_tot_rf', 'rec_tot_mm'], ['RF', 'MM'], ['#e67e22', '#c0392b'], 'Receita Total por Estratégia (/ano)', 'R$ Milhões', 'plot')

def gerar_grafico_rec_adm_estrategia_total(df_rec):
    return _gerar_grafico_basico(df_rec, None, ['rec_adm_rf', 'rec_adm_mm'], ['RF', 'MM'], ['#e67e22', '#c0392b'], 'Receita de Adm por Estratégia (/ano)', 'R$ Milhões', 'plot')

def gerar_grafico_dist_rec_total_componentes(df_rec):
    tot = (df_rec['rec_adm_total'] + df_rec['rec_pfee_total']).replace(0, np.nan)
    df_pct = pd.DataFrame({'adm': (df_rec['rec_adm_total'] / tot).fillna(0), 'pfee': (df_rec['rec_pfee_total'] / tot).fillna(0)})
    return _gerar_grafico_basico(df_pct, None, ['adm', 'pfee'], ['Receita de Adm Total', 'Receita de PFee Total'], ['#bcdff0', '#1a2439'], 'Distribuição de Receita Total (/ano)', '', 'stackplot')

def gerar_grafico_dist_rec_total_estrategia(df_rec):
    tot = (df_rec['rec_tot_rf'] + df_rec['rec_tot_mm']).replace(0, np.nan)
    df_pct = pd.DataFrame({'rf': (df_rec['rec_tot_rf'] / tot).fillna(0), 'mm': (df_rec['rec_tot_mm'] / tot).fillna(0)})
    return _gerar_grafico_basico(df_pct, None, ['rf', 'mm'], ['RF', 'MM'], ['#e67e22', '#c0392b'], 'Distribuição de Receita Total por Estratégia (/ano)', '', 'stackplot')

def gerar_grafico_dist_rec_adm_estrategia(df_rec):
    tot = (df_rec['rec_adm_rf'] + df_rec['rec_adm_mm']).replace(0, np.nan)
    df_pct = pd.DataFrame({'rf': (df_rec['rec_adm_rf'] / tot).fillna(0), 'mm': (df_rec['rec_adm_mm'] / tot).fillna(0)})
    return _gerar_grafico_basico(df_pct, None, ['rf', 'mm'], ['RF', 'MM'], ['#e67e22', '#c0392b'], 'Distribuição de Receita de Adm por Estratégia (/ano)', '', 'stackplot')

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
        if cnpj not in df_pl.columns or cnpj not in df_cotas.columns or t_pfee == 0: continue
        nome = dicionario_masters.get(cnpj, dicionario_fics.get(cnpj, ''))
        pfee = (df_pl[cnpj] * ((df_cotas[cnpj] / df_cotas[cnpj].shift(252)) - 1).clip(lower=0).fillna(0) * t_pfee) / 1e6
        if _eh_rf(nome): rec_pfee_rf += pfee
        else: rec_pfee_mm += pfee

    pl_aj_f = (df_pl[cnpjs_f].multiply(pd.Series([dict_vol_target.get(c, 0.06) for c in cnpjs_f], index=cnpjs_f), axis=1) / 0.06).sum(axis=1) / 1e6 if cnpjs_f else pd.Series(0, index=dates)
    pl_aj_m = (df_pl[cnpjs_m].multiply(pd.Series([dict_vol_target.get(c, 0.06) for c in cnpjs_m], index=cnpjs_m), axis=1) / 0.06).sum(axis=1) / 1e6 if cnpjs_m else pd.Series(0, index=dates)
    
    df_rec['rec_pfee_rf'] = rec_pfee_rf
    df_rec['rec_pfee_mm'] = rec_pfee_mm
    df_rec['taxa_efetiva_total'] = ((df_rec['rec_adm_total'] + df_rec['rec_pfee_total']) / (pl_aj_f + pl_aj_m).replace(0, np.nan)).fillna(0)
    return df_rec

def gerar_grafico_rec_pfee_estrategia(df_pfee):
    return _gerar_grafico_basico(df_pfee, None, ['rec_pfee_rf', 'rec_pfee_mm'], ['RF', 'MM'], ['#e67e22', '#c0392b'], 'Receita de PFee por Estratégia (/ano)', 'R$ Milhões', 'plot')

def gerar_grafico_taxa_efetiva_total(df_pfee):
    fig, ax = plt.subplots(figsize=(5.5, 3.5))
    ax.plot(df_pfee.index, df_pfee['taxa_efetiva_total'].values, color='#111111', linewidth=2.0)
    ax.set_title('Receita Total (/ano) / PL Ajustado = taxa efetiva', fontsize=9, fontweight='bold', color='#111111')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f"{y*100:.2f}%".replace('.', ',')))
    MESES_PT = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"{MESES_PT[mdates.num2date(x).month]}-{str(mdates.num2date(x).year)[-2:]}"))
    ax.tick_params(axis='x', rotation=45, labelsize=7)
    ax.tick_params(axis='y', labelsize=8)
    ax.yaxis.grid(True, linestyle='-', color='#e0e0e0', zorder=0); ax.xaxis.grid(False)
    for s in ['top', 'right', 'left']: ax.spines[s].set_visible(False)
    fig.tight_layout()
    buffer = io.BytesIO()
    fig.savefig(buffer, format='png', dpi=120, transparent=True, bbox_inches='tight')
    plt.close(fig)
    return f"<img src='data:image/png;base64,{base64.b64encode(buffer.getvalue()).decode()}' style='max-width:100%; height:auto;' />"

def gerar_grafico_dist_pfee_estrategia(df_pfee):
    tot = (df_pfee['rec_pfee_rf'] + df_pfee['rec_pfee_mm']).replace(0, np.nan)
    df_pct = pd.DataFrame({'rf': (df_pfee['rec_pfee_rf'] / tot).fillna(0), 'mm': (df_pfee['rec_pfee_mm'] / tot).fillna(0)})
    return _gerar_grafico_basico(df_pct, None, ['rf', 'mm'], ['RF', 'MM'], ['#e67e22', '#c0392b'], 'Distribuição de Receita de PFee por Estratégia (/ano)', '', 'stackplot')

def gerar_grafico_dist_taxa_efetiva_estrategia(df_pfee):
    tot = (df_pfee['rec_tot_rf'] + df_pfee['rec_tot_mm']).replace(0, np.nan)
    df_pct = pd.DataFrame({'rf': (df_pfee['rec_tot_rf'] / tot).fillna(0), 'mm': (df_pfee['rec_tot_mm'] / tot).fillna(0)})
    return _gerar_grafico_basico(df_pct, None, ['rf', 'mm'], ['RF', 'MM'], ['#e67e22', '#c0392b'], 'Receita Total (/ano) / PL Ajustado por Estratégia', '', 'stackplot')


# ==========================================
# FUNÇÃO PRINCIPAL BUILD
# ==========================================
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
        inner_content = p.body.decode_contents()
        master_html += f'<!-- ===== PÁGINA {idx+1} ===== -->\n'
        master_html += f'<div class="page-section" id="pagina-{idx+1}">\n'
        master_html += f'{inner_content}\n'
        master_html += '</div>\n\n'
    
    master_html += '</body></html>'

    with open(r'c:\Users\vncsa\OneDrive\Área de Trabalho\ItauAsset\Andamento\Acompanhamento\Relatorio_Final_2.html', 'w', encoding='utf-8') as f:
        f.write(master_html)

if __name__ == "__main__":
    build()