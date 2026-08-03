import pandas as pd
import numpy as np

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


