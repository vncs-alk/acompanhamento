SET NOCOUNT ON;
SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;

DECLARE @DataRef date = ?;
DECLARE @CodigoFundo varchar(100) = ?;
DECLARE @CarteiraJson nvarchar(max) = ?;


------------------------------------------------------------
-- 1. CARTEIRA
------------------------------------------------------------

IF OBJECT_ID('tempdb..#carteira') IS NOT NULL
    DROP TABLE #carteira;

CREATE TABLE #carteira
(
    isin varchar(50) NOT NULL,
    peso decimal(28,12) NOT NULL,
    CONSTRAINT PK_Carteira PRIMARY KEY CLUSTERED (isin)
);


WITH CarteiraBruta AS
(
    SELECT
        UPPER(LTRIM(RTRIM(j.isin))) AS isin,

        TRY_CONVERT(
            decimal(28,12),
            TRY_CONVERT(
                float,
                REPLACE(
                    REPLACE(
                        NULLIF(LTRIM(RTRIM(j.peso)), ''),
                        '%',
                        ''
                    ),
                    ',',
                    '.'
                )
            )
        ) AS peso

    FROM OPENJSON(@CarteiraJson)
    WITH
    (
        isin varchar(50) '$.isin',
        peso nvarchar(100) '$.peso'
    ) AS j

    WHERE NULLIF(LTRIM(RTRIM(j.isin)), '') IS NOT NULL
)

INSERT INTO #carteira
(
    isin,
    peso
)
SELECT
    isin,
    SUM(peso)
FROM CarteiraBruta
WHERE peso IS NOT NULL
GROUP BY isin;


------------------------------------------------------------
-- 2. ANBIMA
------------------------------------------------------------

WITH AnbimaReferencia AS
(
    SELECT
        DataReferencia AS [date],

        MAX(CASE
            WHEN Indice = 'IMA-B 5+'
            THEN [Yield]
        END) AS yield_ima_b5_mais,

        MAX(CASE
            WHEN Indice = 'IMA-B 5'
            THEN [Yield]
        END) AS yield_ima_b5,

        MAX(CASE
            WHEN Indice = 'IMA-B'
            THEN [Yield]
        END) AS yield_ima_b

    FROM ANBIMA_IMA_GERAL WITH (NOLOCK)

    WHERE DataReferencia = @DataRef
      AND Indice IN
      (
          'IMA-B 5+',
          'IMA-B 5',
          'IMA-B'
      )

    GROUP BY DataReferencia
),


------------------------------------------------------------
-- 3. CADASTRO DOS ATIVOS
--
-- Agora busca somente:
-- ISIN + DATA EXATA
------------------------------------------------------------

CadastraAtivos AS
(
    SELECT
        C.isin,
        C.peso,

        CAST(X.[date] AS date) AS data_informacao,
        CONVERT(varchar(200), X.codigo_IAM_ativo) AS codigo_IAM_ativo,

        X.ativo,
        X.tipo,
        X.tipo_ativo_exposicao,
        X.fator_risco,
        X.fator_risco_mercado,
        X.setor_iam,
        X.grupo_economico_iam,
        X.indexador,
        X.RATING_RANK_EQ_L_BR,
        X.modified_duration,
        X.spread_cdi,
        X.spread_ipca,
        X.spread_ntnb,
        X.spread_equivalente

    FROM #carteira AS C

    LEFT JOIN VW_FUNDOS_POSICAO_EXPLODIDA_DOWN AS X WITH (NOLOCK)
        ON X.isin = C.isin
       AND X.[date] = @DataRef
),


------------------------------------------------------------
-- 4. BASE DE POSIÇÕES
------------------------------------------------------------

BasePosicoes AS
(
    SELECT
        @DataRef AS [date],
        @CodigoFundo AS codigo_IAM_fundo,

        C.isin,
        C.codigo_IAM_ativo,

        COALESCE(
            NULLIF(LTRIM(RTRIM(C.ativo)), ''),
            C.isin
        ) AS ativo,

        C.tipo,
        C.tipo_ativo_exposicao,
        C.fator_risco,
        C.fator_risco_mercado,

        COALESCE(
            NULLIF(LTRIM(RTRIM(C.setor_iam)), ''),
            'Sem Setor'
        ) AS setor,

        COALESCE(
            NULLIF(LTRIM(RTRIM(C.grupo_economico_iam)), ''),
            'Sem Grupo'
        ) AS grupo_economico,

        CASE
            WHEN NULLIF(LTRIM(RTRIM(C.indexador)), '') IS NULL
                THEN 'Sem Indexador'

            WHEN UPPER(C.indexador) LIKE '%CDI%'
                THEN 'CDI'

            ELSE LTRIM(RTRIM(C.indexador))
        END AS indexador,


        ----------------------------------------------------
        -- RATING
        ----------------------------------------------------

        C.RATING_RANK_EQ_L_BR,

        CASE C.RATING_RANK_EQ_L_BR
            WHEN 'AAA'  THEN 22
            WHEN 'AA+'  THEN 21
            WHEN 'AA'   THEN 20
            WHEN 'AA-'  THEN 19
            WHEN 'A+'   THEN 18
            WHEN 'A'    THEN 17
            WHEN 'A-'   THEN 16
            WHEN 'BBB+' THEN 15
            WHEN 'BBB'  THEN 14
            WHEN 'BBB-' THEN 13
            WHEN 'BB+'  THEN 12
            WHEN 'BB'   THEN 11
            WHEN 'BB-'  THEN 10
            WHEN 'B+'   THEN 9
            WHEN 'B'    THEN 8
            WHEN 'B-'   THEN 7
            WHEN 'CCC+' THEN 6
            WHEN 'CCC'  THEN 5
            WHEN 'CCC-' THEN 4
            WHEN 'CC'   THEN 3
            WHEN 'C'    THEN 2
            WHEN 'RD'   THEN 1
            WHEN 'D'    THEN 0
            ELSE NULL
        END AS rating_externo,


        CASE C.RATING_RANK_EQ_L_BR
            WHEN 'AAA'  THEN 1
            WHEN 'AA+'  THEN 2
            WHEN 'AA'   THEN 3
            WHEN 'AA-'  THEN 4
            WHEN 'A+'   THEN 5
            WHEN 'A'    THEN 6
            WHEN 'A-'   THEN 7
            WHEN 'BBB+' THEN 8
            WHEN 'BBB'  THEN 9
            WHEN 'BBB-' THEN 10
            WHEN 'BB+'  THEN 11
            WHEN 'BB'   THEN 12
            WHEN 'BB-'  THEN 13
            WHEN 'B+'   THEN 14
            WHEN 'B'    THEN 15
            WHEN 'B-'   THEN 16
            WHEN 'CCC+' THEN 17
            WHEN 'CCC'  THEN 18
            WHEN 'CCC-' THEN 19
            WHEN 'CC'   THEN 20
            WHEN 'C'    THEN 21
            ELSE 99
        END AS ordem_rating,


        ----------------------------------------------------
        -- CLASSIFICAÇÃO DO INSTRUMENTO
        ----------------------------------------------------

        CASE

            WHEN C.tipo IN
            (
                'CaixaCPR',
                'Tx. Adm',
                'Despesa',
                'CaixaCPROutros',
                'Tx. Pfee'
            )
                THEN 'Caixa e CPR'

            WHEN C.fator_risco = 'Credito'
             AND C.ativo IN ('Fund', 'FIDC')
                THEN 'FIDC'

            WHEN C.fator_risco = 'Credito'
             AND C.ativo LIKE '%CDB%'
                THEN 'CDB'

            WHEN C.fator_risco = 'Credito'
             AND C.ativo LIKE '%LFSN%'
                THEN 'LF Subordinada Nível II'

            WHEN C.fator_risco = 'Credito'
             AND C.ativo LIKE '%LFSS%'
                THEN 'LF Subordinada'

            WHEN C.fator_risco = 'Credito'
             AND C.ativo LIKE '%LFX%'
                THEN 'LF'

            WHEN C.fator_risco = 'Credito'
             AND C.ativo LIKE '%DEBINC%'
                THEN 'Debenture Incentivada'

            WHEN C.fator_risco = 'Credito'
             AND C.ativo LIKE '%DEB%'
                THEN 'Debenture'

            WHEN UPPER(LTRIM(RTRIM(C.tipo))) = 'FUTURE'
                THEN
                    'Futuro'
                    + ISNULL(
                        CONVERT(varchar(200), C.tipo_ativo_exposicao),
                        ''
                    )

            ELSE
                COALESCE(
                    NULLIF(LTRIM(RTRIM(C.ativo)), ''),
                    C.isin,
                    'Nao classificado'
                )

        END AS instrumento,


        ----------------------------------------------------
        -- PESOS
        ----------------------------------------------------

        C.peso AS perc_pl,
        C.peso AS perc_nav,


        ----------------------------------------------------
        -- DURATION / SPREADS
        ----------------------------------------------------

        TRY_CONVERT(
            decimal(28,12),
            C.modified_duration
        ) AS modified_duration,

        TRY_CONVERT(
            decimal(28,12),
            C.spread_cdi
        ) AS spread_cdi_ativo,

        TRY_CONVERT(
            decimal(28,12),
            C.spread_ipca
        ) AS spread_ipca_ativo,

        TRY_CONVERT(
            decimal(28,12),
            C.spread_ntnb
        ) AS spread_ntnb_ativo,

        TRY_CONVERT(
            decimal(28,12),
            C.spread_equivalente
        ) AS spread_equivalente_ativo,


        ----------------------------------------------------
        -- DEFASAGEM
        ----------------------------------------------------

        DATEDIFF(
            day,
            C.data_informacao,
            @DataRef
        ) AS defasagem_dias,


        CASE
            WHEN C.codigo_IAM_ativo IS NULL THEN 0
            ELSE 1
        END AS isin_encontrado_iam,


        ----------------------------------------------------
        -- IMA-B
        ----------------------------------------------------

        AR.yield_ima_b5_mais,
        AR.yield_ima_b5,
        AR.yield_ima_b

    FROM CadastraAtivos AS C

    LEFT JOIN AnbimaReferencia AS AR
        ON AR.[date] = @DataRef
),


------------------------------------------------------------
-- 5. INDICADORES BASE
------------------------------------------------------------

IndicadoresBase AS
(
    SELECT
        [date],
        codigo_IAM_fundo,


        ----------------------------------------------------
        -- EXPOSIÇÃO DE CRÉDITO
        ----------------------------------------------------

        SUM(
            CASE
                WHEN fator_risco = 'Credito'
                    THEN ISNULL(perc_pl, 0)
                ELSE 0
            END
        ) AS perc_pl_credito,


        SUM(
            CASE
                WHEN fator_risco = 'Credito'
                 AND tipo = 'Fixed Income'
                    THEN ISNULL(perc_pl, 0)
                ELSE 0
            END
        ) AS perc_pl_credito_sem_cotas,


        SUM(
            CASE
                WHEN tipo IN
                (
                    'CaixaCPR',
                    'CaixaCPROutros',
                    'Despesa',
                    'Fixed Income'
                )
                AND fator_risco = 'Credito'
                AND ISNULL(ativo, '') <> 'Bonds'
                AND spread_cdi_ativo IS NOT NULL
                    THEN ISNULL(perc_pl, 0)
                ELSE 0
            END
        ) AS perc_pl_credito_tem_spread,


        SUM(
            CASE
                WHEN tipo IN
                (
                    'CaixaCPR',
                    'CaixaCPROutros',
                    'Despesa',
                    'Fixed Income'
                )
                AND ISNULL(ativo, '') <> 'Bonds'
                    THEN ISNULL(perc_pl, 0)
                ELSE 0
            END
        ) AS perc_pl_considerado,


        ----------------------------------------------------
        -- DEBIN
        ----------------------------------------------------

        SUM(
            CASE
                WHEN ativo = 'DEBIN'
                  OR instrumento = 'Debenture Incentivada'
                    THEN ISNULL(perc_pl, 0)
                ELSE 0
            END
        ) AS perc_debin,


        ----------------------------------------------------
        -- MODIFIED DURATION
        ----------------------------------------------------

        SUM(
            CASE
                WHEN fator_risco = 'Credito'
                    THEN
                        ISNULL(perc_pl, 0)
                        * ISNULL(modified_duration, 0)
                        / 252.0
                ELSE 0
            END
        ) AS moddur_credito_ponderada,


        SUM(
            CASE
                WHEN indexador NOT IN ('CDI', 'CDIK')
                  OR indexador IS NULL
                    THEN
                        ISNULL(perc_pl, 0)
                        * ISNULL(modified_duration, 0)
                        / 252.0
                ELSE 0
            END
        ) AS moddur_fundo,


        ----------------------------------------------------
        -- SPREAD CDI
        ----------------------------------------------------

        SUM(
            CASE
                WHEN ISNULL(spread_cdi_ativo, 0) > 15
                    THEN 0
                ELSE
                    ISNULL(spread_cdi_ativo, 0)
                    * ISNULL(perc_pl, 0)
            END
        ) AS spread_cdi_ponderado,


        SUM(
            CASE
                WHEN ISNULL(spread_cdi_ativo, 0) > 15
                    THEN ISNULL(spread_cdi_ativo, 0)
                ELSE 0
            END
        ) AS spread_cdi_ativo,


        SUM(
            CASE
                WHEN ISNULL(spread_cdi_ativo, 0) > 15
                    THEN
                        ISNULL(spread_cdi_ativo, 0)
                        * ISNULL(perc_pl, 0)
                ELSE 0
            END
        ) AS spread_cdi_credito_ponderado,


        ----------------------------------------------------
        -- SPREAD EQUIVALENTE
        ----------------------------------------------------

        SUM(
            CASE
                WHEN COALESCE(
                    spread_equivalente_ativo,
                    spread_cdi_ativo,
                    0
                ) > 15
                    THEN 0

                ELSE
                    COALESCE(
                        spread_equivalente_ativo,
                        spread_cdi_ativo,
                        0
                    )
                    * ISNULL(perc_pl, 0)
            END
        ) AS spread_equivalente_ponderado,


        SUM(
            CASE
                WHEN COALESCE(
                    spread_equivalente_ativo,
                    spread_cdi_ativo,
                    0
                ) > 15

                    THEN COALESCE(
                        spread_equivalente_ativo,
                        spread_cdi_ativo,
                        0
                    )

                ELSE 0
            END
        ) AS spread_equivalente_ativo,


        SUM(
            CASE
                WHEN fator_risco = 'Credito'
                    THEN
                        ISNULL(perc_pl, 0)
                        * COALESCE(
                            spread_equivalente_ativo,
                            spread_cdi_ativo,
                            0
                        )
                ELSE 0
            END
        ) AS spread_equivalente_credito_ponderado,


        ----------------------------------------------------
        -- PRÊMIO
        ----------------------------------------------------

        SUM(
            CASE
                WHEN ISNULL(spread_cdi_ativo, 0) > 15
                    THEN 0

                WHEN fator_risco_mercado = 'Inflacao'
                    THEN
                        COALESCE(
                            spread_ntnb_ativo,
                            spread_equivalente_ativo,
                            0
                        )

                ELSE
                    ISNULL(spread_cdi_ativo, 0)

            END
            * ISNULL(perc_pl, 0)
        ) AS premio_ponderado,


        SUM(
            CASE
                WHEN ISNULL(spread_cdi_ativo, 0) > 15
                    THEN 0

                WHEN fator_risco_mercado = 'Inflacao'
                    THEN
                        COALESCE(
                            spread_ntnb_ativo,
                            spread_equivalente_ativo,
                            0
                        )

                ELSE
                    ISNULL(spread_cdi_ativo, 0)

            END
        ) AS premio_ativo,


        SUM(
            CASE
                WHEN fator_risco = 'Credito'
                    THEN
                        ISNULL(perc_pl, 0)
                        * COALESCE(
                            spread_ntnb_ativo,
                            spread_equivalente_ativo,
                            0
                        )
                ELSE 0
            END
        ) AS premio_credito_ponderado,


        ----------------------------------------------------
        -- IMA-B
        ----------------------------------------------------

        MAX(yield_ima_b5_mais) AS yield_ima_b5_mais,
        MAX(yield_ima_b5) AS yield_ima_b5,
        MAX(yield_ima_b) AS yield_ima_b

    FROM BasePosicoes

    GROUP BY
        [date],
        codigo_IAM_fundo
),


------------------------------------------------------------
-- 6. INDICADORES FINAIS
------------------------------------------------------------

Indicadores AS
(
    SELECT
        [date],
        codigo_IAM_fundo,

        perc_pl_credito,
        perc_pl_credito_sem_cotas,
        perc_pl_credito_tem_spread,
        perc_pl_considerado,
        perc_debin,

        moddur_credito_ponderada
            / NULLIF(
                perc_pl_credito_sem_cotas,
                0
            ) AS moddur_credito,

        moddur_fundo,

        spread_equivalente_ponderado
            / NULLIF(
                perc_pl_considerado,
                0
            ) AS spread_equivalente,

        spread_equivalente_credito_ponderado
            / NULLIF(
                perc_pl_credito_tem_spread,
                0
            ) AS spread_equivalente_credito,

        spread_cdi_ponderado
            / NULLIF(
                perc_pl_credito_tem_spread,
                0
            ) AS spread_cdi,

        spread_cdi_credito_ponderado
            / NULLIF(
                perc_pl_credito_tem_spread,
                0
            ) AS spread_cdi_credito,

        premio_ponderado
            / NULLIF(
                perc_pl_considerado,
                0
            ) AS premio,

        premio_credito_ponderado
            / NULLIF(
                perc_pl_credito_tem_spread,
                0
            ) AS premio_credito,

        yield_ima_b5_mais,
        yield_ima_b5,
        yield_ima_b

    FROM IndicadoresBase
)


------------------------------------------------------------
-- 7. RESULTADO FINAL
------------------------------------------------------------

SELECT
    B.[date],
    B.codigo_IAM_fundo,
    B.isin,
    B.codigo_IAM_ativo,

    B.ativo,
    B.tipo,
    B.tipo_ativo_exposicao,
    B.fator_risco,
    B.fator_risco_mercado,

    B.instrumento,
    B.setor,
    B.grupo_economico,
    B.indexador,

    B.RATING_RANK_EQ_L_BR,
    B.ordem_rating,

    B.perc_pl,
    B.perc_nav,

    B.modified_duration,

    B.spread_cdi_ativo,
    B.spread_ipca_ativo,
    B.spread_ntnb_ativo,
    B.spread_equivalente_ativo,


    --------------------------------------------------------
    -- CAMPOS NÃO CALCULADOS
    --------------------------------------------------------

    CAST(NULL AS varchar(200)) AS benchmark,
    CAST(NULL AS varchar(200)) AS benchmark_ajustado,


    --------------------------------------------------------
    -- INDICADORES
    --------------------------------------------------------

    I.perc_pl_credito,
    I.perc_pl_credito_sem_cotas,
    I.perc_pl_credito_tem_spread,
    I.perc_pl_considerado,

    I.spread_equivalente,
    I.spread_equivalente_credito,

    CAST(NULL AS decimal(28,12))
        AS spread_ntnb_equivalente,

    CAST(NULL AS decimal(28,12))
        AS spread_equivalente_credito_extra,

    CAST(NULL AS varchar(200))
        AS ntnb_equivalente_credito,

    I.spread_cdi,
    I.spread_cdi_credito,

    I.premio,
    I.premio_credito,

    I.moddur_credito,
    I.moddur_fundo,


    --------------------------------------------------------
    -- CAMPOS NÃO CALCULADOS
    --------------------------------------------------------

    CAST(NULL AS decimal(28,12))
        AS patrimonio,

    I.perc_debin,

    CAST(NULL AS decimal(28,12))
        AS taxa_administracao_total,

    CAST(NULL AS decimal(28,12))
        AS taxa_performance,

    CAST(NULL AS varchar(50))
        AS indice_imab_referencia,

    CAST(NULL AS decimal(28,12))
        AS yield_imab_referencia,


    --------------------------------------------------------
    -- IMA-B
    --------------------------------------------------------

    I.yield_ima_b5_mais,
    I.yield_ima_b5,
    I.yield_ima_b,


    --------------------------------------------------------
    -- CAMPOS NÃO CALCULADOS
    --------------------------------------------------------

    CAST(NULL AS decimal(28,12))
        AS carrego_liquido,

    CAST(NULL AS decimal(28,12))
        AS carrego_liquido_b_equivalente,

    CAST(NULL AS decimal(28,12))
        AS carrego_liquido_imab,


    --------------------------------------------------------
    -- CONTROLE
    --------------------------------------------------------

    B.defasagem_dias,
    B.isin_encontrado_iam

FROM BasePosicoes AS B

LEFT JOIN Indicadores AS I
    ON I.[date] = B.[date]
   AND I.codigo_IAM_fundo = B.codigo_IAM_fundo

ORDER BY
    B.instrumento,
    B.grupo_economico,
    B.setor,
    B.codigo_IAM_ativo,
    B.isin

OPTION (RECOMPILE);
