-- Views de KPI: reproduzem as regras do painel (somente motoristas/caminhoes).
--
-- Indicadores por dia:
--   MVALID  = rotas validas               = SUM(val_jl)
--   MJL     = rotas dentro da JL (<=10:20) = SUM(bate_jl)
--   REAL    = MJL / MVALID
--   TML     = tempo medio de liberacao     = AVG(tml_cap_min)
--   %TML    = % com liberacao <= 30 min    = SUM(bate_tml) / SUM(val_tml)
--   %TR     = % com tempo em rota <= 09:20 = SUM(tr_ok) / SUM(val_jl)
--   %TI     = % com tempo interno <= 30min = SUM(ti_ok) / SUM(val_jl)

CREATE VIEW vw_kpi_diario AS
SELECT c.data_id,
       c.data,
       c.data_br,
       c.ano,
       c.mes,
       c.semana_iso,
       c.dia_semana,
       SUM(f.val_jl)                                               AS mvalid,
       SUM(f.bate_jl)                                              AS mjl,
       ROUND(1.0 * SUM(f.bate_jl) / NULLIF(SUM(f.val_jl), 0), 4)   AS real_jl,
       ROUND(AVG(f.tml_cap_min), 1)                                AS tml_medio_min,
       ROUND(1.0 * SUM(f.bate_tml) / NULLIF(SUM(f.val_tml), 0), 4) AS pct_tml,
       ROUND(1.0 * SUM(f.tr_ok) / NULLIF(SUM(f.val_jl), 0), 4)     AS pct_tr,
       ROUND(1.0 * SUM(f.ti_ok) / NULLIF(SUM(f.val_jl), 0), 4)     AS pct_ti
FROM fato_jornada f
JOIN dim_calendario  c ON c.data_id  = f.data_id
JOIN dim_colaborador d ON d.colab_id = f.colab_id
WHERE d.tipo = 'Motorista'
GROUP BY c.data_id
ORDER BY c.data_id;

CREATE VIEW vw_kpi_mensal AS
SELECT c.ano,
       c.mes,
       COUNT(DISTINCT c.data_id)                                   AS dias,
       SUM(f.val_jl)                                               AS mvalid,
       SUM(f.bate_jl)                                              AS mjl,
       ROUND(1.0 * SUM(f.bate_jl) / NULLIF(SUM(f.val_jl), 0), 4)   AS real_jl,
       ROUND(AVG(f.tml_cap_min), 1)                                AS tml_medio_min,
       ROUND(1.0 * SUM(f.bate_tml) / NULLIF(SUM(f.val_tml), 0), 4) AS pct_tml,
       ROUND(1.0 * SUM(f.tr_ok) / NULLIF(SUM(f.val_jl), 0), 4)     AS pct_tr,
       ROUND(1.0 * SUM(f.ti_ok) / NULLIF(SUM(f.val_jl), 0), 4)     AS pct_ti
FROM fato_jornada f
JOIN dim_calendario  c ON c.data_id  = f.data_id
JOIN dim_colaborador d ON d.colab_id = f.colab_id
WHERE d.tipo = 'Motorista'
GROUP BY c.ano, c.mes
ORDER BY c.ano, c.mes;
