-- Modelo estrela (star schema) do data warehouse de produtividade de rotas.
-- Gerado/aplicado por warehouse.py (SQLite). Aqui fica documentado o modelo.
--
-- Grao do fato: uma rota por motorista/ajudante por dia.
-- Dimensoes: calendario (tempo) e colaborador (quem).

CREATE TABLE dim_calendario (
    data_id       INTEGER PRIMARY KEY,   -- AAAAMMDD
    data          TEXT NOT NULL,         -- AAAA-MM-DD (ISO)
    data_br       TEXT NOT NULL,         -- DD/MM/AAAA
    ano           INTEGER NOT NULL,
    mes           INTEGER NOT NULL,
    dia           INTEGER NOT NULL,
    trimestre     INTEGER NOT NULL,
    dia_semana    INTEGER NOT NULL,      -- 1=dom .. 7=sab
    nome_dia      TEXT NOT NULL,
    semana_iso    INTEGER NOT NULL,
    fim_de_semana INTEGER NOT NULL       -- 0/1
);

CREATE TABLE dim_colaborador (
    colab_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo       TEXT NOT NULL,
    tipo         TEXT NOT NULL,          -- Motorista / Ajudante
    nome         TEXT,
    cpf_mascara  TEXT,                   -- PII mascarada: 123.***.***-99
    UNIQUE (codigo, tipo)
);

CREATE TABLE fato_jornada (
    fato_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    data_id      INTEGER NOT NULL REFERENCES dim_calendario (data_id),
    colab_id     INTEGER NOT NULL REFERENCES dim_colaborador (colab_id),
    tml_min      REAL,   -- liberacao (min)
    tr_min       REAL,   -- tempo em rota (min)
    ti_min       REAL,   -- tempo interno (min)
    jl_min       REAL,   -- jornada liquida total (min) = tml + tr + ti
    previsto_min REAL,   -- tempo previsto da rota (min)
    tml_cap_min  REAL,   -- liberacao para media (limitada a 5h)
    val_jl       INTEGER NOT NULL DEFAULT 0,   -- rota valida (liberacao 0-4h)
    bate_jl      INTEGER NOT NULL DEFAULT 0,   -- JL total <= 10:20
    val_tml      INTEGER NOT NULL DEFAULT 0,
    bate_tml     INTEGER NOT NULL DEFAULT 0,   -- liberacao <= 30 min
    tr_ok        INTEGER NOT NULL DEFAULT 0,   -- tempo em rota <= 09:20
    ti_ok        INTEGER NOT NULL DEFAULT 0    -- tempo interno <= 30 min
);
CREATE INDEX ix_fato_data  ON fato_jornada (data_id);
CREATE INDEX ix_fato_colab ON fato_jornada (colab_id);

-- Resultados das checagens de qualidade de dados (uma linha por verificacao).
CREATE TABLE dq_resultados (
    verificacao  TEXT,
    valor        TEXT,
    status       TEXT,     -- OK / ALERTA
    detalhe      TEXT,
    executado_em TEXT
);
