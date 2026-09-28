# -*- coding: utf-8 -*-
"""
Camada de engenharia de dados do sistema de produtividade de rotas.

Constroi um data warehouse (modelo estrela) em SQLite a partir das jornadas
processadas e exporta tabelas prontas para ferramentas de BI (Power BI,
Metabase, Looker Studio etc.).

Pipeline:
    Extract / Transform  ->  jornada_liquida.processar_arquivos (2 ART + ponto)
    Load                 ->  modelo estrela em SQLite
    Data Quality         ->  checagens de integridade, cobertura e faixa
    Export               ->  CSV das dimensoes, do fato e das views de KPI

Modelo estrela:
    dim_calendario(data_id, data, ano, mes, dia, dia_semana, semana_iso, ...)
    dim_colaborador(colab_id, codigo, tipo, nome, cpf_mascara)
    fato_jornada(fato_id, data_id, colab_id, medidas..., flags...)

Views de KPI (mesma regra do painel do Excel):
    vw_kpi_diario, vw_kpi_mensal

Uso via linha de comando:
    python warehouse.py --2art <arquivo> --ponto <arquivo> [--saida-db dw.db]
"""

import csv
import datetime as _dt
import os
import sqlite3

import jornada_liquida as core


# metas dos indicadores (minutos), iguais as do painel
METAS = {
    "jl_max_min": 620,     # jornada liquida total <= 10:20
    "tml_max_min": 30,     # liberacao <= 30 min
    "tr_max_min": 560,     # tempo em rota <= 09:20
    "ti_max_min": 30,      # tempo interno <= 30 min
    "real_meta": 0.70,     # % de rotas dentro da JL
    "libera_valida_max": 240,   # rota valida = liberacao entre 0 e 4h
}

_NOMES_DIA = ["seg", "ter", "qua", "qui", "sex", "sab", "dom"]

SCHEMA_SQL = """
CREATE TABLE dim_calendario (
    data_id      INTEGER PRIMARY KEY,   -- AAAAMMDD
    data         TEXT NOT NULL,         -- AAAA-MM-DD (ISO)
    data_br      TEXT NOT NULL,         -- DD/MM/AAAA
    ano          INTEGER NOT NULL,
    mes          INTEGER NOT NULL,
    dia          INTEGER NOT NULL,
    trimestre    INTEGER NOT NULL,
    dia_semana   INTEGER NOT NULL,      -- 1=dom .. 7=sab
    nome_dia     TEXT NOT NULL,
    semana_iso   INTEGER NOT NULL,
    fim_de_semana INTEGER NOT NULL      -- 0/1
);

CREATE TABLE dim_colaborador (
    colab_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo       TEXT NOT NULL,
    tipo         TEXT NOT NULL,         -- Motorista / Ajudante
    nome         TEXT,
    cpf_mascara  TEXT,                  -- PII mascarada: 123.***.***-99
    UNIQUE (codigo, tipo)
);

CREATE TABLE fato_jornada (
    fato_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    data_id      INTEGER NOT NULL REFERENCES dim_calendario (data_id),
    colab_id     INTEGER NOT NULL REFERENCES dim_colaborador (colab_id),
    tml_min      REAL,   -- liberacao (min)
    tr_min       REAL,   -- tempo em rota (min)
    ti_min       REAL,   -- tempo interno (min)
    jl_min       REAL,   -- jornada liquida total (min)
    previsto_min REAL,   -- tempo previsto da rota (min)
    tml_cap_min  REAL,   -- liberacao para media (limitada a 5h)
    val_jl       INTEGER NOT NULL DEFAULT 0,
    bate_jl      INTEGER NOT NULL DEFAULT 0,
    val_tml      INTEGER NOT NULL DEFAULT 0,
    bate_tml     INTEGER NOT NULL DEFAULT 0,
    tr_ok        INTEGER NOT NULL DEFAULT 0,
    ti_ok        INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX ix_fato_data  ON fato_jornada (data_id);
CREATE INDEX ix_fato_colab ON fato_jornada (colab_id);

CREATE TABLE dq_resultados (
    verificacao  TEXT,
    valor        TEXT,
    status       TEXT,     -- OK / ALERTA
    detalhe      TEXT,
    executado_em TEXT
);
"""

# Views de KPI: reproduzem as regras do painel (so motoristas/caminhoes).
VIEWS_SQL = """
CREATE VIEW vw_kpi_diario AS
SELECT c.data_id,
       c.data,
       c.data_br,
       c.ano,
       c.mes,
       c.semana_iso,
       c.dia_semana,
       SUM(f.val_jl)                                              AS mvalid,
       SUM(f.bate_jl)                                             AS mjl,
       ROUND(1.0 * SUM(f.bate_jl) / NULLIF(SUM(f.val_jl), 0), 4)  AS real_jl,
       ROUND(AVG(f.tml_cap_min), 1)                               AS tml_medio_min,
       ROUND(1.0 * SUM(f.bate_tml) / NULLIF(SUM(f.val_tml), 0), 4) AS pct_tml,
       ROUND(1.0 * SUM(f.tr_ok) / NULLIF(SUM(f.val_jl), 0), 4)    AS pct_tr,
       ROUND(1.0 * SUM(f.ti_ok) / NULLIF(SUM(f.val_jl), 0), 4)    AS pct_ti
FROM fato_jornada f
JOIN dim_calendario  c ON c.data_id  = f.data_id
JOIN dim_colaborador d ON d.colab_id = f.colab_id
WHERE d.tipo = 'Motorista'
GROUP BY c.data_id
ORDER BY c.data_id;

CREATE VIEW vw_kpi_mensal AS
SELECT c.ano,
       c.mes,
       COUNT(DISTINCT c.data_id)                                  AS dias,
       SUM(f.val_jl)                                              AS mvalid,
       SUM(f.bate_jl)                                             AS mjl,
       ROUND(1.0 * SUM(f.bate_jl) / NULLIF(SUM(f.val_jl), 0), 4)  AS real_jl,
       ROUND(AVG(f.tml_cap_min), 1)                               AS tml_medio_min,
       ROUND(1.0 * SUM(f.bate_tml) / NULLIF(SUM(f.val_tml), 0), 4) AS pct_tml,
       ROUND(1.0 * SUM(f.tr_ok) / NULLIF(SUM(f.val_jl), 0), 4)    AS pct_tr,
       ROUND(1.0 * SUM(f.ti_ok) / NULLIF(SUM(f.val_jl), 0), 4)    AS pct_ti
FROM fato_jornada f
JOIN dim_calendario  c ON c.data_id  = f.data_id
JOIN dim_colaborador d ON d.colab_id = f.colab_id
WHERE d.tipo = 'Motorista'
GROUP BY c.ano, c.mes
ORDER BY c.ano, c.mes;
"""


def _mascara_cpf(cpf):
    """PII: mostra so os 3 primeiros e 2 ultimos digitos (123.***.***-99)."""
    d = core.cpf11(cpf)
    if not core.is_cpf_valido(d):
        return None
    return "%s.***.***-%s" % (d[:3], d[9:11])


def _abrir(caminho_db):
    if os.path.exists(caminho_db):
        os.remove(caminho_db)
    conn = sqlite3.connect(caminho_db)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.executescript(SCHEMA_SQL)
    conn.executescript(VIEWS_SQL)
    return conn


def _dim_calendario_row(dobj):
    data_id = dobj.year * 10000 + dobj.month * 100 + dobj.day
    excel_wd = (dobj.weekday() + 1) % 7 + 1     # 1=dom .. 7=sab (igual ao painel)
    return (
        data_id,
        dobj.isoformat(),
        "%02d/%02d/%04d" % (dobj.day, dobj.month, dobj.year),
        dobj.year, dobj.month, dobj.day,
        (dobj.month - 1) // 3 + 1,
        excel_wd,
        _NOMES_DIA[dobj.weekday()],
        dobj.isocalendar()[1],
        1 if dobj.weekday() >= 5 else 0,
    )


def _carregar(conn, linhas):
    """Popula dimensoes e fato a partir das linhas processadas (CABECALHO)."""
    cur = conn.cursor()
    cal_vistas = set()
    colab_id = {}      # (codigo, tipo) -> colab_id
    n = 0
    for l in linhas:
        dobj = core._data_obj(l[0] if l else "")
        if dobj is None:
            continue
        cal = _dim_calendario_row(dobj)
        data_id = cal[0]
        if data_id not in cal_vistas:
            cur.execute(
                "INSERT OR IGNORE INTO dim_calendario VALUES (?,?,?,?,?,?,?,?,?,?,?)", cal)
            cal_vistas.add(data_id)

        tipo = l[1] if len(l) > 1 else ""
        codigo = str(l[2]) if len(l) > 2 else ""
        nome = (l[core._IX_NOME] if len(l) > core._IX_NOME and l[core._IX_NOME]
                else (l[3] if len(l) > 3 else "")) or ""
        cpf = l[4] if len(l) > 4 else ""
        chave = (codigo, tipo)
        if chave not in colab_id:
            cur.execute(
                "INSERT OR IGNORE INTO dim_colaborador (codigo, tipo, nome, cpf_mascara) "
                "VALUES (?,?,?,?)", (codigo, tipo, nome, _mascara_cpf(cpf)))
            cur.execute(
                "SELECT colab_id FROM dim_colaborador WHERE codigo=? AND tipo=?",
                (codigo, tipo))
            colab_id[chave] = cur.fetchone()[0]

        ind = core.indicadores_rota(l)
        cur.execute(
            "INSERT INTO fato_jornada "
            "(data_id, colab_id, tml_min, tr_min, ti_min, jl_min, previsto_min, "
            " tml_cap_min, val_jl, bate_jl, val_tml, bate_tml, tr_ok, ti_ok) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (data_id, colab_id[chave], ind["tml_min"], ind["tr_min"], ind["ti_min"],
             ind["jl_min"], ind["previsto_min"], ind["tml_cap_min"],
             ind["val_jl"], ind["bate_jl"], ind["val_tml"], ind["bate_tml"],
             ind["tr_ok"], ind["ti_ok"]))
        n += 1
    conn.commit()
    return n


def _rodar_qualidade(conn):
    """Checagens de qualidade de dados. Grava em dq_resultados e devolve lista."""
    cur = conn.cursor()
    agora = _dt.datetime.now().isoformat(timespec="seconds")
    res = []

    def q(sql, *p):
        cur.execute(sql, p)
        return cur.fetchone()[0]

    total = q("SELECT COUNT(*) FROM fato_jornada")
    res.append(("registros_fato", total, "OK", "linhas na tabela fato"))

    mot = q("SELECT COUNT(*) FROM fato_jornada f JOIN dim_colaborador d "
            "ON d.colab_id=f.colab_id WHERE d.tipo='Motorista'")
    res.append(("registros_motorista", mot, "OK", "rotas de motorista"))

    cols = q("SELECT COUNT(*) FROM dim_colaborador")
    sem_cpf = q("SELECT COUNT(*) FROM dim_colaborador WHERE cpf_mascara IS NULL")
    pct_cpf = (100.0 * (cols - sem_cpf) / cols) if cols else 0.0
    res.append(("colaboradores_com_cpf",
                "%.1f%%" % pct_cpf,
                "OK" if pct_cpf >= 90 else "ALERTA",
                "%d de %d colaboradores com CPF valido" % (cols - sem_cpf, cols)))

    sem_ponto = q("SELECT COUNT(*) FROM fato_jornada WHERE jl_min IS NULL")
    pct_sp = (100.0 * sem_ponto / total) if total else 0.0
    res.append(("rotas_sem_ponto",
                "%.1f%%" % pct_sp,
                "OK" if pct_sp <= 15 else "ALERTA",
                "%d rotas sem espelho de ponto casado" % sem_ponto))

    dup = q("SELECT COUNT(*) FROM (SELECT data_id, colab_id, COUNT(*) c "
            "FROM fato_jornada GROUP BY data_id, colab_id HAVING c > 1)")
    res.append(("chaves_duplicadas", dup,
                "OK" if dup == 0 else "ALERTA",
                "pares (dia, colaborador) repetidos no fato"))

    orfaos = q("SELECT COUNT(*) FROM fato_jornada f "
               "LEFT JOIN dim_calendario c ON c.data_id=f.data_id "
               "WHERE c.data_id IS NULL")
    res.append(("fatos_orfaos", orfaos,
                "OK" if orfaos == 0 else "ALERTA",
                "fatos sem data na dimensao (integridade referencial)"))

    fora = q("SELECT COUNT(*) FROM fato_jornada f JOIN dim_colaborador d "
             "ON d.colab_id=f.colab_id WHERE d.tipo='Motorista' "
             "AND f.tml_min IS NOT NULL AND (f.tml_min <= 0 OR f.tml_min > ?)",
             METAS["libera_valida_max"])
    res.append(("liberacao_fora_de_faixa", fora,
                "OK" if fora == 0 else "ALERTA",
                "rotas com liberacao <=0 ou >4h (excluidas do valido)"))

    dmin = q("SELECT MIN(data) FROM dim_calendario")
    dmax = q("SELECT MAX(data) FROM dim_calendario")
    res.append(("cobertura_periodo", "%s a %s" % (dmin, dmax), "OK",
                "menor e maior data carregada"))

    cur.executemany(
        "INSERT INTO dq_resultados "
        "(verificacao, valor, status, detalhe, executado_em) VALUES (?,?,?,?,?)",
        [(v, str(val), st, det, agora) for (v, val, st, det) in res])
    conn.commit()
    return res


def _exportar_csv(conn, pasta):
    """Exporta tabelas e views de KPI para CSV (UTF-8) prontos para BI."""
    if not os.path.isdir(pasta):
        os.makedirs(pasta)
    alvos = ["dim_calendario", "dim_colaborador", "fato_jornada",
             "dq_resultados", "vw_kpi_diario", "vw_kpi_mensal"]
    cur = conn.cursor()
    gerados = []
    for nome in alvos:
        cur.execute("SELECT * FROM %s" % nome)
        cols = [d[0] for d in cur.description]
        caminho = os.path.join(pasta, nome + ".csv")
        with open(caminho, "w", newline="", encoding="utf-8-sig") as f:
            wtr = csv.writer(f)
            wtr.writerow(cols)
            wtr.writerows(cur.fetchall())
        gerados.append(caminho)
    return gerados


def construir_e_exportar(linhas_mot, linhas_aju, caminho_db, exportar_csv=True):
    """Roda o pipeline completo e devolve um resumo (registros, DQ, caminhos)."""
    conn = _abrir(caminho_db)
    try:
        n = _carregar(conn, list(linhas_mot) + list(linhas_aju))
        dq = _rodar_qualidade(conn)
        pasta_csv = None
        if exportar_csv:
            pasta_csv = os.path.splitext(caminho_db)[0] + "_bi"
            _exportar_csv(conn, pasta_csv)
    finally:
        conn.close()
    return {"registros": n, "dq": dq, "db": caminho_db, "csv": pasta_csv}


def _main():
    import argparse
    import armazenamento

    ap = argparse.ArgumentParser(
        description="Pipeline ETL -> data warehouse de produtividade de rotas.")
    ap.add_argument("--2art", dest="art", required=True,
                    help="2 ART do periodo (CSV ou Excel).")
    ap.add_argument("--ponto", dest="ponto", required=True,
                    help="Espelho de ponto/jornada (Excel do Pontomais).")
    ap.add_argument("--saida-db", dest="db", default="warehouse.db",
                    help="Arquivo SQLite de saida (padrao: warehouse.db).")
    ap.add_argument("--sem-csv", dest="sem_csv", action="store_true",
                    help="Nao exportar os CSV para BI.")
    args = ap.parse_args()

    dados = armazenamento.carregar()
    lm, la, conf = core.processar_arquivos(
        dados["motoristas"], dados["ajudantes"], dados["colaboradores"],
        args.art, args.ponto)

    res = construir_e_exportar(lm, la, args.db, exportar_csv=not args.sem_csv)

    print("Data warehouse gerado: %s" % res["db"])
    print("  Registros no fato: %d  (motoristas + ajudantes)" % res["registros"])
    print("  Casos de conferencia: %d" % len(conf))
    if res["csv"]:
        print("  CSV para BI em: %s" % res["csv"])
    print("\nQualidade de dados:")
    for v, val, st, det in res["dq"]:
        print("  [%-7s] %-26s %-14s %s" % (st, v, str(val), det))


if __name__ == "__main__":
    _main()
