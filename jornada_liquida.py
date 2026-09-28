# -*- coding: utf-8 -*-
"""
Jornada Líquida - Analisador de tempos Motorista/Ajudante
==========================================================
Cruza os horarios do 2 ART (Promax) com o espelho de ponto (Pontomais),
usando o CPF como chave de ligacao entre os dois sistemas, e gera um
relatorio em Excel com as diferencas de horario por pessoa e por dia.

Cadeia de ligacao:
    2 ART (codigo do motorista/ajudante)
        -> cadastro Promax (codigo -> CPF)
            -> cadastro de colaboradores do Pontomais (CPF -> nome completo)
                -> espelho de jornada do Pontomais (nome -> entrada/saida/intervalo)

Roda como programa (.exe) sem precisar instalar Python.
"""

import csv
import os
import re
import unicodedata
import datetime as _dt

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList


# ---------------------------------------------------------------------------
# Utilidades de normalizacao
# ---------------------------------------------------------------------------
def only_digits(s):
    """Mantem apenas digitos de um texto."""
    return re.sub(r"\D", "", s or "")


def cpf11(s):
    """Normaliza CPF para 11 digitos.

    O Promax as vezes corta o zero a esquerda (o campo fica com 10 digitos),
    entao completamos com zero para conseguir casar com o CPF do Pontomais.
    """
    d = only_digits(s)
    if len(d) == 10:
        d = "0" + d
    return d


def norm_txt(s):
    """Texto em maiusculas, sem acento e sem espacos duplicados (para comparar nomes)."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().upper()


_CPF_FAKE = {
    "00000000000", "11111111111", "22222222222", "33333333333",
    "44444444444", "55555555555", "66666666666", "77777777777",
    "88888888888", "99999999999",
}


def is_cpf_valido(cpf):
    """CPF utilizavel para casar (11 digitos e nao e um CPF 'fake' de cadastro)."""
    return len(cpf) == 11 and cpf not in _CPF_FAKE


# ---------------------------------------------------------------------------
# Tempo (HH:MM)
# ---------------------------------------------------------------------------
def hhmm(x):
    """Retorna 'HH:MM' se o texto for um horario valido, senao ''."""
    x = (x or "")
    x = str(x).strip()
    m = re.match(r"^(\d{1,2}):(\d{2})", x)
    if not m:
        return ""
    return "%02d:%02d" % (int(m.group(1)), int(m.group(2)))


def to_min(t):
    """'HH:MM' -> minutos (int) ou None."""
    t = hhmm(t)
    if not t:
        return None
    h, m = t.split(":")
    return int(h) * 60 + int(m)


def fmt_min(m):
    """minutos (int, pode ser negativo) -> 'HH:MM' ou ''."""
    if m is None:
        return ""
    sinal = "-" if m < 0 else ""
    m = abs(int(round(m)))
    return "%s%02d:%02d" % (sinal, m // 60, m % 60)


def diff_min(a, b):
    """minutos de a - b, tratando None."""
    ma, mb = to_min(a), to_min(b)
    if ma is None or mb is None:
        return None
    return ma - mb


# ---------------------------------------------------------------------------
# Leitura dos cadastros do Promax (motorista / ajudante)
# ---------------------------------------------------------------------------
def carregar_cadastro_promax(caminho):
    """Le um CSV de cadastro do Promax (separador ';').

    Detecta as colunas de Codigo, Nome e CPF pelo cabecalho.
    Retorna: dict codigo(str, sem zeros a esquerda) -> {'nome', 'cpf'}
    """
    with open(caminho, "r", encoding="latin-1", newline="") as f:
        linhas = list(csv.reader(f, delimiter=";"))
    if not linhas:
        return {}

    cab = [norm_txt(c) for c in linhas[0]]

    def achar(*nomes):
        for i, c in enumerate(cab):
            for n in nomes:
                if c == n:
                    return i
        for i, c in enumerate(cab):
            for n in nomes:
                if n in c:
                    return i
        return None

    i_cod = achar("CODIGO") or 0
    i_nome = achar("NOME", "NOME AJUDANTE", "NOME MOTORISTA")
    i_cpf = achar("CPF")

    if i_nome is None:
        i_nome = 1
    if i_cpf is None:
        raise ValueError("Nao encontrei a coluna CPF no arquivo:\n%s" % caminho)

    d = {}
    for row in linhas[1:]:
        if not row or len(row) <= max(i_cod, i_nome, i_cpf):
            continue
        cod = (row[i_cod] or "").strip().lstrip("0") or "0"
        if cod == "0" and not (row[i_cod] or "").strip():
            continue
        d[cod] = {
            "nome": norm_txt(row[i_nome]),
            "cpf": cpf11(row[i_cpf]),
        }
    return d


# ---------------------------------------------------------------------------
# Leitura dos colaboradores do Pontomais (CPF -> nome)
# ---------------------------------------------------------------------------
def _iter_linhas_xlsx(caminho):
    """Le todas as linhas de um .xlsx (sem read_only, para pegar todas as colunas)."""
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb.active
    linhas = list(ws.iter_rows(values_only=True))
    wb.close()
    return linhas


def carregar_colaboradores_pontomais(caminho):
    """Le o 'Relatorio de Colaboradores' do Pontomais.

    Estrutura em blocos por colaborador:
        Colaborador | NOME
        Nome | PIS | Cargo | Equipe | Turno | CPF | E-mail | ...
        NOME | pis | cargo | ...   | cpf | ...
        Resumo | Totais
        Total | 1

    Retorna: dict cpf(11) -> nome_normalizado
    """
    linhas = _iter_linhas_xlsx(caminho)
    cpf2nome = {}
    i = 0
    while i < len(linhas):
        r = linhas[i]
        if r and str(r[0]).strip() == "Nome":
            # essa linha e o cabecalho; descobrir indice do CPF
            cab = [norm_txt(c) for c in r]
            i_cpf = None
            for k, c in enumerate(cab):
                if c == "CPF":
                    i_cpf = k
                    break
            if i_cpf is None:
                i_cpf = 5  # posicao padrao observada
            if i + 1 < len(linhas):
                dr = linhas[i + 1]
                nome = norm_txt(dr[0]) if dr else ""
                cpf = cpf11(dr[i_cpf]) if (dr and len(dr) > i_cpf) else ""
                if nome and is_cpf_valido(cpf):
                    cpf2nome[cpf] = nome
            i += 2
            continue
        i += 1
    return cpf2nome


# ---------------------------------------------------------------------------
# Leitura do espelho de jornada do Pontomais (nome + data -> horarios)
# ---------------------------------------------------------------------------
def _data_br_para_chave(txt):
    """'Qua, 22/07/2026' -> '22072026' (mesma chave do 2 ART). Retorna '' se nao achar."""
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})", str(txt or ""))
    if not m:
        return ""
    return m.group(1) + m.group(2) + m.group(3)


def carregar_jornada_pontomais(caminho):
    """Le o 'Relatorio de Jornada' (espelho de ponto) do Pontomais.

    Retorna: dict nome_normalizado -> { chave_data: {e1,s1,e2,s2,intervalo,obs} }
    """
    linhas = _iter_linhas_xlsx(caminho)
    jornada = {}
    nome_atual = None
    for i, r in enumerate(linhas):
        if not r:
            continue
        c0 = str(r[0]).strip() if r[0] is not None else ""
        if c0 == "Colaborador":
            nome_atual = norm_txt(r[1]) if len(r) > 1 else None
            continue
        if c0 in ("Data", "TOTAIS", "Resumo", "Total", ""):
            continue
        # linha de dados de um dia: comeca com dia da semana + data
        chave = _data_br_para_chave(c0)
        if chave and nome_atual:
            def get(idx):
                return r[idx] if len(r) > idx else None
            jornada.setdefault(nome_atual, {})[chave] = {
                "e1": get(1), "s1": get(2), "e2": get(3), "s2": get(4),
                "intervalo": get(7), "obs": get(12),
            }
    return jornada


def entrada_saida_do_dia(dia):
    """Do registro de um dia, retorna (entrada, saida)."""
    entrada = hhmm(dia.get("e1"))
    saida = hhmm(dia.get("s2")) or hhmm(dia.get("s1"))
    return entrada, saida


# ---------------------------------------------------------------------------
# Leitura do 2 ART (Promax)
# ---------------------------------------------------------------------------
# indices das colunas no 2 ART (arquivo posicional exportado como CSV ';')
_2ART = {
    "data": 0, "mapa": 9, "hrsai": 20, "hrentr": 21,
    "cdmot": 30, "cdaju1": 31, "cdaju2": 32, "tempoprev": 25,
}


def _serial_para_ddmmyyyy(v):
    """Data do 2 ART (numero de serie do Excel, date/datetime ou texto) -> 'DDMMYYYY'."""
    if v is None:
        return ""
    if isinstance(v, _dt.datetime):
        d = v.date()
    elif isinstance(v, _dt.date):
        d = v
    elif isinstance(v, (int, float)):
        d = (_dt.datetime(1899, 12, 30) + _dt.timedelta(days=float(v))).date()
    else:
        return only_digits(str(v))
    return "%02d%02d%04d" % (d.day, d.month, d.year)


def _fracao_para_hhmm(v):
    """Hora do 2 ART (fracao do dia do Excel, time/datetime ou 'HH:MM') -> 'HH:MM'."""
    if v is None:
        return ""
    if isinstance(v, _dt.datetime):
        return "%02d:%02d" % (v.hour, v.minute)
    if isinstance(v, _dt.time):
        return "%02d:%02d" % (v.hour, v.minute)
    if isinstance(v, (int, float)):
        t = int(round(float(v) * 1440))
        return "%02d:%02d" % ((t // 60) % 24, t % 60)
    return hhmm(str(v))


def _cod_2art(v):
    """Codigo do 2 ART (numero ou texto) -> string sem zeros a esquerda."""
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.lstrip("0")


def _e_xlsx(caminho):
    try:
        with open(caminho, "rb") as f:
            return f.read(2) == b"PK"
    except Exception:
        return False


def _carregar_2art_xlsx(caminho):
    """Le o 2 ART quando ele vem em Excel (xlsx), mesmo com extensao .csv.
    Datas viram numero de serie e horas viram fracao do dia - convertidos aqui.
    """
    import io
    with open(caminho, "rb") as f:
        wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True, read_only=True)
    ws = wb.active
    saida = []
    primeira = True
    for row in ws.iter_rows(values_only=True):
        if primeira:
            primeira = False
            continue
        if not row or len(row) <= _2ART["cdaju2"]:
            continue
        hs = _fracao_para_hhmm(row[_2ART["hrsai"]])
        he = _fracao_para_hhmm(row[_2ART["hrentr"]])
        if not hs or not he or (hs == "00:00" and he == "00:00"):
            continue
        data = _serial_para_ddmmyyyy(row[_2ART["data"]])
        mapa = str(row[_2ART["mapa"]] if row[_2ART["mapa"]] is not None else "").strip()
        tprev = _fracao_para_hhmm(row[_2ART["tempoprev"]])

        cod_mot = _cod_2art(row[_2ART["cdmot"]])
        if cod_mot and cod_mot != "0":
            saida.append({"tipo": "MOT", "cod": cod_mot, "mapa": mapa, "data": data,
                          "hrsai": hs, "hrentr": he, "tempoprev": tprev})
        for key in ("cdaju1", "cdaju2"):
            cod = _cod_2art(row[_2ART[key]])
            if cod and cod != "0":
                saida.append({"tipo": "AJU", "cod": cod, "mapa": mapa, "data": data,
                              "hrsai": hs, "hrentr": he, "tempoprev": tprev})
    return saida


def carregar_2art(caminho):
    """Le o 2 ART (CSV do Promax OU Excel/xlsx). Retorna lista de dicts:
       {tipo:'MOT'/'AJU', cod, mapa, data, hrsai, hrentr, tempoprev}
    Ignora linhas sem rota (HrSai e HrEntr zerados) e codigos vazios/00000.
    """
    if _e_xlsx(caminho):
        return _carregar_2art_xlsx(caminho)

    with open(caminho, "r", encoding="latin-1", newline="") as f:
        linhas = list(csv.reader(f, delimiter=";"))
    saida = []
    for row in linhas[1:]:
        if len(row) <= _2ART["cdaju2"]:
            continue
        hs = hhmm(row[_2ART["hrsai"]])
        he = hhmm(row[_2ART["hrentr"]])
        if not hs or not he or (hs == "00:00" and he == "00:00"):
            continue
        data = only_digits(row[_2ART["data"]])
        mapa = (row[_2ART["mapa"]] or "").strip()
        tprev = hhmm(row[_2ART["tempoprev"]])  # tempo previsto de rota (HH:MM)

        cod_mot = (row[_2ART["cdmot"]] or "").strip().lstrip("0")
        if cod_mot and cod_mot != "0":
            saida.append({"tipo": "MOT", "cod": cod_mot, "mapa": mapa, "data": data,
                          "hrsai": hs, "hrentr": he, "tempoprev": tprev})
        for key in ("cdaju1", "cdaju2"):
            cod = (row[_2ART[key]] or "").strip().lstrip("0")
            if cod and cod != "0":
                saida.append({"tipo": "AJU", "cod": cod, "mapa": mapa, "data": data,
                              "hrsai": hs, "hrentr": he, "tempoprev": tprev})
    return saida


# ---------------------------------------------------------------------------
# Processamento / cruzamento
# ---------------------------------------------------------------------------
def data_br(chave):
    """'22072026' -> '22/07/2026'."""
    if len(chave) == 8:
        return "%s/%s/%s" % (chave[0:2], chave[2:4], chave[4:8])
    return chave


def processar(cad_mot, cad_aju, colab_pontomais, jornada, art_linhas):
    """Cruza tudo e devolve (linhas_motorista, linhas_ajudante, conferencia).

    Agrupa por (tipo, codigo, data): pega a saida de caminhao mais cedo e a
    entrada mais tarde do dia (caso a pessoa tenha mais de um mapa).
    """
    grupos = {}  # (tipo, cod, data) -> {mapas, hrsai_min, hrentr_max, tempoprev_soma}
    for e in art_linhas:
        k = (e["tipo"], e["cod"], e["data"])
        g = grupos.setdefault(k, {"mapas": [], "hrsai": e["hrsai"],
                                  "hrentr": e["hrentr"], "tprev": 0, "tem_prev": False})
        if e["mapa"] and e["mapa"] not in g["mapas"]:
            g["mapas"].append(e["mapa"])
        if to_min(e["hrsai"]) is not None and (to_min(g["hrsai"]) is None or to_min(e["hrsai"]) < to_min(g["hrsai"])):
            g["hrsai"] = e["hrsai"]
        if to_min(e["hrentr"]) is not None and (to_min(g["hrentr"]) is None or to_min(e["hrentr"]) > to_min(g["hrentr"])):
            g["hrentr"] = e["hrentr"]
        tp = to_min(e.get("tempoprev"))
        if tp:  # soma o tempo previsto de cada mapa da pessoa no dia
            g["tprev"] += tp
            g["tem_prev"] = True

    linhas_mot, linhas_aju, conferencia = [], [], []

    for (tipo, cod, data), g in sorted(grupos.items(), key=lambda x: (x[0][0], int(x[0][1]) if x[0][1].isdigit() else 0)):
        cad = (cad_mot if tipo == "MOT" else cad_aju).get(cod)
        tipo_txt = "Motorista" if tipo == "MOT" else "Ajudante"
        mapas = ", ".join(g["mapas"])
        hrsai = g["hrsai"]
        hrentr = g["hrentr"]

        if not cad:
            conferencia.append([data_br(data), tipo_txt, cod, "", "", "", mapas,
                                "Codigo nao encontrado no cadastro do Promax"])
            continue

        nome_promax = cad["nome"]
        cpf = cad["cpf"]

        # nome no ponto via CPF
        nome_ponto = ""
        obs = ""
        dia = None
        if not is_cpf_valido(cpf):
            obs = "CPF invalido/generico no Promax - nao da para localizar no ponto"
        else:
            nome_ponto = colab_pontomais.get(cpf, "")
            if not nome_ponto:
                obs = "CPF nao localizado no cadastro do Pontomais"
            else:
                dias = jornada.get(nome_ponto)
                if not dias:
                    obs = "Colaborador sem espelho de jornada no arquivo"
                else:
                    dia = dias.get(data) or (list(dias.values())[0] if len(dias) == 1 else None)
                    if dia is None:
                        obs = "Sem registro de jornada para o dia %s" % data_br(data)

        if dia is None:
            # nao conseguimos tempos do ponto -> vai para conferencia,
            # mas ainda mostramos o que temos (inclusive o nome do Promax)
            conferencia.append([data_br(data), tipo_txt, cod, nome_promax,
                                cpf, nome_ponto, mapas, obs])
            # mesmo assim geramos a linha na aba principal com o que da (rota),
            # deixando os campos de ponto vazios e a observacao preenchida.

        entrada, saida = (entrada_saida_do_dia(dia) if dia else ("", ""))
        intervalo = hhmm(dia.get("intervalo")) if dia else ""

        espera_antes = diff_min(hrsai, entrada)          # saida caminhao - entrada ponto
        tempo_rota = diff_min(hrentr, hrsai)             # entrada caminhao - saida caminhao
        alm_min = to_min(intervalo) or 0
        # jornada liquida = tempo em rota - almoco (precisa do ponto p/ saber o almoco)
        rota_liquida = ((tempo_rota - alm_min)
                        if (dia is not None and tempo_rota is not None) else None)
        espera_apos = diff_min(saida, hrentr)            # saida ponto - entrada caminhao
        jornada_total = None
        if to_min(saida) is not None and to_min(entrada) is not None:
            jornada_total = to_min(saida) - to_min(entrada) - alm_min

        # tempo previsto de rota (2 ART) x rota liquida (real).
        # So compara quando ha ponto valido (dia != None), pois a rota liquida
        # depende do almoco do ponto; sem isso a comparacao seria enganosa.
        tempoprev = g["tprev"] if g["tem_prev"] else None
        dif_prev = None       # rota liquida - previsto  (negativo = fez mais rapido)
        cumpriu = ""
        if dia is not None and tempoprev is not None and rota_liquida is not None:
            dif_prev = rota_liquida - tempoprev
            cumpriu = "Sim" if dif_prev <= 0 else "Nao"

        linha = [
            data_br(data), tipo_txt, cod, nome_promax, cpf, nome_ponto, mapas,
            entrada,                 # Entrada Ponto
            hrsai,                   # Saida Caminhao
            fmt_min(espera_antes),   # Espera antes da rota
            hrentr,                  # Chegada Caminhao
            fmt_min(tempo_rota),     # Tempo em Rota
            intervalo,               # Almoco
            fmt_min(rota_liquida),   # Rota Liquida
            saida,                   # Saida Ponto
            fmt_min(espera_apos),    # Espera apos a rota
            fmt_min(jornada_total),  # Jornada Total
            fmt_min(tempoprev),      # Tempo Previsto
            fmt_min(dif_prev),       # Dif vs Previsto (Rota Liq - Previsto)
            cumpriu,                 # Cumpriu?
            obs,
        ]
        (linhas_mot if tipo == "MOT" else linhas_aju).append(linha)

    return linhas_mot, linhas_aju, conferencia


# ---------------------------------------------------------------------------
# Geracao do Excel
# ---------------------------------------------------------------------------
CABECALHO = [
    "Data", "Tipo", "Codigo", "Nome (Promax)", "CPF", "Nome no Ponto", "Mapa(s)",
    "Entrada Ponto", "Saida Caminhao", "Espera antes da rota",
    "Chegada Caminhao", "Tempo em Rota", "Almoco", "Jornada Liquida",
    "Saida Ponto", "Espera apos a rota", "Jornada Total (ponto)",
    "Tempo Previsto", "Jornada Liq - Previsto", "Cumpriu Previsto", "Observacao",
]

CAB_CONF = ["Data", "Tipo", "Codigo", "Nome (Promax)", "CPF", "Nome no Ponto",
            "Mapa(s)", "Motivo"]


def _estilizar_aba(ws, cabecalho, linhas, cor_cab):
    fill = PatternFill("solid", fgColor=cor_cab)
    font_cab = Font(bold=True, color="FFFFFF", size=10)
    thin = Side(style="thin", color="D0D0D0")
    borda = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.append(cabecalho)
    for j, _ in enumerate(cabecalho, start=1):
        c = ws.cell(row=1, column=j)
        c.fill = fill
        c.font = font_cab
        c.alignment = center
        c.border = borda

    for linha in linhas:
        ws.append(linha)

    # bordas + alinhamento nas celulas de dados
    for i in range(2, ws.max_row + 1):
        for j in range(1, len(cabecalho) + 1):
            c = ws.cell(row=i, column=j)
            c.border = borda
            if j not in (4, 6, 21):  # nomes e observacao ficam a esquerda
                c.alignment = Alignment(horizontal="center", vertical="center")

    # larguras
    larguras = {
        1: 11, 2: 11, 3: 8, 4: 26, 5: 15, 6: 30, 7: 14,
        8: 11, 9: 12, 10: 13, 11: 13, 12: 12, 13: 9, 14: 12,
        15: 11, 16: 13, 17: 12, 18: 13, 19: 14, 20: 9, 21: 40,
    }
    for j in range(1, len(cabecalho) + 1):
        ws.column_dimensions[get_column_letter(j)].width = larguras.get(j, 14)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(cabecalho)), max(ws.max_row, 1))


# colunas da aba Dados (fonte UNICA de graficos e da tabela de detalhe).
# 1-12: numericas (minutos) p/ os graficos.  13+: texto HH:MM p/ a tabela.
DADOS_CAB = [
    "Tipo", "Codigo", "Nome", "Data",
    "Espera antes (min)", "Tempo rota (min)", "Almoco (min)",
    "Rota liquida (min)", "Espera apos (min)", "Jornada (min)",
    "Tempo previsto (min)", "Economia prev (min)",
    # colunas de exibicao (HH:MM) para a tabela de detalhe dinamica
    "Entrada", "Saida Cam", "Espera Antes", "Chegada", "Espera Apos",
    "Jornada Liquida", "Almoco", "Jorn Total", "Tempo Previsto",
    "Dif vs Prev", "Cumpriu",
    # periodos e valores (fracao do dia) para as abas acumuladas
    "Ano", "Mes", "Semana", "JornLiqFrac", "PrevFrac", "CumpriuNum",
    # fracoes do dia p/ os graficos mostrarem HH:MM
    "EspAntesFrac", "EspAposFrac", "EconomiaFrac", "AlmocoFrac",
    # metas (1/0) p/ o painel geral: liberacao (TML) e tempo interno (TI) <= 30min
    "LiberOK", "TIOK",
    # data como serial real do Excel (p/ as formulas do painel Produtividade)
    "DataSerial",
    # flags dos indicadores (regras de produtividade) - col 37+ = AK+
    "VAL_JL", "BATE_JL", "VAL_TML", "BATE_TML", "TR_OK", "TI_OK2", "TMLcapFrac",
]

# nomes dos campos (row fields) da tabela de detalhe dinamica, na ordem exibida
DET_CAMPOS = [
    "Nome", "Data", "Tipo", "Entrada", "Saida Cam", "Espera Antes", "Chegada",
    "Espera Apos", "Jornada Liquida", "Almoco", "Jorn Total", "Tempo Previsto",
    "Dif vs Prev", "Cumpriu",
]


_MESES = ["", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
          "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def _frac(minutos):
    """Minutos -> fracao do dia (para o Excel exibir como [h]:mm). None -> None."""
    return (minutos / 1440.0) if minutos is not None else None


def _data_obj(data_str):
    """'DD/MM/YYYY' -> datetime.date (serial real do Excel) ou None."""
    m = re.match(r"(\d{2})/(\d{2})/(\d{4})", str(data_str or ""))
    if not m:
        return None
    try:
        return _dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def _periodos(data_str):
    """'DD/MM/YYYY' -> (ano, mes 'MM - Mmm/AAAA', semana_do_mes 'Sem N')."""
    m = re.match(r"(\d{2})/(\d{2})/(\d{4})", str(data_str or ""))
    if not m:
        return ("", "", "")
    d = _dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    mes = "%02d - %s/%d" % (d.month, _MESES[d.month], d.year)
    sem_mes = "Semana %d" % (((d.day - 1) // 7) + 1)   # semana do mes (1 a 5)
    return (str(d.year), mes, sem_mes)


def indicadores_rota(linha):
    """Indicadores de uma rota/dia (mesma regra do painel e do warehouse).

    TML = liberacao (espera antes)   TR = tempo em rota   TI = tempo interno
    (espera apos)   JL = TML + TR + TI (jornada bruta).  Metas: JL <= 10:20,
    TML <= 30min, TR <= 09:20, TI <= 30min, rota valida = liberacao entre 0 e 4h.
    Retorna dict com minutos e flags 0/1 (fonte unica p/ Excel e SQL).
    """
    def g(i):
        v = linha[i] if len(linha) > i else ""
        return "" if v is None else v
    tml = _min_assinado(g(_IX_ESP_ANTES))
    tr = _min_assinado(g(_IX_ROTA))
    ti = _min_assinado(g(_IX_ESP_APOS))
    jl = (tml + tr + ti) if None not in (tml, tr, ti) else None
    prev = _min_assinado(g(_IX_TEMPOPREV))
    val_jl = 1 if (tml is not None and 0 < tml <= 240) else 0
    bate_jl = 1 if (val_jl and jl is not None and jl <= 620
                    and ti is not None and ti > 1) else 0
    val_tml = 1 if (tml is not None and 0 < tml < 240) else 0
    bate_tml = 1 if (val_tml and tml <= 30) else 0
    tr_ok = 1 if (val_jl and tr is not None and tr <= 560) else 0
    ti_ok = 1 if (val_jl and ti is not None and ti <= 30) else 0
    tml_cap = tml if (val_tml and tml <= 300) else None
    return {"tml_min": tml, "tr_min": tr, "ti_min": ti, "jl_min": jl,
            "previsto_min": prev, "val_jl": val_jl, "bate_jl": bate_jl,
            "val_tml": val_tml, "bate_tml": bate_tml, "tr_ok": tr_ok,
            "ti_ok": ti_ok, "tml_cap_min": tml_cap}


def _aba_dados(wb, linhas_mot, linhas_aju):
    """Aba plana (fonte da Tabela Dinamica): minutos p/ graficos + HH:MM p/ detalhe."""
    ws = wb.create_sheet("Dados")
    ws.append(DADOS_CAB)
    for l in (linhas_mot + linhas_aju):
        def g(i):
            v = l[i] if len(l) > i else ""
            return "" if v is None else v
        prev = _min_assinado(g(_IX_TEMPOPREV))
        rliq = _min_assinado(g(_IX_ROTA_LIQ))
        dif = _min_assinado(g(_IX_DIFPREV))
        economia = (-dif) if dif is not None else None
        ano, mes, semana = _periodos(g(0))
        cumpriu_txt = str(g(_IX_CUMPRIU))
        cumpriu_num = (1.0 if cumpriu_txt == "Sim"
                       else (0.0 if cumpriu_txt == "Nao" else None))
        esp_antes = _min_assinado(g(_IX_ESP_ANTES))
        esp_apos = _min_assinado(g(_IX_ESP_APOS))
        liber_ok = (1.0 if esp_antes <= 30 else 0.0) if esp_antes is not None else None
        ti_ok = (1.0 if esp_apos <= 30 else 0.0) if esp_apos is not None else None
        # flags dos indicadores (fonte unica: indicadores_rota) p/ o painel
        ind = indicadores_rota(l)
        val_jl, bate_jl = ind["val_jl"], ind["bate_jl"]
        val_tml, bate_tml = ind["val_tml"], ind["bate_tml"]
        tr_ok, ti_ok2 = ind["tr_ok"], ind["ti_ok"]
        tml_cap = _frac(ind["tml_cap_min"])
        ws.append([
            g(1), g(2), g(_IX_NOME) or g(3), g(0),
            _min_assinado(g(_IX_ESP_ANTES)), _min_assinado(g(_IX_ROTA)),
            _min_assinado(g(_IX_ALMOCO)), rliq,
            _min_assinado(g(_IX_ESP_APOS)), _min_assinado(g(_IX_JORNADA)),
            prev, economia,
            # HH:MM (texto) p/ a tabela de detalhe
            str(g(7)), str(g(8)), str(g(_IX_ESP_ANTES)), str(g(10)),
            str(g(_IX_ESP_APOS)), str(g(_IX_ROTA_LIQ)), str(g(_IX_ALMOCO)),
            str(g(_IX_JORNADA)), str(g(_IX_TEMPOPREV)), str(g(_IX_DIFPREV)),
            str(g(_IX_CUMPRIU)),
            # periodos + valores (fracao do dia) p/ acumular
            ano, mes, semana,
            (rliq / 1440.0) if rliq is not None else None,
            (prev / 1440.0) if prev is not None else None,
            cumpriu_num,
            # fracoes p/ os graficos (HH:MM) - negativo nao faz sentido p/ media
            _frac(max(0, esp_antes) if esp_antes is not None else None),
            _frac(max(0, esp_apos) if esp_apos is not None else None),
            _frac(economia),
            _frac(_min_assinado(g(_IX_ALMOCO))),
            # metas (1/0)
            liber_ok, ti_ok,
            # data real (serial do Excel) p/ COUNTIFS/AVERAGEIFS do painel
            _data_obj(g(0)),
            # flags dos indicadores p/ o painel de produtividade
            val_jl, bate_jl, val_tml, bate_tml, tr_ok, ti_ok2, tml_cap,
        ])
    for j, _ in enumerate(DADOS_CAB, start=1):
        ws.column_dimensions[get_column_letter(j)].width = 15
    ws.sheet_state = "hidden"
    return ws


def gerar_excel(caminho_saida, linhas_mot, linhas_aju, conferencia, periodo=""):
    wb = openpyxl.Workbook()

    # 1a aba: Dashboard (fallback estatico; a versao interativa e montada depois via Excel)
    wsd = wb.active
    wsd.title = "Dashboard"

    ws = wb.create_sheet("Motoristas")
    _estilizar_aba(ws, CABECALHO, linhas_mot, "1F4E78")

    ws2 = wb.create_sheet("Ajudantes")
    _estilizar_aba(ws2, CABECALHO, linhas_aju, "548235")

    ws3 = wb.create_sheet("Conferencia")
    _estilizar_aba_conf(ws3, conferencia)

    _aba_dados(wb, linhas_mot, linhas_aju)

    _montar_dashboard(wb, wsd, linhas_mot, linhas_aju, conferencia, periodo)

    wb.save(caminho_saida)


# ---------------------------------------------------------------------------
# Dashboard (aba com graficos dentro do Excel)
# ---------------------------------------------------------------------------
def _min_assinado(s):
    """'HH:MM' ou '-HH:MM' -> minutos (int, pode ser negativo) ou None."""
    if s is None:
        return None
    m = re.match(r"^(-?)(\d{1,2}):(\d{2})$", str(s).strip())
    if not m:
        return None
    v = int(m.group(2)) * 60 + int(m.group(3))
    return -v if m.group(1) == "-" else v


def _media(vals):
    v = [x for x in vals if x is not None]
    return (sum(v) / len(v)) if v else None


def _rotulos_valor():
    """DataLabelList que mostra APENAS o valor (sem nome de serie/categoria)."""
    dl = DataLabelList()
    dl.showVal = True
    dl.showSerName = False
    dl.showCatName = False
    dl.showLegendKey = False
    dl.showPercent = False
    dl.showBubbleSize = False
    return dl


# indices numericos nas linhas (ver CABECALHO)
_IX_NOME, _IX_ESP_ANTES, _IX_ROTA = 5, 9, 11
_IX_ALMOCO, _IX_ROTA_LIQ, _IX_ESP_APOS, _IX_JORNADA = 12, 13, 15, 16
_IX_TEMPOPREV, _IX_DIFPREV, _IX_CUMPRIU = 17, 18, 19


def _registros(linhas):
    out = []
    for l in linhas:
        def g(i):
            return l[i] if len(l) > i else ""
        out.append({
            "nome": g(_IX_NOME) or g(3),
            "espera_antes": _min_assinado(g(_IX_ESP_ANTES)),
            "rota": _min_assinado(g(_IX_ROTA)),
            "almoco": _min_assinado(g(_IX_ALMOCO)),
            "rota_liq": _min_assinado(g(_IX_ROTA_LIQ)),
            "espera_apos": _min_assinado(g(_IX_ESP_APOS)),
            "jornada": _min_assinado(g(_IX_JORNADA)),
        })
    return out


def _montar_dashboard(wb, ws, linhas_mot, linhas_aju, conferencia, periodo):
    mot = _registros(linhas_mot)
    aju = _registros(linhas_aju)
    todos = mot + aju

    azul, verde, cinza = "1F4E78", "548235", "6B7280"

    # ---- titulo ----
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    for col in "BCDEFGHIJKLMNOPQR":
        ws.column_dimensions[col].width = 12

    ws["B2"] = "JORNADA LIQUIDA - DASHBOARD"
    ws["B2"].font = Font(bold=True, size=20, color=azul)
    sub = "Gerado em %s" % _dt.datetime.now().strftime("%d/%m/%Y %H:%M")
    if periodo:
        sub = "Período: %s   ·   %s" % (periodo, sub)
    ws["B3"] = sub
    ws["B3"].font = Font(size=10, color=cinza)

    # ---- KPIs ----
    def hhmm(m):
        return fmt_min(int(round(m))) if m is not None else "-"

    kpis = [
        ("PESSOAS NO DIA", str(len(todos)), "%d motoristas · %d ajudantes" % (len(mot), len(aju)), azul),
        ("MÉDIA JORNADA LÍQUIDA", hhmm(_media([r["jornada"] for r in todos])), "trabalho no dia", azul),
        ("MÉDIA ROTA LÍQUIDA", hhmm(_media([r["rota_liq"] for r in todos])), "rota − almoço", verde),
        ("MÉDIA ESPERA ANTES", hhmm(_media([r["espera_antes"] for r in todos])), "chegar → caminhão sair", azul),
        ("MÉDIA ESPERA APÓS", hhmm(_media([r["espera_apos"] for r in todos])), "caminhão volta → sair", azul),
        ("CASOS P/ CONFERIR", str(len(conferencia)), "aba Conferência", "C00000" if conferencia else verde),
    ]
    # cada KPI ocupa 2 colunas; 6 KPIs a partir da coluna B (B,D,F,H,J,L)
    col0 = 2
    for i, (lbl, val, hint, cor) in enumerate(kpis):
        c = col0 + i * 2
        cl = get_column_letter(c)
        cr = get_column_letter(c + 1)
        ws.merge_cells("%s5:%s5" % (cl, cr))
        ws.merge_cells("%s6:%s6" % (cl, cr))
        ws.merge_cells("%s7:%s7" % (cl, cr))
        cel_l = ws["%s5" % cl]
        cel_l.value = lbl
        cel_l.font = Font(bold=True, size=8, color=cinza)
        cel_l.alignment = Alignment(horizontal="center")
        cel_v = ws["%s6" % cl]
        cel_v.value = val
        cel_v.font = Font(bold=True, size=18, color=cor)
        cel_v.alignment = Alignment(horizontal="center")
        cel_h = ws["%s7" % cl]
        cel_h.value = hint
        cel_h.font = Font(size=8, color=cinza)
        cel_h.alignment = Alignment(horizontal="center")
        fill = PatternFill("solid", fgColor="F2F5F9")
        for rr in (5, 6, 7):
            ws["%s%d" % (cl, rr)].fill = fill
            ws["%s%d" % (cr, rr)].fill = fill

    # ---- planilha auxiliar (fonte dos graficos), oculta ----
    calc = wb.create_sheet("_calc")
    calc.sheet_state = "hidden"

    def escrever_top(titulo_col, campo, dados, linha_ini):
        """Escreve tabela (nome, minutos) dos 10 maiores. Retorna faixas p/ grafico."""
        rows = sorted([r for r in dados if r[campo] is not None],
                      key=lambda r: r[campo], reverse=True)[:10]
        # gravar em ordem crescente para o MAIOR aparecer no TOPO do grafico de barras
        rows = list(reversed(rows))
        calc.cell(row=linha_ini, column=1, value="Nome")
        calc.cell(row=linha_ini, column=2, value=titulo_col)
        for k, r in enumerate(rows, start=1):
            calc.cell(row=linha_ini + k, column=1, value=r["nome"])
            calc.cell(row=linha_ini + k, column=2, value=int(round(r[campo])))
        n = len(rows)
        return linha_ini, n

    def add_bar(anchor, titulo, li, n, cor):
        ch = BarChart()
        ch.type = "bar"
        ch.title = titulo
        ch.height = 8.5
        ch.width = 15
        ch.legend = None
        if n > 0:
            data = Reference(calc, min_col=2, min_row=li, max_row=li + n)
            cats = Reference(calc, min_col=1, min_row=li + 1, max_row=li + n)
            ch.add_data(data, titles_from_data=True)
            ch.set_categories(cats)
            ch.dataLabels = _rotulos_valor()
            try:
                ch.series[0].graphicalProperties.solidFill = cor
            except Exception:
                pass
        ch.x_axis.title = "minutos"
        ch.y_axis.delete = False
        ch.x_axis.delete = False
        ws.add_chart(ch, anchor)

    li1, n1 = escrever_top("Espera antes (min)", "espera_antes", todos, 1)
    li2, n2 = escrever_top("Espera apos (min)", "espera_apos", todos, 20)
    li3, n3 = escrever_top("Jornada liquida (min)", "jornada", todos, 40)

    add_bar("B9", "Maiores esperas para iniciar a rota (min)", li1, n1, azul)
    add_bar("J9", "Maiores esperas após a rota (min)", li2, n2, verde)
    add_bar("B28", "Maiores jornadas líquidas (min)", li3, n3, "2F6FB0")

    # ---- comparativo Motorista x Ajudante ----
    comp_ini = 60
    metrics = [
        ("Jornada líquida", "jornada"), ("Rota líquida", "rota_liq"),
        ("Espera antes", "espera_antes"), ("Espera após", "espera_apos"),
        ("Almoço", "almoco"),
    ]
    calc.cell(row=comp_ini, column=1, value="Métrica")
    calc.cell(row=comp_ini, column=2, value="Motorista")
    calc.cell(row=comp_ini, column=3, value="Ajudante")
    for k, (nome, campo) in enumerate(metrics, start=1):
        calc.cell(row=comp_ini + k, column=1, value=nome)
        mv = _media([r[campo] for r in mot])
        av = _media([r[campo] for r in aju])
        calc.cell(row=comp_ini + k, column=2, value=int(round(mv)) if mv is not None else 0)
        calc.cell(row=comp_ini + k, column=3, value=int(round(av)) if av is not None else 0)

    ch = BarChart()
    ch.type = "col"
    ch.title = "Comparativo Motoristas × Ajudantes (médias, min)"
    ch.height = 8.5
    ch.width = 15
    data = Reference(calc, min_col=2, max_col=3, min_row=comp_ini, max_row=comp_ini + len(metrics))
    cats = Reference(calc, min_col=1, min_row=comp_ini + 1, max_row=comp_ini + len(metrics))
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.dataLabels = _rotulos_valor()
    try:
        ch.series[0].graphicalProperties.solidFill = azul
        ch.series[1].graphicalProperties.solidFill = verde
    except Exception:
        pass
    ch.y_axis.title = "minutos"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ws.add_chart(ch, "J28")


def _estilizar_aba_conf(ws, conferencia):
    fill = PatternFill("solid", fgColor="C00000")
    font_cab = Font(bold=True, color="FFFFFF", size=10)
    thin = Side(style="thin", color="D0D0D0")
    borda = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.append(CAB_CONF)
    for j, _ in enumerate(CAB_CONF, start=1):
        c = ws.cell(row=1, column=j)
        c.fill = fill
        c.font = font_cab
        c.alignment = center
        c.border = borda
    for linha in conferencia:
        ws.append(linha)
    for i in range(2, ws.max_row + 1):
        for j in range(1, len(CAB_CONF) + 1):
            ws.cell(row=i, column=j).border = borda
    larg = {1: 11, 2: 11, 3: 8, 4: 26, 5: 15, 6: 30, 7: 14, 8: 50}
    for j in range(1, len(CAB_CONF) + 1):
        ws.column_dimensions[get_column_letter(j)].width = larg.get(j, 14)
    ws.freeze_panes = "A2"
    if ws.max_row == 1:
        ws.append(["", "", "", "", "", "", "", "Tudo certo! Nenhum caso para conferir."])


# ---------------------------------------------------------------------------
# Fluxo completo (usado pela interface)
# ---------------------------------------------------------------------------
def processar_arquivos(cad_mot, cad_aju, colab, caminho_2art, caminho_jornada):
    """Le os dois arquivos do dia e devolve (linhas_mot, linhas_aju, conferencia)."""
    jornada = carregar_jornada_pontomais(caminho_jornada)
    art = carregar_2art(caminho_2art)
    return processar(cad_mot, cad_aju, colab, jornada, art)


def _top_nomes(linhas, idx, n=10, menor=False):
    """Nomes dos n maiores (ou menores, se menor=True) valores da coluna idx."""
    regs = []
    for l in linhas:
        nome = (l[_IX_NOME] if len(l) > _IX_NOME and l[_IX_NOME]
                else (l[3] if len(l) > 3 else ""))
        v = _min_assinado(l[idx] if len(l) > idx else "")
        if nome and v is not None:
            regs.append((nome, v))
    regs.sort(key=lambda x: x[1], reverse=not menor)
    ordem = []
    for nm, _ in regs:
        if nm not in ordem:
            ordem.append(nm)
        if len(ordem) >= n:
            break
    return ordem


# colunas da tabela de detalhe do dashboard (valores em HH:MM, legiveis)
DET_CAB = ["Nome", "Data", "Tipo", "Entrada Ponto", "Saida Caminhao", "Espera antes",
           "Chegada Caminhao", "Espera apos", "Jornada Liquida", "Almoco", "Jorn Total",
           "Tempo Previsto", "Jornada Liq - Prev", "Cumpriu"]


def _detalhe(linhas_mot, linhas_aju):
    """Linhas da tabela de detalhe (HH:MM), ordenadas pela jornada (maior 1o)."""
    rows = []
    for l in (linhas_mot + linhas_aju):
        def g(i):
            v = l[i] if len(l) > i else ""
            return "" if v is None else str(v)
        nome = g(_IX_NOME) or g(3)
        rows.append([
            nome, g(0), g(1), g(7), g(8), g(_IX_ESP_ANTES), g(10),
            g(_IX_ESP_APOS), g(_IX_ROTA_LIQ), g(_IX_ALMOCO), g(_IX_JORNADA),
            g(_IX_TEMPOPREV), g(_IX_DIFPREV), g(_IX_CUMPRIU),
        ])
    rows.sort(key=lambda r: (_min_assinado(r[10]) if _min_assinado(r[10]) is not None else -10 ** 9),
              reverse=True)
    return rows


def _kpis(linhas_mot, linhas_aju):
    """Cartoes de KPI para o topo do dashboard: lista de (rotulo, valor)."""
    todas = linhas_mot + linhas_aju

    def media_fmt(idx):
        vs = [_min_assinado(l[idx]) for l in todas if len(l) > idx]
        vs = [x for x in vs if x is not None]
        return fmt_min(round(sum(vs) / len(vs))) if vs else "--"

    sim = sum(1 for l in todas if len(l) > _IX_CUMPRIU and l[_IX_CUMPRIU] == "Sim")
    nao = sum(1 for l in todas if len(l) > _IX_CUMPRIU and l[_IX_CUMPRIU] == "Nao")
    tot = sim + nao
    cumpr = ("%d/%d (%d%%)" % (sim, tot, round(100 * sim / tot))) if tot else "--"
    return [
        ("PESSOAS", str(len(todas))),
        ("MEDIA JORNADA LIQ.", media_fmt(_IX_ROTA_LIQ)),   # rota - almoco
        ("MEDIA PREVISTO", media_fmt(_IX_TEMPOPREV)),
        ("ESPERA ANTES", media_fmt(_IX_ESP_ANTES)),
        ("ESPERA APOS", media_fmt(_IX_ESP_APOS)),
        ("CUMPRIRAM PREV.", cumpr),
    ]


def gerar_relatorio(cad_mot, cad_aju, colab, caminho_2art,
                    caminho_jornada, caminho_saida):
    """Gera o relatorio usando cadastros JA carregados (em memoria).

    So le do disco os dois arquivos do dia: o 2 ART e o espelho de jornada.
    Tenta montar o dashboard interativo (Excel); se nao der, mantem o estatico.
    Retorna um dicionario com estatisticas.
    """
    linhas_mot, linhas_aju, conferencia = processar_arquivos(
        cad_mot, cad_aju, colab, caminho_2art, caminho_jornada)

    try:
        gerar_excel(caminho_saida, linhas_mot, linhas_aju, conferencia)
    except PermissionError:
        raise RuntimeError(
            "Nao consegui salvar o arquivo:\n%s\n\n"
            "Ele parece estar ABERTO no Excel. Feche o arquivo (ou escolha "
            "outro nome) e gere o relatorio de novo." % caminho_saida)

    # periodo (data) a partir das linhas
    periodo = ""
    for l in (linhas_mot + linhas_aju):
        if l and l[0]:
            periodo = l[0]
            break

    kpis = _kpis(linhas_mot, linhas_aju)
    detalhe = (DET_CAMPOS, _detalhe(linhas_mot, linhas_aju))

    interativo = False
    try:
        import dashboard_excel
        interativo = dashboard_excel.montar(caminho_saida, periodo, detalhe, kpis)
    except Exception:
        interativo = False

    # camada de engenharia de dados: data warehouse (modelo estrela) + CSV p/ BI
    warehouse_db = None
    dq = None
    try:
        import warehouse
        warehouse_db = os.path.splitext(caminho_saida)[0] + ".db"
        res_dw = warehouse.construir_e_exportar(linhas_mot, linhas_aju, warehouse_db)
        dq = res_dw.get("dq")
    except Exception:
        warehouse_db = None

    return {
        "motoristas": len(linhas_mot),
        "ajudantes": len(linhas_aju),
        "conferencia": len(conferencia),
        "cad_mot": len(cad_mot),
        "cad_aju": len(cad_aju),
        "colaboradores": len(colab),
        "interativo": interativo,
        "warehouse": warehouse_db,
        "dq": dq,
    }


def executar(caminho_mot, caminho_aju, caminho_colab, caminho_2art,
             caminho_jornada, caminho_saida):
    """Versao que le tudo de arquivos (mantida para uso via script/teste)."""
    cad_mot = carregar_cadastro_promax(caminho_mot)
    cad_aju = carregar_cadastro_promax(caminho_aju)
    colab = carregar_colaboradores_pontomais(caminho_colab)
    return gerar_relatorio(cad_mot, cad_aju, colab, caminho_2art,
                           caminho_jornada, caminho_saida)
