# -*- coding: utf-8 -*-
"""
Monta um dashboard INTERATIVO dentro do Excel (Tabela Dinamica + Segmentacao/
Slicers + Graficos Dinamicos), via automacao do Excel (pywin32/COM).

Layout:
  - Filtros a esquerda:  Slicer "Tipo" (afeta tudo) e Slicer "Nome" (afeta a
    tabela de detalhe).
  - Graficos a direita (2x2): Top 10 espera antes, Top 10 espera apos,
    Top 10 jornada, e Comparativo Motorista x Ajudante.
  - Tabela de detalhe (Nome x metricas) embaixo, ordenavel/filtravel.

Precisa do Microsoft Excel instalado. Se nao houver (ou der erro), o chamador
mantem o dashboard estatico de fallback.

Funcao principal:  montar(caminho_xlsx, periodo="", tops=None) -> bool
    tops = {"antes":[nomes...], "apos":[...], "jornada":[...]}  (Top 10 por metrica)
"""

import os


# constantes do Excel
xlDatabase = 1
xlRowField = 1
xlAverage = -4106
xlTopCount = 1                # tipo de filtro "10 Maiores" (PivotFilters.Add2)
xlHideNoData = 4             # SlicerCache.CrossFilterType: esconde itens sem dados
xlPageField = 3              # campo de pagina (filtro) da dinamica
xlSum = -4157
xlCount = -4112
xlDescending = 2
xlUp = -4162
xlColumnClustered = 51
xlBarClustered = 57
xlCategory = 1
xlValue = 2
xlLow = -4134


def _disponivel():
    try:
        import win32com.client  # noqa
        return True
    except Exception:
        return False


def _rgb(r, g, b):
    return r + (g << 8) + (b << 16)


# paleta escura (estilo BI)
BG = _rgb(24, 16, 34)        # fundo geral (roxo bem escuro)
PANEL = _rgb(40, 28, 56)     # paineis / cards
HEADER = _rgb(66, 46, 96)    # cabecalho da tabela de detalhe
ORANGE = _rgb(245, 140, 40)
RED = _rgb(232, 58, 74)
GREEN = _rgb(90, 200, 130)
TXT = _rgb(236, 236, 242)    # texto claro
MUT = _rgb(150, 150, 170)    # texto suave
GRID = _rgb(70, 60, 90)      # linhas de grade discretas


def montar(caminho_xlsx, periodo="", detalhe=None, kpis=None):
    """Constroi o dashboard interativo. Retorna True se conseguiu, False senao.
    detalhe = (cabecalho, linhas) da tabela de detalhe (HH:MM)."""
    if not _disponivel():
        return False
    det_cab, det_linhas = (detalhe if detalhe else ([], []))
    kpis = kpis or []

    import pythoncom
    import win32com.client as w

    pythoncom.CoInitialize()
    xl = None
    wb = None
    try:
        xl = w.dynamic.Dispatch("Excel.Application")
        for prop in ("DisplayAlerts", "ScreenUpdating", "EnableEvents"):
            try:
                setattr(xl, prop, False)
            except Exception:
                pass
        try:
            xl.Calculation = -4135  # xlCalculationManual: acelera muito a montagem
        except Exception:
            pass

        wb = xl.Workbooks.Open(os.path.abspath(caminho_xlsx))

        for nm in ("Dashboard", "_calc", "Pivots"):
            try:
                wb.Worksheets(nm).Delete()
            except Exception:
                pass

        dados = wb.Worksheets("Dados")
        dados.Visible = True
        last = dados.Cells(dados.Rows.Count, 1).End(xlUp).Row
        if last < 2:
            return False
        lastcol = dados.Cells(1, dados.Columns.Count).End(-4159).Column  # xlToLeft
        rng = dados.Range(dados.Cells(1, 1), dados.Cells(last, lastcol))

        piv = wb.Worksheets.Add()
        piv.Name = "Pivots"
        dash = wb.Worksheets.Add()
        dash.Name = "Dashboard"
        dash.Move(Before=wb.Worksheets(1))

        cache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=rng)

        # ---- PAINEL de produtividade: indicadores de JL por dia ----
        _montar_produtividade(wb, dados, dash, pythoncom, periodo)
        piv.Visible = False

        # ---- abas ACUMULADAS por pessoa (Ano / Mes / Semana / Diario) ----
        try:
            _montar_acumulados(wb, cache, pythoncom)
        except Exception:
            if os.environ.get("JL_DEBUG"):
                import traceback
                traceback.print_exc()

        try:
            dash.Range("A1").Select()
            xl.ActiveWindow.DisplayGridlines = False
        except Exception:
            pass
        try:
            xl.Calculation = -4105  # volta p/ automatico antes de salvar
        except Exception:
            pass

        wb.Save()
        wb.Close(SaveChanges=False)
        wb = None
        return True
    except Exception:
        if os.environ.get("JL_DEBUG"):
            import traceback
            traceback.print_exc()
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
                wb = None
        except Exception:
            pass
        return False
    finally:
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _montar_acumulados(wb, cache, pythoncom):
    """Cria 4 abas de acumulado por pessoa (Ano / Mes / Semana / Diario), cada
    uma com uma Tabela Dinamica (Nome x totais) e slicer(s) para escolher o
    periodo. Compartilha o mesmo cache dos dados; slicers ligados so ao seu pivot.
    """
    abas = [
        ("Acumulado Ano", ["Ano"],
         "ACUMULADO POR ANO - selecione o ANO no filtro"),
        ("Acumulado Mes", ["Ano", "Mes"],
         "ACUMULADO POR MES - selecione ANO e MES no filtro"),
        ("Acumulado Semana", ["Ano", "Mes", "Semana"],
         "ACUMULADO POR SEMANA DO MES - selecione ANO, MES e SEMANA no filtro"),
        ("Por Dia", ["Ano", "Mes", "Data"],
         "POR DIA - selecione ANO, MES e DIA no filtro"),
    ]
    for nome, campos, titulo in abas:
        ws = wb.Worksheets.Add()
        try:
            ws.Name = nome
        except Exception:
            pass
        try:
            ws.Range("A1").Value = titulo
            ws.Range("A1").Font.Size = 13
            ws.Range("A1").Font.Bold = True
            ws.Range("A2").Value = ("Jornada Liq / Previsto = TOTAL de horas no periodo. "
                                    "Cumpriu % = dias em que a jornada liq ficou <= previsto.")
            ws.Range("A2").Font.Size = 9
        except Exception:
            pass

        pt = cache.CreatePivotTable(
            TableDestination=ws.Range("A16"),
            TableName="pt" + nome.replace(" ", ""))
        try:
            pt.ManualUpdate = True
        except Exception:
            pass
        pt.PivotFields("Nome").Orientation = xlRowField
        d1 = pt.AddDataField(pt.PivotFields("Data"), "Dias", xlCount)
        d2 = pt.AddDataField(pt.PivotFields("JornLiqFrac"), "Jornada Liq (total)", xlSum)
        d3 = pt.AddDataField(pt.PivotFields("PrevFrac"), "Previsto (total)", xlSum)
        d4 = pt.AddDataField(pt.PivotFields("CumpriuNum"), "Cumpriu %", xlAverage)
        try:
            d2.NumberFormat = '[h]"h"mm"min"'   # horas cheias: 425h20min
            d3.NumberFormat = '[h]"h"mm"min"'
            d4.NumberFormat = "0%"
        except Exception:
            pass
        try:
            pt.RowAxisLayout(1)       # tabular
            pt.ColumnGrand = False
        except Exception:
            pass
        try:
            pt.PivotFields("Nome").AutoSort(xlDescending, "Jornada Liq (total)")
        except Exception:
            pass
        try:
            pt.TableStyle2 = "PivotStyleMedium9"
        except Exception:
            pass
        try:
            pt.ManualUpdate = False
        except Exception:
            pass
        try:
            ws.Columns(1).ColumnWidth = 32
            for j in range(2, 6):
                ws.Columns(j).ColumnWidth = 15
        except Exception:
            pass

        # slicers de periodo (no topo)
        left = 10
        for campo in campos:
            try:
                sc = wb.SlicerCaches.Add2(pt, campo)
                try:
                    sc.CrossFilterType = xlHideNoData  # esconde itens fora do que ja foi filtrado
                except Exception:
                    pass
                sc.Slicers.Add(ws, pythoncom.Empty,
                               "Sl_" + nome.replace(" ", "") + "_" + campo,
                               campo, 45, left, 150, 150)
                left += 160
            except Exception:
                pass


# cores tema claro (layout de indicadores diarios)
CL_GREEN_FILL = _rgb(198, 239, 206)
CL_GREEN_FONT = _rgb(0, 97, 0)
CL_RED_FILL = _rgb(255, 199, 206)
CL_RED_FONT = _rgb(156, 0, 6)
CL_HDR = _rgb(31, 56, 100)          # cabecalho azul escuro
CL_HDRTX = _rgb(255, 255, 255)
CL_TITLE = _rgb(197, 90, 17)        # laranja do titulo
CL_SUB = _rgb(31, 78, 120)          # azul do subtitulo
CL_YELLOW = _rgb(255, 242, 204)     # caixa do seletor de mes
CL_SUMFILL = _rgb(217, 225, 242)    # celulas de resumo
CL_METAFILL = _rgb(242, 242, 242)   # linha de metas

xlExpression = 2


def _montar_produtividade(wb, dados, dash, pythoncom, periodo):
    """Painel de produtividade (Gestao da Jornada Liquida):
    tabela dia-a-dia (so motoristas/caminhoes) dirigida por formulas, com
    seletor de MES, resumo no topo (metas x realizado) e mini-tabelas de JL
    por Weekday / WeekNum / Road. Interativo: troca o mes no dropdown e tudo
    recalcula via COUNTIFS/AVERAGEIFS sobre a aba Dados.
    """
    # separador decimal local (pt-BR=','); os codigos de NumberFormat neste
    # Excel sao interpretados no idioma local -> montar "0,00%" corretamente
    try:
        dec = dash.Application.International(3)  # xlDecimalSeparator
    except Exception:
        dec = ","
    pct2 = "0" + dec + "00%"     # ex.: "0,00%"

    # meses disponiveis (coluna 25 = "Mes") p/ o dropdown
    lastd = dados.Cells(dados.Rows.Count, 1).End(xlUp).Row
    meses = []
    try:
        if lastd >= 3:
            vals = dados.Range(dados.Cells(2, 25), dados.Cells(lastd, 25)).Value
            meses = sorted({row[0] for row in vals if row and row[0]})
        else:
            v = dados.Cells(2, 25).Value
            if v:
                meses = [v]
    except Exception:
        pass
    if not meses:
        meses = [""]

    # ---------- cabecalho / titulo ----------
    dash.Range("A1").Value = "Produtividade de Rotas"
    dash.Range("A3").Value = "Gestao da Jornada Liquida - indicadores diarios"
    dash.Range("A5").Value = "Resultado JL Mês"
    dash.Range("A6").Value = ">>Mês<<"
    try:
        dash.Range("A1").Font.Size = 26
        dash.Range("A1").Font.Bold = True
        dash.Range("A1").Font.Color = CL_TITLE
        dash.Range("A3").Font.Size = 14
        dash.Range("A3").Font.Bold = True
        dash.Range("A3").Font.Color = CL_SUB
        dash.Range("A5").Font.Bold = True
        dash.Range("A5").Font.Color = CL_SUB
        dash.Range("A6").Font.Size = 9
    except Exception:
        pass

    # ---------- seletor de mes (dropdown) ----------
    # lista de meses escondida na coluna X; validacao aponta p/ ela
    # IMPORTANTE: formatar como TEXTO senao o Excel converte "07 - Jul/2026" em data
    try:
        dash.Columns(24).NumberFormat = "@"
        dash.Range("B6").NumberFormat = "@"
    except Exception:
        pass
    for i, m in enumerate(meses, start=1):
        dash.Cells(i, 24).Value = m       # coluna X = 24
    sel = dash.Range("B6")
    sel.Value = meses[-1]
    try:
        sel.Interior.Color = CL_YELLOW
        sel.Font.Bold = True
        sel.Font.Size = 12
        sel.HorizontalAlignment = -4108   # center
        sel.Borders.LineStyle = 1
        dv = sel.Validation
        try:
            dv.Delete()
        except Exception:
            pass
        dv.Add(3, 1, 1, "=$X$1:$X$%d" % len(meses))   # xlValidateList
    except Exception:
        pass

    # helpers Ano / MesNum a partir do texto "MM - Mmm/AAAA" (coluna Z, escondida)
    dash.Range("Z1").Formula = '=IFERROR(VALUE(RIGHT($B$6,4)),YEAR(TODAY()))'
    dash.Range("Z2").Formula = '=IFERROR(VALUE(LEFT($B$6,2)),1)'

    # ---------- linha de resumo (metas x realizado) ----------
    dash.Range("D8").Value = "Meta"
    dash.Range("D9").Formula = "=0.7"
    dash.Range("E9").Formula = "=SUM(E12:E42)"
    dash.Range("F9").Formula = "=SUM(F12:F42)"
    dash.Range("G9").Formula = '=IFERROR(AVERAGE(G12:G42),"-")'
    dash.Range("I9").Formula = '=IFERROR(AVERAGE(I12:I42),"-")'
    dash.Range("J9").Formula = '=IFERROR(AVERAGE(J12:J42),"-")'
    dash.Range("K9").Formula = '=IFERROR(AVERAGE(K12:K42),"-")'   # TR = % rota <= 9:20
    dash.Range("N9").Formula = '=IFERROR(AVERAGE(N12:N42),"-")'
    # metas (linha 10): REAL 70% | TML 30min | %TML 85% | TR 70% | TI 30min(limiar)
    dash.Range("G10").Formula = "=0.7"
    dash.Range("I10").Formula = "=TIME(0,30,0)"
    dash.Range("J10").Formula = "=0.85"
    dash.Range("K10").Formula = "=0.7"
    dash.Range("N10").Formula = "=TIME(0,30,0)"
    try:
        dash.Range("D8").Font.Bold = True
        for a in ("D9", "E9", "F9", "G9", "I9", "J9", "K9", "N9"):
            dash.Range(a).Interior.Color = CL_SUMFILL
            dash.Range(a).Font.Bold = True
        for a in ("G10", "I10", "J10", "K10", "N10"):
            dash.Range(a).Interior.Color = CL_METAFILL
            dash.Range(a).Font.Color = _rgb(120, 120, 120)
        dash.Range("D9").NumberFormat = "0%"
        dash.Range("E9:F9").NumberFormat = "0"
        dash.Range("G9").NumberFormat = pct2
        dash.Range("I9").NumberFormat = "[h]:mm:ss"
        dash.Range("J9").NumberFormat = pct2
        dash.Range("K9").NumberFormat = pct2
        dash.Range("N9").NumberFormat = pct2
        dash.Range("G10").NumberFormat = "0%"
        dash.Range("I10").NumberFormat = "[h]:mm:ss"
        dash.Range("J10").NumberFormat = "0%"
        dash.Range("K10").NumberFormat = "0%"
        dash.Range("N10").NumberFormat = "[h]:mm:ss"
    except Exception:
        pass

    # ---------- cabecalho da tabela (linha 11) ----------
    heads = ["Sem", "Dia", "Dia Sem", "META", "MVALID", "MJL", "REAL", "VAR",
             "TML", "%TML", "TR", "PNP", "REPASSE", "TI"]
    for j, h in enumerate(heads, start=1):
        c = dash.Cells(11, j)
        c.Value = h
        try:
            c.Interior.Color = CL_HDR
            c.Font.Color = CL_HDRTX
            c.Font.Bold = True
            c.HorizontalAlignment = -4108
        except Exception:
            pass

    # ---------- formulas dia-a-dia (linhas 12..42) ----------
    linhas = []
    for r in range(12, 43):
        d = r - 11
        O = "$O%d" % r
        E = "$E%d" % r
        F = "$F%d" % r
        G = "$G%d" % r
        D = "$D%d" % r
        base = 'Dados!$A:$A,"Motorista",Dados!$AJ:$AJ,%s' % O

        # colunas de flags no Dados: AK=VAL_JL AL=BATE_JL AM=VAL_TML
        # AN=BATE_TML AO=TR_OK AP=TI_OK2 AQ=TMLcapFrac
        def somaif(col):                         # soma um flag (0/1) do dia
            return '=IF(%s="","",SUMIFS(Dados!%s,%s))' % (O, col, base)

        def media(col):                          # media (fracao do dia) -> HH:MM
            return ('=IF(%s="","",IF(%s=0,"-",IFERROR(AVERAGEIFS(Dados!%s,%s),"-")))'
                    % (O, E, col, base))

        def razao(num, den):                     # % = soma(num)/soma(den)
            return ('=IF(%s="","",IF(%s=0,"-",IFERROR('
                    'SUMIFS(Dados!%s,%s)/SUMIFS(Dados!%s,%s),"-")))'
                    % (O, E, num, base, den, base))

        linhas.append((
            '=IF(%s="","",ISOWEEKNUM(%s))' % (O, O),
            '=IF(%s="","",TEXT(DAY(%s),"00")&"/"&TEXT(MONTH(%s),"00")&"/"&$Z$1)' % (O, O, O),
            '=IF(%s="","",WEEKDAY(%s))' % (O, O),
            '=IF(%s="","",0.7)' % O,
            somaif("$AK:$AK"),                   # E MVALID = SUM(VAL_JL)
            somaif("$AL:$AL"),                   # F MJL    = SUM(BATE_JL)
            '=IF(%s="","",IF(%s=0,"-",%s/%s))' % (O, E, F, E),   # G REAL = MJL/MVALID
            '=IF(%s="","",IF(%s=0,"-",%s-%s))' % (O, E, G, D),   # H VAR
            media("$AQ:$AQ"),                    # I TML  = media liberacao (cap 5h)
            razao("$AN:$AN", "$AM:$AM"),         # J %TML = BATE_TML / VAL_TML
            razao("$AO:$AO", "$AK:$AK"),         # K TR   = TR_OK / VAL_JL  (% <= 9:20)
            "",                                  # L PNP
            "",                                  # M REPASSE
            razao("$AP:$AP", "$AK:$AK"),         # N TI   = TI_OK2 / VAL_JL (% <= 30min)
            '=IF(MONTH(DATE($Z$1,$Z$2,%d))=$Z$2,DATE($Z$1,$Z$2,%d),"")' % (d, d),
        ))
    dash.Range("A12:O42").Formula = tuple(linhas)

    # formatos das colunas
    try:
        dash.Range("A12:A42").NumberFormat = "0"
        dash.Range("B12:B42").NumberFormat = "@"
        dash.Range("C12:C42").NumberFormat = "0"
        dash.Range("D12:D42").NumberFormat = "0%"
        dash.Range("E12:F42").NumberFormat = "0"
        dash.Range("G12:H42").NumberFormat = pct2
        dash.Range("I12:I42").NumberFormat = "[h]:mm:ss"
        dash.Range("J12:J42").NumberFormat = pct2
        dash.Range("K12:K42").NumberFormat = pct2          # TR = % rota <= 9:20
        dash.Range("N12:N42").NumberFormat = pct2
        dash.Range("A12:N42").HorizontalAlignment = -4108   # center
        dash.Range("B12:B42").HorizontalAlignment = -4108
    except Exception:
        pass

    # ---------- formatacao condicional (verde cumpre / vermelho falha) ----------
    # ATENCAO 1: em late-binding argumentos NOMEADOS nao funcionam -> posicionais.
    # ATENCAO 2: FormatConditions.Add interpreta a formula na LINGUA LOCAL do
    # Excel. Em pt-BR: AND->E, ISNUMBER->ENUM, separador ',' -> ';'. Tentamos a
    # versao US e, se falhar, a pt-BR (robusto em qualquer idioma).
    def cf(addr, op_green, op_red, meta):
        rng = dash.Range(addr)
        top = addr.split(":")[0]                 # ex.: "G12"
        for op, fill, font in ((op_green, CL_GREEN_FILL, CL_GREEN_FONT),
                               (op_red, CL_RED_FILL, CL_RED_FONT)):
            variantes = (
                '=AND(ISNUMBER(%s),%s%s%s)' % (top, top, op, meta),
                '=E(ÉNÚM(%s);%s%s%s)' % (top, top, op, meta),
            )
            for f in variantes:
                try:
                    fc = rng.FormatConditions.Add(xlExpression, pythoncom.Empty, f)
                    fc.Interior.Color = fill
                    fc.Font.Color = font
                    break
                except Exception:
                    continue

    cf("G12:G42", ">=", "<", "$G$10")     # REAL vs meta 70%
    cf("H12:H42", ">=", "<", "0")         # VAR: positivo verde
    cf("I12:I42", "<=", ">", "$I$10")     # TML: menor que 30min verde
    cf("J12:J42", ">=", "<", "$J$10")     # %TML vs meta 85%
    cf("K12:K42", ">=", "<", "$K$10")     # TR vs meta 70%
    cf("N12:N42", ">=", "<", "$J$10")     # TI vs meta 85%

    # ---------- mini-tabelas de JL (Weekday / WeekNum / Road) ----------
    def mini_head(rr, titulo):
        a = dash.Cells(rr, 17)   # coluna Q
        b = dash.Cells(rr, 18)   # coluna R
        a.Value = titulo
        b.Value = "JL"
        for c in (a, b):
            try:
                c.Interior.Color = CL_HDR
                c.Font.Color = CL_HDRTX
                c.Font.Bold = True
            except Exception:
                pass

    # Dia da semana
    mini_head(11, "Dia Sem")
    dias = [("seg", 2), ("ter", 3), ("qua", 4), ("qui", 5),
            ("sex", 6), ("sáb", 7), ("dom", 1)]
    for k, (nome, wd) in enumerate(dias):
        rr = 12 + k
        dash.Cells(rr, 17).Value = nome
        # ponderado: soma MJL / soma MVALID daquele dia da semana (igual a planilha)
        dash.Cells(rr, 18).Formula = (
            '=IFERROR(SUMIFS($F$12:$F$42,$C$12:$C$42,%d)/'
            'SUMIFS($E$12:$E$42,$C$12:$C$42,%d),"-")' % (wd, wd))
        dash.Cells(rr, 18).NumberFormat = pct2

    # Semana (numero)
    mini_head(20, "Semana")
    for k in range(6):
        rr = 21 + k
        dash.Cells(rr, 17).Formula = (
            '=IFERROR(ISOWEEKNUM(DATE($Z$1,$Z$2,1))+%d,"")' % k)
        dash.Cells(rr, 17).NumberFormat = "0"
        # ponderado por numero da semana ISO
        dash.Cells(rr, 18).Formula = (
            '=IFERROR(SUMIFS($F$12:$F$42,$A$12:$A$42,$Q%d)/'
            'SUMIFS($E$12:$E$42,$A$12:$A$42,$Q%d),"-")' % (rr, rr))
        dash.Cells(rr, 18).NumberFormat = pct2

    # Revenda (total ponderado) + totais
    mini_head(28, "Revenda")
    dash.Cells(28, 19).Value = "MJL"
    dash.Cells(28, 20).Value = "MVALID"
    try:
        for cc in (19, 20):
            dash.Cells(28, cc).Interior.Color = CL_HDR
            dash.Cells(28, cc).Font.Color = CL_HDRTX
            dash.Cells(28, cc).Font.Bold = True
    except Exception:
        pass
    dash.Cells(29, 17).Value = 1
    dash.Cells(29, 18).Formula = '=IFERROR(SUM($F$12:$F$42)/SUM($E$12:$E$42),"-")'
    dash.Cells(29, 18).NumberFormat = pct2
    dash.Cells(29, 19).Formula = "=SUM($F$12:$F$42)"
    dash.Cells(29, 20).Formula = "=SUM($E$12:$E$42)"

    # cor verde/vermelho tambem nas mini-tabelas de JL
    cf("R12:R18", ">=", "<", "$G$10")
    cf("R21:R26", ">=", "<", "$G$10")
    cf("R29:R29", ">=", "<", "$G$10")

    # ---------- larguras / esconder colunas auxiliares ----------
    try:
        larg = {1: 6, 2: 11, 3: 7, 4: 7, 5: 8, 6: 7, 7: 9, 8: 9, 9: 10,
                10: 8, 11: 7, 12: 6, 13: 9, 14: 8, 17: 10, 18: 9, 19: 8, 20: 9}
        for col, w in larg.items():
            dash.Columns(col).ColumnWidth = w
        dash.Columns(15).Hidden = True   # O = data auxiliar
        dash.Columns(24).Hidden = True   # X = lista de meses
        dash.Columns(26).Hidden = True   # Z = Ano/MesNum
    except Exception:
        pass
    try:
        dash.Range("A12:N42").Borders.LineStyle = 1
        dash.Range("A12:N42").Borders.Color = _rgb(200, 200, 200)
    except Exception:
        pass
