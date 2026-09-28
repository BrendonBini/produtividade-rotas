# -*- coding: utf-8 -*-
"""
Gera dados de exemplo 100% sinteticos (sem nenhuma informacao real) para que o
projeto possa ser executado de ponta a ponta por qualquer pessoa:

    cadastros_base.json    -> base de motoristas / ajudantes / colaboradores
    exemplos/2art_exemplo.csv    -> 2 ART (roteirizacao) no formato posicional
    exemplos/ponto_exemplo.xlsx  -> espelho de ponto no formato do relatorio

Os CPFs comecam em 900... de proposito, para deixar claro que sao ficticios.
Rode: python exemplos/gerar_dados_exemplo.py
"""

import csv
import datetime as _dt
import json
import os
import random

import openpyxl

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA = os.path.join(RAIZ, "exemplos")

random.seed(42)

_PRENOMES = ["ANTONIO", "JOSE", "CARLOS", "PAULO", "MARCOS", "LUIZ", "PEDRO",
             "RAFAEL", "BRUNO", "FELIPE", "RODRIGO", "EDUARDO", "FABIO",
             "GUSTAVO", "LEANDRO", "ANDERSON", "TIAGO", "VINICIUS", "MARCELO",
             "DIEGO", "RENATO", "SERGIO", "ROBERTO", "WESLEY", "DOUGLAS"]
_SOBRENOMES = ["SILVA", "SANTOS", "OLIVEIRA", "SOUZA", "PEREIRA", "LIMA",
               "FERREIRA", "COSTA", "RIBEIRO", "ALVES", "GOMES", "MARTINS",
               "ROCHA", "CARVALHO", "ALMEIDA", "NUNES", "MENDES", "BARBOSA"]


def _hhmm(minutos):
    minutos = int(round(minutos)) % (24 * 60)
    return "%02d:%02d" % (minutos // 60, minutos % 60)


def _dias_uteis(inicio, n):
    dias, d = [], inicio
    while len(dias) < n:
        if d.weekday() < 5:      # seg-sex
            dias.append(d)
        d += _dt.timedelta(days=1)
    return dias


def _nomes(qtd, usados):
    saida = []
    while len(saida) < qtd:
        nome = "%s %s %s" % (random.choice(_PRENOMES), random.choice(_SOBRENOMES),
                             random.choice(_SOBRENOMES))
        if nome not in usados:
            usados.add(nome)
            saida.append(nome)
    return saida


def gerar():
    if not os.path.isdir(PASTA):
        os.makedirs(PASTA)
    usados = set()
    seq_cpf = 90000000000

    def novo_cpf():
        nonlocal seq_cpf
        seq_cpf += 137
        return "%011d" % seq_cpf

    # --- cadastros ---
    motoristas, ajudantes, colaboradores = {}, {}, {}
    nomes_mot = _nomes(25, usados)
    nomes_aju = _nomes(15, usados)
    mot_cods, aju_cods = [], []
    for i, nome in enumerate(nomes_mot):
        cod = str(101 + i)
        cpf = novo_cpf()
        motoristas[cod] = {"nome": nome, "cpf": cpf}
        colaboradores[cpf] = nome
        mot_cods.append((cod, cpf, nome))
    for i, nome in enumerate(nomes_aju):
        cod = str(201 + i)
        cpf = novo_cpf()
        ajudantes[cod] = {"nome": nome, "cpf": cpf}
        colaboradores[cpf] = nome
        aju_cods.append((cod, cpf, nome))

    with open(os.path.join(RAIZ, "cadastros_base.json"), "w", encoding="utf-8") as f:
        json.dump({"motoristas": motoristas, "ajudantes": ajudantes,
                   "colaboradores": colaboradores}, f, ensure_ascii=False, indent=2)

    # --- 2 ART (posicional, ';') e ponto ---
    dias = _dias_uteis(_dt.date(2026, 9, 1), 10)
    linhas_2art = []
    # ponto: nome -> lista de (data, entrada, s1, e2, s2, intervalo)
    ponto = {}
    mapa_seq = 390000

    for dia in dias:
        data_br = dia.strftime("%d/%m/%Y")
        # sorteia quais motoristas rodam no dia (a maioria)
        for cod, cpf, nome in mot_cods:
            if random.random() < 0.12:      # folga eventual
                continue
            mapa_seq += 1
            hrsai = random.randint(6 * 60 + 30, 7 * 60 + 30)
            tempo_rota = random.randint(7 * 60, 10 * 60 + 40)
            hrentr = hrsai + tempo_rota
            liberacao = random.randint(8, 48)          # TML
            interno = random.randint(3, 45)            # TI
            almoco = random.randint(45, 70)
            entrada = hrsai - liberacao
            saida = hrentr + interno
            previsto = tempo_rota + random.randint(-40, 40)

            # ajudantes desse mapa (0 a 2)
            ajus = random.sample(aju_cods, random.randint(0, 2))
            row = [""] * 33
            row[0] = data_br
            row[9] = str(mapa_seq)
            row[20] = _hhmm(hrsai)
            row[21] = _hhmm(hrentr)
            row[25] = _hhmm(previsto)
            row[30] = cod
            row[31] = ajus[0][0] if len(ajus) > 0 else ""
            row[32] = ajus[1][0] if len(ajus) > 1 else ""
            linhas_2art.append(row)

            ponto.setdefault(nome, []).append(
                (dia, entrada, entrada + 4 * 60, entrada + 4 * 60 + almoco, saida, almoco))
            for _, _, nome_aju in ajus:
                # ajudante bate ponto parecido com o motorista do mapa
                ea = entrada + random.randint(-10, 10)
                sa = saida + random.randint(-10, 10)
                ponto.setdefault(nome_aju, []).append(
                    (dia, ea, ea + 4 * 60, ea + 4 * 60 + almoco, sa, almoco))

    # escreve 2 ART (latin-1, ';')
    cab = ["Data"] + ["c%d" % i for i in range(1, 33)]
    with open(os.path.join(PASTA, "2art_exemplo.csv"), "w", newline="",
              encoding="latin-1") as f:
        wtr = csv.writer(f, delimiter=";")
        wtr.writerow(cab)
        wtr.writerows(linhas_2art)

    # escreve ponto (xlsx no formato do relatorio de jornada)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Jornada"
    _DIAS = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"]
    for nome, regs in ponto.items():
        ws.append(["Colaborador", nome])
        ws.append(["Data", "Entrada", "Saida", "Entrada", "Saida", "", "",
                   "Intervalo", "", "", "", "", "Obs"])
        for (dia, e1, s1, e2, s2, almoco) in sorted(regs):
            rot = "%s, %s" % (_DIAS[dia.weekday()], dia.strftime("%d/%m/%Y"))
            ws.append([rot, _hhmm(e1), _hhmm(s1), _hhmm(e2), _hhmm(s2), "", "",
                       _hhmm(almoco), "", "", "", "", ""])
        ws.append(["TOTAIS", ""])
    wb.save(os.path.join(PASTA, "ponto_exemplo.xlsx"))

    print("Gerado:")
    print("  cadastros_base.json  (%d motoristas, %d ajudantes, %d colaboradores)"
          % (len(motoristas), len(ajudantes), len(colaboradores)))
    print("  exemplos/2art_exemplo.csv  (%d linhas)" % len(linhas_2art))
    print("  exemplos/ponto_exemplo.xlsx  (%d colaboradores)" % len(ponto))


if __name__ == "__main__":
    gerar()
