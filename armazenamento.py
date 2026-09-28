# -*- coding: utf-8 -*-
"""
Armazenamento local dos cadastros (fica salvo entre as aberturas do programa).

- Base do usuario:   %LOCALAPPDATA%\\JornadaLiquida\\cadastros.json
- Semente embutida:  cadastros_base.json (vai dentro do .exe)

Na PRIMEIRA vez que o programa roda num PC, se ainda nao existe a base do
usuario, ele copia a semente embutida. Assim o .exe ja "vem com" os cadastros
e voce nao precisa importar de novo em outro computador.

Estrutura:
{
  "motoristas":    { "665": {"nome": "...", "cpf": "21454044810"}, ... },
  "ajudantes":     { "196": {"nome": "...", "cpf": "39521528893"}, ... },
  "colaboradores": { "39521528893": "ADELITON JOSE MACHADO", ... }
}
"""

import os
import sys
import json
import shutil


def pasta_dados():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    d = os.path.join(base, "JornadaLiquida")
    os.makedirs(d, exist_ok=True)
    return d


def caminho_dados():
    return os.path.join(pasta_dados(), "cadastros.json")


def _caminho_semente():
    """Caminho da base embutida (dentro do .exe do PyInstaller ou ao lado do .py)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "cadastros_base.json")


def _normalizar(dados):
    dados = dados or {}
    dados.setdefault("motoristas", {})
    dados.setdefault("ajudantes", {})
    dados.setdefault("colaboradores", {})
    return dados


def _ler_json(caminho):
    with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)


def carregar():
    """Le os cadastros salvos. Na primeira vez, semeia da base embutida no .exe."""
    p = caminho_dados()
    if not os.path.isfile(p):
        semente = _caminho_semente()
        if os.path.isfile(semente):
            try:
                shutil.copyfile(semente, p)
            except Exception:
                pass
    dados = {}
    if os.path.isfile(p):
        try:
            dados = _ler_json(p)
        except Exception:
            dados = {}
    return _normalizar(dados)


def salvar(dados):
    """Grava os cadastros no disco."""
    with open(caminho_dados(), "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)


def exportar_para(caminho, dados):
    """Salva uma copia da base num arquivo escolhido (para levar a outro PC)."""
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(_normalizar(dados), f, ensure_ascii=False, indent=2)


def importar_de(caminho):
    """Le uma base de um arquivo .json exportado. Retorna o dicionario normalizado."""
    return _normalizar(_ler_json(caminho))
