# -*- coding: utf-8 -*-
"""
Gera um dashboard de BI web (HTML autocontido) a partir do data warehouse.

Le as views de KPI do SQLite, monta um JSON e injeta em um template com
graficos interativos (Chart.js). O resultado e um unico arquivo .html que roda
em qualquer navegador e pode ser publicado no GitHub Pages, sem backend.

Uso:
    python bi_web.py --db warehouse.db --saida docs/index.html
"""

import datetime as _dt
import json
import os
import sqlite3

_NOMES_MES = ["", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
              "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
_ORDEM_DIA = [("seg", 2), ("ter", 3), ("qua", 4), ("qui", 5),
              ("sex", 6), ("sab", 7), ("dom", 1)]


def _coletar(caminho_db):
    """Le as views/tabelas do DW e devolve o dicionario de dados do dashboard."""
    conn = sqlite3.connect(caminho_db)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    diario = [dict(r) for r in cur.execute(
        "SELECT * FROM vw_kpi_diario ORDER BY data_id")]
    mensal = [dict(r) for r in cur.execute(
        "SELECT * FROM vw_kpi_mensal ORDER BY ano, mes")]
    dq = [dict(r) for r in cur.execute(
        "SELECT verificacao, valor, status, detalhe FROM dq_resultados")]
    conn.close()

    meses = {}
    for m in mensal:
        chave = "%02d/%d" % (m["mes"], m["ano"])
        rotulo = "%s/%d" % (_NOMES_MES[m["mes"]], m["ano"])
        meses[chave] = {
            "rotulo": rotulo,
            "kpi": {
                "real": m["real_jl"], "tml": m["tml_medio_min"],
                "ptml": m["pct_tml"], "ptr": m["pct_tr"], "pti": m["pct_ti"],
                "mvalid": m["mvalid"], "mjl": m["mjl"], "dias": m["dias"],
            },
            "dias": [], "weekday": [], "semana": [],
        }

    for d in diario:
        chave = "%02d/%d" % (d["mes"], d["ano"])
        if chave not in meses:
            continue
        meses[chave]["dias"].append({
            "d": d["data_br"][:5], "mvalid": d["mvalid"], "mjl": d["mjl"],
            "real": d["real_jl"], "tml": d["tml_medio_min"],
            "ptml": d["pct_tml"], "ptr": d["pct_tr"], "pti": d["pct_ti"],
            "wd": d["dia_semana"], "sem": d["semana_iso"],
        })

    # recortes ponderados (soma MJL / soma MVALID), iguais aos do painel
    for chave, mm in meses.items():
        by_wd, by_sem = {}, {}
        for dia in mm["dias"]:
            a = by_wd.setdefault(dia["wd"], [0, 0])
            a[0] += dia["mjl"] or 0
            a[1] += dia["mvalid"] or 0
            b = by_sem.setdefault(dia["sem"], [0, 0])
            b[0] += dia["mjl"] or 0
            b[1] += dia["mvalid"] or 0
        mm["weekday"] = [
            {"nome": nome, "real": (by_wd[wd][0] / by_wd[wd][1]) if wd in by_wd and by_wd[wd][1] else None}
            for nome, wd in _ORDEM_DIA]
        mm["semana"] = [
            {"n": s, "real": (by_sem[s][0] / by_sem[s][1]) if by_sem[s][1] else None}
            for s in sorted(by_sem)]

    return {"meses": meses, "dq": dq,
            "gerado_em": _dt.datetime.now().strftime("%d/%m/%Y %H:%M")}


def gerar_html(caminho_db, caminho_html):
    """Gera o dashboard web a partir do warehouse SQLite."""
    dados = _coletar(caminho_db)
    pasta = os.path.dirname(os.path.abspath(caminho_html))
    if pasta and not os.path.isdir(pasta):
        os.makedirs(pasta)
    html = _TEMPLATE.replace("__DADOS__", json.dumps(dados, ensure_ascii=False))
    with open(caminho_html, "w", encoding="utf-8") as f:
        f.write(html)
    return caminho_html


_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Produtividade de Rotas - BI</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root{
    --bg:#0f1420; --panel:#171d2e; --panel2:#1f2740; --line:#2a3350;
    --tx:#e8ecf6; --mut:#8b95b5; --green:#37c98a; --red:#ff6b7a;
    --orange:#f5993c; --blue:#4aa8ff; --purple:#b98cff; --meta:#e0c060;
    --shadow:0 6px 20px rgba(0,0,0,.35);
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--tx);
    font-family:"Segoe UI",system-ui,Roboto,Arial,sans-serif;padding:24px;}
  h1{margin:0;font-size:26px;letter-spacing:.3px}
  .sub{color:var(--mut);margin:4px 0 18px;font-size:14px}
  .top{display:flex;flex-wrap:wrap;gap:16px;align-items:flex-end;justify-content:space-between}
  select{background:var(--panel2);color:var(--tx);border:1px solid var(--line);
    border-radius:8px;padding:8px 12px;font-size:14px}
  label{color:var(--mut);font-size:12px;text-transform:uppercase;letter-spacing:.6px;display:block;margin-bottom:6px}
  .cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:14px;margin:18px 0}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:12px;
    padding:16px;box-shadow:var(--shadow)}
  .card .rot{color:var(--mut);font-size:12px;text-transform:uppercase;letter-spacing:.6px}
  .card .val{font-size:28px;font-weight:700;margin-top:6px}
  .card .meta{font-size:12px;margin-top:4px;color:var(--mut)}
  .pill{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600}
  .ok{background:rgba(55,201,138,.15);color:var(--green)}
  .bad{background:rgba(255,107,122,.15);color:var(--red)}
  .grid{display:grid;grid-template-columns:2fr 1fr;gap:16px}
  .grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
  .box{background:var(--panel);border:1px solid var(--line);border-radius:12px;
    padding:16px 16px 8px;box-shadow:var(--shadow);margin-bottom:16px}
  .box h3{margin:0 0 10px;font-size:14px;font-weight:600;color:var(--tx)}
  .box .cv{position:relative;height:260px}
  .dq{display:flex;flex-wrap:wrap;gap:8px}
  .dq .item{background:var(--panel2);border:1px solid var(--line);border-radius:8px;
    padding:8px 12px;font-size:12px}
  .dq .item b{display:block;color:var(--mut);font-weight:500;font-size:11px;text-transform:uppercase}
  footer{color:var(--mut);font-size:12px;margin-top:18px}
  @media(max-width:900px){.grid,.grid2{grid-template-columns:1fr}}
</style>
</head>
<body>
  <div class="top">
    <div>
      <h1>Produtividade de Rotas</h1>
      <div class="sub">Gestao da Jornada Liquida - indicadores de motoristas por dia</div>
    </div>
    <div>
      <label for="mes">Periodo</label>
      <select id="mes"></select>
    </div>
  </div>

  <div class="cards" id="cards"></div>

  <div class="grid">
    <div class="box"><h3>JL dentro da meta (REAL) por dia</h3><div class="cv"><canvas id="cReal"></canvas></div></div>
    <div class="box"><h3>Volume por dia (MVALID x MJL)</h3><div class="cv"><canvas id="cVol"></canvas></div></div>
  </div>

  <div class="box"><h3>Metas de tempo por dia: liberacao / rota / interno (%)</h3><div class="cv"><canvas id="cTempos"></canvas></div></div>

  <div class="grid2">
    <div class="box"><h3>JL por dia da semana</h3><div class="cv"><canvas id="cWd"></canvas></div></div>
    <div class="box"><h3>JL por semana</h3><div class="cv"><canvas id="cSem"></canvas></div></div>
  </div>

  <div class="box"><h3>Qualidade de dados (pipeline)</h3><div class="dq" id="dq"></div></div>

  <footer>Fonte: data warehouse (modelo estrela em SQLite) gerado pelo pipeline. Dados de exemplo sinteticos. Atualizado em <span id="ger"></span>.</footer>

<script>
const DADOS = __DADOS__;
const META = {real:0.70, ptml:0.85, ptr:0.70, pti:0.85, tml_min:30};
const C = {green:"#37c98a", red:"#ff6b7a", orange:"#f5993c", blue:"#4aa8ff",
           purple:"#b98cff", meta:"#e0c060", grid:"#2a3350", mut:"#8b95b5"};
Chart.defaults.color = C.mut;
Chart.defaults.font.family = "Segoe UI, system-ui, Arial";
Chart.defaults.plugins.legend.labels.boxWidth = 12;

const pct = v => v==null ? "-" : (v*100).toFixed(1).replace(".",",")+"%";
const hhmm = m => m==null ? "-" : String(Math.floor(m/60)).padStart(1,"0")+":"+String(Math.round(m%60)).padStart(2,"0");
const cor = (v, meta, maior=true) => v==null ? C.mut : ((maior? v>=meta : v<=meta) ? C.green : C.red);

let charts = {};
function novo(id, cfg){ if(charts[id]) charts[id].destroy(); charts[id]=new Chart(document.getElementById(id), cfg); }
function metaLine(valor, label){ return {type:"line", label:label, data:null, borderColor:C.meta,
  borderDash:[6,4], borderWidth:1.5, pointRadius:0, fill:false}; }

function gradeOpt(fmtY){ return {responsive:true, maintainAspectRatio:false,
  scales:{x:{grid:{color:C.grid}}, y:{grid:{color:C.grid}, ticks:{callback:fmtY}}},
  plugins:{legend:{display:true, position:"top"}}}; }

function render(chave){
  const m = DADOS.meses[chave];
  const dias = m.dias;
  const labels = dias.map(d=>d.d);

  // cards
  const k = m.kpi;
  const cards = [
    ["REAL (JL na meta)", pct(k.real), "meta "+pct(META.real), k.real>=META.real],
    ["Caminhoes validos", k.mvalid, k.dias+" dias", true],
    ["Dentro da JL", k.mjl, "de "+k.mvalid+" rotas", true],
    ["TML medio", hhmm(k.tml), "meta "+META.tml_min+" min", (k.tml!=null&&k.tml<=META.tml_min)],
    ["% TML na meta", pct(k.ptml), "meta "+pct(META.ptml), k.ptml>=META.ptml],
    ["% TR na meta", pct(k.ptr), "meta "+pct(META.ptr), k.ptr>=META.ptr],
    ["% TI na meta", pct(k.pti), "meta "+pct(META.pti), k.pti>=META.pti],
  ];
  document.getElementById("cards").innerHTML = cards.map(c=>`
    <div class="card"><div class="rot">${c[0]}</div>
    <div class="val" style="color:${c[3]?C.green:C.red}">${c[1]}</div>
    <div class="meta">${c[2]} <span class="pill ${c[3]?'ok':'bad'}">${c[3]?'na meta':'abaixo'}</span></div></div>`).join("");

  // REAL por dia
  novo("cReal", {data:{labels, datasets:[
      {type:"bar", label:"REAL", data:dias.map(d=>d.real),
       backgroundColor:dias.map(d=>cor(d.real,META.real)), borderRadius:4},
      {type:"line", label:"meta 70%", data:dias.map(()=>META.real),
       borderColor:C.meta, borderDash:[6,4], borderWidth:1.5, pointRadius:0}
    ]}, options:gradeOpt(v=>Math.round(v*100)+"%")});

  // Volume MVALID x MJL
  novo("cVol", {type:"bar", data:{labels, datasets:[
      {label:"MVALID", data:dias.map(d=>d.mvalid), backgroundColor:C.blue, borderRadius:4},
      {label:"MJL", data:dias.map(d=>d.mjl), backgroundColor:C.green, borderRadius:4}
    ]}, options:gradeOpt(v=>v)});

  // %TML / %TR / %TI
  novo("cTempos", {type:"line", data:{labels, datasets:[
      {label:"% TML", data:dias.map(d=>d.ptml), borderColor:C.orange, tension:.3, pointRadius:2},
      {label:"% TR", data:dias.map(d=>d.ptr), borderColor:C.purple, tension:.3, pointRadius:2},
      {label:"% TI", data:dias.map(d=>d.pti), borderColor:C.blue, tension:.3, pointRadius:2}
    ]}, options:gradeOpt(v=>Math.round(v*100)+"%")});

  // por dia da semana
  novo("cWd", {type:"bar", data:{labels:m.weekday.map(w=>w.nome), datasets:[
      {label:"REAL", data:m.weekday.map(w=>w.real),
       backgroundColor:m.weekday.map(w=>cor(w.real,META.real)), borderRadius:4}
    ]}, options:{...gradeOpt(v=>Math.round(v*100)+"%"), plugins:{legend:{display:false}}}});

  // por semana
  novo("cSem", {type:"bar", data:{labels:m.semana.map(s=>"S"+s.n), datasets:[
      {label:"REAL", data:m.semana.map(s=>s.real),
       backgroundColor:m.semana.map(s=>cor(s.real,META.real)), borderRadius:4}
    ]}, options:{...gradeOpt(v=>Math.round(v*100)+"%"), plugins:{legend:{display:false}}}});
}

// DQ
document.getElementById("dq").innerHTML = DADOS.dq.map(d=>`
  <div class="item"><b>${d.verificacao}</b>${d.valor}
  <span class="pill ${d.status==='OK'?'ok':'bad'}">${d.status}</span></div>`).join("");
document.getElementById("ger").textContent = DADOS.gerado_em;

// seletor de mes
const sel = document.getElementById("mes");
const chaves = Object.keys(DADOS.meses);
sel.innerHTML = chaves.map(c=>`<option value="${c}">${DADOS.meses[c].rotulo}</option>`).join("");
sel.value = chaves[chaves.length-1];
sel.onchange = ()=>render(sel.value);
render(sel.value);
</script>
</body>
</html>
"""


def _main():
    import argparse
    ap = argparse.ArgumentParser(description="Gera o dashboard de BI web a partir do data warehouse.")
    ap.add_argument("--db", required=True, help="Arquivo SQLite do data warehouse.")
    ap.add_argument("--saida", default="docs/index.html", help="HTML de saida.")
    args = ap.parse_args()
    caminho = gerar_html(args.db, args.saida)
    print("Dashboard web gerado: %s" % caminho)


if __name__ == "__main__":
    _main()
