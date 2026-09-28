# -*- coding: utf-8 -*-
"""
Jornada Líquida - Interface grafica (com cadastros salvos)

- Os cadastros de Motoristas, Ajudantes e Colaboradores (ponto) ficam SALVOS
  no computador entre as aberturas do programa.
- No dia a dia voce so importa o 2 ART e o Espelho de Jornada e clica em Gerar.
"""

import os
import threading
import traceback
import datetime as _dt

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import jornada_liquida as core
import armazenamento


APP_TITULO = "Jornada Líquida - Analisador de Tempos"
COR_FUNDO = "#f4f6f8"
COR_PRIMARIA = "#1F4E78"
COR_OK = "#548235"
COR_ERRO = "#C00000"

CSV_T = [("Arquivos CSV", "*.csv"), ("Todos", "*.*")]
XLSX_T = [("Arquivos Excel", "*.xlsx"), ("Todos", "*.*")]
ART_T = [("2 ART (CSV ou Excel)", "*.csv;*.xlsx"), ("Todos", "*.*")]


# ===========================================================================
# Aba de cadastro de pessoas (Motorista / Ajudante): codigo, nome, CPF
# ===========================================================================
class AbaPessoas(ttk.Frame):
    def __init__(self, master, app, chave, titulo):
        super().__init__(master, padding=16)
        self.app = app
        self.chave = chave        # "motoristas" ou "ajudantes"
        self.titulo = titulo

        ttk.Label(self, text="Cadastro de %s" % titulo, style="Secao.TLabel").pack(anchor="w")
        ttk.Label(self,
                  text="Importe o cadastro do Promax (CSV) ou inclua manualmente. "
                       "Fica salvo para as proximas vezes.",
                  style="Desc.TLabel").pack(anchor="w", pady=(0, 8))

        # --- barra de acoes ---
        barra = ttk.Frame(self)
        barra.pack(fill="x")
        ttk.Button(barra, text="Importar cadastro (CSV do Promax)...",
                   command=self.importar).pack(side="left")
        ttk.Button(barra, text="Remover selecionado",
                   command=self.remover).pack(side="left", padx=(8, 0))
        ttk.Button(barra, text="Limpar tudo",
                   command=self.limpar_tudo).pack(side="left", padx=(8, 0))
        self.lbl_total = ttk.Label(barra, text="", style="Campo.TLabel")
        self.lbl_total.pack(side="right")

        # --- formulario de inclusao/edicao manual ---
        form = ttk.LabelFrame(self, text=" Incluir / editar manualmente ", padding=10)
        form.pack(fill="x", pady=10)
        ttk.Label(form, text="Código:").grid(row=0, column=0, sticky="w")
        self.e_cod = ttk.Entry(form, width=12)
        self.e_cod.grid(row=0, column=1, padx=(4, 16), sticky="w")
        ttk.Label(form, text="Nome:").grid(row=0, column=2, sticky="w")
        self.e_nome = ttk.Entry(form, width=34)
        self.e_nome.grid(row=0, column=3, padx=(4, 16), sticky="w")
        ttk.Label(form, text="CPF:").grid(row=0, column=4, sticky="w")
        self.e_cpf = ttk.Entry(form, width=18)
        self.e_cpf.grid(row=0, column=5, padx=(4, 16), sticky="w")
        ttk.Button(form, text="Adicionar / Salvar",
                   command=self.salvar_manual).grid(row=0, column=6)
        ttk.Button(form, text="Limpar campos",
                   command=self.limpar_campos).grid(row=0, column=7, padx=(6, 0))

        # --- lista ---
        cx = ttk.Frame(self)
        cx.pack(fill="both", expand=True)
        cols = ("codigo", "nome", "cpf")
        self.tree = ttk.Treeview(cx, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("codigo", text="Código")
        self.tree.heading("nome", text="Nome")
        self.tree.heading("cpf", text="CPF")
        self.tree.column("codigo", width=90, anchor="center")
        self.tree.column("nome", width=340, anchor="w")
        self.tree.column("cpf", width=160, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(cx, command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind("<<TreeviewSelect>>", self._ao_selecionar)

        self.recarregar()

    # ---- dados ----
    def _dic(self):
        return self.app.dados[self.chave]

    def recarregar(self):
        self.tree.delete(*self.tree.get_children())
        dic = self._dic()
        for cod in sorted(dic.keys(), key=lambda c: (int(c) if c.isdigit() else 9e9, c)):
            reg = dic[cod]
            self.tree.insert("", "end", iid=cod,
                             values=(cod, reg.get("nome", ""), reg.get("cpf", "")))
        self.lbl_total.configure(text="%d cadastrados" % len(dic))
        self.app.atualizar_resumo()

    # ---- acoes ----
    def importar(self):
        cam = filedialog.askopenfilename(title="Cadastro do Promax (CSV)", filetypes=CSV_T)
        if not cam:
            return
        try:
            novos = core.carregar_cadastro_promax(cam)
        except Exception as e:
            messagebox.showerror("Erro ao importar", str(e))
            return
        dic = self._dic()
        dic.update(novos)  # adiciona/atualiza por codigo
        self.app.salvar()
        self.recarregar()
        messagebox.showinfo("Importado",
                            "%d registros importados/atualizados.\nTotal agora: %d."
                            % (len(novos), len(dic)))

    def salvar_manual(self):
        cod = self.e_cod.get().strip().lstrip("0") or self.e_cod.get().strip()
        nome = core.norm_txt(self.e_nome.get())
        cpf = core.cpf11(self.e_cpf.get())
        if not cod:
            messagebox.showwarning("Falta o código", "Informe o código.")
            return
        if not nome:
            messagebox.showwarning("Falta o nome", "Informe o nome.")
            return
        if cpf and len(cpf) != 11:
            if not messagebox.askyesno("CPF incompleto",
                                       "O CPF nao tem 11 digitos. Salvar mesmo assim?"):
                return
        self._dic()[cod] = {"nome": nome, "cpf": cpf}
        self.app.salvar()
        self.recarregar()
        self.limpar_campos()

    def remover(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Selecione", "Selecione um registro na lista para remover.")
            return
        cod = sel[0]
        if messagebox.askyesno("Remover", "Remover o código %s do cadastro?" % cod):
            self._dic().pop(cod, None)
            self.app.salvar()
            self.recarregar()

    def limpar_tudo(self):
        if not self._dic():
            return
        if messagebox.askyesno("Limpar tudo",
                               "Apagar TODOS os %d cadastros de %s?"
                               % (len(self._dic()), self.titulo)):
            self.app.dados[self.chave] = {}
            self.app.salvar()
            self.recarregar()

    def limpar_campos(self):
        self.e_cod.delete(0, "end")
        self.e_nome.delete(0, "end")
        self.e_cpf.delete(0, "end")
        self.e_cod.focus_set()

    def _ao_selecionar(self, _evt):
        sel = self.tree.selection()
        if not sel:
            return
        cod = sel[0]
        reg = self._dic().get(cod, {})
        self.limpar_campos()
        self.e_cod.insert(0, cod)
        self.e_nome.insert(0, reg.get("nome", ""))
        self.e_cpf.insert(0, reg.get("cpf", ""))


# ===========================================================================
# Aba de colaboradores do ponto: CPF -> nome
# ===========================================================================
class AbaColaboradores(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=16)
        self.app = app

        ttk.Label(self, text="Colaboradores do Ponto (Pontomais)", style="Secao.TLabel").pack(anchor="w")
        ttk.Label(self,
                  text="Liga o CPF ao nome que aparece no espelho de ponto. Importe o "
                       "'Relatorio de Colaboradores' (com CPF) ou inclua manualmente. Fica salvo.",
                  style="Desc.TLabel").pack(anchor="w", pady=(0, 8))

        barra = ttk.Frame(self)
        barra.pack(fill="x")
        ttk.Button(barra, text="Importar colaboradores (Excel do Pontomais)...",
                   command=self.importar).pack(side="left")
        ttk.Button(barra, text="Remover selecionado",
                   command=self.remover).pack(side="left", padx=(8, 0))
        ttk.Button(barra, text="Limpar tudo",
                   command=self.limpar_tudo).pack(side="left", padx=(8, 0))
        self.lbl_total = ttk.Label(barra, text="", style="Campo.TLabel")
        self.lbl_total.pack(side="right")

        form = ttk.LabelFrame(self, text=" Incluir / editar manualmente ", padding=10)
        form.pack(fill="x", pady=10)
        ttk.Label(form, text="CPF:").grid(row=0, column=0, sticky="w")
        self.e_cpf = ttk.Entry(form, width=18)
        self.e_cpf.grid(row=0, column=1, padx=(4, 16), sticky="w")
        ttk.Label(form, text="Nome (igual ao do ponto):").grid(row=0, column=2, sticky="w")
        self.e_nome = ttk.Entry(form, width=40)
        self.e_nome.grid(row=0, column=3, padx=(4, 16), sticky="w")
        ttk.Button(form, text="Adicionar / Salvar",
                   command=self.salvar_manual).grid(row=0, column=4)
        ttk.Button(form, text="Limpar campos",
                   command=self.limpar_campos).grid(row=0, column=5, padx=(6, 0))

        cx = ttk.Frame(self)
        cx.pack(fill="both", expand=True)
        cols = ("cpf", "nome")
        self.tree = ttk.Treeview(cx, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("cpf", text="CPF")
        self.tree.heading("nome", text="Nome no Ponto")
        self.tree.column("cpf", width=160, anchor="center")
        self.tree.column("nome", width=420, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(cx, command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind("<<TreeviewSelect>>", self._ao_selecionar)

        self.recarregar()

    def _dic(self):
        return self.app.dados["colaboradores"]

    def recarregar(self):
        self.tree.delete(*self.tree.get_children())
        dic = self._dic()
        for cpf in sorted(dic.keys(), key=lambda c: dic[c]):
            self.tree.insert("", "end", iid=cpf, values=(cpf, dic[cpf]))
        self.lbl_total.configure(text="%d cadastrados" % len(dic))
        self.app.atualizar_resumo()

    def importar(self):
        cam = filedialog.askopenfilename(title="Colaboradores (Excel do Pontomais)", filetypes=XLSX_T)
        if not cam:
            return
        try:
            novos = core.carregar_colaboradores_pontomais(cam)
        except Exception as e:
            messagebox.showerror("Erro ao importar", str(e))
            return
        if not novos:
            messagebox.showwarning("Nada importado",
                                   "Nao encontrei CPFs nesse arquivo.\n\n"
                                   "Confira se voce exportou o 'Relatorio de Colaboradores' "
                                   "do Pontomais COM a coluna CPF marcada.")
            return
        self._dic().update(novos)
        self.app.salvar()
        self.recarregar()
        messagebox.showinfo("Importado",
                            "%d colaboradores importados/atualizados.\nTotal agora: %d."
                            % (len(novos), len(self._dic())))

    def salvar_manual(self):
        cpf = core.cpf11(self.e_cpf.get())
        nome = core.norm_txt(self.e_nome.get())
        if len(cpf) != 11:
            messagebox.showwarning("CPF invalido", "Informe um CPF com 11 digitos.")
            return
        if not nome:
            messagebox.showwarning("Falta o nome", "Informe o nome.")
            return
        self._dic()[cpf] = nome
        self.app.salvar()
        self.recarregar()
        self.limpar_campos()

    def remover(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Selecione", "Selecione um registro para remover.")
            return
        cpf = sel[0]
        if messagebox.askyesno("Remover", "Remover o CPF %s?" % cpf):
            self._dic().pop(cpf, None)
            self.app.salvar()
            self.recarregar()

    def limpar_tudo(self):
        if not self._dic():
            return
        if messagebox.askyesno("Limpar tudo",
                               "Apagar TODOS os %d colaboradores?" % len(self._dic())):
            self.app.dados["colaboradores"] = {}
            self.app.salvar()
            self.recarregar()

    def limpar_campos(self):
        self.e_cpf.delete(0, "end")
        self.e_nome.delete(0, "end")
        self.e_cpf.focus_set()

    def _ao_selecionar(self, _evt):
        sel = self.tree.selection()
        if not sel:
            return
        cpf = sel[0]
        self.limpar_campos()
        self.e_cpf.insert(0, cpf)
        self.e_nome.insert(0, self._dic().get(cpf, ""))


# ===========================================================================
# Aba de geracao do relatorio
# ===========================================================================
class AbaGerar(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=16)
        self.app = app
        self.c_2art = tk.StringVar()
        self.c_jorn = tk.StringVar()

        ttk.Label(self, text="Gerar Relatório do Dia", style="Secao.TLabel").pack(anchor="w")
        ttk.Label(self,
                  text="Com os cadastros ja salvos, importe apenas os dois arquivos do dia "
                       "e clique em Gerar.",
                  style="Desc.TLabel").pack(anchor="w", pady=(0, 8))

        # resumo dos cadastros salvos
        self.lbl_resumo = ttk.Label(self, text="", style="Campo.TLabel", justify="left")
        self.lbl_resumo.pack(anchor="w", pady=(0, 10))

        cx = ttk.LabelFrame(self, text=" Arquivos do dia ", padding=12)
        cx.pack(fill="x")
        self._campo(cx, "3. 2 ART do dia (CSV ou Excel do Promax)", self.c_2art, ART_T, 0)
        self._campo(cx, "5. Espelho de Jornada / ponto (Excel do Pontomais)", self.c_jorn, XLSX_T, 1)

        self.btn = ttk.Button(self, text="Gerar Relatório em Excel",
                              style="Gerar.TButton", command=self._gerar)
        self.btn.pack(fill="x", ipady=8, pady=(12, 0))

        ttk.Label(self, text="Resultado:", style="Campo.TLabel").pack(anchor="w", pady=(12, 2))
        lc = ttk.Frame(self)
        lc.pack(fill="both", expand=True)
        self.log = tk.Text(lc, height=10, wrap="word", font=("Consolas", 9),
                           bg="#ffffff", fg="#222222", relief="solid", bd=1)
        self.log.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(lc, command=self.log.yview)
        sb.pack(side="right", fill="y")
        self.log.configure(yscrollcommand=sb.set, state="disabled")

        self.atualizar_resumo()

    def _campo(self, parent, titulo, var, tipos, linha):
        ttk.Label(parent, text=titulo, style="Campo.TLabel").grid(
            row=linha * 2, column=0, columnspan=2, sticky="w", pady=(4, 0))
        e = ttk.Entry(parent, textvariable=var)
        e.grid(row=linha * 2 + 1, column=0, sticky="we", ipady=3, pady=(0, 6))
        ttk.Button(parent, text="Selecionar...",
                   command=lambda: self._pick(var, tipos)).grid(
            row=linha * 2 + 1, column=1, padx=(6, 0), pady=(0, 6))
        parent.columnconfigure(0, weight=1)

    def _pick(self, var, tipos):
        cam = filedialog.askopenfilename(title="Selecione o arquivo", filetypes=tipos)
        if cam:
            var.set(cam)

    def atualizar_resumo(self):
        d = self.app.dados
        self.lbl_resumo.configure(
            text="Cadastros salvos:   Motoristas: %d      Ajudantes: %d      Colaboradores (ponto): %d"
                 % (len(d["motoristas"]), len(d["ajudantes"]), len(d["colaboradores"])))

    # ---- log ----
    def _esc(self, txt, cor=None):
        self.log.configure(state="normal")
        if cor:
            tag = "c_" + cor
            self.log.tag_configure(tag, foreground=cor)
            self.log.insert("end", txt + "\n", tag)
        else:
            self.log.insert("end", txt + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    def _limpar(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _gerar(self):
        d = self.app.dados
        if not d["motoristas"] and not d["ajudantes"]:
            messagebox.showwarning("Sem cadastro",
                                   "Nao ha motoristas nem ajudantes cadastrados.\n\n"
                                   "Va nas abas de cadastro e importe (ou inclua) primeiro.")
            return
        if not d["colaboradores"]:
            messagebox.showwarning("Sem colaboradores",
                                   "Nao ha colaboradores do ponto cadastrados.\n\n"
                                   "Va na aba 'Colaboradores (Ponto)' e importe primeiro.")
            return
        cam_2art = self.c_2art.get().strip()
        cam_jorn = self.c_jorn.get().strip()
        for nome, cam in (("2 ART", cam_2art), ("Espelho de Jornada", cam_jorn)):
            if not cam:
                messagebox.showwarning("Falta arquivo", "Selecione o arquivo de %s." % nome)
                return
            if not os.path.isfile(cam):
                messagebox.showerror("Nao encontrado", "Arquivo de %s nao existe:\n%s" % (nome, cam))
                return

        saida = filedialog.asksaveasfilename(
            title="Salvar relatorio como",
            defaultextension=".xlsx",
            initialfile="Jornada_Liquida_%s.xlsx" % _dt.datetime.now().strftime("%Y-%m-%d"),
            filetypes=[("Arquivo Excel", "*.xlsx")])
        if not saida:
            return

        self.btn.configure(state="disabled")
        self._limpar()
        self._esc("Processando, aguarde...\n")
        threading.Thread(target=self._rodar, args=(cam_2art, cam_jorn, saida), daemon=True).start()

    def _rodar(self, cam_2art, cam_jorn, saida):
        try:
            d = self.app.dados
            stats = core.gerar_relatorio(d["motoristas"], d["ajudantes"], d["colaboradores"],
                                         cam_2art, cam_jorn, saida)
            self.after(0, self._ok, stats, saida)
        except Exception as e:
            self.after(0, self._erro, e, traceback.format_exc())

    def _ok(self, stats, saida):
        self._esc("Concluido com sucesso!\n", COR_OK)
        self._esc("  Motoristas: %d" % stats["motoristas"], COR_PRIMARIA)
        self._esc("  Ajudantes:  %d" % stats["ajudantes"], COR_OK)
        conf = stats["conferencia"]
        self._esc("  Conferencia (verificar): %d" % conf, COR_ERRO if conf else COR_OK)
        self._esc("")
        if stats.get("interativo"):
            self._esc("Dashboard interativo (filtros + graficos dinamicos) montado.", COR_OK)
        else:
            self._esc("Dashboard estatico gerado (Excel nao disponivel p/ o interativo).", COR_PRIMARIA)
        self._esc("")
        self._esc("Salvo em:")
        self._esc("  " + saida)
        if stats.get("warehouse"):
            self._esc("")
            self._esc("Data warehouse (BI) gerado:", COR_OK)
            self._esc("  " + stats["warehouse"])
            self._esc("  + CSVs para BI em: " + os.path.splitext(stats["warehouse"])[0] + "_bi")
            alertas = [d for d in (stats.get("dq") or []) if d[2] != "OK"]
            if alertas:
                self._esc("  Qualidade de dados: %d alerta(s) - ver aba dq_resultados" % len(alertas),
                          COR_ERRO)
            else:
                self._esc("  Qualidade de dados: todas as checagens OK", COR_OK)
        self.btn.configure(state="normal")
        if messagebox.askyesno("Pronto!",
                               "Relatorio gerado!\n\nMotoristas: %d   Ajudantes: %d   Conferencia: %d\n\n"
                               "Abrir agora?" % (stats["motoristas"], stats["ajudantes"], conf)):
            try:
                os.startfile(saida)
            except Exception:
                pass

    def _erro(self, e, tb):
        self._esc("ERRO ao processar:", COR_ERRO)
        self._esc(str(e), COR_ERRO)
        self._esc("")
        self._esc(tb)
        self.btn.configure(state="normal")
        messagebox.showerror("Erro", "Ocorreu um erro:\n\n%s" % e)


# ===========================================================================
# Janela principal
# ===========================================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITULO)
        self.geometry("900x720")
        self.minsize(820, 660)
        self.configure(bg=COR_FUNDO)

        self.dados = armazenamento.carregar()
        self._estilos()
        self._menu()

        topo = ttk.Frame(self, padding=(20, 16, 20, 6))
        topo.pack(fill="x")
        ttk.Label(topo, text="Jornada Líquida", style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(topo,
                  text="Cadastros ficam salvos. No dia a dia, use a aba 'Gerar Relatório' "
                       "com o 2 ART e o espelho de ponto.",
                  style="Sub.TLabel").pack(anchor="w")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=(6, 12))

        self.aba_gerar = AbaGerar(nb, self)
        self.aba_mot = AbaPessoas(nb, self, "motoristas", "Motoristas")
        self.aba_aju = AbaPessoas(nb, self, "ajudantes", "Ajudantes")
        self.aba_colab = AbaColaboradores(nb, self)

        nb.add(self.aba_gerar, text="  Gerar Relatório  ")
        nb.add(self.aba_mot, text="  Motoristas  ")
        nb.add(self.aba_aju, text="  Ajudantes  ")
        nb.add(self.aba_colab, text="  Colaboradores (Ponto)  ")

    def _menu(self):
        barra = tk.Menu(self)
        m = tk.Menu(barra, tearoff=0)
        m.add_command(label="Exportar base (para levar a outro PC)...",
                      command=self._exportar_base)
        m.add_command(label="Importar base de um arquivo...",
                      command=self._importar_base)
        barra.add_cascade(label="Base de cadastros", menu=m)
        self.config(menu=barra)

    def _exportar_base(self):
        cam = filedialog.asksaveasfilename(
            title="Exportar base de cadastros",
            defaultextension=".json",
            initialfile="cadastros_jornada_liquida.json",
            filetypes=[("Base (JSON)", "*.json")])
        if not cam:
            return
        try:
            armazenamento.exportar_para(cam, self.dados)
            messagebox.showinfo(
                "Base exportada",
                "Base salva em:\n%s\n\nCopie esse arquivo para o outro PC e use "
                "'Importar base' la." % cam)
        except Exception as e:
            messagebox.showerror("Erro", "Nao consegui exportar:\n%s" % e)

    def _importar_base(self):
        cam = filedialog.askopenfilename(
            title="Importar base de cadastros",
            filetypes=[("Base (JSON)", "*.json"), ("Todos", "*.*")])
        if not cam:
            return
        if not messagebox.askyesno(
                "Importar base",
                "Isto vai SUBSTITUIR os cadastros atuais deste PC pelos do arquivo.\n\n"
                "Deseja continuar?"):
            return
        try:
            self.dados = armazenamento.importar_de(cam)
            self.salvar()
            self.aba_mot.recarregar()
            self.aba_aju.recarregar()
            self.aba_colab.recarregar()
            self.atualizar_resumo()
            messagebox.showinfo(
                "Base importada",
                "Cadastros carregados:\nMotoristas: %d   Ajudantes: %d   Colaboradores: %d"
                % (len(self.dados["motoristas"]), len(self.dados["ajudantes"]),
                   len(self.dados["colaboradores"])))
        except Exception as e:
            messagebox.showerror("Erro", "Nao consegui importar:\n%s" % e)

    def _estilos(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure("TFrame", background=COR_FUNDO)
        st.configure("TLabel", background=COR_FUNDO)
        st.configure("TLabelframe", background=COR_FUNDO)
        st.configure("TLabelframe.Label", background=COR_FUNDO, foreground=COR_PRIMARIA)
        st.configure("Titulo.TLabel", foreground=COR_PRIMARIA, font=("Segoe UI", 18, "bold"))
        st.configure("Sub.TLabel", foreground="#555555", font=("Segoe UI", 10))
        st.configure("Secao.TLabel", foreground=COR_PRIMARIA, font=("Segoe UI", 12, "bold"))
        st.configure("Campo.TLabel", foreground="#222222", font=("Segoe UI", 10, "bold"))
        st.configure("Desc.TLabel", foreground="#888888", font=("Segoe UI", 9))
        st.configure("Gerar.TButton", font=("Segoe UI", 12, "bold"))

    def salvar(self):
        armazenamento.salvar(self.dados)

    def atualizar_resumo(self):
        if hasattr(self, "aba_gerar"):
            self.aba_gerar.atualizar_resumo()


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
