#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import os
import time
import argparse
import importlib.util

# ----------------------------------------------------------------------
# Verificação de dependências ANTES dos imports pesados
# ----------------------------------------------------------------------
DEPENDENCIAS = {
    "tkinter":    "Interface gráfica (Tcl/Tk)",
    "pdfplumber": "Leitura de PDF",
    "PIL":        "Renderização de imagem (Pillow)",
}

def verificar_dependencias():
    faltando = []
    for modulo, descricao in DEPENDENCIAS.items():
        if importlib.util.find_spec(modulo) is None:
            faltando.append((modulo, descricao))
    if not faltando:
        return True
    print("\n[ERRO] Dependências ausentes:\n")
    for modulo, descricao in faltando:
        print(f"  • {modulo} — {descricao}")
    print("\nComo resolver:")
    print("  pip install pdfplumber pillow")
    print("  Para tkinter no Windows: reinstale o Python marcando 'tcl/tk and IDLE'.")
    print("  Para tkinter no Linux:   sudo apt install python3-tk")
    return False

def verificar_openpyxl():
    if importlib.util.find_spec("openpyxl") is None:
        print("\n[ERRO] openpyxl não encontrado — necessário para exportar XLSX.")
        print("  pip install openpyxl")
        return False
    return True

if not verificar_dependencias():
    sys.exit(1)

import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import pdfplumber
from PIL import Image, ImageTk
import re
import csv

# ----------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------
def safe_print(*args, **kwargs):
    try:
        print(*args, **kwargs)
    except Exception:
        pass

def _corpo_mascara(mascara):
    regex = re.escape(mascara)
    regex = re.sub(r'[dD]{2}', r'\d{2}', regex)
    regex = re.sub(r'[yY]{4}|[aA]{4}', r'\d{4}', regex)
    regex = re.sub(r'[yY]{2}|[aA]{2}', r'\d{2}', regex)
    regex = re.sub(r'[mM]{4}', r'[a-zA-ZáàâãéèêíïóôõúüçÇÁÀÂÃÉÈÊÍÏÓÔÕÚÜ]{4,}', regex)
    regex = re.sub(r'[mM]{3}', r'[a-zA-ZáàâãéèêíïóôõúüçÇÁÀÂÃÉÈÊÍÏÓÔÕÚÜ]{3}', regex)
    regex = re.sub(r'[mM]{2}', r'\d{2}', regex)
    return regex

def converter_mascaras_para_regex(mascaras_str):
    partes = [m.strip() for m in re.split(r'[,;]', mascaras_str) if m.strip()]
    if not partes:
        return None
    if len(partes) == 1:
        return re.compile(r'^' + _corpo_mascara(partes[0]))
    corpos = [_corpo_mascara(m) for m in partes]
    return re.compile(r'^(?:' + '|'.join(corpos) + ')')

REGEX_DATA_AUTO = re.compile(
    r'^(?:'
    r'\d{1,2}[/.\-]\d{1,2}(?:[/.\-]\d{2,4})?'
    r'|\d{1,2}\s+(?:seg|ter|qua|qui|sex|s[áa]b|dom)\b'
    r'|\d{1,2}\s+de\s+\S+'
    r'|\d{1,2}\s*[-–—]'
    r'|\d{1,2}(?=\s|$)'
    r')',
    re.IGNORECASE | re.UNICODE
)

def parse_paginas(pag_str):
    pag_inicio = 1
    pag_fim = 999999
    pag_str = (pag_str or "").strip()
    if pag_str and pag_str.lower() != "1-end":
        if "-" in pag_str:
            partes = pag_str.split("-")
            pag_inicio = int(partes[0])
            if partes[1].lower() != "end" and partes[1] != "":
                pag_fim = int(partes[1])
        else:
            pag_inicio = int(pag_str)
            pag_fim = int(pag_str)
    return pag_inicio, pag_fim

def inferir_formato(caminho):
    return "xlsx" if caminho.lower().endswith(".xlsx") else "csv"

def garantir_extensao(caminho, formato):
    ext = ".xlsx" if formato == "xlsx" else ".csv"
    caminho_lower = caminho.lower()
    if caminho_lower.endswith(ext):
        return caminho
    for outra in (".csv", ".xlsx"):
        if caminho_lower.endswith(outra):
            caminho = caminho[:-len(outra)]
            break
    return caminho + ext

# ----------------------------------------------------------------------
# Seletor de arquivo customizado (ordenação cronológica)
# ----------------------------------------------------------------------
class FilePickerDialog(tk.Toplevel):
    def __init__(self, parent, title="Selecionar arquivo",
                 initialdir=None, filetypes=None,
                 sort_by="date", descending=True):
        super().__init__(parent)
        self.title(title)
        self.geometry("880x560")
        self.minsize(720, 420)
        self.transient(parent)
        self.grab_set()

        self.filetypes = filetypes or [("Todos os arquivos", "*")]
        self.selected_path = None
        self.sort_by = sort_by
        self.descending = descending
        self.current_dir = os.path.abspath(initialdir or os.path.expanduser("~"))

        top = tk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=6)
        tk.Button(top, text="↑ Subir", command=self.go_up, width=9).pack(side=tk.LEFT)
        self.lbl_path = tk.Label(top, text="", anchor="w", fg="#333", bg="#f0f0f0", relief=tk.SUNKEN)
        self.lbl_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)

        filtro_frame = tk.Frame(self)
        filtro_frame.pack(fill=tk.X, padx=8, pady=(0, 6))
        tk.Label(filtro_frame, text="Filtrar por nome:").pack(side=tk.LEFT)
        self.entry_filtro = tk.Entry(filtro_frame, width=30)
        self.entry_filtro.pack(side=tk.LEFT, padx=6)
        self.entry_filtro.bind("<KeyRelease>", lambda e: self.refresh())

        tree_frame = tk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=8)
        cols = ("name", "modified", "size")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("name", text="Nome", command=lambda: self.set_sort("name"))
        self.tree.heading("modified", text="Modificado", command=lambda: self.set_sort("date"))
        self.tree.heading("size", text="Tamanho", command=lambda: self.set_sort("size"))
        self.tree.column("name", width=460, anchor="w")
        self.tree.column("modified", width=170, anchor="center")
        self.tree.column("size", width=110, anchor="e")
        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Return>", lambda e: self.confirm())

        bot = tk.Frame(self)
        bot.pack(fill=tk.X, padx=8, pady=10)
        tk.Button(bot, text="Cancelar", command=self.cancel, width=14).pack(side=tk.RIGHT, padx=4)
        tk.Button(bot, text="Abrir", command=self.confirm, width=14, bg="#2e7d32", fg="white", font=("Arial", 10, "bold")).pack(side=tk.RIGHT, padx=4)

        self.refresh()
        self.wait_window()

    def extensions_match(self, filename):
        for _, pattern in self.filetypes:
            if pattern in ("*", "*.*"):
                return True
            for ext in pattern.split(";"):
                ext = ext.strip()
                if ext.startswith("*.") and filename.lower().endswith(ext[1:].lower()):
                    return True
        return False

    def set_sort(self, key):
        if self.sort_by == key:
            self.descending = not self.descending
        else:
            self.sort_by = key
            self.descending = (key != "name")
        self.refresh()

    def refresh(self):
        self.lbl_path.config(text=self.current_dir)
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        try:
            entries = list(os.scandir(self.current_dir))
        except PermissionError:
            messagebox.showerror("Erro", "Sem permissão para acessar esta pasta.")
            return

        filtro = self.entry_filtro.get().strip().lower() if hasattr(self, "entry_filtro") else ""
        dirs, files = [], []
        for e in entries:
            try:
                if e.name.startswith("."): continue
                if filtro and filtro not in e.name.lower(): continue
                if e.is_dir():
                    dirs.append((e.name, e.stat().st_mtime, 0))
                elif e.is_file() and self.extensions_match(e.name):
                    st = e.stat()
                    files.append((e.name, st.st_mtime, st.st_size))
            except OSError:
                continue

        def key(item):
            name, mtime, size = item
            if self.sort_by == "name": return name.lower()
            if self.sort_by == "size": return size
            return mtime

        dirs.sort(key=key, reverse=self.descending)
        files.sort(key=key, reverse=self.descending)
        for name, _, _ in dirs:
            self.tree.insert("", "end", values=("📁 " + name, "", ""))
        for name, mtime, size in files:
            mtime_str = time.strftime("%d/%m/%Y %H:%M", time.localtime(mtime))
            self.tree.insert("", "end", values=(name, mtime_str, self._fmt_size(size)))

        arrow = " ▼" if self.descending else " ▲"
        self.tree.heading("name", text="Nome" + (arrow if self.sort_by == "name" else ""))
        self.tree.heading("modified", text="Modificado" + (arrow if self.sort_by == "date" else ""))
        self.tree.heading("size", text="Tamanho" + (arrow if self.sort_by == "size" else ""))

    @staticmethod
    def _fmt_size(n):
        for unit in ("B", "KB", "MB", "GB"):
            if n < 1024: return f"{n:.1f} {unit}"
            n /= 1024
        return f"{n:.1f} TB"

    def go_up(self):
        parent = os.path.dirname(self.current_dir)
        if parent and parent != self.current_dir:
            self.current_dir = parent
            self.refresh()

    def on_double_click(self, event):
        sel = self.tree.selection()
        if not sel: return
        name = self.tree.item(sel[0], "values")[0]
        if name.startswith("📁 "):
            self.current_dir = os.path.join(self.current_dir, name[2:])
            self.refresh()
        else:
            self.confirm()

    def confirm(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione um arquivo.")
            return
        name = self.tree.item(sel[0], "values")[0]
        if name.startswith("📁 "):
            self.current_dir = os.path.join(self.current_dir, name[2:])
            self.refresh()
            return
        self.selected_path = os.path.join(self.current_dir, name)
        self.destroy()

    def cancel(self):
        self.selected_path = None
        self.destroy()

# ----------------------------------------------------------------------
# NOVO: Janela para ordenar as colunas
# ----------------------------------------------------------------------
class ColumnOrderDialog(tk.Toplevel):
    def __init__(self, parent, nomes_colunas):
        super().__init__(parent)
        self.title("Definir Ordem das Colunas")
        self.geometry("380x420")
        self.minsize(380, 420)
        self.transient(parent)
        self.grab_set()
        self.resultado = None

        tk.Label(self, text="Defina a ordem de saída (a primeira da lista\nserá a coluna 1 do arquivo):",
                 fg="blue", font=("Arial", 10, "bold")).pack(pady=10)

        frame_lista = tk.Frame(self)
        frame_lista.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        self.listbox = tk.Listbox(frame_lista, selectmode=tk.SINGLE, exportselection=False, font=("Arial", 11))
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll = tk.Scrollbar(frame_lista, orient=tk.VERTICAL, command=self.listbox.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scroll.set)

        self.indices = list(range(len(nomes_colunas)))
        for nome in nomes_colunas:
            self.listbox.insert(tk.END, nome)

        frame_botoes = tk.Frame(self)
        frame_botoes.pack(fill=tk.X, padx=15, pady=15)

        tk.Button(frame_botoes, text="↑ Subir", command=self.subir, width=10, font=("Arial", 10)).pack(side=tk.LEFT, padx=2)
        tk.Button(frame_botoes, text="↓ Descer", command=self.descer, width=10, font=("Arial", 10)).pack(side=tk.LEFT, padx=2)
        tk.Button(frame_botoes, text="✔ Confirmar", command=self.confirmar, bg="#2e7d32", fg="white", width=12, font=("Arial", 10, "bold")).pack(side=tk.RIGHT, padx=2)
        tk.Button(frame_botoes, text="✖ Cancelar", command=self.destroy, width=10, font=("Arial", 10)).pack(side=tk.RIGHT, padx=2)

        self.wait_window()

    def subir(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] == 0: return
        idx = sel[0]
        texto = self.listbox.get(idx)
        indice_orig = self.indices.pop(idx)
        self.listbox.delete(idx)
        self.listbox.insert(idx - 1, texto)
        self.indices.insert(idx - 1, indice_orig)
        self.listbox.selection_set(idx - 1)

    def descer(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] == self.listbox.size() - 1: return
        idx = sel[0]
        texto = self.listbox.get(idx)
        indice_orig = self.indices.pop(idx)
        self.listbox.delete(idx)
        self.listbox.insert(idx + 1, texto)
        self.indices.insert(idx + 1, indice_orig)
        self.listbox.selection_set(idx + 1)

    def confirmar(self):
        self.resultado = self.indices
        self.destroy()

# ----------------------------------------------------------------------
# Janela de configuração inicial
# ----------------------------------------------------------------------
class ConfigWindow:
    def __init__(self, root):
        self.root = root
        self.resultado = None

        self.root.title("Extrator PDF → CSV/XLSX - Configuração")
        self.root.geometry("920x430")
        self.root.resizable(False, False)

        tk.Label(root, text="Configure a extração antes de escolher as colunas:", font=("Arial", 12, "bold"), fg="#333").grid(row=0, column=0, columnspan=3, padx=10, pady=(15, 10), sticky="w")

        tk.Label(root, text="Arquivo PDF:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.entry_pdf = tk.Entry(root, width=60)
        self.entry_pdf.grid(row=1, column=1, padx=5, pady=5, sticky="we")
        tk.Button(root, text="Procurar...", command=self.escolher_pdf, width=12).grid(row=1, column=2, padx=5, pady=5)

        tk.Label(root, text="Arquivo de saída:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.entry_csv = tk.Entry(root, width=60)
        self.entry_csv.grid(row=2, column=1, padx=5, pady=5, sticky="we")
        tk.Button(root, text="Salvar como...", command=self.escolher_csv, width=12).grid(row=2, column=2, padx=5, pady=5)

        tk.Label(root, text="Formato de saída:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        frame_formato = tk.Frame(root)
        frame_formato.grid(row=3, column=1, columnspan=2, sticky="w", padx=5, pady=5)
        self.var_formato = tk.StringVar(value="csv")
        tk.Radiobutton(frame_formato, text="CSV  (separado por ;)", variable=self.var_formato, value="csv", command=self.on_formato_change, font=("Arial", 10)).pack(side=tk.LEFT)
        tk.Radiobutton(frame_formato, text="XLSX  (Excel)", variable=self.var_formato, value="xlsx", command=self.on_formato_change, font=("Arial", 10)).pack(side=tk.LEFT, padx=(20, 0))

        tk.Label(root, text="Intervalo de páginas:").grid(row=4, column=0, sticky="w", padx=10, pady=5)
        frame_pag = tk.Frame(root)
        frame_pag.grid(row=4, column=1, sticky="w", padx=5, pady=5)
        self.entry_paginas = tk.Entry(frame_pag, width=20)
        self.entry_paginas.insert(0, "1-end")
        self.entry_paginas.pack(side=tk.LEFT)
        tk.Label(frame_pag, text="  Ex.: 1-end, 5-30 ou 7", fg="#555").pack(side=tk.LEFT)

        self.var_filtro = tk.BooleanVar(value=False)
        self.var_auto_data = tk.BooleanVar(value=False)

        chk_filtro = tk.Checkbutton(root, text="Filtrar linhas por data", variable=self.var_filtro, command=self.toggle_filtro)
        chk_filtro.grid(row=5, column=0, sticky="w", padx=10, pady=5)

        self.chk_auto = tk.Checkbutton(root, text="Detecção automática (01, 01 qua, 01/01/2010…)", variable=self.var_auto_data, command=self.toggle_filtro, state="disabled")
        self.chk_auto.grid(row=5, column=1, columnspan=2, sticky="w", padx=5, pady=5)

        tk.Label(root, text="Formato(s) da data:").grid(row=6, column=0, sticky="w", padx=10, pady=5)
        self.entry_data = tk.Entry(root, width=48, state="disabled")
        self.entry_data.insert(0, "dd/mm/aaaa")
        self.entry_data.grid(row=6, column=1, padx=5, pady=5, sticky="we")

        self.lbl_dica_data = tk.Label(root, text="", fg="#666", font=("Arial", 8), anchor="w")
        self.lbl_dica_data.grid(row=6, column=2, sticky="w", padx=5, pady=5)

        frame_botoes = tk.Frame(root)
        frame_botoes.grid(row=7, column=0, columnspan=3, pady=25)
        tk.Button(frame_botoes, text="Iniciar", command=self.iniciar, bg="#2e7d32", fg="white", width=16, font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=10)
        tk.Button(frame_botoes, text="Cancelar", command=self.cancelar, width=16, font=("Arial", 10)).pack(side=tk.LEFT, padx=10)

    def toggle_filtro(self):
        if self.var_filtro.get():
            self.chk_auto.config(state="normal")
            if self.var_auto_data.get():
                self.entry_data.config(state="disabled")
                self.lbl_dica_data.config(text="automático")
            else:
                self.entry_data.config(state="normal")
                self.lbl_dica_data.config(text="aceita várias máscaras:\ndd/mm/aaaa, dd mmm, dd")
        else:
            self.chk_auto.config(state="disabled")
            self.var_auto_data.set(False)
            self.entry_data.config(state="disabled")
            self.lbl_dica_data.config(text="")

    def on_formato_change(self):
        caminho = self.entry_csv.get().strip()
        if caminho:
            novo = garantir_extensao(caminho, self.var_formato.get())
            if novo != caminho:
                self.entry_csv.delete(0, tk.END)
                self.entry_csv.insert(0, novo)

    def escolher_pdf(self):
        initial = None
        atual = self.entry_pdf.get().strip()
        if atual and os.path.isdir(os.path.dirname(atual)):
            initial = os.path.dirname(atual)
        elif os.path.isdir(os.path.expanduser("~/Downloads")):
            initial = os.path.expanduser("~/Downloads")

        dlg = FilePickerDialog(self.root, title="Selecione o PDF", initialdir=initial, filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")], sort_by="date", descending=True)
        if dlg.selected_path:
            self.entry_pdf.delete(0, tk.END)
            self.entry_pdf.insert(0, dlg.selected_path)
            if not self.entry_csv.get().strip():
                base, _ = os.path.splitext(dlg.selected_path)
                ext = ".xlsx" if self.var_formato.get() == "xlsx" else ".csv"
                self.entry_csv.insert(0, base + ext)

    def escolher_csv(self):
        formato = self.var_formato.get()
        if formato == "xlsx":
            tipos = [("Planilha Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
            ext = ".xlsx"
        else:
            tipos = [("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")]
            ext = ".csv"
        caminho = filedialog.asksaveasfilename(title=f"Salvar {formato.upper()} como...", defaultextension=ext, filetypes=tipos)
        if caminho:
            self.entry_csv.delete(0, tk.END)
            self.entry_csv.insert(0, garantir_extensao(caminho, formato))

    def iniciar(self):
        pdf = self.entry_pdf.get().strip()
        csv_path = self.entry_csv.get().strip()
        pag_str = self.entry_paginas.get().strip()
        formato = self.var_formato.get()

        if not pdf:
            messagebox.showerror("Erro", "Selecione o arquivo PDF.")
            return
        if not os.path.exists(pdf):
            messagebox.showerror("Erro", f"Arquivo PDF não encontrado:\n{pdf}")
            return
        if not csv_path:
            messagebox.showerror("Erro", "Indique o nome do arquivo de saída.")
            return

        csv_path = garantir_extensao(csv_path, formato)
        self.entry_csv.delete(0, tk.END)
        self.entry_csv.insert(0, csv_path)

        if formato == "xlsx" and importlib.util.find_spec("openpyxl") is None:
            messagebox.showerror("openpyxl ausente", "Para exportar em XLSX é necessário instalar o openpyxl:\n\n    pip install openpyxl\n\nOu troque o formato para CSV.")
            return

        try:
            pag_inicio, pag_fim = parse_paginas(pag_str)
            if pag_inicio < 1 or pag_fim < pag_inicio:
                raise ValueError
        except ValueError:
            messagebox.showerror("Erro", "Intervalo de páginas inválido.\nUse o formato: 1-end, 5-30 ou 7.")
            return

        regex_data = None
        if self.var_filtro.get():
            if self.var_auto_data.get():
                regex_data = REGEX_DATA_AUTO
            else:
                formato_data = self.entry_data.get().strip()
                if not formato_data:
                    messagebox.showerror("Erro", "Informe o formato da data (ex.: dd/mm/aaaa).")
                    return
                try:
                    regex_data = converter_mascaras_para_regex(formato_data)
                except Exception as e:
                    messagebox.showerror("Erro", f"Formato de data inválido: {e}")
                    return
                if regex_data is None:
                    messagebox.showerror("Erro", "Nenhuma máscara de data válida informada.")
                    return

        self.resultado = (pdf, csv_path, pag_inicio, pag_fim, regex_data, formato)
        self.root.quit()
        self.root.destroy()

    def cancelar(self):
        self.resultado = None
        self.root.quit()
        self.root.destroy()

# ----------------------------------------------------------------------
# Seletor visual (colunas + limites Y + exclusões + cabeçalho)
# ----------------------------------------------------------------------
class PDFColumnSelector:
    TOLERANCIA_PX = 10

    def __init__(self, root, pdf_path, pagina_visualizacao):
        self.root = root
        self.pdf_path = pdf_path

        self.limites_colunas = []
        self.limites_linhas = []
        self.retangulos_exclusao = []
        self.regiao_cabecalho = None
        self.ordem_colunas = None # Guarda a ordem final definida pelo usuário

        self.linhas_canvas = []
        self.linhas_horizontais = []
        self.retangulos_canvas = []
        self.cabecalho_canvas = None
        self.rotulos_canvas = [] # Textos com os números das colunas

        self.modo = tk.StringVar(value="vertical")
        self.confirmado = False
        self._drag_inicio = None
        self._drag_id = None

        self.root.title(f"Seletor de Colunas/Linhas - Página {pagina_visualizacao}")

        LARGURA_MAXIMA_TELA = 1200
        ALTURA_MAXIMA_TELA = 800

        with pdfplumber.open(self.pdf_path) as pdf:
            idx_pag = max(0, min(pagina_visualizacao - 1, len(pdf.pages) - 1))
            self.pagina_alvo = pdf.pages[idx_pag]
            img_plumber = self.pagina_alvo.to_image(resolution=150)
            self.largura_pdf = self.pagina_alvo.width
            self.altura_pdf = self.pagina_alvo.height

            img_tmp_path = os.path.join(os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd(), "pdf_preview_dinamico.png")
            try:
                img_plumber.save(img_tmp_path, format="PNG")
            except Exception:
                img_tmp_path = os.path.join(os.getcwd(), "pdf_preview_dinamico.png")
                img_plumber.save(img_tmp_path, format="PNG")

            img_original = Image.open(img_tmp_path)
            largura_original, altura_original = img_original.size

            if largura_original > LARGURA_MAXIMA_TELA:
                proporcao = LARGURA_MAXIMA_TELA / float(largura_original)
                nova_largura = LARGURA_MAXIMA_TELA
                nova_altura = int(float(altura_original) * float(proporcao))
                self.img_redimensionada = img_original.resize((nova_largura, nova_altura), Image.Resampling.LANCZOS)
            else:
                self.img_redimensionada = img_original

            self.largura_img, self.altura_img = self.img_redimensionada.size
            self.fator_escala = self.largura_pdf / self.largura_img

        # --- Barra de modos ---
        frame_modos = tk.Frame(root, bg="#d0d0d0", pady=4)
        frame_modos.pack(side=tk.TOP, fill=tk.X)
        tk.Label(frame_modos, text="Modo:", font=("Arial", 10, "bold"), bg="#d0d0d0").pack(side=tk.LEFT, padx=(10, 5))
        
        tk.Radiobutton(frame_modos, text="│ Coluna (divisória vertical)", variable=self.modo, value="vertical", bg="#d0d0d0", font=("Arial", 10)).pack(side=tk.LEFT, padx=3)
        tk.Radiobutton(frame_modos, text="─ Limite Y (linha horizontal)", variable=self.modo, value="horizontal", bg="#d0d0d0", font=("Arial", 10)).pack(side=tk.LEFT, padx=3)
        tk.Radiobutton(frame_modos, text="▭ Exclusão (área ignorada)", variable=self.modo, value="exclusao", bg="#d0d0d0", font=("Arial", 10)).pack(side=tk.LEFT, padx=3)
        tk.Radiobutton(frame_modos, text="◳ Cabeçalho (Dado Fixo)", variable=self.modo, value="cabecalho", bg="#d0d0d0", font=("Arial", 10, "bold"), fg="#004d00").pack(side=tk.LEFT, padx=3)

        # --- Barra de botões ---
        frame_botoes = tk.Frame(root, bg="#e0e0e0", pady=6)
        frame_botoes.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Button(frame_botoes, text="✔ Iniciar Extração", command=self.iniciar_extracao, bg="#2e7d32", fg="white", font=("Arial", 10, "bold"), width=18).pack(side=tk.RIGHT, padx=8)
        tk.Button(frame_botoes, text="✖ Cancelar", command=self.cancelar, font=("Arial", 10), width=12).pack(side=tk.RIGHT, padx=4)
        tk.Button(frame_botoes, text="🗑 Limpar Tudo", command=self.limpar_tudo, font=("Arial", 10), width=14).pack(side=tk.LEFT, padx=8)
        
        self.lbl_status = tk.Label(frame_botoes, text="nenhum elemento marcado", fg="#333", bg="#e0e0e0", font=("Arial", 10))
        self.lbl_status.pack(side=tk.LEFT, padx=8)

        lbl_instrucoes = tk.Label(root, text=("│ Clique = divisória · ─ Clique = limite Y · ▭ Arraste = área ignorada · ◳ Arraste = dado fixo (ex: mês/ano)\n"
                                              "Passe o mouse sobre um elemento e clique para removê-lo"), fg="blue", font=("Arial", 9, "bold"))
        lbl_instrucoes.pack(side=tk.BOTTOM, pady=4)

        # --- Canvas ---
        frame_principal = tk.Frame(root)
        frame_principal.pack(side=tk.TOP, expand=True, fill=tk.BOTH)
        scroll_x = tk.Scrollbar(frame_principal, orient=tk.HORIZONTAL)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        scroll_y = tk.Scrollbar(frame_principal, orient=tk.VERTICAL)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        largura_janela = min(self.largura_img, LARGURA_MAXIMA_TELA)
        altura_janela = min(self.altura_img, ALTURA_MAXIMA_TELA)

        self.canvas = tk.Canvas(frame_principal, width=largura_janela, height=altura_janela, xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)
        self.canvas.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        scroll_x.config(command=self.canvas.xview)
        scroll_y.config(command=self.canvas.yview)
        self.canvas.config(scrollregion=(0, 0, self.largura_img, self.altura_img))

        self.tk_img = ImageTk.PhotoImage(self.img_redimensionada)
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.tk_img)

        self.canvas.bind("<Button-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.canvas.bind("<Motion>", self.ao_mover_mouse)

        self.root.protocol("WM_DELETE_WINDOW", self.fechar_pela_janela)

    def atualizar_rotulos_colunas(self):
        """Desenha os números (1, 2, 3...) no topo de cada coluna delimitada."""
        for rot_id in self.rotulos_canvas:
            self.canvas.delete(rot_id)
        self.rotulos_canvas.clear()

        xs = [0] + [x for _, x in self.linhas_canvas] + [self.largura_img]
        for i in range(len(xs) - 1):
            centro_x = (xs[i] + xs[i+1]) / 2
            # Desenha um pequeno fundo branco para destacar o número sobre a imagem
            bg_id = self.canvas.create_rectangle(centro_x - 12, 5, centro_x + 12, 25, fill="white", outline="black")
            txt_id = self.canvas.create_text(centro_x, 15, text=str(i+1), font=("Arial", 11, "bold"), fill="blue")
            self.rotulos_canvas.extend([bg_id, txt_id])

    def ao_mover_mouse(self, event):
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        modo = self.modo.get()
        sobre = False

        for line_id, _ in self.linhas_canvas: self.canvas.itemconfig(line_id, fill="red", width=2)
        for line_id, _ in self.linhas_horizontais: self.canvas.itemconfig(line_id, fill="blue", width=2)
        for rect_id, _, _, _, _ in self.retangulos_canvas: self.canvas.itemconfig(rect_id, outline="darkorange", width=2)
        if self.cabecalho_canvas: self.canvas.itemconfig(self.cabecalho_canvas, outline="green", width=2)

        if modo == "vertical":
            for line_id, x_canvas in self.linhas_canvas:
                if abs(x_canvas - x) < self.TOLERANCIA_PX:
                    self.canvas.itemconfig(line_id, fill="orange", width=4)
                    sobre = True
        elif modo == "horizontal":
            for line_id, y_canvas in self.linhas_horizontais:
                if abs(y_canvas - y) < self.TOLERANCIA_PX:
                    self.canvas.itemconfig(line_id, fill="cyan", width=4)
                    sobre = True
        elif modo == "exclusao":
            for rect_id, rx0, ry0, rx1, ry1 in self.retangulos_canvas:
                if rx0 <= x <= rx1 and ry0 <= y <= ry1:
                    self.canvas.itemconfig(rect_id, outline="red", width=3)
                    sobre = True
        elif modo == "cabecalho" and self.cabecalho_canvas:
            cx0, cy0, cx1, cy1 = self.canvas.coords(self.cabecalho_canvas)
            if cx0 <= x <= cx1 and cy0 <= y <= cy1:
                self.canvas.itemconfig(self.cabecalho_canvas, outline="red", width=3)
                sobre = True

        if modo in ("vertical", "horizontal"):
            self.canvas.config(cursor="crosshair")
        elif sobre:
            self.canvas.config(cursor="hand2")
        else:
            self.canvas.config(cursor="")

    def on_press(self, event):
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        modo = self.modo.get()

        if modo == "vertical":
            self._toggle_vertical(x)
        elif modo == "horizontal":
            self._toggle_horizontal(y)
        elif modo == "exclusao":
            for i, (rect_id, rx0, ry0, rx1, ry1) in enumerate(self.retangulos_canvas):
                if rx0 <= x <= rx1 and ry0 <= y <= ry1:
                    self.canvas.delete(rect_id)
                    self.retangulos_canvas.pop(i)
                    self.retangulos_exclusao.pop(i)
                    self.atualizar_status()
                    safe_print("Área de exclusão removida.")
                    return
            self._drag_inicio = (x, y)
            self._drag_id = None
        elif modo == "cabecalho":
            if self.cabecalho_canvas:
                cx0, cy0, cx1, cy1 = self.canvas.coords(self.cabecalho_canvas)
                if cx0 <= x <= cx1 and cy0 <= y <= cy1:
                    self.canvas.delete(self.cabecalho_canvas)
                    self.cabecalho_canvas = None
                    self.regiao_cabecalho = None
                    self.atualizar_status()
                    safe_print("Região de cabeçalho removida.")
                    return
            self._drag_inicio = (x, y)
            self._drag_id = None

    def on_drag(self, event):
        if self.modo.get() not in ("exclusao", "cabecalho") or self._drag_inicio is None:
            return
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        x0, y0 = self._drag_inicio
        
        cor = "darkorange" if self.modo.get() == "exclusao" else "green"
        preenchimento = "yellow" if self.modo.get() == "exclusao" else "lightgreen"

        if self._drag_id is not None:
            self.canvas.coords(self._drag_id, x0, y0, x, y)
        else:
            self._drag_id = self.canvas.create_rectangle(x0, y0, x, y, outline=cor, width=2, fill=preenchimento, stipple="gray25")

    def on_release(self, event):
        modo = self.modo.get()
        if modo not in ("exclusao", "cabecalho") or self._drag_inicio is None:
            return
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        x0, y0 = self._drag_inicio
        self._drag_inicio = None

        if self._drag_id is not None:
            self.canvas.delete(self._drag_id)
            self._drag_id = None
        if abs(x - x0) < 5 or abs(y - y0) < 5:
            return

        x_min, x_max = min(x0, x), max(x0, x)
        y_min, y_max = min(y0, y), max(y0, y)

        if modo == "exclusao":
            rect_id = self.canvas.create_rectangle(x_min, y_min, x_max, y_max, outline="darkorange", width=2, fill="yellow", stipple="gray25")
            self.retangulos_canvas.append((rect_id, x_min, y_min, x_max, y_max))
            self.retangulos_exclusao.append((x_min * self.fator_escala, y_min * self.fator_escala, x_max * self.fator_escala, y_max * self.fator_escala))
            safe_print("Área de exclusão adicionada.")
        
        elif modo == "cabecalho":
            if self.cabecalho_canvas:
                self.canvas.delete(self.cabecalho_canvas)
            
            self.cabecalho_canvas = self.canvas.create_rectangle(x_min, y_min, x_max, y_max, outline="green", width=2, fill="lightgreen", stipple="gray25")
            self.regiao_cabecalho = (x_min * self.fator_escala, y_min * self.fator_escala, x_max * self.fator_escala, y_max * self.fator_escala)
            safe_print("Região de cabeçalho definida.")
            
        self.atualizar_status()

    def _toggle_vertical(self, x):
        mais_proxima = None
        for i, (line_id, x_canvas) in enumerate(self.linhas_canvas):
            dist = abs(x_canvas - x)
            if dist < self.TOLERANCIA_PX:
                if mais_proxima is None or dist < mais_proxima[0]:
                    mais_proxima = (dist, i, line_id, x_canvas)
        if mais_proxima is not None:
            _, idx, line_id, x_canvas = mais_proxima
            self.canvas.delete(line_id)
            self.linhas_canvas.pop(idx)
            x_pdf = x_canvas * self.fator_escala
            if self.limites_colunas:
                idx_l = min(range(len(self.limites_colunas)), key=lambda k: abs(self.limites_colunas[k] - x_pdf))
                self.limites_colunas.pop(idx_l)
            self.atualizar_rotulos_colunas() # <--- Atualiza os números
            self.atualizar_status()
            safe_print(f"Divisória vertical removida em X={x_pdf:.2f}")
            return
        x_pdf = x * self.fator_escala
        self.limites_colunas.append(x_pdf)
        self.limites_colunas.sort()
        line_id = self.canvas.create_line(x, 0, x, self.altura_img, fill="red", width=2)
        self.linhas_canvas.append((line_id, x))
        self.atualizar_rotulos_colunas() # <--- Atualiza os números
        self.atualizar_status()
        safe_print(f"Divisória vertical adicionada em X={x_pdf:.2f}")

    def _toggle_horizontal(self, y):
        mais_proxima = None
        for i, (line_id, y_canvas) in enumerate(self.linhas_horizontais):
            dist = abs(y_canvas - y)
            if dist < self.TOLERANCIA_PX:
                if mais_proxima is None or dist < mais_proxima[0]:
                    mais_proxima = (dist, i, line_id, y_canvas)
        if mais_proxima is not None:
            _, idx, line_id, y_canvas = mais_proxima
            self.canvas.delete(line_id)
            self.linhas_horizontais.pop(idx)
            y_pdf = y_canvas * self.fator_escala
            if self.limites_linhas:
                idx_l = min(range(len(self.limites_linhas)), key=lambda k: abs(self.limites_linhas[k] - y_pdf))
                self.limites_linhas.pop(idx_l)
            self.atualizar_status()
            safe_print(f"Limite horizontal removido em Y={y_pdf:.2f}")
            return
        y_pdf = y * self.fator_escala
        self.limites_linhas.append(y_pdf)
        self.limites_linhas.sort()
        line_id = self.canvas.create_line(0, y, self.largura_img, y, fill="blue", width=2)
        self.linhas_horizontais.append((line_id, y))
        self.atualizar_status()
        safe_print(f"Limite horizontal adicionado em Y={y_pdf:.2f}")

    def atualizar_status(self):
        n_col = len(self.limites_colunas)
        n_lin = len(self.limites_linhas)
        n_exc = len(self.retangulos_exclusao)
        tem_cab = self.regiao_cabecalho is not None
        
        partes = []
        if n_col > 0: partes.append(f"│ {n_col + 1} col")
        if n_lin > 0: partes.append(f"─ {n_lin} linha(s)")
        if n_exc > 0: partes.append(f"▭ {n_exc} exclusão(ões)")
        if tem_cab: partes.append("◳ cabeçalho definido")
        if not partes: partes.append("nenhum elemento marcado")
        
        self.lbl_status.config(text="  •  ".join(partes))

    def limpar_tudo(self):
        if not (self.linhas_canvas or self.linhas_horizontais or self.retangulos_canvas or self.cabecalho_canvas):
            return
        if not messagebox.askyesno("Limpar tudo", "Remover TODAS as marcações?"):
            return
        for line_id, _ in self.linhas_canvas: self.canvas.delete(line_id)
        for line_id, _ in self.linhas_horizontais: self.canvas.delete(line_id)
        for rect_id, _, _, _, _ in self.retangulos_canvas: self.canvas.delete(rect_id)
        if self.cabecalho_canvas: self.canvas.delete(self.cabecalho_canvas)
        
        self.linhas_canvas.clear()
        self.linhas_horizontais.clear()
        self.retangulos_canvas.clear()
        self.limites_colunas.clear()
        self.limites_linhas.clear()
        self.retangulos_exclusao.clear()
        self.regiao_cabecalho = None
        self.cabecalho_canvas = None
        self.atualizar_rotulos_colunas()
        self.atualizar_status()
        safe_print("Todas as marcações foram removidas.")

    def iniciar_extracao(self):
        if not (self.limites_colunas or self.limites_linhas or self.retangulos_exclusao or self.regiao_cabecalho):
            messagebox.showwarning("Nada marcado", "Adicione pelo menos uma marcação na imagem.")
            return

        # Prepara os nomes das colunas para a janela de ordenação
        num_colunas = len(self.limites_colunas) + 1
        nomes = []
        if self.regiao_cabecalho is not None:
            nomes.append("◳ Cabeçalho (Dado Fixo)")
        for i in range(num_colunas):
            nomes.append(f"Coluna {i+1}")

        if len(nomes) > 1:
            dlg = ColumnOrderDialog(self.root, nomes)
            if dlg.resultado is None:
                return # Usuário cancelou a reordenação, permanece na tela
            self.ordem_colunas = dlg.resultado
        else:
            self.ordem_colunas = [0]

        self.confirmado = True
        self.root.destroy()

    def cancelar(self):
        self.confirmado = False
        self.root.destroy()

    def fechar_pela_janela(self):
        if not (self.linhas_canvas or self.linhas_horizontais or self.retangulos_canvas or self.cabecalho_canvas):
            self.cancelar()
            return
        resp = messagebox.askyesno("Confirmar", "Existem marcações pendentes.\nDeseja iniciar a extração?")
        if resp: self.iniciar_extracao()
        else: self.cancelar()

# ----------------------------------------------------------------------
# Extração (CSV ou XLSX)
# ----------------------------------------------------------------------
def processar_pdf_avancado(pdf_path, limites_x, caminho_saida, formato, regex_data,
                           pag_inicio, pag_fim, limites_y=None, exclusoes=None, regiao_cabecalho=None, ordem_colunas=None):
    limites_y = limites_y or []
    exclusoes = exclusoes or []

    if formato == "xlsx":
        from openpyxl import Workbook
        wb = Workbook(write_only=True)
        ws = wb.create_sheet(title="Extração")
        f_csv = None
        escritor_csv = None
    else:
        f_csv = open(caminho_saida, 'w', newline='', encoding='utf-8')
        escritor_csv = csv.writer(f_csv, delimiter=';')
        wb = None
        ws = None

    def escrever_linhas(linhas):
        if formato == "xlsx":
            for linha in linhas: ws.append(linha)
        else:
            escritor_csv.writerows(linhas)

    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_paginas = len(pdf.pages)
            inicio = max(1, min(pag_inicio, total_paginas))
            fim = max(inicio, min(pag_fim, total_paginas))

            if len(limites_y) >= 2:
                y_min_ext = min(limites_y)
                y_max_ext = max(limites_y)
                safe_print(f"Limite Y ativo: extração entre Y={y_min_ext:.1f} e Y={y_max_ext:.1f}")
            else:
                y_min_ext, y_max_ext = 0.0, float('inf')

            safe_print(f"\nIniciando extração ({formato.upper()}): Páginas {inicio} até {fim}")
            if ordem_colunas:
                safe_print(f"Ordem customizada das colunas: {ordem_colunas}")

            ultimo_cabecalho = ""

            for num_pag in range(inicio - 1, fim):
                pagina = pdf.pages[num_pag]
                dados_da_pagina = []

                texto_cabecalho = ""
                if regiao_cabecalho:
                    try:
                        recorte = pagina.within_bbox(regiao_cabecalho)
                        bruto = recorte.extract_text()
                        if bruto:
                            texto_cabecalho = " ".join(bruto.split())
                    except Exception as e:
                        safe_print(f"Aviso: não foi possível extrair cabeçalho na pág {num_pag+1}: {e}")

                    if texto_cabecalho:
                        ultimo_cabecalho = texto_cabecalho
                    else:
                        texto_cabecalho = ultimo_cabecalho
                        if texto_cabecalho:
                            safe_print(f"-> Página {num_pag+1} sem cabeçalho. Reutilizando: '{texto_cabecalho}'")

                largura_total = pagina.width
                limites_completos = [0] + limites_x + [largura_total]

                palavras = pagina.extract_words(x_tolerance=3, y_tolerance=3)
                palavras_filtradas = []
                
                for p in palavras:
                    y_center = (p['top'] + p['bottom']) / 2
                    x_center = (p['x0'] + p['x1']) / 2
                    if y_center < y_min_ext or y_center > y_max_ext:
                        continue
                    
                    excluido = False
                    for (ex0, ey0, ex1, ey1) in exclusoes:
                        if ex0 <= x_center <= ex1 and ey0 <= y_center <= ey1:
                            excluido = True
                            break
                    if excluido: continue
                    
                    if regiao_cabecalho:
                        cx0, cy0, cx1, cy1 = regiao_cabecalho
                        if cx0 <= x_center <= cx1 and cy0 <= y_center <= cy1:
                            continue
                            
                    palavras_filtradas.append(p)

                linhas_por_y = {}
                for p in palavras_filtradas:
                    achou = False
                    for y_existente in linhas_por_y.keys():
                        if abs(y_existente - p['top']) < 5:
                            linhas_por_y[y_existente].append(p)
                            achou = True
                            break
                    if not achou:
                        linhas_por_y[p['top']] = [p]

                for y_top in sorted(linhas_por_y.keys()):
                    palavras_da_linha = sorted(linhas_por_y[y_top], key=lambda x: x['x0'])
                    texto_completo = " ".join([p['text'] for p in palavras_da_linha]).strip()

                    if regex_data:
                        if not regex_data.match(texto_completo): continue
                    else:
                        if not texto_completo: continue

                    campos_linha = [""] * (len(limites_completos) - 1)
                    for p in palavras_da_linha:
                        x_centro = (p['x0'] + p['x1']) / 2
                        for i in range(len(limites_completos) - 1):
                            if limites_completos[i] <= x_centro < limites_completos[i + 1]:
                                if campos_linha[i]:
                                    campos_linha[i] += " " + p['text']
                                else:
                                    campos_linha[i] = p['text']
                                break
                    
                    # Monta a linha base (Cabeçalho + Colunas)
                    if regiao_cabecalho is not None:
                        linha_completa = [texto_cabecalho] + campos_linha
                    else:
                        linha_completa = campos_linha

                    # Aplica a reordenação, se definida
                    if ordem_colunas:
                        linha_completa = [linha_completa[i] for i in ordem_colunas]

                    dados_da_pagina.append(linha_completa)

                if dados_da_pagina:
                    escrever_linhas(dados_da_pagina)
                    safe_print(f"-> Página {num_pag + 1} gravada com sucesso.")

        if formato == "xlsx":
            wb.save(caminho_saida)
    finally:
        if f_csv is not None:
            f_csv.close()

    msg = f"Processamento concluído ({formato.upper()})!\nSalvo em: {caminho_saida}"
    safe_print(msg)
    try:
        messagebox.showinfo("Sucesso", msg)
    except Exception:
        pass

def executar_fluxo(pdf, caminho_saida, pag_inicio, pag_fim, regex_data, formato):
    try:
        root = tk.Tk()
        app = PDFColumnSelector(root, pdf, pag_inicio)
        root.mainloop()
    except Exception as e:
        safe_print(f"Erro ao abrir PDF: {e}")
        return

    if not app.confirmado:
        safe_print("Execução cancelada pelo usuário.")
        return

    processar_pdf_avancado(
        pdf, app.limites_colunas, caminho_saida, formato, regex_data,
        pag_inicio, pag_fim,
        limites_y=app.limites_linhas,
        exclusoes=app.retangulos_exclusao,
        regiao_cabecalho=app.regiao_cabecalho,
        ordem_colunas=app.ordem_colunas
    )

# ----------------------------------------------------------------------
# Entrada
# ----------------------------------------------------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extrator Geométrico de PDF para CSV/XLSX.")
    parser.add_argument("pdf", nargs="?", help="Caminho do PDF de entrada.")
    parser.add_argument("csv", nargs="?", help="Caminho do arquivo de saída (.csv ou .xlsx).")
    parser.add_argument("-f", "--formato", choices=["csv", "xlsx"], default=None, help="Formato de saída.")
    parser.add_argument("-d", "--data-formato", default=None, help="Máscara(s) de data. Ex: 'dd/mm/aaaa, dd mmm'")
    parser.add_argument("--auto-data", action="store_true", help="Detecção automática de datas.")
    parser.add_argument("-p", "--paginas", default="1-end", help="Intervalo de páginas (ex: 5-30).")

    args = parser.parse_args()

    if args.pdf and args.csv:
        if not os.path.exists(args.pdf):
            safe_print(f"Erro: arquivo {args.pdf} não encontrado.")
            sys.exit(1)
        formato = args.formato or inferir_formato(args.csv)
        saida = garantir_extensao(args.csv, formato)
        if formato == "xlsx" and not verificar_openpyxl(): sys.exit(1)
        
        regex_data = REGEX_DATA_AUTO if args.auto_data else (converter_mascaras_para_regex(args.data_formato) if args.data_formato else None)
        try:
            pag_inicio, pag_fim = parse_paginas(args.paginas)
        except ValueError:
            safe_print("Erro: intervalo de páginas inválido.")
            sys.exit(1)

        executar_fluxo(args.pdf, saida, pag_inicio, pag_fim, regex_data, formato)
    else:
        root = tk.Tk()
        config = ConfigWindow(root)
        root.mainloop()
        if config.resultado:
            pdf, csv_path, pag_inicio, pag_fim, regex_data, formato = config.resultado
            executar_fluxo(pdf, csv_path, pag_inicio, pag_fim, regex_data, formato)