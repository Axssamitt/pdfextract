#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import os
import time
import argparse
import importlib.util
import re
import csv

# ----------------------------------------------------------------------
# Verificação de dependências
# ----------------------------------------------------------------------
def verificar_dependencias():
    faltando = []
    if importlib.util.find_spec("PyQt6") is None: faltando.append("PyQt6 (Interface gráfica)")
    if importlib.util.find_spec("pdfplumber") is None: faltando.append("pdfplumber (Leitura de PDF)")
    if importlib.util.find_spec("PIL") is None: faltando.append("Pillow (Renderização de imagem)")

    if not faltando:
        return True
    
    print("\n[ERRO] Dependências ausentes:\n")
    for dep in faltando:
        print(f"  • {dep}")
    print("\nComo resolver:\n  pip install PyQt6 pdfplumber pillow openpyxl")
    return False

def verificar_openpyxl():
    if importlib.util.find_spec("openpyxl") is None:
        print("\n[ERRO] openpyxl não encontrado — necessário para exportar XLSX.")
        return False
    return True

if not verificar_dependencias():
    sys.exit(1)

from PyQt6.QtWidgets import (QApplication, QWidget, QDialog, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QRadioButton, QButtonGroup,
                             QCheckBox, QFileDialog, QMessageBox, QScrollArea, QListWidget)
from PyQt6.QtGui import QPixmap, QPainter, QColor, QPen, QBrush, QFont, QImage
from PyQt6.QtCore import Qt, QPoint, QRect, QSize
import pdfplumber
from PIL import Image

# ----------------------------------------------------------------------
# Utilidades de Dados e Regex
# ----------------------------------------------------------------------
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
    if not partes: return None
    if len(partes) == 1: return re.compile(r'^' + _corpo_mascara(partes[0]))
    corpos = [_corpo_mascara(m) for m in partes]
    return re.compile(r'^(?:' + '|'.join(corpos) + ')')

REGEX_DATA_AUTO = re.compile(
    r'^(?:\d{1,2}[/.\-]\d{1,2}(?:[/.\-]\d{2,4})?|\d{1,2}\s+(?:seg|ter|qua|qui|sex|s[áa]b|dom)\b|\d{1,2}\s+de\s+\S+|\d{1,2}\s*[-–—]|\d{1,2}(?=\s|$))',
    re.IGNORECASE | re.UNICODE
)

def parse_paginas(pag_str):
    pag_inicio, pag_fim = 1, 999999
    pag_str = (pag_str or "").strip()
    if pag_str and pag_str.lower() != "1-end":
        if "-" in pag_str:
            partes = pag_str.split("-")
            pag_inicio = int(partes[0])
            if partes[1].lower() != "end" and partes[1] != "": pag_fim = int(partes[1])
        else:
            pag_inicio = pag_fim = int(pag_str)
    return pag_inicio, pag_fim

def inferir_formato(caminho):
    return "xlsx" if caminho.lower().endswith(".xlsx") else "csv"

def garantir_extensao(caminho, formato):
    ext = ".xlsx" if formato == "xlsx" else ".csv"
    if caminho.lower().endswith(ext): return caminho
    for outra in (".csv", ".xlsx"):
        if caminho.lower().endswith(outra): caminho = caminho[:-len(outra)]
    return caminho + ext

# ----------------------------------------------------------------------
# Janela de Configuração Inicial
# ----------------------------------------------------------------------
class ConfigDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Extrator PDF → CSV/XLSX - Configuração")
        self.setFixedSize(500, 400)
        self.resultado = None
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("<b>Configure a extração antes de escolher as colunas:</b>"))

        # PDF
        h_pdf = QHBoxLayout()
        h_pdf.addWidget(QLabel("Arquivo PDF:"))
        self.txt_pdf = QLineEdit()
        h_pdf.addWidget(self.txt_pdf)
        btn_pdf = QPushButton("Procurar...")
        btn_pdf.clicked.connect(self.escolher_pdf)
        h_pdf.addWidget(btn_pdf)
        layout.addLayout(h_pdf)

        # Saída
        h_out = QHBoxLayout()
        h_out.addWidget(QLabel("Arquivo de saída:"))
        self.txt_out = QLineEdit()
        h_out.addWidget(self.txt_out)
        btn_out = QPushButton("Salvar como...")
        btn_out.clicked.connect(self.escolher_saida)
        h_out.addWidget(btn_out)
        layout.addLayout(h_out)

        # Formato
        h_fmt = QHBoxLayout()
        h_fmt.addWidget(QLabel("Formato de saída:"))
        self.rb_csv = QRadioButton("CSV (separado por ;)")
        self.rb_xlsx = QRadioButton("XLSX (Excel)")
        self.rb_csv.setChecked(True)
        self.rb_csv.toggled.connect(self.on_formato_change)
        h_fmt.addWidget(self.rb_csv)
        h_fmt.addWidget(self.rb_xlsx)
        h_fmt.addStretch()
        layout.addLayout(h_fmt)

        # Páginas
        h_pag = QHBoxLayout()
        h_pag.addWidget(QLabel("Intervalo de páginas:"))
        self.txt_pag = QLineEdit("1-end")
        self.txt_pag.setFixedWidth(100)
        h_pag.addWidget(self.txt_pag)
        h_pag.addWidget(QLabel("Ex.: 1-end, 5-30 ou 7"))
        h_pag.addStretch()
        layout.addLayout(h_pag)

        # Filtro Data
        self.chk_filtro = QCheckBox("Filtrar linhas por data")
        self.chk_filtro.toggled.connect(self.toggle_filtro)
        layout.addWidget(self.chk_filtro)

        h_data = QHBoxLayout()
        h_data.addWidget(QLabel("Formato(s) da data:"))
        self.txt_data = QLineEdit("dd/mm/aaaa")
        self.txt_data.setEnabled(False)
        h_data.addWidget(self.txt_data)
        layout.addLayout(h_data)

        self.chk_auto = QCheckBox("Detecção automática (01, 01 qua, 01/01/2010…)")
        self.chk_auto.setEnabled(False)
        self.chk_auto.toggled.connect(self.toggle_filtro)
        layout.addWidget(self.chk_auto)

        layout.addStretch()

        # Botões
        h_btn = QHBoxLayout()
        btn_ok = QPushButton("Iniciar")
        btn_ok.clicked.connect(self.iniciar)
        btn_ok.setStyleSheet("background-color: #2e7d32; color: white; font-weight: bold; padding: 5px 20px;")
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.clicked.connect(self.reject)
        h_btn.addStretch()
        h_btn.addWidget(btn_cancel)
        h_btn.addWidget(btn_ok)
        layout.addLayout(h_btn)

    def toggle_filtro(self):
        ativo = self.chk_filtro.isChecked()
        self.chk_auto.setEnabled(ativo)
        if ativo:
            self.txt_data.setEnabled(not self.chk_auto.isChecked())
        else:
            self.txt_data.setEnabled(False)

    def on_formato_change(self):
        caminho = self.txt_out.text().strip()
        if caminho:
            fmt = "xlsx" if self.rb_xlsx.isChecked() else "csv"
            self.txt_out.setText(garantir_extensao(caminho, fmt))

    def escolher_pdf(self):
        caminho, _ = QFileDialog.getOpenFileName(self, "Selecione o PDF", "", "Arquivos PDF (*.pdf);;Todos os arquivos (*)")
        if caminho:
            self.txt_pdf.setText(caminho)
            if not self.txt_out.text():
                base, _ = os.path.splitext(caminho)
                ext = ".xlsx" if self.rb_xlsx.isChecked() else ".csv"
                self.txt_out.setText(base + ext)

    def escolher_saida(self):
        fmt = "xlsx" if self.rb_xlsx.isChecked() else "csv"
        filtro = "Planilha Excel (*.xlsx)" if fmt == "xlsx" else "Arquivos CSV (*.csv)"
        caminho, _ = QFileDialog.getSaveFileName(self, f"Salvar {fmt.upper()} como...", "", filtro)
        if caminho:
            self.txt_out.setText(garantir_extensao(caminho, fmt))

    def iniciar(self):
        pdf = self.txt_pdf.text().strip()
        out = self.txt_out.text().strip()
        pag = self.txt_pag.text().strip()
        fmt = "xlsx" if self.rb_xlsx.isChecked() else "csv"

        if not pdf or not os.path.exists(pdf):
            QMessageBox.critical(self, "Erro", "Selecione um arquivo PDF válido.")
            return
        if not out:
            QMessageBox.critical(self, "Erro", "Informe o arquivo de saída.")
            return
        if fmt == "xlsx" and not verificar_openpyxl():
            QMessageBox.critical(self, "Erro", "Instale o openpyxl para exportar XLSX:\n\npip install openpyxl")
            return

        try:
            p_ini, p_fim = parse_paginas(pag)
        except ValueError:
            QMessageBox.critical(self, "Erro", "Intervalo de páginas inválido.")
            return

        regex_data = None
        if self.chk_filtro.isChecked():
            if self.chk_auto.isChecked():
                regex_data = REGEX_DATA_AUTO
            else:
                mascaras = self.txt_data.text().strip()
                regex_data = converter_mascaras_para_regex(mascaras)
                if not regex_data:
                    QMessageBox.critical(self, "Erro", "Máscara de data inválida.")
                    return

        self.resultado = (pdf, out, p_ini, p_fim, regex_data, fmt)
        self.accept()

# ----------------------------------------------------------------------
# Janela de Ordenação de Colunas
# ----------------------------------------------------------------------
class ColumnOrderDialog(QDialog):
    def __init__(self, nomes):
        super().__init__()
        self.setWindowTitle("Definir Ordem das Colunas")
        self.setFixedSize(300, 400)
        self.resultado = None
        
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("A primeira da lista será a Coluna 1:"))

        self.lista = QListWidget()
        for nome in nomes:
            self.lista.addItem(nome)
        layout.addWidget(self.lista)

        h_btn = QHBoxLayout()
        btn_up = QPushButton("↑ Subir")
        btn_up.clicked.connect(self.subir)
        btn_down = QPushButton("↓ Descer")
        btn_down.clicked.connect(self.descer)
        h_btn.addWidget(btn_up)
        h_btn.addWidget(btn_down)
        layout.addLayout(h_btn)

        h_ok = QHBoxLayout()
        btn_ok = QPushButton("Confirmar")
        btn_ok.clicked.connect(self.confirmar)
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.clicked.connect(self.reject)
        h_ok.addWidget(btn_cancel)
        h_ok.addWidget(btn_ok)
        layout.addLayout(h_ok)

    def subir(self):
        row = self.lista.currentRow()
        if row > 0:
            item = self.lista.takeItem(row)
            self.lista.insertItem(row - 1, item)
            self.lista.setCurrentRow(row - 1)

    def descer(self):
        row = self.lista.currentRow()
        if row < self.lista.count() - 1:
            item = self.lista.takeItem(row)
            self.lista.insertItem(row + 1, item)
            self.lista.setCurrentRow(row + 1)

    def confirmar(self):
        self.resultado = [self.lista.item(i).text() for i in range(self.lista.count())]
        self.accept()

# ----------------------------------------------------------------------
# Widget do PDF (Canvas do Qt)
# ----------------------------------------------------------------------
class PDFCanvas(QWidget):
    def __init__(self, pixmap, fator_escala, parent=None):
        super().__init__(parent)
        self.pixmap = pixmap
        self.fator_escala = fator_escala
        self.setFixedSize(pixmap.size())
        self.setMouseTracking(True)

        self.linhas_v = [] # [(x_canvas, x_pdf)]
        self.linhas_h = [] # [(y_canvas, y_pdf)]
        self.exclusoes = [] # [(QRect_canvas, QRect_pdf)]
        self.cabecalho = None # QRect_canvas, QRect_pdf

        self.modo = "vertical"
        self.hover_item = None
        
        self._drag_inicio = None
        self._drag_atual = None

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.drawPixmap(0, 0, self.pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Desenha Linhas Verticais
        for i, (x_c, x_p) in enumerate(self.linhas_v):
            if self.modo == "vertical" and self.hover_item == i:
                painter.setPen(QPen(QColor("orange"), 4))
            else:
                painter.setPen(QPen(QColor("red"), 2))
            painter.drawLine(int(x_c), 0, int(x_c), self.height())

        # Desenha Linhas Horizontais
        for i, (y_c, y_p) in enumerate(self.linhas_h):
            if self.modo == "horizontal" and self.hover_item == i:
                painter.setPen(QPen(QColor("cyan"), 4))
            else:
                painter.setPen(QPen(QColor("blue"), 2))
            painter.drawLine(0, int(y_c), self.width(), int(y_c))

        # Desenha Exclusões
        for i, (rect_c, rect_p) in enumerate(self.exclusoes):
            if self.modo == "exclusao" and self.hover_item == i:
                painter.setPen(QPen(QColor("red"), 3))
                painter.setBrush(QBrush(QColor(255, 0, 0, 50)))
            else:
                painter.setPen(QPen(QColor("darkorange"), 2))
                painter.setBrush(QBrush(QColor(255, 200, 0, 50)))
            painter.drawRect(rect_c)

        # Desenha Cabeçalho
        if self.cabecalho:
            rect_c, rect_p = self.cabecalho
            if self.modo == "cabecalho" and self.hover_item == "header":
                painter.setPen(QPen(QColor("red"), 3))
                painter.setBrush(QBrush(QColor(0, 255, 0, 80)))
            else:
                painter.setPen(QPen(QColor("green"), 2))
                painter.setBrush(QBrush(QColor(144, 238, 144, 80)))
            painter.drawRect(rect_c)

        # Desenha a área de drag em andamento
        if self._drag_inicio and self._drag_atual:
            rect = QRect(self._drag_inicio, self._drag_atual).normalized()
            cor = QColor("darkorange") if self.modo == "exclusao" else QColor("green")
            painter.setPen(QPen(cor, 2, Qt.PenStyle.DashLine))
            painter.setBrush(QBrush(QColor(cor.red(), cor.green(), cor.blue(), 50)))
            painter.drawRect(rect)

        # Desenha Rótulos das Colunas
        if self.linhas_v:
            xs = [0] + [x for x, _ in self.linhas_v] + [self.width()]
            painter.setFont(QFont("Arial", 10, QFont.Weight.Bold))
            for i in range(len(xs) - 1):
                centro = (xs[i] + xs[i+1]) // 2
                painter.setPen(QPen(QColor("black"), 1))
                painter.setBrush(QBrush(QColor("white")))
                painter.drawRect(centro - 12, 5, 24, 20)
                painter.setPen(QPen(QColor("blue"), 1))
                painter.drawText(QRect(centro - 12, 5, 24, 20), Qt.AlignmentFlag.AlignCenter, str(i+1))

    def mousePressEvent(self, event):
        pos = event.pos()
        x, y = pos.x(), pos.y()

        if self.modo == "vertical":
            self._toggle_vertical(x)
        elif self.modo == "horizontal":
            self._toggle_horizontal(y)
        elif self.modo == "exclusao":
            for i, (rect_c, rect_p) in enumerate(self.exclusoes):
                if rect_c.contains(pos):
                    self.exclusoes.pop(i)
                    self.update()
                    return
            self._drag_inicio = pos
            self._drag_atual = pos
        elif self.modo == "cabecalho":
            if self.cabecalho and self.cabecalho[0].contains(pos):
                self.cabecalho = None
                self.update()
                return
            self._drag_inicio = pos
            self._drag_atual = pos

    def mouseMoveEvent(self, event):
        pos = event.pos()
        if self._drag_inicio:
            self._drag_atual = pos
            self.update()
            return

        self.hover_item = None
        if self.modo == "vertical":
            for i, (x, _) in enumerate(self.linhas_v):
                if abs(x - pos.x()) < 10: self.hover_item = i; break
        elif self.modo == "horizontal":
            for i, (y, _) in enumerate(self.linhas_h):
                if abs(y - pos.y()) < 10: self.hover_item = i; break
        elif self.modo == "exclusao":
            for i, (rect, _) in enumerate(self.exclusoes):
                if rect.contains(pos): self.hover_item = i; break
        elif self.modo == "cabecalho" and self.cabecalho:
            if self.cabecalho[0].contains(pos): self.hover_item = "header"

        self.update()

    def mouseReleaseEvent(self, event):
        if not self._drag_inicio: return
        rect_c = QRect(self._drag_inicio, event.pos()).normalized()
        self._drag_inicio = None

        if rect_c.width() < 5 or rect_c.height() < 5:
            self.update()
            return

        rect_p = QRect(
            int(rect_c.x() * self.fator_escala), int(rect_c.y() * self.fator_escala),
            int(rect_c.width() * self.fator_escala), int(rect_c.height() * self.fator_escala)
        )

        if self.modo == "exclusao":
            self.exclusoes.append((rect_c, rect_p))
        elif self.modo == "cabecalho":
            self.cabecalho = (rect_c, rect_p)
        
        self._drag_atual = None
        self.update()

    def _toggle_vertical(self, x):
        for i, (x_c, x_p) in enumerate(self.linhas_v):
            if abs(x_c - x) < 10:
                self.linhas_v.pop(i)
                self.update()
                return
        self.linhas_v.append((x, x * self.fator_escala))
        self.linhas_v.sort(key=lambda item: item[0])
        self.update()

    def _toggle_horizontal(self, y):
        for i, (y_c, y_p) in enumerate(self.linhas_h):
            if abs(y_c - y) < 10:
                self.linhas_h.pop(i)
                self.update()
                return
        self.linhas_h.append((y, y * self.fator_escala))
        self.linhas_h.sort(key=lambda item: item[0])
        self.update()

# ----------------------------------------------------------------------
# Janela Principal do Seletor
# ----------------------------------------------------------------------
class PDFSelectorWindow(QWidget):
    def __init__(self, pdf_path, pagina):
        super().__init__()
        self.pdf_path = pdf_path
        self.confirmado = False
        self.ordem_colunas = None
        
        self.setWindowTitle(f"Seletor - Página {pagina}")
        self.resize(1000, 800)
        
        self.initUI()
        self.carregar_pdf(pagina)

    def initUI(self):
        layout = QVBoxLayout(self)

        # Barra de Modos
        h_modos = QHBoxLayout()
        h_modos.addWidget(QLabel("<b>Modo:</b>"))
        self.bg_modos = QButtonGroup(self)
        
        rb_v = QRadioButton("│ Coluna")
        rb_h = QRadioButton("─ Limite Y")
        rb_e = QRadioButton("▭ Exclusão")
        rb_c = QRadioButton("◳ Cabeçalho")
        rb_v.setChecked(True)
        
        for rb, val in [(rb_v, "vertical"), (rb_h, "horizontal"), (rb_e, "exclusao"), (rb_c, "cabecalho")]:
            self.bg_modos.addButton(rb)
            rb.toggled.connect(lambda checked, v=val: self.mudar_modo(v) if checked else None)
            h_modos.addWidget(rb)
        
        h_modos.addStretch()
        layout.addLayout(h_modos)

        # Scroll Area para o PDF
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("background-color: #888;")
        layout.addWidget(self.scroll)

        # Barra de Status e Botões
        h_status = QHBoxLayout()
        self.lbl_status = QLabel("Nenhum elemento marcado")
        h_status.addWidget(self.lbl_status)
        h_status.addStretch()

        btn_clear = QPushButton("🗑 Limpar Tudo")
        btn_clear.clicked.connect(self.limpar_tudo)
        h_status.addWidget(btn_clear)

        btn_cancel = QPushButton("✖ Cancelar")
        btn_cancel.clicked.connect(self.reject)
        h_status.addWidget(btn_cancel)

        btn_ok = QPushButton("✔ Iniciar Extração")
        btn_ok.setStyleSheet("background-color: #2e7d32; color: white; font-weight: bold;")
        btn_ok.clicked.connect(self.iniciar_extracao)
        h_status.addWidget(btn_ok)

        layout.addLayout(h_status)

    def mudar_modo(self, modo):
        if hasattr(self, 'canvas'):
            self.canvas.modo = modo
            self.canvas.hover_item = None
            self.canvas.update()

    def carregar_pdf(self, pagina):
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                idx = max(0, min(pagina - 1, len(pdf.pages) - 1))
                page = pdf.pages[idx]
                
                img_obj = page.to_image(resolution=150)
                img_path = os.path.join(os.getcwd(), "pdf_preview_qt.png")
                img_obj.save(img_path, format="PNG")
                
                pil_img = Image.open(img_path)
                
                # CORREÇÃO AQUI: Converte para RGBA antes de extrair os bytes
                pil_img = pil_img.convert("RGBA")
                
                data = pil_img.tobytes("raw", "RGBA")
                qimg = QImage(data, pil_img.width, pil_img.height, QImage.Format.Format_RGBA8888)
                pixmap = QPixmap.fromImage(qimg)
                
                if pil_img.width > 1200:
                    pixmap = pixmap.scaledToWidth(1200, Qt.TransformationMode.SmoothTransformation)
                
                fator = page.width / pixmap.width()
                
                self.canvas = PDFCanvas(pixmap, fator)
                self.scroll.setWidget(self.canvas)
                
                # Conecta os sinais para atualizar o status automaticamente
                # Precisamos recriar o laço de eventos, ou podemos usar um timer. 
                # Para simplificar, o status será atualizado quando o mouse for solto.
                
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar PDF:\n{e}")
            self.reject()

    def atualizar_status(self):
        n_col = len(self.canvas.linhas_v) + 1
        n_lin = len(self.canvas.linhas_h)
        n_exc = len(self.canvas.exclusoes)
        tem_cab = self.canvas.cabecalho is not None
        
        partes = []
        if len(self.canvas.linhas_v) > 0: partes.append(f"│ {n_col} col")
        if n_lin > 0: partes.append(f"─ {n_lin} linha(s)")
        if n_exc > 0: partes.append(f"▭ {n_exc} exclusão(ões)")
        if tem_cab: partes.append("◳ cabeçalho definido")
        
        self.lbl_status.setText(" • ".join(partes) if partes else "Nenhum elemento marcado")

    def limpar_tudo(self):
        if QMessageBox.question(self, "Limpar", "Remover todas as marcações?") == QMessageBox.StandardButton.Yes:
            self.canvas.linhas_v.clear()
            self.canvas.linhas_h.clear()
            self.canvas.exclusoes.clear()
            self.canvas.cabecalho = None
            self.canvas.update()
            self.atualizar_status()

    def iniciar_extracao(self):
        if not (self.canvas.linhas_v or self.canvas.linhas_h or self.canvas.exclusoes or self.canvas.cabecalho):
            QMessageBox.warning(self, "Aviso", "Adicione pelo menos uma marcação.")
            return

        num_colunas = len(self.canvas.linhas_v) + 1
        nomes = []
        if self.canvas.cabecalho: nomes.append("◳ Cabeçalho (Dado Fixo)")
        for i in range(num_colunas): nomes.append(f"Coluna {i+1}")

        if len(nomes) > 1:
            dlg = ColumnOrderDialog(nomes)
            if dlg.exec() == QDialog.DialogCode.Accepted:
                self.ordem_colunas = dlg.resultado
                self.confirmado = True
                self.close()
        else:
            self.confirmado = True
            self.close()

    def reject(self):
        self.confirmado = False
        self.close()

    def closeEvent(self, event):
        if not self.confirmado and (self.canvas.linhas_v or self.canvas.cabecalho):
            resp = QMessageBox.question(self, "Confirmar", "Existem marcações pendentes. Deseja iniciar a extração?")
            if resp == QMessageBox.StandardButton.Yes:
                self.iniciar_extracao()
                if self.confirmado: event.accept()
                else: event.ignore()
            else:
                event.accept()
        else:
            event.accept()

# ----------------------------------------------------------------------
# Extração Lógica
# ----------------------------------------------------------------------
def processar_pdf_avancado(pdf_path, limites_x, caminho_saida, formato, regex_data,
                           pag_inicio, pag_fim, limites_y, exclusoes, regiao_cabecalho, ordem_colunas):
    if formato == "xlsx":
        from openpyxl import Workbook
        wb = Workbook(write_only=True)
        ws = wb.create_sheet(title="Extração")
        f_csv, escritor_csv = None, None
    else:
        f_csv = open(caminho_saida, 'w', newline='', encoding='utf-8')
        escritor_csv = csv.writer(f_csv, delimiter=';')
        wb, ws = None, None

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

            y_min_ext = min(limites_y) if len(limites_y) >= 2 else 0
            y_max_ext = max(limites_y) if len(limites_y) >= 2 else float('inf')

            ultimo_cabecalho = ""

            for num_pag in range(inicio - 1, fim):
                pagina = pdf.pages[num_pag]
                dados_pag = []

                texto_cab = ""
                if regiao_cabecalho:
                    try:
                        rx, ry, rw, rh = regiao_cabecalho
                        bbox = (rx, ry, rx + rw, ry + rh)
                        bruto = pagina.within_bbox(bbox).extract_text()
                        if bruto: texto_cab = " ".join(bruto.split())
                    except Exception:
                        pass

                    if texto_cab: ultimo_cabecalho = texto_cab
                    else: texto_cab = ultimo_cabecalho

                largura_total = pagina.width
                limites_completos = [0] + limites_x + [largura_total]
                palavras = pagina.extract_words(x_tolerance=3, y_tolerance=3)
                
                palavras_filtradas = []
                for p in palavras:
                    y_c = (p['top'] + p['bottom']) / 2
                    x_c = (p['x0'] + p['x1']) / 2
                    if y_c < y_min_ext or y_c > y_max_ext: continue
                    
                    if any(ex.x() <= x_c <= ex.x()+ex.width() and ex.y() <= y_c <= ex.y()+ex.height() for ex in exclusoes):
                        continue
                    
                    if regiao_cabecalho:
                        rx, ry, rw, rh = regiao_cabecalho
                        if rx <= x_c <= rx + rw and ry <= y_c <= ry + rh:
                            continue
                    palavras_filtradas.append(p)

                linhas_por_y = {}
                for p in palavras_filtradas:
                    achou = False
                    for y_ex in linhas_por_y:
                        if abs(y_ex - p['top']) < 5:
                            linhas_por_y[y_ex].append(p)
                            achou = True
                            break
                    if not achou: linhas_por_y[p['top']] = [p]

                for y_top in sorted(linhas_por_y):
                    p_linha = sorted(linhas_por_y[y_top], key=lambda x: x['x0'])
                    txt_completo = " ".join([p['text'] for p in p_linha]).strip()

                    if regex_data and not regex_data.match(txt_completo): continue
                    if not txt_completo: continue

                    campos = [""] * (len(limites_completos) - 1)
                    for p in p_linha:
                        x_c = (p['x0'] + p['x1']) / 2
                        for i in range(len(limites_completos) - 1):
                            if limites_completos[i] <= x_c < limites_completos[i + 1]:
                                campos[i] += (" " + p['text'] if campos[i] else p['text'])
                                break
                    
                    linha_completa = [texto_cab] + campos if regiao_cabecalho else campos

                    if ordem_colunas:
                        dic_cols = {}
                        if regiao_cabecalho: dic_cols["◳ Cabeçalho (Dado Fixo)"] = texto_cab
                        for i, c in enumerate(campos): dic_cols[f"Coluna {i+1}"] = c
                        linha_completa = [dic_cols[nome] for nome in ordem_colunas]

                    dados_pag.append(linha_completa)

                if dados_pag: escrever_linhas(dados_pag)

        if formato == "xlsx": wb.save(caminho_saida)
    finally:
        if f_csv: f_csv.close()

    QMessageBox.information(None, "Sucesso", f"Extração concluída!\nSalvo em: {caminho_saida}")

# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------
def main():
    app = QApplication(sys.argv)
    
    config = ConfigDialog()
    if config.exec() == QDialog.DialogCode.Accepted:
        pdf, out, p_ini, p_fim, regex, fmt = config.resultado
        
        seletor = PDFSelectorWindow(pdf, p_ini)
        seletor.show()
        app.exec()

        if seletor.confirmado:
            lim_x = [x for x, _ in seletor.canvas.linhas_v]
            lim_y = [y for y, _ in seletor.canvas.linhas_h]
            exc_pdf = [rect_p for _, rect_p in seletor.canvas.exclusoes]
            
            cab = None
            if seletor.canvas.cabecalho:
                r_pdf = seletor.canvas.cabecalho[1]
                cab = (r_pdf.x(), r_pdf.y(), r_pdf.width(), r_pdf.height())
            
            processar_pdf_avancado(pdf, lim_x, out, fmt, regex, p_ini, p_fim, lim_y, exc_pdf, cab, seletor.ordem_colunas)
    else:
        print("Cancelado.")

if __name__ == "__main__":
    main()