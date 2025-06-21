
# IMPORTAÇÕES E CONFIGURAÇÕES INICIAIS
import sys
import os
import shutil
import tempfile
import pandas as pd
import pyautogui
import time
import keyboard
import qdarkstyle
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QFileDialog,
    QVBoxLayout, QComboBox, QCheckBox, QGroupBox, QMessageBox,
    QProgressBar, QTabWidget, QSpinBox, QDialog, QListWidget,
    QListWidgetItem, QTableWidget, QTableWidgetItem, QHeaderView,
    QRadioButton, QHBoxLayout
)
from PyQt5.QtCore import QTimer, QSettings, Qt

# CONFIGURAÇÕES DO TEMA
class Tema:
    def __init__(self):
        self.settings = QSettings("MeuApp", "Infer32")
    def salvar_tema(self, tema):
        self.settings.setValue("tema", tema)
    def carregar_tema(self):
        return self.settings.value("tema", "Claro")

# DIÁLOGO DE CONFIRMAÇÃO DE COLUNA
class ConfirmacaoDialog(QDialog):
    def __init__(self, coluna, tempo=10):
        super().__init__()
        self.setWindowTitle("Confirmar Coluna")
        self.setFixedSize(400, 120)
        self.tempo_restante = tempo
        self.iniciar = False
        self.cancelar = False

        self.label = QLabel(f"Selecione a célula para '{coluna}' e pressione F9, aguarde {self.tempo_restante}s ou clique em Pular")
        self.btn_confirmar = QPushButton(f"Iniciar ({self.tempo_restante}s)")
        self.btn_confirmar.clicked.connect(self.iniciar_automatico)
        self.btn_pular = QPushButton("Pular para a próxima coluna")
        self.btn_pular.clicked.connect(self.pular_coluna)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.btn_confirmar)
        layout.addWidget(self.btn_pular)
        self.setLayout(layout)

        self.timer = QTimer()
        self.timer.timeout.connect(self.atualizar_tempo)
        self.timer.start(1000)
        self.checar_teclado()

    def atualizar_tempo(self):
        self.tempo_restante -= 1
        if self.tempo_restante <= 0:
            self.iniciar = True
            self.accept()
        else:
            self.btn_confirmar.setText(f"Iniciar ({self.tempo_restante}s)")

    def iniciar_automatico(self):
        self.iniciar = True
        self.accept()

    def pular_coluna(self):
        self.cancelar = True
        self.reject()

    def checar_teclado(self):
        def escutar():
            while not self.iniciar and not self.cancelar:
                if keyboard.is_pressed('F9'):
                    self.iniciar = True
                    self.accept()
                    break
                if keyboard.is_pressed('F12'):
                    self.cancelar = True
                    self.reject()
                    break
                time.sleep(0.1)
        import threading
        threading.Thread(target=escutar, daemon=True).start()

# CLASSE PRINCIPAL DO SISTEMA
class Infer32Automation(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Preenchimento Infer-32")
        self.resize(800, 600)
        self.cancelar = False
        self.linhas_ignoradas = set()
        self.arquivo_excel = None
        self.arquivo_iw3 = None
        self.settings = QSettings("MeuApp", "Infer32")

        self.tabs = QTabWidget()
        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.tabs.addTab(self.tab1, "Preenchimento")
        self.tabs.addTab(self.tab2, "Gerenciar Linhas")

        self.init_tab1()
        self.init_tab2()

        layout = QVBoxLayout()
        layout.addWidget(self.tabs)
        self.setLayout(layout)

        self.restaurar_configuracoes()

    def init_tab1(self):
        self.label_excel = QLabel("Arquivo Excel não selecionado")
        self.label_iw3 = QLabel("Arquivo IW3 não selecionado")
        self.btn_excel = QPushButton("Selecionar Excel")
        self.btn_iw3 = QPushButton("Selecionar IW3")
        self.btn_iniciar = QPushButton("Iniciar Preenchimento")
        self.btn_cancelar = QPushButton("Cancelar")
        self.btn_cancelar.setEnabled(False)
        self.progress = QProgressBar()

        self.tipo_combo = QComboBox()
        self.tipo_combo.addItems(["Selecione", "Casa", "Apartamento", "Terreno", "Comercial"])
        self.tipo_combo.currentIndexChanged.connect(self.sugerir_variaveis)

        self.checkboxes = {}
        variaveis = [
            "Área", "Valor", "Setor Urbano", "Padrão de Acabamento", "Área de Lazer",
            "Vagas de Garagem", "Estado de Conservação (E.C.)", "Dormitórios", "Suítes", "Banheiros"
        ]
        box_layout = QVBoxLayout()
        for var in variaveis:
            cb = QCheckBox(var)
            self.checkboxes[var] = cb
            box_layout.addWidget(cb)
        self.group_variaveis = QGroupBox("Variáveis")
        self.group_variaveis.setLayout(box_layout)

        layout = QVBoxLayout()
        layout.addWidget(self.label_excel)
        layout.addWidget(self.btn_excel)
        layout.addWidget(self.label_iw3)
        layout.addWidget(self.btn_iw3)
        layout.addWidget(QLabel("Tipo de imóvel:"))
        layout.addWidget(self.tipo_combo)
        layout.addWidget(self.group_variaveis)
        layout.addWidget(self.btn_iniciar)
        layout.addWidget(self.btn_cancelar)
        layout.addWidget(self.progress)
        self.tab1.setLayout(layout)

        self.btn_excel.clicked.connect(self.carregar_excel)
        self.btn_iw3.clicked.connect(self.carregar_iw3)
        self.btn_iniciar.clicked.connect(self.iniciar_preenchimento)
        self.btn_cancelar.clicked.connect(self.cancelar_execucao)

    def init_tab2(self):
        self.radio_todas = QRadioButton("Usar todas as linhas")
        self.radio_definir = QRadioButton("Definir número de linhas")
        self.radio_definir.setChecked(True)
        self.spin = QSpinBox()
        self.spin.setRange(1, 1000)
        self.spin.setValue(10)

        layout = QVBoxLayout()
        layout.addWidget(self.radio_todas)
        layout.addWidget(self.radio_definir)
        layout.addWidget(QLabel("Número de linhas a considerar:"))
        layout.addWidget(self.spin)
        self.tab2.setLayout(layout)

    def carregar_excel(self):
        file, _ = QFileDialog.getOpenFileName(self, "Selecionar Excel", "", "Excel Files (*.xlsx)")
        if file:
            self.arquivo_excel = file
            self.label_excel.setText(os.path.basename(file))

    def carregar_iw3(self):
        file, _ = QFileDialog.getOpenFileName(self, "Selecionar IW3", "", "IW3 Files (*.IW3)")
        if file:
            self.arquivo_iw3 = file
            self.label_iw3.setText(os.path.basename(file))

    def sugerir_variaveis(self):
        tipo = self.tipo_combo.currentText()
        for cb in self.checkboxes.values():
            cb.setChecked(False)
        if tipo == "Casa":
            for var in ["Área", "Setor Urbano", "Padrão de Acabamento", "Dormitórios", "Banheiros"]:
                self.checkboxes[var].setChecked(True)
        elif tipo == "Apartamento":
            for var in ["Área", "Valor", "Dormitórios", "Suítes", "Banheiros"]:
                self.checkboxes[var].setChecked(True)
        elif tipo == "Terreno":
            for var in ["Área", "Setor Urbano", "Valor"]:
                self.checkboxes[var].setChecked(True)
        elif tipo == "Comercial":
            for var in ["Área", "Valor", "Padrão de Acabamento"]:
                self.checkboxes[var].setChecked(True)

    def iniciar_preenchimento(self):
        if not self.arquivo_excel or not self.arquivo_iw3:
            QMessageBox.warning(self, "Atenção", "Selecione ambos os arquivos.")
            return
        try:
            df = pd.read_excel(self.arquivo_excel, header=5)
            if self.radio_definir.isChecked():
                df = df.head(self.spin.value())

            colunas = [col for col, cb in self.checkboxes.items() if cb.isChecked()]
            if not colunas:
                colunas = list(self.checkboxes.keys())

            self.cancelar = False
            self.btn_cancelar.setEnabled(True)
            self.progress.setMaximum(len(df))
            self.progress.setValue(0)

            for coluna in colunas:
                if coluna not in df.columns:
                    continue
                dialog = ConfirmacaoDialog(coluna)
                if dialog.exec_() == QDialog.Rejected:
                    continue
                for i, val in enumerate(df[coluna].dropna()):
                    if self.cancelar:
                        break
                    pyautogui.typewrite(str(val))
                    pyautogui.press("enter")
                    time.sleep(0.3)
                    self.progress.setValue(i + 1)
            self.btn_cancelar.setEnabled(False)
            QMessageBox.information(self, "Finalizado", "Processo concluído.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))

    def cancelar_execucao(self):
        self.cancelar = True

    def restaurar_configuracoes(self):
        tema = self.settings.value("tema", "Claro")
        if tema == "Escuro":
            QApplication.instance().setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())

# EXECUÇÃO
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Infer32Automation()
    window.show()
    sys.exit(app.exec_())
