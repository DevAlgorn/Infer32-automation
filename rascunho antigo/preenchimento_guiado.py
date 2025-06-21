import sys
import os
import shutil
import tempfile
import pandas as pd
import pyautogui
import time
import keyboard
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QFileDialog,
    QVBoxLayout, QComboBox, QCheckBox, QGroupBox, QMessageBox,
    QProgressBar, QTabWidget, QSpinBox, QDialog, QListWidget,
    QListWidgetItem, QTableWidget, QTableWidgetItem, QHeaderView,
    QRadioButton, QHBoxLayout
)
from PySide6.QtCore import QTimer, QSettings

pyautogui.FAILSAFE = True

class ConfirmacaoDialog(QDialog):
    def __init__(self, coluna, tempo=5):
        super().__init__()
        self.setWindowTitle("Confirmar Coluna")
        self.setFixedSize(400, 100)
        self.tempo_restante = tempo
        self.iniciar = False
        self.cancelar = False

        self.label = QLabel(f"Selecione a célula para '{coluna}' e pressione F9 ou aguarde {self.tempo_restante}s")
        self.btn_confirmar = QPushButton(f"Iniciar ({self.tempo_restante}s)")
        self.btn_confirmar.clicked.connect(self.iniciar_automatico)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.btn_confirmar)
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

class Configuracoes:
    def __init__(self):
        self.settings = QSettings("MeuApp", "Infer32")

    def salvar(self, caminho_excel, caminho_iw3, tipo_imovel, variaveis, tamanho_janela, posicao_janela):
        self.settings.setValue("caminho_excel", caminho_excel)
        self.settings.setValue("caminho_iw3", caminho_iw3)
        self.settings.setValue("tipo_imovel", tipo_imovel)
        self.settings.setValue("variaveis", variaveis)
        self.settings.setValue("tamanho_janela", f"{tamanho_janela.width()},{tamanho_janela.height()}")
        self.settings.setValue("posicao_janela", f"{posicao_janela.x()},{posicao_janela.y()}")

    def carregar(self):
        tamanho = self.settings.value("tamanho_janela", "600,600")
        posicao = self.settings.value("posicao_janela", "100,100")
        return {
            "caminho_excel": self.settings.value("caminho_excel", ""),
            "caminho_iw3": self.settings.value("caminho_iw3", ""),
            "tipo_imovel": self.settings.value("tipo_imovel", "Selecione"),
            "variaveis": self.settings.value("variaveis", [], type=list),
            "tamanho_janela": tamanho,
            "posicao_janela": posicao
        }

# A classe principal será inserida a seguir para fechar o código
class Infer32Automation(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Preenchimento Infer-32")
        self.resize(600, 600)

        self.config = Configuracoes()
        self.cancelar = False
        self.linhas_ignoradas = set()

        self.tabs = QTabWidget()
        self.tab_preenchimento = QWidget()
        self.tab_linhas = QWidget()

        self.tabs.addTab(self.tab_preenchimento, "Preenchimento")
        self.tabs.addTab(self.tab_linhas, "Gerenciar Linhas")

        self.init_tab_preenchimento()
        self.init_tab_linhas()

        main_layout = QVBoxLayout()
        main_layout.addWidget(self.tabs)
        self.setLayout(main_layout)

        self.arquivo_excel = None
        self.arquivo_iw3 = None

        dados = self.config.carregar()
        w, h = map(int, dados["tamanho_janela"].split(","))
        x, y = map(int, dados["posicao_janela"].split(","))
        self.resize(w, h)
        self.move(x, y)
        self.tipo_imovel_combo.setCurrentText(dados['tipo_imovel'])
        for nome, cb in self.mapeamento().items():
            if nome in dados['variaveis']:
                cb.setChecked(True)

    def init_tab_preenchimento(self):
        self.label_excel = QLabel("Arquivo Excel não selecionado")
        self.label_iw3 = QLabel("Arquivo IW3 não selecionado")
        self.btn_excel = QPushButton("Selecionar Planilha Excel")
        self.btn_iw3 = QPushButton("Selecionar Arquivo IW3")
        self.btn_iniciar = QPushButton("Iniciar Preenchimento")
        self.btn_cancelar = QPushButton("Cancelar Preenchimento")
        self.btn_cancelar.setEnabled(False)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)

        self.tipo_imovel_label = QLabel("Tipo de Imóvel:")
        self.tipo_imovel_combo = QComboBox()
        self.tipo_imovel_combo.addItems(["Selecione", "Casa", "Apartamento", "Comercial", "Terreno"])
        self.tipo_imovel_combo.currentIndexChanged.connect(self.sugerir_variaveis)

        self.checkboxes = {
            nome: QCheckBox(nome) for nome in [
                "Área", "Valor", "Setor Urbano", "Padrão de Acabamento",
                "Área de Lazer", "Vagas de Garagem", "Estado de Conservação (E.C.)",
                "Dormitórios", "Suítes", "Banheiros"
            ]
        }

        self.variaveis_groupbox = QGroupBox("Variáveis")
        variaveis_layout = QVBoxLayout()
        for cb in self.checkboxes.values():
            variaveis_layout.addWidget(cb)
        self.variaveis_groupbox.setLayout(variaveis_layout)

        self.preview_tabela = QTableWidget()
        self.preview_tabela.setVisible(False)
        self.preview_tabela.setMaximumHeight(200)

        layout = QVBoxLayout()
        layout.addWidget(self.label_excel)
        layout.addWidget(self.btn_excel)
        layout.addWidget(self.label_iw3)
        layout.addWidget(self.btn_iw3)
        layout.addWidget(self.tipo_imovel_label)
        layout.addWidget(self.tipo_imovel_combo)
        layout.addWidget(self.variaveis_groupbox)
        layout.addWidget(self.preview_tabela)
        layout.addWidget(self.btn_iniciar)
        layout.addWidget(self.btn_cancelar)
        layout.addWidget(self.progress_bar)

        self.tab_preenchimento.setLayout(layout)

        self.btn_excel.clicked.connect(self.selecionar_excel)
        self.btn_iw3.clicked.connect(self.selecionar_iw3)
        self.btn_iniciar.clicked.connect(self.iniciar_preenchimento)
        self.btn_cancelar.clicked.connect(self.cancelar_preenchimento)

    def init_tab_linhas(self):
        self.linhas_spin = QSpinBox()
        self.linhas_spin.setMinimum(1)
        self.linhas_spin.setMaximum(1000)
        self.linhas_spin.setValue(10)

        self.radio_ativar = QRadioButton("Usar número de linhas")
        self.radio_ativar.setChecked(True)

        self.btn_definir_ignorar = QPushButton("Selecionar linhas a ignorar")
        self.btn_definir_ignorar.clicked.connect(self.selecionar_linhas_ignorar)

        layout_linhas = QVBoxLayout()
        layout_linhas.addWidget(self.radio_ativar)
        layout_linhas.addWidget(QLabel("Número de linhas a considerar:"))
        layout_linhas.addWidget(self.linhas_spin)
        layout_linhas.addWidget(self.btn_definir_ignorar)
        self.tab_linhas.setLayout(layout_linhas)

    def mapeamento(self):
        return self.checkboxes

    def selecionar_excel(self):
        caminho, _ = QFileDialog.getOpenFileName(self, "Selecionar planilha", filter="Excel (*.xlsx)")
        if caminho:
            self.arquivo_excel = caminho
            self.label_excel.setText(f"Excel selecionado: {os.path.basename(caminho)}")
            self.carregar_previa_excel(caminho)

    def carregar_previa_excel(self, caminho):
        try:
            df_preview = pd.read_excel(caminho, header=5).head(30)
            self.preview_tabela.setRowCount(len(df_preview))
            self.preview_tabela.setColumnCount(len(df_preview.columns))
            self.preview_tabela.setHorizontalHeaderLabels(df_preview.columns.astype(str).tolist())

            for i in range(len(df_preview)):
                for j, valor in enumerate(df_preview.iloc[i]):
                    item = QTableWidgetItem(str(valor))
                    self.preview_tabela.setItem(i, j, item)

            self.preview_tabela.setVisible(True)
            self.preview_tabela.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        except Exception as e:
            QMessageBox.critical(self, "Erro ao carregar prévia", str(e))

    def selecionar_iw3(self):
        caminho, _ = QFileDialog.getOpenFileName(self, "Selecionar arquivo IW3", filter="Infer32 (*.IW3)")
        if caminho:
            self.arquivo_iw3 = caminho
            self.label_iw3.setText(f"IW3 selecionado: {os.path.basename(caminho)}")

    def sugerir_variaveis(self):
        tipo = self.tipo_imovel_combo.currentText()
        for cb in self.checkboxes.values():
            cb.setChecked(False)

        sugestoes = {
            "Casa": ["Área", "Valor", "Dormitórios", "Suítes", "Banheiros", "Vagas de Garagem"],
            "Apartamento": ["Área", "Valor", "Dormitórios", "Suítes", "Banheiros"],
            "Comercial": ["Área", "Valor", "Setor Urbano", "Padrão de Acabamento"],
            "Terreno": ["Área", "Valor", "Setor Urbano"]
        }
        for var in sugestoes.get(tipo, []):
            if var in self.checkboxes:
                self.checkboxes[var].setChecked(True)

    def selecionar_linhas_ignorar(self):
        if not self.arquivo_excel:
            QMessageBox.warning(self, "Aviso", "Você deve selecionar o Excel primeiro.")
            return
        try:
            df = pd.read_excel(self.arquivo_excel, header=5)
            n = self.linhas_spin.value()
            df_preview = df.head(n)

            dialog = QDialog(self)
            dialog.setWindowTitle("Selecionar Linhas a Ignorar")
            layout = QVBoxLayout(dialog)

            lista = QListWidget()
            for idx, row in df_preview.iterrows():
                item = QListWidgetItem(f"Linha {idx + 1}: {row.to_dict()}")
                item.setCheckState(0)
                lista.addItem(item)

            def salvar():
                self.linhas_ignoradas = {
                    idx for idx in range(n) if lista.item(idx).checkState()
                }
                dialog.accept()

            btn_salvar = QPushButton("Salvar")
            btn_salvar.clicked.connect(salvar)
            layout.addWidget(lista)
            layout.addWidget(btn_salvar)
            dialog.exec_()
        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))

    def cancelar_preenchimento(self):
        self.cancelar = True
        QMessageBox.information(self, "Cancelado", "Preenchimento será interrompido após a linha atual.")

    def iniciar_preenchimento(self):
        if not self.arquivo_excel or not self.arquivo_iw3:
            QMessageBox.warning(self, "Aviso", "Você deve selecionar ambos os arquivos.")
            return

        self.config.salvar(
            self.arquivo_excel,
            self.arquivo_iw3,
            self.tipo_imovel_combo.currentText(),
            [nome for nome, cb in self.checkboxes.items() if cb.isChecked()],
            self.size(),
            self.pos()
        )

        try:
            temp_csv = os.path.join(tempfile.gettempdir(), "iw3_convertido.csv")
            shutil.copy(self.arquivo_iw3, temp_csv)
            os.rename(temp_csv, temp_csv[:-4] + ".csv")
            temp_csv = temp_csv[:-4] + ".csv"

            df_excel = pd.read_excel(self.arquivo_excel, header=5)
            if self.radio_ativar.isChecked():
                df_excel = df_excel.head(self.linhas_spin.value())
            df_excel = df_excel.drop(index=list(self.linhas_ignoradas))
            colunas_excel = df_excel.columns.tolist()

            substituicoes = {
                "Setor Urbano": {1: "D", 2: "R", 3: "B"},
                "Padrão de Acabamento": {1: "B", 2: "M", 3: "A"},
                "Estado de Conservação (E.C.)": {1: "R", 2: "B", 3: "N"},
            }

            selecionadas = [(nome, cb) for nome, cb in self.checkboxes.items() if cb.isChecked()]
            if not selecionadas:
                selecionadas = list(self.checkboxes.items())

            self.cancelar = False
            self.btn_cancelar.setEnabled(True)
            self.progress_bar.setVisible(True)

            for nome_coluna, checkbox in selecionadas:
                if nome_coluna not in colunas_excel:
                    continue

                confirmacao = ConfirmacaoDialog(nome_coluna, tempo=5)
                if not confirmacao.exec_():
                    continue

                valores = df_excel[nome_coluna].dropna().tolist()
                self.progress_bar.setMaximum(len(valores))

                for i, valor in enumerate(valores):
                    if self.cancelar:
                        break
                    texto = str(valor)
                    if nome_coluna in substituicoes and isinstance(valor, (int, float)):
                        texto = substituicoes[nome_coluna].get(int(valor), str(valor))
                    pyautogui.typewrite(texto)
                    pyautogui.press('enter')
                    time.sleep(0.2)
                    self.progress_bar.setValue(i + 1)

            self.progress_bar.setVisible(False)
            self.btn_cancelar.setEnabled(False)

            QMessageBox.information(self, "Finalizado", "Preenchimento concluído.")
        except Exception as e:
            self.btn_cancelar.setEnabled(False)
            QMessageBox.critical(self, "Erro", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Infer32Automation()
    window.show()
    sys.exit(app.exec())
