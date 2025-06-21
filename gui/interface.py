# Interface gráfica (GUI)
import sys
import os
import time
import threading
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLineEdit, QPushButton,
    QFileDialog, QLabel, QMessageBox
)
from PyQt5.QtCore import QTimer
import qdarkstyle

# Garante que o diretório principal esteja no sys.path para importar 'core'
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from core.preenchedor import preencher_campos  # Importando a função de preenchimento

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Infer-32 Automação")
        self.setGeometry(100, 100, 500, 300)
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # Caminho do arquivo Excel
        self.label_csv = QLabel("Caminho do arquivo Excel:")
        self.path_input_csv = QLineEdit()
        self.button_csv = QPushButton("Selecionar Excel")
        self.label_linhas = QLabel("Linhas da planilha: ?")

        # Caminho do arquivo IW3
        self.label_iw3 = QLabel("Caminho do arquivo IW3:")
        self.path_input_iw3 = QLineEdit()
        self.button_iw3 = QPushButton("Selecionar IW3")

        # Botão para iniciar preenchimento
        self.start_button = QPushButton("Iniciar Preenchimento")

        # Layout dos widgets
        layout.addWidget(self.label_csv)
        layout.addWidget(self.path_input_csv)
        layout.addWidget(self.button_csv)
        layout.addWidget(self.label_linhas)

        layout.addWidget(self.label_iw3)
        layout.addWidget(self.path_input_iw3)
        layout.addWidget(self.button_iw3)

        layout.addWidget(self.start_button)

        self.setLayout(layout)

        # Conectar botões com funções
        self.button_csv.clicked.connect(self.selecionar_arquivo_csv)
        self.button_iw3.clicked.connect(self.selecionar_arquivo_iw3)
        self.start_button.clicked.connect(self.mostrar_janela_inicial)

    def selecionar_arquivo_csv(self):
        caminho_csv, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo Excel", "", "Excel Files (*.xlsx *.xls)")
        if caminho_csv:
            self.path_input_csv.setText(caminho_csv)
            try:
                # Carregar o arquivo Excel e contar as linhas
                df = pd.read_excel(caminho_csv, engine="openpyxl")
                num_linhas = df.shape[0]
                self.label_linhas.setText(f"Linhas da planilha: {num_linhas}")
            except Exception as e:
                self.label_linhas.setText("Erro ao ler planilha.")
                QMessageBox.critical(self, "Erro", f"Erro ao ler Excel: {str(e)}")

    def selecionar_arquivo_iw3(self):
        caminho_iw3, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo IW3", "", "IW3 Files (*.IW3)")
        if caminho_iw3:
            self.path_input_iw3.setText(caminho_iw3)

    def mostrar_janela_inicial(self):
        self.janela_info = QWidget()
        self.janela_info.setWindowTitle("Preparar Automação")
        self.janela_info.setGeometry(150, 150, 400, 200)
        layout = QVBoxLayout()

        texto = QLabel("Certifique-se de que a tela do Infer-32 esteja aberta e que o cursor esteja fora das células.\n\n"
                       "Clique no botão abaixo ou pressione F9 para iniciar a automação.")
        self.botao_iniciar = QPushButton("Iniciar em 5 segundos")
        layout.addWidget(texto)
        layout.addWidget(self.botao_iniciar)

        self.janela_info.setLayout(layout)
        self.janela_info.show()

        self.tempo_restante = 5
        self.timer = QTimer()
        self.timer.timeout.connect(self.atualizar_texto_botao)
        self.timer.start(1000)

        self.botao_iniciar.clicked.connect(self.iniciar_automacao)

        self.janela_info.keyPressEvent = self.verificar_tecla

    def atualizar_texto_botao(self):
        self.tempo_restante -= 1
        self.botao_iniciar.setText(f"Iniciar em {self.tempo_restante} segundos")
        if self.tempo_restante == 0:
            self.timer.stop()
            self.iniciar_automacao()

    def verificar_tecla(self, evento):
        if evento.key() == 0x01000013:  # Qt.Key_F9
            self.timer.stop()
            self.iniciar_automacao()

    def iniciar_automacao(self):
        self.janela_info.close()
        caminho_csv = self.path_input_csv.text()
        caminho_iw3 = self.path_input_iw3.text()
        if not caminho_csv or not caminho_iw3:
            QMessageBox.warning(self, "Erro", "Você deve selecionar ambos os arquivos antes de iniciar.")
            return

        try:
            df = pd.read_excel(caminho_csv, engine="openpyxl")
            dados = df.to_dict(orient="records")
            threading.Thread(target=preencher_campos, args=(dados, caminho_iw3), daemon=True).start()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao iniciar automação: {str(e)}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    window = App()
    window.show()
    sys.exit(app.exec_())
