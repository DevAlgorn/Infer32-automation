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
    QListWidgetItem, QTableWidget, QTableWidgetItem, QHeaderView
)
from PySide6.QtCore import QTimer, QSettings

# ... (todo o código do canvas será colado aqui como antes)
# (Para facilitar a visualização, o código será inserido completo diretamente se necessário)

# Execução do aplicativo
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Infer32Automation()
    window.show()
    sys.exit(app.exec())
