# gui/teste_interface.py
from PyQt5.QtWidgets import QApplication, QLabel, QWidget
import qdarkstyle
import sys

app = QApplication(sys.argv)
app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())

window = QWidget()
window.setWindowTitle("Teste de Interface")
window.setGeometry(100, 100, 300, 200)

label = QLabel("Interface carregada com sucesso!", parent=window)
label.move(50, 80)

window.show()
sys.exit(app.exec_())