from Cut001_QT介面 import *
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog


app = QApplication(sys.argv)
window = MainWindow()
window.show()
app.exec_()

