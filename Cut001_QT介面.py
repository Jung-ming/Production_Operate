import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget, \
    QMessageBox, QProgressDialog
from PyQt5.QtCore import Qt
from Cut03_文件處理 import 文件處理


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # 創建一個 QVBoxLayout 佈局
        layout = QVBoxLayout()

        # 添加一個 QLabel 顯示標題
        title_label = QLabel("請選擇功能", self)
        title_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title_label, alignment=Qt.AlignCenter)
        layout.addSpacing(160)

        # 創建一個 QWidget 作為佈局的容器
        container = QWidget(self)
        container.setLayout(layout)
        # 將容器設置為中央控件
        self.setCentralWidget(container)
        self.setStyleSheet("""
                    QLabel {
                        color: #FF0000;
                        font-size: 12px;
                    }
                    QMainWindow {
                        background-color: #ECECEC;
                    }
                    QMainWindow {
                        border: 2px solid #FF0000;
                        border-radius: 10px;
                     }
                """)

        # 介面標題與大小
        self.setWindowTitle("排程自動化")
        self.setGeometry(300, 300, 300, 200)

        # 按鈕設置與大小
        # 在這個上下文中，self 是一個特殊的參數，它指向正在創建的類的實例。
        # 在 MainWindow 類的方法中，self 用於引用類的實例本身。
        button = QPushButton("選擇檔案", self)
        button.setGeometry(50, 50, 200, 30)
        # 添加一個 QLabel 顯示選擇的文件路徑
        self.file_label = QLabel(self)
        # 自動換行設定
        self.file_label.setWordWrap(True)
        layout.addWidget(self.file_label, alignment=Qt.AlignCenter)
        # 按鈕連接函數
        button.clicked.connect(self.selectFile)

    def selectFile(self):
        # 創建一個 QFileDialog 的實例，用於顯示文件對話框。
        文件選擇視窗 = QFileDialog()
        # 使用變數 file_path 來接收文件路徑，而 _ 變數表示我們不關心文件類型
        文件路徑, _ = 文件選擇視窗.getOpenFileName(self, "選擇檔案")
        if 文件路徑:
            self.file_label.setText(f"選擇的文件：{文件路徑}")
            try:

                文件處理(文件路徑)
                # dialog = QProgressDialog("文件處理中，請稍後...", None, 0, 0, self)
                # dialog.setWindowModality(Qt.WindowModal)
                # dialog.setWindowTitle("Loading")
                # dialog.setCancelButton(None)
                # dialog.show()
                #
                # dialog.close()
                QMessageBox.information(self, '結果', '文件處理完成!')
            except Exception as e:
                # 異常處理
                QMessageBox.critical(self, "錯誤", f"發生錯誤：{str(e)}")
                print(f"錯誤：{str(e)}")
                # 可選擇終止程序
                sys.exit()

# 這是 Python 中的慣用語法，表示如果這個程式碼是直接被執行而不是被當作模組引入，則執行下面的程式碼塊
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
