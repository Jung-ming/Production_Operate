import sys
import traceback

from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget, \
    QMessageBox, QProgressDialog, QCheckBox, QHBoxLayout, QSpacerItem, QSizePolicy
from PyQt5.QtCore import Qt, QDate, QSettings
from PyQt5 import QtWidgets, QtCore
from Cut03_文件處理 import *


class 主介面(QMainWindow):
    def __init__(self):
        super().__init__()
        self.文件選擇 = 子介面_文件選擇()
        self.完成工單剪下 = 子介面_完成工單剪下()
        self.開工工單貼上 = 子介面_開工工單貼上()
        self.線別設定 = 線別設定介面()
        # 判斷使用者是否有選擇功能執行，若有，就會在執行功能後輸出文件
        self.執行確認 = False
        self.initUI()

    def initUI(self):
        # 創建一個 QVBoxLayout 佈局
        layout = QtWidgets.QGridLayout()

        self.setStyleSheet("""
                            QLabel {
                                color: #FF0000;
                                font-size: 12px;
                            }
                            QMainWindow {
                                background-color: #FDFEAA;
                            }
                        """)

        # 介面標題與大小
        self.setWindowTitle("排程自動化")
        self.setGeometry(300, 300, 1000, 500)

        # 添加一個 QLabel 顯示標題
        title_label = QLabel("Production 自動化", self)
        title_label.setStyleSheet("font-size: 24px; font-weight: bold;")

        self.執行按鈕 = QPushButton("執行")
        self.執行按鈕.clicked.connect(self.Runcheck)
        self.執行按鈕.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        self.執行按鈕.setFixedSize(150, 30)

        self.線別設定紐 = QPushButton("線別設定")
        self.線別設定紐.clicked.connect(self.showset)
        self.線別設定紐.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        self.線別設定紐.setFixedSize(150, 30)

        layout.addWidget(title_label, 0, 3)
        layout.addWidget(self.完成工單剪下, 1, 1, 1, 5)
        layout.addWidget(self.開工工單貼上, 2, 1, 1, 5)
        layout.addWidget(self.線別設定紐, 3, 2)
        layout.addWidget(self.文件選擇, 4, 1, 1, 5)
        layout.addWidget(self.執行按鈕, 5, 3)

        # 创建一个 QWidget 作为布局的容器
        container = QWidget(self)
        container.setLayout(layout)
        container.setGeometry(0, 0, 1000, 500)

    def Runcheck(self):
        try:
            # 判斷是否選擇文件，未選擇文件就會直接返回，故下面仍然使用if
            if not self.文件選擇.檔案選擇確認:
                QMessageBox.warning(self, "警告", "請先選擇文件！")
                return
            待處理文件 = self.文件選擇.待處理文件
            # 判斷是否勾選功能，若勾選則執行對應功能
            if self.完成工單剪下.勾選框_完成工單剪下.isChecked():
                data_DIP, data_SMT = 文件處理_剪下去(待處理文件, self.完成工單剪下.日期選擇.剪下日期)
                待處理文件 = [data_DIP, data_SMT]
                self.執行確認 = True
            # DIP客戶名單, SMT客戶名單, 抓取四零四, 抓取其他客戶
            if self.開工工單貼上.勾選框_開工工單貼上_四零四.isChecked() \
                    or self.開工工單貼上.勾選框_開工工單貼上_其他客戶.isChecked():
                data_DIP, data_SMT = 文件處理_貼上去(待處理文件=待處理文件,
                                              起始日期=self.開工工單貼上.日期選擇.起始日期,
                                              結束日期=self.開工工單貼上.日期選擇.結束日期,
                                              DIP客戶名單=self.線別設定.子介面_2201.data_list,
                                              SMT客戶名單=[self.線別設定.子介面_2101.data_list,
                                                       self.線別設定.子介面_2102.data_list,
                                                       self.線別設定.子介面_2103.data_list,
                                                       self.線別設定.子介面_2105.data_list],
                                              抓取四零四=self.開工工單貼上.勾選框_開工工單貼上_四零四.isChecked(),
                                              抓取其他客戶=self.開工工單貼上.勾選框_開工工單貼上_其他客戶.isChecked())
                待處理文件 = [data_DIP, data_SMT]
                self.執行確認 = True

            # 判斷是否執行任何功能，若有則輸出文件，無則警告使用者選擇功能
            if self.執行確認:
                文件輸出(待處理文件)
                QMessageBox.information(self, '結果', '文件處理完成!')
                self.執行確認 = False
            else:
                QMessageBox.warning(self, "警告", "未選擇任何功能！")
                return

        except Exception as e:
            error_message = traceback.format_exc()
            QMessageBox.warning(self, "警告", f"錯誤 : {error_message}")
            # QMessageBox.warning(self, "警告", f"錯誤 : {e}")

    def showset(self):
        self.線別設定.show()


class 子介面_文件選擇(QMainWindow):
    def __init__(self):
        super().__init__()
        self.檔案選擇確認 = False
        self.待處理文件 = None
        self.initUI()

    def initUI(self):
        # 創建一個 水平佈局
        layout = QHBoxLayout()

        # 添加一個
        button_檔案選擇 = QPushButton("檔案選擇", self)
        button_檔案選擇.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        button_檔案選擇.setFixedSize(150, 30)
        button_檔案選擇.clicked.connect(self.selectFile)

        self.選擇文件顯示 = QLabel(self)
        self.選擇文件顯示.setStyleSheet(
            "font-size: 15px;border: 3px groove black; background-color: white; padding: 3px;")
        self.選擇文件顯示.setMinimumSize(600, 10)

        layout.addWidget(button_檔案選擇)
        layout.addWidget(self.選擇文件顯示)
        layout.setAlignment(Qt.AlignLeft)

        # 創建一個 QWidget 作為佈局的容器
        container = QWidget(self)
        container.setLayout(layout)
        container.setGeometry(40, 40, 1000, 54)

    def selectFile(self):
        # 創建一個 QFileDialog 的實例，用於顯示文件對話框。
        文件選擇視窗 = QFileDialog()

        # 使用變數 file_path 來接收文件路徑，而 _ 變數表示我們不關心文件類型
        文件路徑, _ = 文件選擇視窗.getOpenFileName(self, "選擇檔案")
        self.檔案選擇確認 = 文件路徑
        self.選擇文件顯示.setText(f"{文件路徑}")
        if self.檔案選擇確認:
            data_DIP, data_SMT = 文件讀取(self.檔案選擇確認)
            self.待處理文件 = [data_DIP, data_SMT]


class 子介面_完成工單剪下(QMainWindow):
    def __init__(self):
        super().__init__()
        self.日期選擇 = DateRangePicker()
        self.initUI()

    def initUI(self):
        # 創建一個 水平佈局
        layout = QHBoxLayout()

        self.文字 = QLabel('完成工單剪下')
        self.文字.setStyleSheet("font-size: 18px;font-family: 新細明體;font-weight: bold")

        # 添加剪下功能勾選
        self.勾選框_完成工單剪下 = QCheckBox("剪")
        self.勾選框_完成工單剪下.setChecked(True)
        self.勾選框_完成工單剪下.setStyleSheet("font-size: 18px;font-family: 新細明體;font-weight: bold")

        layout.addSpacing(100)
        layout.addWidget(self.文字)
        layout.addSpacing(50)
        layout.addWidget(self.勾選框_完成工單剪下)
        layout.addSpacing(50)
        layout.addWidget(self.日期選擇.剪下判斷日期)

        # 置左對齊
        layout.setAlignment(Qt.AlignLeft)

        # 創建一個 QWidget 作為佈局的容器
        container = QWidget(self)
        container.setLayout(layout)
        container.setGeometry(40, 40, 1000, 80)


class 子介面_開工工單貼上(QMainWindow):
    def __init__(self):
        super().__init__()
        self.日期選擇 = DateRangePicker()
        self.initUI()

    def initUI(self):
        # 創建一個 水平佈局
        layout = QHBoxLayout()

        self.文字 = QLabel('開工工單貼上')
        self.文字.setStyleSheet("font-size: 18px;font-family: 新細明體;font-weight: bold")

        self.勾選框_開工工單貼上_四零四 = QCheckBox("四零四")
        self.勾選框_開工工單貼上_四零四.setStyleSheet("font-size: 18px;font-family: 新細明體;font-weight: bold")

        self.勾選框_開工工單貼上_其他客戶 = QCheckBox("其他客戶")
        self.勾選框_開工工單貼上_其他客戶.setChecked(True)
        self.勾選框_開工工單貼上_其他客戶.setStyleSheet("font-size: 18px;font-family: 新細明體;font-weight: bold")

        layout.addSpacing(100)
        layout.addWidget(self.文字)
        layout.addSpacing(50)
        layout.addWidget(self.勾選框_開工工單貼上_四零四)
        layout.addSpacing(30)
        layout.addWidget(self.勾選框_開工工單貼上_其他客戶)
        layout.addSpacing(50)
        layout.addWidget(self.日期選擇.start_date_edit)
        layout.addWidget(self.日期選擇.end_date_edit)

        # 置左對齊
        layout.setAlignment(Qt.AlignLeft)

        # 創建一個 QWidget 作為佈局的容器
        container = QWidget(self)
        container.setLayout(layout)
        container.setGeometry(40, 40, 1000, 80)


class 線別設定子介面_客戶名單(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.data_list = []  # 存储数据的列表
        self.initUI()

    def initUI(self):
        self.setWindowTitle('客戶列表')
        self.setGeometry(450, 400, 350, 400)
        layout = QtWidgets.QVBoxLayout()

        self.提示字串 = QtWidgets.QLabel('請添加或移除客戶')
        self.提示字串.setStyleSheet(
            "font-size: 14px;font-family: 新細明體")
        self.提示字串.setAlignment(QtCore.Qt.AlignCenter)

        self.添加內容 = QtWidgets.QLabel(self)
        self.添加內容.setStyleSheet(
            "font-size: 15px;border: 1px groove black; background-color: white; padding: 1px;")

        # Create ListWidget
        self.list_widget = QtWidgets.QListWidget()

        # 创建一个文本输入框
        self.combo_box = QtWidgets.QComboBox(self)
        self.combo_box.setEditable(True)

        self.客戶列表 = ['01-益網', '02-Syndiant', '03-台灣先能', '04-偉台旺', '05-禾蒼', '06-台灣旭電', '07-易心', '08-新宿', '09-添馨貿易', '0A-禾鈶', '0B-友視達', '0C-日商山下', '0D-協欣', '0E-SCT', '0F-入家生技', '0G-井藤', '0H-泰利', '0I-合豐自動化', '0J-海派世通', '0K-KFA', '0M-鴻懋電子', '0N-鼎群科技', '0O-雲煮易', '0P-Algas', '0Q-嘉倫提', '0R-FIME', '0S-高譽', '0T-富茂電子', '0U-南京資訊', '0V-智慧平台', '0X-捷萌', '0Y-均利', '0Z-光網', '11-發意思', '12-頂福', '13-台優電機', '14-晨宇創新', '15-祥瑞精密', '16-Nata Vision', '17-暐禾實業', '18-識驊科技', '19-瑞仕普', '1A-凱鈺科技', '1B-Jacques', '1C-匯星', '1D-向興資訊', '1E-HP TWN', '1E-1-HP USA', '1F-歐英克', '1G-Cyoptics', '1H-昇楓', '1I-倍微', '1J-捷創精密', '1K-陽光智慧', '1L-捷普', '1M-利達微控', '1N-新綠', '1O-丞集', '1P-榮順軒', '1Q-陸普', '1R-艾米特', '1S-慶鈺實業', '1T-昱科', '1U-廣盛', '1V-佳易', '1W-展連', '1X-其達', '1Y-連易科技', '1Z-勤創電子', '21-群暉', '22-赫揚', '23-其陽科技', '24-登譽', '25-宗懋科技', '26-鴻發國際', '27-佳景', '28-弘道', '29-創億科技', '2A-成欣', '2B-鼎眾', '2B-1-鼎眾投資', '2C-通用拉鏈', '2E-營邦', '2F-中興電工', '2G-鋰想', '2H-家榮', '2I-SLS', '2J-承偉', '2K-連營', '2L-思創', '2M-城堡岩石', '2N-大通', '2O-圜達', '2P-茗莞', '2Q-詠笠科技', '2R-寰永', '2S-奕能科技', '2T-凱恩', '2U-Button', '2V-INTECH', '2W-森富科技', '2X-歐美聖系統', '2Y-H2O', '2Z-元誠', '30-台灣亮明', '31-佳格儀器', '32-ARRIS', '33-美商齊闊', '34-撼訊', '35-華騰', '36-福華電子', '37-SymLink', '38-FORTUNE', '39-鈊象', '3B-亞瑟萊特', '3C-庫康', '3D-威泓京業', '3E-視惟', '3F-衣騰科技', '3G-攸泰', '3H-詰泰科技', '3I-玖鼎電力', '3J-TK USA', '3J-1-TK TWN', '3K-世均', '3L-創新', '3M-東亞通信', '3N-SnapAV', '3O-宇泉能源', '3P-曜隆', '3Q-世仰', '3R-紫微科技', '3S-連易通', '3T-永贏網業', '3V-掌宇', '3W-廣睿', '3Y-楊啟雄', '3Z-東碩', '40-日笠', '41-朋諾電腦', '42-昱源', '44-Asentria', '45-歐霖光通', '45-1-Owlink', '45-2-歐英克', '46-美商定誼(DT)', '47-威歐', '48-慶暐醫療', '49-冠宇', '4A-凱健', '4B-啟翔', '4C-蔚藍視界', '4D-Omnisense', '4E-HDL', '4F-GT JIGS', '4G-YSC', '4H-主向位', '4I-奇邑科技', '4J-捷智科技', '4J-捷智科技', '4K-研能科技', '4L-華羚', '4M-立象科技', '4N-Winstronics', '4O-鴻齡科技', '4Q-佰才邦', '4R-BAICELLS北美子公司', '4S-銓盛', '4U-Mersintel', '4V-ABIT', '4W-立訊', '4X-群光電能', '4Y-Creative', '4Z-鋒霖科技', '50-V-Squared', '51-廣安', '52-日晶', '53-傳承光電', '54-雙鴻科技', '55-鎧力', '56-Sensiway', '57-誠釱科技', '58-奔馬', '59-四零四', '5A-冠大', '5B-HPE', '5C-楓憲', '5D-美商安邁', '5E-興暘', '5F-圓剛', '5G-夏德', '5H-Onlogic', '5I-EERO', '5J-明躍', '5K-紐沃科技', '5L-宜星科技', '5M-思靈客', '5N-優力', '5O-恩星', '5O-恩星', '5P-美商晶典', '5Q-B&B', '5R-帝希', '5S-融程電訊', '5U-瑞旺', '5V-慶旺', '5W-柏德瑞克', '5Z-鉉鴻', '60-迪託設計', '61-Railcorp Corporation', '62-恩德', '63-逢霖工業', '64-大訊科技', '65-貝斯美德', '66-維田科技', '67-奇揚', '68-佳達', '69-鴻海', '69-1-鴻海新竹', '6A-友立森', '6B-恆智', '6C-一元素', '6D-見臻', '6E-光寶', '6F-雙愛電子', '6G-iGrid', '6H-Tiam', '6I-加雲', '6J-研揚', '6K-倚天酷碁', '6L-點子建', '6M-磐儀', '6N-諾內', '6O-矽谷能源', '6Q-璞恆', '6R-得邁斯', '6S-亞美科', '6T-吉禾發', '6U-聰泰', '6V-華謙', '6W-豐和', '6X-波動光', '6Y-艾肯創科', '6Z-高麟國際', '70-艾緯科技', '71-聚拓', '72-亮捷電子', '73-縱橫', '74-科邦', '75-資正電子', '76-高格亞翼', '77-慧榮', '77-1-慧榮新竹', '78-醫揚', '79-CVT', '7A-美銓', '7B-璿智', '7C-詮欣', '7D-朝富科技', '7E-奇磊', '7G-吉鴻', '7H-Diamond Systems', '7I-永辰', '7J-悅明達', '7K-酷碼', '7M-D-Crypt', '7N-佐臻', '7O-兆赫', '7P-PrecisionOT', '7Q-邑昇', '7R-Verdigris', '80-虹堡科技', '81-奈鋒', '82-晉福', '83-DTR-USA', '85-日煬科技', '86-亞元', '87-富鴻網', '88-艾易', '89-世達', '90-立亞特', '91-晶瑞光電', '92-金鵬微控', '93-栢昇興業', '94-奇科數位', '95-nQueue Billback LL', '96-博音科技', '97-Samsara', '98-聯強國際', '99-帝光科技', 'B5-安佳環球', 'B6-唐佑', 'BV-Scollar', 'C1-Ajile', 'C9-沃爾斯', 'CA-HSL', 'CB-Tripplite', 'CU-艾而卡', 'D1-茁邇', 'D6-Arconas', 'E1-寶晟', 'E4-崇林', 'E6-Crea', 'E9-圓展', 'H6-筠鼎', 'I5-宇喬', 'J7-Logic Controls', 'K4-興匯']
        completer = QtWidgets.QCompleter(self.客戶列表)
        completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.combo_box.setCompleter(completer)

        # 将选项添加到下拉式选择菜单中
        self.combo_box.addItems(self.客戶列表)

        # 連結鍵盤事件處理函數
        self.combo_box.installEventFilter(self)

        # 创建一个添加按钮
        self.add_button = QtWidgets.QPushButton("添加", self)

        # Create Remove button
        remove_button = QtWidgets.QPushButton("移除")
        remove_button.clicked.connect(self.removeItem)

        layout.addWidget(self.提示字串)
        layout.addWidget(self.list_widget)
        layout.addWidget(self.添加內容)
        layout.addWidget(self.combo_box)
        layout.addWidget(self.add_button)
        layout.addWidget(remove_button)
        # 连接添加按钮的点击事件到槽函数
        self.add_button.clicked.connect(self.additem)

        # 创建一个 QWidget 作为布局的容器
        container = QtWidgets.QWidget(self)
        container.setLayout(layout)
        container.setGeometry(0, 0, 350, 400)

    def eventFilter(self, obj, event):
        if obj is self.combo_box and event.type() == QtCore.QEvent.KeyPress:
            if event.key() == QtCore.Qt.Key_Enter:
                self.additem()
                return True
        return super().eventFilter(obj, event)

    def additem(self):
        try:
            text = self.combo_box.currentText()
            items = self.list_widget.findItems(text, QtCore.Qt.MatchExactly)

            # 檢查是否出現選單以外的項目
            if text not in self.客戶列表:
                QtWidgets.QMessageBox.warning(self, "警告", "選單無此項目，請確認輸入內容是否正確！")
                return
            # 檢查欄位是否空白
            if not text:
                QtWidgets.QMessageBox.warning(self, "警告", "欄位不可空白！")
                return

            # 檢查項目是否已經存在
            if items:
                QtWidgets.QMessageBox.warning(self, "警告", "項目已存在！")
                return
            self.list_widget.addItem(text)
            self.data_list.append(text)
            self.updateLabel()
        except Exception as e:
            print(e)

    def removeItem(self):
        try:
            # Get selected items
            selected_items = self.list_widget.selectedItems()

            # Remove selected items from ListWidget
            for item in selected_items:
                self.list_widget.takeItem(self.list_widget.row(item))
                self.data_list.remove(item.text())

            self.updateLabel()
        except Exception as e:
            print(e)

    def updateLabel(self):
        名單 = ''
        for item in self.data_list:
            if not 名單:
                名單 += item
            else:
                名單 = 名單 + '、' + item
        self.添加內容.setText(f'{名單}')


class 線別設定介面(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.子介面_2201 = 線別設定子介面_客戶名單()
        self.子介面_2101 = 線別設定子介面_客戶名單()
        self.子介面_2102 = 線別設定子介面_客戶名單()
        self.子介面_2103 = 線別設定子介面_客戶名單()
        self.子介面_2105 = 線別設定子介面_客戶名單()
        self.initUI()

    def initUI(self):
        self.setGeometry(800, 400, 600, 300)
        self.setWindowTitle('線別設定')
        layout = QtWidgets.QGridLayout()

        self.DIP_label = QtWidgets.QLabel('DIP')
        self.DIP_label.setStyleSheet(
            "font-size: 15px;font-family: Times New Roman")
        self.DIP_label.setAlignment(QtCore.Qt.AlignCenter)

        self.SMT_label = QtWidgets.QLabel('SMT')
        self.SMT_label.setStyleSheet(
            "font-size: 15px;font-family: Times New Roman")
        self.SMT_label.setAlignment(QtCore.Qt.AlignCenter)

        button_2201 = QtWidgets.QPushButton("2201")
        button_2201.clicked.connect(self.showsetting2201)
        button_2101 = QtWidgets.QPushButton("2101")
        button_2101.clicked.connect(self.showsetting2101)
        button_2102 = QtWidgets.QPushButton("2102")
        button_2102.clicked.connect(self.showsetting2102)
        button_2103 = QtWidgets.QPushButton("2103")
        button_2103.clicked.connect(self.showsetting2103)
        button_2105 = QtWidgets.QPushButton("2105")
        button_2105.clicked.connect(self.showsetting2105)

        reset_button_2201 = QtWidgets.QPushButton("重設")
        reset_button_2201.clicked.connect(self.Reset2201)
        reset_button_2101 = QtWidgets.QPushButton("重設")
        reset_button_2101.clicked.connect(self.Reset2101)
        reset_button_2102 = QtWidgets.QPushButton("重設")
        reset_button_2102.clicked.connect(self.Reset2102)
        reset_button_2103 = QtWidgets.QPushButton("重設")
        reset_button_2103.clicked.connect(self.Reset2103)
        reset_button_2105 = QtWidgets.QPushButton("重設")
        reset_button_2105.clicked.connect(self.Reset2105)

        # DIP部分
        layout.addWidget(self.DIP_label, 0, 0)
        layout.addWidget(button_2201, 1, 0)
        layout.addWidget(self.子介面_2201.添加內容, 1, 1, 1, 2)
        layout.addWidget(reset_button_2201, 1, 3)

        # SMT部分
        layout.addWidget(self.SMT_label, 3, 0)
        layout.addWidget(button_2101, 4, 0)
        layout.addWidget(self.子介面_2101.添加內容, 4, 1, 1, 2)
        layout.addWidget(reset_button_2101, 4, 3)
        layout.addWidget(button_2102, 5, 0)
        layout.addWidget(self.子介面_2102.添加內容, 5, 1, 1, 2)
        layout.addWidget(reset_button_2102, 5, 3)
        layout.addWidget(button_2103, 6, 0)
        layout.addWidget(self.子介面_2103.添加內容, 6, 1, 1, 2)
        layout.addWidget(reset_button_2103, 6, 3)
        layout.addWidget(button_2105, 7, 0)
        layout.addWidget(self.子介面_2105.添加內容, 7, 1, 1, 2)
        layout.addWidget(reset_button_2105, 7, 3)

        # 创建一个 QWidget 作为布局的容器
        container = QtWidgets.QWidget(self)
        container.setLayout(layout)
        container.setGeometry(0, 0, 600, 300)

        self.loadSettings()

    def showsetting2201(self):
        self.子介面_2201.show()

    def showsetting2101(self):
        self.子介面_2101.show()

    def showsetting2102(self):
        self.子介面_2102.show()

    def showsetting2103(self):
        self.子介面_2103.show()

    def showsetting2105(self):
        self.子介面_2105.show()

    def Reset2201(self):
        self.子介面_2201.data_list.clear()
        self.子介面_2201.updateLabel()
        self.子介面_2201.list_widget.clear()

    def Reset2101(self):
        self.子介面_2101.data_list.clear()
        self.子介面_2101.updateLabel()
        self.子介面_2101.list_widget.clear()

    def Reset2102(self):
        self.子介面_2102.data_list.clear()
        self.子介面_2102.updateLabel()
        self.子介面_2102.list_widget.clear()

    def Reset2103(self):
        self.子介面_2103.data_list.clear()
        self.子介面_2103.updateLabel()
        self.子介面_2103.list_widget.clear()

    def Reset2105(self):
        self.子介面_2105.data_list.clear()
        self.子介面_2105.updateLabel()
        self.子介面_2105.list_widget.clear()

    def closeEvent(self, event):
        self.saveSettings()
        event.accept()

    def saveSettings(self):
        settings = QtCore.QSettings("OrganizationName", "AppName")
        settings.setValue("data_list_2201", self.子介面_2201.data_list)
        settings.setValue("data_list_2101", self.子介面_2101.data_list)
        settings.setValue("data_list_2102", self.子介面_2102.data_list)
        settings.setValue("data_list_2103", self.子介面_2103.data_list)
        settings.setValue("data_list_2105", self.子介面_2105.data_list)
        # settings.setValue("list_widget_2201", self.子介面_2201.list_widget)
        # settings.setValue("list_widget_2101", self.子介面_2101.list_widget)
        # settings.setValue("list_widget_2102", self.子介面_2102.list_widget)
        # settings.setValue("list_widget_2103", self.子介面_2103.list_widget)

    def loadSettings(self):
        settings = QtCore.QSettings("OrganizationName", "AppName")
        self.子介面_2201.data_list = settings.value("data_list_2201", [])
        self.子介面_2101.data_list = settings.value("data_list_2101", [])
        self.子介面_2102.data_list = settings.value("data_list_2102", [])
        self.子介面_2103.data_list = settings.value("data_list_2103", [])
        self.子介面_2105.data_list = settings.value("data_list_2105", [])
        # self.子介面_2201.list_widget = settings.value("list_widget_2201", [])
        # self.子介面_2101.list_widget = settings.value("list_widget_2101", [])
        # self.子介面_2102.list_widget = settings.value("list_widget_2102", [])
        # self.子介面_2103.list_widget = settings.value("list_widget_2103", [])
        self.子介面_2201.updateLabel()
        self.子介面_2101.updateLabel()
        self.子介面_2102.updateLabel()
        self.子介面_2103.updateLabel()
        self.子介面_2105.updateLabel()

        for item in self.子介面_2201.data_list:
            self.子介面_2201.list_widget.addItem(item)
        for item in self.子介面_2101.data_list:
            self.子介面_2101.list_widget.addItem(item)
        for item in self.子介面_2102.data_list:
            self.子介面_2102.list_widget.addItem(item)
        for item in self.子介面_2103.data_list:
            self.子介面_2103.list_widget.addItem(item)
        for item in self.子介面_2105.data_list:
            self.子介面_2105.list_widget.addItem(item)


class DateRangePicker(QtWidgets.QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        初始日期 = QtCore.QDate.currentDate()
        起始日期 = 初始日期.addDays(-6)
        結束日期 = 初始日期.addDays(7)
        if 起始日期.dayOfWeek() == 7:
            起始日期.addDays(-1)
        if 結束日期.dayOfWeek() == 7:
            結束日期.addDays(1)

        self.剪下日期 = 初始日期.toString("yyyy/MM/dd 00:00")
        self.起始日期 = 起始日期.addDays(-6).toString("yyyy/MM/dd 00:00")
        self.結束日期 = 結束日期.addDays(7).toString("yyyy/MM/dd 00:00")

        # 創建起始日期的 QDateEdit
        # QtWidgets.QDateEdit() 函數內可設置預設日期
        self.剪下判斷日期 = self.General_createDateEdit(初始日期)
        self.start_date_edit = self.General_createDateEdit(起始日期)
        self.end_date_edit = self.General_createDateEdit(結束日期)

        # 連接按鈕的點擊事件到槽函數
        self.start_date_edit.dateChanged.connect(self.ChangdateRange)
        self.end_date_edit.dateChanged.connect(self.ChangdateRange)
        self.剪下判斷日期.dateChanged.connect(self.Changdate_cut)

    def General_createDateEdit(self, date):
        # 建立通用的日期選擇框
        date_edit = QtWidgets.QDateEdit(date)
        date_edit.setCalendarPopup(True)
        date_edit.setFixedWidth(110)
        date_edit.setStyleSheet("font-size: 18px;font-Family: Times New Roman")
        return date_edit

    def ChangdateRange(self):
        # 獲取選擇的起始日期和結束日期
        self.起始日期 = self.start_date_edit.date().toString("yyyy/MM/dd 00:00")
        self.結束日期 = self.end_date_edit.date().toString("yyyy/MM/dd 00:00")
        print(self.起始日期, self.結束日期)

    def Changdate_cut(self):
        self.剪下日期 = self.剪下判斷日期.date().toString("yyyy/MM/dd 00:00")
        print(self.剪下日期)


# 這是 Python 中的慣用語法，表示如果這個程式碼是直接被執行而不是被當作模組引入，則執行下面的程式碼塊
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = 主介面()
    window.show()
    sys.exit(app.exec_())
