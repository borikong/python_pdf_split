import sys
from PyQt5.QtWidgets import *
import pdf_cut

class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def initUI(self):
        self.statusBar=QStatusBar()

        self.pdf_root = ""
        self.save_root = ""

        get_pdf_btn = QPushButton('pdf 가져오기(*)') #버튼 텍스트, 부모 클래스(&Button1->단축키 alt+b)
        get_save_btn = QPushButton('저장 위치 선택(*)')
        split_btn=QPushButton('영수증 첨부지 별 PDF 쪼개기')

        get_pdf_btn.clicked.connect(self.request_pdf_root)
        get_save_btn.clicked.connect(self.request_save_root)
        split_btn.clicked.connect(self.split_pdf)

        self.pdf_root_qle = QLineEdit()
        self.save_root_qle = QLineEdit()
        self.get_month_qle=QLineEdit()

        self.pdf_root_qle.setDisabled(True)
        self.save_root_qle.setDisabled(True)

        groupbox1=QGroupBox("PDF 원본")
        groupbox2 = QGroupBox("저장 위치 선택")
        groupbox3=QGroupBox("몇월인가요?(숫자만 입력)")

        hbox1 = QHBoxLayout()
        hbox2 = QHBoxLayout()
        hbox3=QHBoxLayout()

        hbox3.addWidget(self.get_month_qle)
        groupbox3.setLayout(hbox3)

        self.vbox=QVBoxLayout()

        hbox1.addWidget(get_pdf_btn)
        hbox1.addWidget(self.pdf_root_qle)
        groupbox1.setLayout(hbox1)
        self.vbox.addWidget(groupbox1)

        hbox2.addWidget(get_save_btn)
        hbox2.addWidget(self.save_root_qle)
        groupbox2.setLayout(hbox2)
        self.vbox.addWidget(groupbox2)

        self.vbox.addWidget(groupbox3)

        self.vbox.addWidget(split_btn)
        self.vbox.addWidget(self.statusBar)

        self.setLayout(self.vbox)
        self.setWindowTitle('영수증 첨부지 별 PDF 쪼개기 시스템')
        self.resize(500, 200)
        self.center()
        self.show()

    def request_pdf_root(self):
        self.pdf_root=pdf_cut.get_pdf_root()
        print(self.pdf_root[-3:])
        if self.pdf_root[-3:] =="pdf" or self.pdf_root[-3:]=='':
            self.pdf_root_qle.setText(self.pdf_root)
        else:
            buttonReply = QMessageBox.warning(
                self, 'warning', "PDF파일을 업로드 해 주세요.",
                QMessageBox.Yes
            )
            self.statusBar.showMessage('PDF파일을 업로드 해 주세요.')
            self.pdf_root_qle.setText('')

    def request_save_root(self):
        self.save_root=pdf_cut.get_save_root()
        self.save_root_qle.setText(self.save_root)

    def split_pdf(self):
        self.statusBar.showMessage('PDF 쪼개는 중...')
        qApp.processEvents()
        pdf_cut.split_pdf(self.get_month_qle.text(),self.pdf_root,self.save_root)
        self.statusBar.showMessage('PDF 쪼개기 성공!!')

    def warning_prompt(self,text):
        buttonReply = QMessageBox.warning(
            self, 'warning', text,
            QMessageBox.Yes
        )
        self.statusBar.showMessage(text)

if __name__ == '__main__':
   global hwp_root, excel_root, save_root

   app = QApplication(sys.argv)
   ex = MyApp()
   sys.exit(app.exec_())