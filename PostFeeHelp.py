# PySide6 로드
from PySide6.QtWidgets import *
# ui 로드용
from PySide6 import QtUiTools
from ui_loader import load_ui

class PostFeeHelp(QDialog):
    def __init__(self):
        super().__init__()
        # UI 로드
        load_ui('PostFeeHelp.ui', self)
        # 윈도우 보이기
        self.show()

        # [버튼 시그널 생성]
        # 프로그램 종료
        self.btn_closewindow.clicked.connect(self.closewindow)

    def closewindow(self): #창 닫기
        self.close()