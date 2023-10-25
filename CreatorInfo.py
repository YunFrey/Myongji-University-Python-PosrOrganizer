# Pyside6 로드
from PySide6.QtWidgets import *
# ui 로드용
from ui_loader import load_ui

class CreatorInfo(QDialog):
    def __init__(self):
        super().__init__()
        #UI 로드
        load_ui('CreatorInfo.ui', self)

        # 창 닫기 버튼 시그널
        self.btn_closewindow.clicked.connect(self.closewindow)

        # Modal 루프 윈도우 생성
        self.exec()


    def closewindow(self): #윈도우 닫기
        self.close()