import sys
from PyQt5.QtCore import QCoreApplication
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QWidget, QPushButton, QToolTip
from PyQt5.QtWidgets import QAction, qApp



class MainMenu(QMainWindow):
    #생성자 상속
    def __init__(self):
        super().__init__()
        self.initUI() #UI구성 실행

    #UI 구성 설정
    def initUI(self):
        self.setWindowTitle('[임시]Post Organizer') #윈도우 제목
        self.setWindowIcon(QIcon('web.png')) #아이콘
        self.setGeometry(300, 300, 1600, 900)

        # 상태바 표시
        selfshowMessage('Idle')  # 상태바 텍스트

        #메인화면 종료버튼
        sys_main_quit = QPushButton('종료', self)
        sys_main_quit.clicked.connect(QCoreApplication.instance().quit)
        sys_main_quit.move(50,50)
        #메인화면 종료버튼 사이즈
        sys_main_quit.resize(sys_main_quit.sizeHint())
        #메인화면 종료버튼 툴팁
        sys_main_quit.setToolTip('프로그램을 종료합니다.')

        # 메뉴바-프로그램 종료
        menu_bar_exit = QAction('Exit', self)
        menu_bar_exit.setShortcut('Ctrl+Q')
        menu_bar_exit.setStatusTip('프로그램 종료')
        menu_bar_exit.triggered.connect(qApp.quit)

        #메뉴바 그룹 생성
        menu_bar = self.menuBar()
        menu_bar.setNativeMenuBar(False)
        filemenu = menu_bar.addMenu('&File')
        filemenu.addAction(menu_bar_exit)




        self.show() #창 보이기




#메인
if __name__ == '__main__':
   app = QApplication(sys.argv)
   ex = MainMenu()
   sys.exit(app.exec_())


