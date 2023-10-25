# PySide6 로드
from PySide6.QtWidgets import *
# ui 로드용
from PySide6 import QtUiTools
from ui_loader import load_ui
# 윈도우 로드
from PostFeeHelp import PostFeeHelp
# 브라우저 기능 로드
import webbrowser
# Pandas 불러오기
import pandas as pd

# SQLite DB 불러오기
import sqlite3
from sqlite3 import Error

# 파일 관리자 로드
import os


class PostFeeMenu(QMainWindow):
    #초기 생성자
    def __init__(self):
        super().__init__()
        #UI 로드
        load_ui('PostFeeMenu.ui', self)
        #윈도우 보이기
        self.show()

        # [버튼 시그널 생성]

        # DB 갱신 버튼
        self.btn_renewdb.clicked.connect(self.renewdb)

        # 우편요금제도 사이트 방문
        self.action_visitpostfeeinfo.triggered.connect(self.visitpostfeepage)

        # 우편물 태그 요령 안내
        self.action_posttaginfo.triggered.connect(self.openposttaghelp)

        # 창 닫기
        self.action_closewindow.triggered.connect(self.closewindow)

        #[창 인스턴스 생성 시 데이터 로드]
        self.loadtabledb1()


# [여기부터는 각 실행 메소드들이 담겨있음]
    def closewindow(self): #창 닫기
        self.close()

    def visitpostfeepage(self): # 우편요금제도 사이트 방문
        visiturl = self.edit_postfeeaddress.toPlainText() #입력부의 URL를 받아
        webbrowser.open_new(visiturl) #webbrowser 에서 해당 변수 url 열기

    def openposttaghelp(self): #우편물 태그 요령 사이트 방문
        visiturl = str(self.edit_postfeeaddress.toPlainText().replace('131.do','213.do?pSiteIdx=125'))
        print(visiturl)
        webbrowser.open_new(visiturl)  # webbrowser 에서 해당 변수 url 열기

    def loadtabledb1(self): #QTableWidget에 DB데이터 넣기

        #DB파일 없으면 StatusBar 알림 추가
        if os.path.isfile("postinformation.db"):
            postdb = sqlite3.connect("postinformation.db")
            cursor = postdb.cursor()
        else:
            self.statusBar.showMessage('DB가 없으므로 DB파일을 생성합니다.')
            postdb = sqlite3.connect("postinformation.db")
            cursor = postdb.cursor()

        # DB 로드전 테이블 초기화--------------------------------------------------------------
        self.table_standardpost.setRowCount(0)  # 갱신 전 테이블 초기화
        self.table_nonstandardpost.setRowCount(0)  # 갱신 전 테이블 초기화
        self.table_registeredpost.setRowCount(0)  # 갱신 전 테이블 초기화
        # ---------------------------------------------------------------------------------

        # 규격테이블 채우기
        tablerow = 0

        try:  # 창 열자마자 DB 불러오기, 실패 시 오류 출력
            # postage 존재하는지 Execute
            cursor.execute('SELECT * FROM postage LIMIT 1')
            # 규격테이블 채우기
            tablerow = 0
            self.table_standardpost.clear()
            for row in cursor.execute('SELECT "중량", "보통우편요금" FROM postage WHERE "내용" = "규격우편물"'):
                self.table_standardpost.setRowCount(tablerow + 1)
                self.table_standardpost.setItem(tablerow, 0, QTableWidgetItem(str(row[0])))
                self.table_standardpost.setItem(tablerow, 1, QTableWidgetItem(str(row[1])))
                tablerow += 1
            names = ('무게', '요금')
            self.table_standardpost.setHorizontalHeaderLabels(names)

            # 규격외테이블 채우기
            tablerow = 0
            self.table_nonstandardpost.clear()
            for row in cursor.execute('SELECT "중량", "보통우편요금" FROM postage WHERE "내용" = "규격외우편물"'):
                self.table_nonstandardpost.setRowCount(tablerow + 1)
                self.table_nonstandardpost.setItem(tablerow, 0, QTableWidgetItem(str(row[0])))
                self.table_nonstandardpost.setItem(tablerow, 1, QTableWidgetItem(str(row[1])))
                tablerow += 1
            names = ('무게', '요금')
            self.table_nonstandardpost.setHorizontalHeaderLabels(names)

            # 등기취급수수료 채우기
            tablerow = 0
            self.table_registeredpost.clear()
            for row in cursor.execute('SELECT "수수료액" FROM postage_special WHERE "종 별" = "등기취급" LIMIT 1'):
                self.table_registeredpost.setRowCount(tablerow + 1)
                self.table_registeredpost.setItem(tablerow, 0, QTableWidgetItem(str(row[0])))
                tablerow += 1
            names = ('수수료액','')
            self.table_registeredpost.setHorizontalHeaderLabels(names)

            # 내용증명 채우기
            tablerow = 0
            self.table_proofofcontent.clear()
            for row in cursor.execute('SELECT "단 위", "수수료액" FROM postage_special WHERE "종 별" = "내용증명"'):
                self.table_proofofcontent.setRowCount(tablerow + 1)
                self.table_proofofcontent.setItem(tablerow, 0, QTableWidgetItem(str(row[0])))
                self.table_proofofcontent.setItem(tablerow, 1, QTableWidgetItem(str(row[1])))
                tablerow += 1
            names = ('단위', '수수료액')
            self.table_proofofcontent.setHorizontalHeaderLabels(names)

            # 익일특급수수료 채우기
            tablerow = 0
            self.table_nextdaypost.clear()
            for row in cursor.execute('SELECT "수수료액" FROM postage_special WHERE "종 별.1" = "익일특급"'):
                self.table_nextdaypost.setRowCount(tablerow + 1)
                self.table_nextdaypost.setItem(tablerow, 0, QTableWidgetItem(str(row[0])))
                tablerow += 1
            names = ('수수료액', '')
            self.table_nextdaypost.setHorizontalHeaderLabels(names)

            # 일반소포 채우기
            tablerow = 0
            self.table_normalpackage.clear()
            for row in cursor.execute('SELECT * FROM postage_package_normal'):
                self.table_normalpackage.setRowCount(len(row) - 4)
                for i in range(2, len(row) - 2):
                    self.table_normalpackage.setItem(tablerow, 0, QTableWidgetItem(str(cursor.description[i][0])))
                    self.table_normalpackage.setItem(tablerow, 1, QTableWidgetItem(str(row[i])))
                    tablerow += 1
            names = ('형태', '요금')
            self.table_normalpackage.setHorizontalHeaderLabels(names)
            self.table_normalpackage.resizeColumnsToContents()

            # 등기소포 채우기
            tablerow = 0
            self.table_registeredpackage.clear()
            for row in cursor.execute('SELECT * FROM postage_package_registered LIMIT 1'):
                self.table_registeredpackage.setRowCount(len(row) - 5)
                for i in range(3, len(row) - 2):
                    self.table_registeredpackage.setItem(tablerow, 0,
                                                            QTableWidgetItem(str(cursor.description[i][0])))
                    self.table_registeredpackage.setItem(tablerow, 1, QTableWidgetItem(str(row[i])))
                    tablerow += 1
            names = ('형태', '요금')
            self.table_registeredpackage.setHorizontalHeaderLabels(names)
            self.table_registeredpackage.resizeColumnsToContents()


            # 소포감액 채우기(기술한계로 요금즉납만 처리)
            tablerow = 0
            self.table_packagediscount.clear()
            for row in cursor.execute('SELECT * FROM postage_package_discount LIMIT 1'):
                self.table_packagediscount.setRowCount(len(row) - 4)
                for i in range(3, 7):
                    self.table_packagediscount.setItem(tablerow, 0, QTableWidgetItem(str(cursor.description[i][0])))
                    self.table_packagediscount.setItem(tablerow, 1, QTableWidgetItem(str(row[i])))
                    tablerow += 1
                break  # 한줄만
            names = ('감액률', '조건')
            self.table_packagediscount.setHorizontalHeaderLabels(names)

        except Error as e:
            self.statusBar.showMessage('테이블 없음, DB를 갱신해주세요')
            print('디버그 : ', str(e))
        # 테이블 로드 종료
        postdb.close()


    def renewdb(self): #DB 갱신
        #progessBar 초기화
        self.progressBar.reset()
        self.statusBar.showMessage('갱신중')
        #사이트 url 얻어오기
        feetable_url = str(self.edit_postfeeaddress.toPlainText())
        #디버스 : url 표시
        print(feetable_url)
        #SQLite DB연결
        postdb = sqlite3.connect("postinformation.db")
        # 10% 진행됨
        self.progressBar.setValue(10)
        # 커서 획득
        cursor = postdb.cursor()
        try:
            # 요금표 테이블 추출 -----------------------------------------------------------------
            try:
                htmldata = pd.read_html(feetable_url)
                # 디버그 : html 읽은 내용 콘솔에 출력
                print('요금표 테이블 추출\n', htmldata)

                # 요금정보 postage 테이블에 저장
                df = htmldata[0]
                df.to_sql('postage', postdb, if_exists='replace')

                # 30% 진행됨
                self.progressBar.setValue(30)
            except:
                self.statusBar.showMessage('에러 : 인터넷 연결 또는 주소를 확인해 주세요')
                #progressBar 초기화
                self.progressBar.reset()



            # 요금특이사항 추출 -----------------------------------------------------------------
            try:
                htmldata = pd.read_html(feetable_url.replace('131', '198'))
            except:
                self.statusBar.showMessage('에러 : 인터넷 연결 또는 주소를 확인해 주세요')
            # 디버그 : html 읽은 내용 콘솔에 출력
            print('요금특이사항 추출\n', htmldata)
            # 요금특이사항 postage_special 테이블에 저장
            df = htmldata[0]
            df.to_sql('postage_special', postdb, if_exists='replace')

            # 30% 진행됨
            self.progressBar.setValue(60)

            # 소포내용 추출 --------------------------------------------------------------------
            try:
                htmldata = pd.read_html(feetable_url.replace('131', '201'))
            except:
                self.statusBar.showMessage('에러 : 인터넷 연결 또는 주소를 확인해 주세요')
            # 디버그 : html 읽은 내용 콘솔에 출력
            print('소포내용 추출\n',htmldata)
            # 소포내용 postage_package 테이블에 저장
            df = htmldata[0]
            df.to_sql('postage_package_registered', postdb, if_exists='replace')
            df = htmldata[1]
            df.to_sql('postage_package_normal', postdb, if_exists='replace')

            df = htmldata[2]
            df.to_sql('postage_package_visit', postdb, if_exists='replace')
            df = htmldata[3]
            df.to_sql('postage_package_add', postdb, if_exists='replace')
            df = htmldata[4]
            df.to_sql('postage_package_discount', postdb, if_exists='replace')

            # 90% 진행됨
            self.progressBar.setValue(90)


            # 변경사항 저장 -----------------------------------------------------------------
            postdb.commit()
            # 연결 종료
            postdb.close()
            # DB 작성 완료했으니 이제 테이블위젯 갱신
            self.loadtabledb1()
            # -----------------------------------------------------------------------------

            # 100% 진행됨
            self.progressBar.setValue(100)

        except TypeError as e:
            print(str(e))
            self.progressBar.reset()
            self.statusBar.showMessage('오류발생, 주소를 확인하세요')

        self.statusBar.showMessage('갱신 완료')