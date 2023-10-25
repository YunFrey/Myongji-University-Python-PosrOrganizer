#######################################################
# 기본 라이브러리 로드

# 파일관리자 로드
import sys
import time
import os
# SQLite 로드
import sqlite3
print("SQLite 버젼 :",sqlite3.version)
print("SQLite DB 버젼:",sqlite3.sqlite_version)

#######################################################
# 프로그램 실행 전 필수유틸리티 설치여부 확인
try:
    # 테마 로드
    import qdarktheme
    # PySide6 로드
    from PySide6.QtWidgets import *
    from PySide6.QtCore import *
    from PySide6.QtGui import *
    # Designer 로 만든 UI파일 로드용 라이브러리
    from ui_loader import load_ui
    # matplot 로드
    import matplotlib.pyplot as plt

except:
    # 프로그램 실행에 필요한 라이브러리 설치
    import pip
    print("시스템 : PySide6 와 qDarkTheme 필수 라이브러리를 설치합니다.")
    # 설치
    pip.main(['install', '-r', 'requirements/requirements.txt'])
    # 설치완료
    print("시스템 : 설치가 완료되었습니다. 프로그램을 다시 시작해주세요.")
    print("시스템 : 추가로 메인화면의 '필수 라이브러리 설치' 에서 추가 라이브러리를 설치해주세요.")

########################################################
# 메인메뉴 클래스 정의
class MainMenu(QMainWindow):
    def __init__(self):
        super().__init__()
        # UI 테마 로드
        qdarktheme.setup_theme('auto')

        # UI 디자인 불러오기
        load_ui('MainMenu.ui', self)

        # 윈도우 프레임 숨기기
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setWindowFlags(Qt.FramelessWindowHint)

        # 앱 위젯 로드 시험
        try:
            self.loadwidget()
        except sqlite3.OperationalError as e:
            print('시스템 : 위젯이 불러와지지 않음(DB가 없습니다.)')
        except ValueError as e:
            print('시스템 : DB에 레코드가 없습니다.')
        except Exception as e:
            print('시스템 : 위젯 로드 실패, 이유 :', e)



        # 윈도우 보이기
        self.show()

        ###########################################################################################
        #[버튼 시그널 연결]
        # 프로그램 종료
        self.btn_quitwindow.clicked.connect(self.quitwindow)
        # 우편물 요금제 확인
        self.btn_openpostfee.clicked.connect(self.openpostfee)
        # 우편물 접수
        self.btn_submitpost.clicked.connect(self.submitpost)
        # 우편물 조회 및 결재
        self.btn_checkpostlist.clicked.connect(self.sortpostlist)
        # 우편물 할인대상 조회
        self.btn_open_checkgrouppost.clicked.connect(self.checkgrouppost)
        # 우편물 사전접수 자동화 시작
        self.btn_execute_web_posting.clicked.connect(self.writepostdata)
        # 위젯 새로고침/불러오기
        self.btn_refreshwidget.clicked.connect(self.loadwidget)

        #[메뉴]
        # 제작자 정보
        self.action_showcreator.triggered.connect(self.opencreatorinfo)
        # 프로그램 종료
        self.action_programexit.triggered.connect(self.quitwindow)
        # 국내우편요금제도 도움말
        self.action_openpostfeehelp.triggered.connect(self.openpostfeehelp)
        # 필요 라이브러리 설치
        self.action_installrequirements.triggered.connect(self.installrequirements)

        #[상태바]
        self.mainmenu_statusBar.showMessage('Idle')

        # 반려우편물 확인
        self.btn_checkreturn.clicked.connect(self.opencheckreturn)

        #[Debug : 우편물요금리셋]
        self.action_checkfeeavailable.triggered.connect(self.resetpostdata)
        #[Debug : 우편접수리스트리셋]
        self.action_reset_postlist.triggered.connect(self.resetpostlist)
        #[Debug : 묶음우편물리스트리셋]
        self.action_reset_grouplist.triggered.connect(self.resetgrouplist)
        #[Debug : 반려우편물리스트리셋]
        self.action_reset_returnpostlist.triggered.connect(self.resetreturnlist)

    ###########################################################################################
    # 창 이동을 위한 마우스이벤트 핸들링 (코드 출처 : https://blog.naver.com/varofla_blog/222344023916)
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.offset = event.pos()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        try:
            #창 밖이 아니고 버튼 왼쪽을 눌렀을 때
            if self.offset is not None and event.buttons() == Qt.LeftButton:
                #현재 위치에서 이벤트 위치 - 오프셋만큼만 움직이기
                self.move(self.pos() + event.pos() - self.offset)
            else:
                super().mouseMoveEvent(event)
        except:
            pass

    def mouseReleaseEvent(self, event):
        self.offset = None
        super().mouseReleaseEvent(event)

    ###########################################################################################
    # 버튼 클릭시 각 기능별 창을 띄우는 함수
    # 제작자 정보 창 생성
    def opencreatorinfo(self):
        from CreatorInfo import CreatorInfo
        self.creatorinfo_win = CreatorInfo()

    # 우편물 요금창 생성
    def openpostfee(self):
        from PostFeeMenu import PostFeeMenu
        self.postfeemenu_win = PostFeeMenu()
        #modal 이면 exec(), nonmodal 이면 show()

    # 우편물 접수창 생성
    def submitpost(self):
        from SubmitPostMenu import SubmitPostMenu
        self.submitpostmenu_win = SubmitPostMenu()

    # 우편물 조회 및 결재창 생성
    def sortpostlist(self):
        from SortPostMenu import SortPostMenu
        self.SortPostMenu = SortPostMenu()

    # 우편물 묶음그룹 조회창 생성
    def checkgrouppost(self):
        from CheckGroupPost import CheckGroupPost
        self.CheckGroupPost = CheckGroupPost()

    # 우편물 자동접수 창 생성
    def writepostdata(self):
        from WritePostData import WritePostData
        self.WritePostData = WritePostData()
        print('자동접수 창 종료')

    # 우편물 반려 조회창 생성
    def opencheckreturn(self):
        from ReturnPostMenu import ReturnPostMenu

        #창 클래스 구성
        class msg_entryid(QDialog):
            def __init__(self):
                super().__init__()

                #윈도우 설정
                self.setWindowIcon(QIcon('title_big.png'))
                self.setWindowTitle('사원번호 입력')
                self.resize(300,100)
                # Label 추가
                self.desc_label = QLabel('사원번호를 입력하세요')
                # 중앙정렬
                self.desc_label.setAlignment(Qt.AlignCenter)
                # LineEdit 추가
                self.entrybox = QLineEdit()
                self.entrybox.move(115, 30)
                # 버튼 추가
                self.btn_closewin = QPushButton("OK")
                self.btn_closewin.resize(150, 50)
                # 레이아웃 생성 및 연결
                self.layout = QVBoxLayout()
                self.layout.addWidget(self.desc_label)
                self.layout.addWidget(self.entrybox)
                self.layout.addWidget(self.btn_closewin)
                self.setLayout(self.layout)

                #버튼 시그널 연결
                self.btn_closewin.clicked.connect(self.close_rtn_window)

                #창 생성 및 루프
                self.exec()




            #반려우편창을 닫는 기능
            def close_rtn_window(self):
                #entrybox 내용 불러오기
                input = str(self.entrybox.text())
                print('input : ', input)
                # DB 불러오기
                empdb = sqlite3.connect("returnpostlist.db")
                empcursor = empdb.cursor()
                query = 'SELECT id FROM postlist WHERE 사원번호 = "%s"' % input
                try:
                    empcursor.execute(query)
                except sqlite3.OperationalError as e:
                    #id가 공란이거나 다른 DB오류 발생 시
                    print('debug : ', e)
                    # 오류 메세지박스 생성(단순객체) #
                    msg_noid = QMessageBox()  # 메세지박스 생성
                    msg_noid.warning(self, "오류", '사원번호를 입력해 주세요.')  # 메세지박스 경고옵션 설정
                    output_noid = msg_noid.show()  # 메세지박스 실행
                    if output_noid == QMessageBox.Ok:  # OK 누르면 닫기
                        msg_noid.close()
                finally:
                    #문제없을 시
                    res = empcursor.fetchall()
                    # DB 닫기
                    empdb.close()
                    # ID가 존재할 경우
                    if len(res) >= 1 :
                        print('창 생성')
                        # 창 닫기
                        self.close()
                        # 창 생성
                        self.returnpostmenu = ReturnPostMenu(input)

                    else :
                        #ID가 없을 경우
                        # 오류 메세지박스 생성(단순객체) #
                        msg_noid = QMessageBox()  # 메세지박스 생성
                        msg_noid.warning(self, "오류", '반려된 우편물이 없습니다.')  # 메세지박스 경고옵션 설정
                        output_noid = msg_noid.show()  # 메세지박스 실행
                        if output_noid == QMessageBox.Ok:  # OK 누르면 닫기
                            msg_noid.close()
                #DB 정리
                empdb.close()
                #창 닫기
                self.close()
        #객체 생성
        self.window = msg_entryid()

    # 국내우편요금제도 도움말 생성
    def openpostfeehelp(self):
        from PostFeeHelp import PostFeeHelp
        self.postfeehelp_win = PostFeeHelp()

    ###########################################################################################
    # 기능함수
    ###########################################################################################
    # 위젯 로드
    def loadwidget(self):
        if os.path.isfile("postlist.db"):  # DB파일이 있을 경우
            print('DB 파일 찾음')
            listdb = sqlite3.connect("postlist.db")
            cursor = listdb.cursor()
        else:  # DB파일이 없을 경우
            print('DB 파일 찾지 못함')
            self.statusBar().showMessage('DB가 없으므로 DB파일을 생성합니다.')
            listdb = sqlite3.connect("postlist.db")
            cursor = listdb.cursor()
            cursor.execute(
                "CREATE TABLE postlist(id integer PRIMARY KEY AUTOINCREMENT, 접수날짜 text, 사원번호 text, 접수자명 text, 부서명 text, 보내는사람 text, 받는사람 text, 주소 text, 우편번호 text, 제목 text, 수량 integer, 종류 text, 긴급여부 text, 가로길이 real, 세로길이 real, 우편중량 real, 전화번호 text, 높이 real, 메모 text, 할인타입 text, 할인타입그룹 integer, 결재여부 , 반려여부 boolean, 비용 text)")
            listdb.commit()
        # 위젯 로드
        try:
            # 1번 위젯
            cursor.execute("SELECT COUNT(*) FROM postlist")
            countall = cursor.fetchone()[0]
            self.widget_countall.setText(str(countall) + '건')

            # 2번 위젯
            cursor.execute("SELECT COUNT(*) FROM postlist WHERE 결재여부 IS NULL")
            countsubmit = cursor.fetchone()[0]
            self.widget_countsubmit.setText(str(countsubmit) + '건')

            # 3번 위젯
            cursor.execute("SELECT COUNT(*) FROM postlist WHERE 결재여부 = '1'")
            countdone = cursor.fetchone()[0]
            self.widget_countdone.setText(str(countdone) + '건')

            # 4번 위젯(그림)
            # 피규어 생성
            ratio = [countsubmit, countdone]
            print(ratio, type(ratio))
            # plt 생성
            plt.figure(figsize=(1.5, 1.5), facecolor='none')
            plt.pie(ratio, autopct='%.1f%%')
            # 이미지 저장
            image_file = 'img/graph.png'
            plt.savefig(image_file)
            # 라벨에 그림 할당
            pixmap = QPixmap(image_file)
            self.widget_submitdone_pie.setPixmap(pixmap)
            print('4번 위젯 완료')

            # 5번 위젯(그림)
            cursor.execute("SELECT 부서명, COUNT(부서명) FROM postlist GROUP BY 부서명")
            res = cursor.fetchall()
            print('db', res)
            # 변수 초기화
            labels = []
            counts = []
            # 각각 넣어주기
            for row in res:
                print(row[0])
                label = row[0]
                print(row[1])
                count = row[1]
                labels.append(label)
                counts.append(count)
            print('분류 카운트 :', counts, type(counts))
            # plt 생성
            plt.figure(figsize=(1.5, 1.5), facecolor='none')
            plt.pie(counts, autopct='%d')
            # 이미지 저장
            image_file = 'img/graph.png'
            plt.savefig(image_file)
            # 라벨에 그림 할당
            pixmap = QPixmap(image_file)
            self.widget_sortdepart.setPixmap(pixmap)

            # 6번 위젯(그림)
            cursor.execute("SELECT 종류, COUNT(종류) FROM postlist GROUP BY 종류")
            res = cursor.fetchall()
            print('db', res)
            # 변수 초기화
            labels = []
            counts = []
            # 각각 넣어주기
            for row in res:
                print(row[0])
                label = row[0]
                print(row[1])
                count = row[1]
                labels.append(label)
                counts.append(count)
            print('분류 카운트 :', counts, type(counts))
            # plt 생성
            plt.figure(figsize=(1.5, 1.5), facecolor='none')
            plt.pie(counts, autopct='%d')
            # 이미지 저장
            image_file = 'img/graph.png'
            plt.savefig(image_file)
            # 라벨에 그림 할당
            pixmap = QPixmap(image_file)
            self.widget_submitsort_pie.setPixmap(pixmap)

            # 7번 위젯
            if os.path.isfile("returnpostlist.db"):  # DB파일이 있을 경우
                listdb = sqlite3.connect("returnpostlist.db")
                cursor = listdb.cursor()
            else:  # DB파일이 없을 경우
                self.statusBar().showMessage('반려 DB가 없으므로 DB파일을 생성합니다.')
                listdb = sqlite3.connect("returnpostlist.db")
                cursor = listdb.cursor()
                cursor.execute(
                    "CREATE TABLE postlist(id integer, 접수날짜 text, 사원번호 text, 접수자명 text, 부서명 text, 보내는사람 text, 받는사람 text, 주소 text, 우편번호 text, 제목 text, 수량 integer, 종류 text, 긴급여부 boolean, 가로길이 real, 세로길이 real, 우편중량 real, 전화번호 text, 높이 real, 메모 text, 할인타입 text, 할인타입그룹 integer, 결재여부 , 반려여부 boolean, 비용 text)")
                cursor.commit()
            cursor.execute("SELECT COUNT(id) FROM postlist")
            countrtnres = cursor.fetchone()[0]
            self.widget_returnedpost.setText(str(countrtnres) + '건')

            # 8번 위젯(그림)
            listdb = sqlite3.connect("grouppostlist.db")
            cursor = listdb.cursor()
            cursor.execute("SELECT COUNT(DISTINCT 할인타입그룹) FROM postlist;")
            countgroup = cursor.fetchone()[0]
            print('group :', countgroup)
            self.widget_groupcount.setText(str(countgroup) + '건')

            # 9번 위젯
            cursor.execute("SELECT SUM(비용) FROM postlist WHERE 결재여부 = '1'")
            countfeeres = cursor.fetchone()[0]
            self.widget_feecount.setText(str(countfeeres) + '원')
        except:
            print('위젯 로드 실패')
        finally:
            listdb.close()


    # 프로그램 종료
    def quitwindow(self):
        QCoreApplication.instance().quit()

    # 우편요금 DB 리셋
    def resetpostdata(self):
        try:
            os.remove('postinformation.db') # 파일 삭제
            msg_reset = QMessageBox()  # 메세지박스 생성
            msg_reset.information(self, "알림", "리셋완료") # 메세지박스 설정
            output_reset = msg_reset.show()  # 메세지박스 실행
            if output_reset == QMessageBox.Ok: #OK 누르면 닫기
                msg_reset.close()
                self.mainmenu_statusBar.showMessage('리셋완료')
        except FileNotFoundError as e:
            print('지을 파일 없음 :',e)
            self.mainmenu_statusBar.showMessage(str(e))
            # 오류 메세지박스 생성 #
            msg_resetfail = QMessageBox()  # 메세지박스 생성
            msg_resetfail.warning(self, "오류", '이미 리셋되어 있습니다.')  # 메세지박스 경고옵션 설정
            output_resetfail = msg_resetfail.show()  # 메세지박스 실행
            if output_resetfail == QMessageBox.Ok:  # OK 누르면 닫기
                msg_resetfail.close()
        except PermissionError as e:
            print('오류 : ',e)
            self.mainmenu_statusBar.showMessage(str(e))
            # 오류 메세지박스 생성 #
            msg_permission = QMessageBox()  # 메세지박스 생성
            msg_permission.warning(self, "오류", '다른 프로세스가 파일을 사용중입니다.')  # 메세지박스 경고옵션 설정
            output_resetfail = msg_permission.show()  # 메세지박스 실행
            if output_resetfail == QMessageBox.Ok:  # OK 누르면 닫기
                msg_permission.close()

    # 우편접수리스트 DB 리셋
    def resetpostlist(self):
        try:
            os.remove('postlist.db')
            msg_reset = QMessageBox()
            msg_reset.information(self, "알림", "리셋완료")
            # 오류 메세지박스 생성(단순객체) #
            output_reset = msg_reset.show()
            if output_reset == QMessageBox.Ok:
                msg_reset.close()
                self.mainmenu_statusBar.showMessage('리셋완료')
        except FileNotFoundError as e:
            print('지을 파일 없음 :',e)
            self.mainmenu_statusBar.showMessage(str(e))
            # 오류 메세지박스 생성(단순객체) #
            msg_resetfail = QMessageBox()  # 메세지박스 생성
            msg_resetfail.warning(self, "오류", '이미 리셋되어 있습니다.')  # 메세지박스 경고옵션 설정
            output_resetfail = msg_resetfail.show()  # 메세지박스 실행
            if output_resetfail == QMessageBox.Ok:  # OK 누르면 닫기
                msg_resetfail.close()
        except PermissionError as e:
            print('오류 : ',e)
            self.mainmenu_statusBar.showMessage(str(e))
            # 오류 메세지박스 생성(단순객체) #
            msg_permission = QMessageBox()  # 메세지박스 생성
            msg_permission.warning(self, "오류", '다른 프로세스가 파일을 사용중입니다.')  # 메세지박스 경고옵션 설정
            output_resetfail = msg_permission.show()  # 메세지박스 실행
            if output_resetfail == QMessageBox.Ok:  # OK 누르면 닫기
                msg_permission.close()
    # 묶음우편물 DB 리셋
    def resetgrouplist(self):
        try:
            os.remove('grouppostlist.db')
            msg_reset = QMessageBox()
            msg_reset.information(self, "알림", "리셋완료")
            # 오류 메세지박스 생성(단순객체) #
            output_reset = msg_reset.show()
            if output_reset == QMessageBox.Ok:
                msg_reset.close()
                self.mainmenu_statusBar.showMessage('리셋완료')
        except FileNotFoundError as e:
            print('지을 파일 없음 :',e)
            self.mainmenu_statusBar.showMessage(str(e))
            # 오류 메세지박스 생성(단순객체) #
            msg_resetfail = QMessageBox()  # 메세지박스 생성
            msg_resetfail.warning(self, "오류", '이미 리셋되어 있습니다.')  # 메세지박스 경고옵션 설정
            output_resetfail = msg_resetfail.show()  # 메세지박스 실행
            if output_resetfail == QMessageBox.Ok:  # OK 누르면 닫기
                msg_resetfail.close()
        except PermissionError as e:
            print('오류 : ',e)
            self.mainmenu_statusBar.showMessage(str(e))
            # 오류 메세지박스 생성(단순객체) #
            msg_permission = QMessageBox()  # 메세지박스 생성
            msg_permission.warning(self, "오류", '다른 프로세스가 파일을 사용중입니다.')  # 메세지박스 경고옵션 설정
            output_resetfail = msg_permission.show()  # 메세지박스 실행
            if output_resetfail == QMessageBox.Ok:  # OK 누르면 닫기
                msg_permission.close()
    # 반려우편물 DB 리셋
    def resetreturnlist(self):

        try:
            os.remove('returnpostlist.db')
            msg_reset = QMessageBox()
            msg_reset.information(self, "알림", "리셋완료")
            # 오류 메세지박스 생성(단순객체) #
            output_reset = msg_reset.show()
            if output_reset == QMessageBox.Ok:
                msg_reset.close()
                self.mainmenu_statusBar.showMessage('리셋완료')
        except FileNotFoundError as e:
            print('지을 파일 없음 :', e)
            self.mainmenu_statusBar.showMessage(str(e))
            # 오류 메세지박스 생성(단순객체) #
            msg_resetfail = QMessageBox()  # 메세지박스 생성
            msg_resetfail.warning(self, "오류", '이미 리셋되어 있습니다.')  # 메세지박스 경고옵션 설정
            output_resetfail = msg_resetfail.show()  # 메세지박스 실행
            if output_resetfail == QMessageBox.Ok:  # OK 누르면 닫기
                msg_resetfail.close()
        except PermissionError as e:
            print('오류 : ', e)
            self.mainmenu_statusBar.showMessage(str(e))
            # 오류 메세지박스 생성(단순객체) #
            msg_permission = QMessageBox()  # 메세지박스 생성
            msg_permission.warning(self, "오류", '다른 프로세스가 파일을 사용중입니다.')  # 메세지박스 경고옵션 설정
            output_resetfail = msg_permission.show()  # 메세지박스 실행
            if output_resetfail == QMessageBox.Ok:  # OK 누르면 닫기
                msg_permission.close()

    # 필수라이브러리 설치
    def installrequirements(self):
        # 프로그램 실행에 필요한 라이브러리 설치
        import pip
        # 설치 시작 메세지
        self.statusBar().showMessage('시스템 : 애드온 라이브러리를 설치합니다.')
        # 설치
        pip.main(['install', '-r', 'requirements/requirements_addon.txt'])
        #설치 완료 시
        msg_completed = QMessageBox()  # 메세지박스 생성
        msg_completed.information(self, "알림", "설치완료")  # 메세지박스 설정
        output_completed = msg_completed.show()  # 메세지박스 실행
        if output_completed == QMessageBox.Ok:  # OK 누르면 닫기
            msg_completed.close()
    ###########################################################################################


########################################################
# 메인 코드
if __name__ == '__main__':
    #프로그램 실행
   app = QApplication(sys.argv) #프로그램 실행 클래스
   MainMenu_win = MainMenu() #윈도우 객체 생성
   app.exec() #이벤트루프 진입

