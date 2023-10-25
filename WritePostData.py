#PySide6 라이브러리
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *
#Pandas 라이브러리
import pandas as pd
#xlsx writer 라이브러리
import xlsxwriter
#SQLite 라이브러리
import sqlite3
#파일관리자 라이브러리
import os
# 시간 라이브러리
import time
# Selenium 과 크롬 드라이버 자동다운 라이브러리
from get_chrome_driver import GetChromeDriver
# Selenium
from selenium import webdriver
# 웹 항목 찾는 라이브러리 호출
from selenium.webdriver.common.by import By
# 드라이버 대기 밎 조건 라이브러리 호출
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class WritePostData(QWidget):
    def __init__(self):
        super().__init__()

        ## SQlite DB 불러와서 DF로 옮기기 ##
        # SQLite DB 연결
        groupdb = sqlite3.connect('grouppostlist.db')
        # SQL 쿼리 실행
        query = "SELECT * FROM groupedlist"
        try:
            df = pd.read_sql_query(query, groupdb)
        except pd.errors.DatabaseError as e:
            print('오류 : ', e)
            # 오류 메세지박스 생성(단순객체) #
            msg_nodb = QMessageBox()  # 메세지박스 생성
            msg_nodb.warning(self, "오류", 'DB에 레코드가 존재하지 않습니다.')  # 메세지박스 경고옵션 설정
            output_nodb = msg_nodb.show()  # 메세지박스 실행
            if output_nodb == QMessageBox.Ok:  # OK 누르면 닫기
                msg_nodb.close()
            #DB 객체 정리
            groupdb.close()
            self.close()
        else:
            # 데이터가 존재할 경우
            # 연결 종료
            groupdb.close()
            print('[DF]\n', df)
            # 할인타입그룹의 Unique 개수 저장
            groupcount = df['할인타입그룹'].nunique(dropna=True)  # None 값은 세지 않음
            print('할인타입그룹 UNIQUE 개수 : ', groupcount)
            # 접수할 우편그룹 선택 창 생성(매개변수에 UNIQUE 개수와 DB 입력)
            self.window = Msg_SelectGroup(groupcount, groupdb)
            ################################################################
            ## 창 닫히면 SQLite 불러와서 entry_var 의 그룹만 저장

            # 엔트리가 다 입력되었는지 체크(임시)
            print('[entry var]')
            print('entry var :', self.entry_var)
            if self.entry_var is None:
                pass  # 아무것도 하지 않음
            else:
                ## 웹서핑 시작 ##
                ## entry_var 에 선택한 그룹값이 있는 경우 Driver 불러오기 ##
                try:
                    self.driver = webdriver.Chrome()
                    self.driver.get(
                        "https://service.epost.go.kr/front.commonpostplus.RetrieveAcceptPlus.postal?gubun=1")
                    print('드라이버 불러와짐')
                except:
                    # 크롬 버젼에 맞는 드라이버가 없을 때 실행
                    # 코드 출처 : https://pypi.org/project/get-chrome-driver/
                    # Downloads ChromeDriver for the installed Chrome version on the machine
                    # Adds the downloaded ChromeDriver to path
                    get_driver = GetChromeDriver()
                    get_driver.install()
                    print('크롬 드라이버 업데이트 완료')
                else:
                    print('웹 탐색 시작')
                    # 웹이 로드되길 대기(10초까지)
                    self.driver.implicitly_wait(time_to_wait=10)
                    print('로드 완료')
                    # 웹 자동화 시작
                    try:
                        self.startwriting()
                    except Exception as e:
                        # 오류 발생 시
                        print("Chrome Driver 에서 오류 발생 :", e)
                finally:
                    # 7초 대기
                    print('웹드라이버가 종료됩니다 (3초 후)')
                    time.sleep(3)
                    # driver 가 켜져있을 경우
                    if self.driver != None:
                        self.driver.quit()
                    # 클래스 종료
                    self.close()


    ################################################################
    #사이트에 정보 입력
    ################################################################
    def startwriting(self):
        # SQLite DB 연결
        groupdb = sqlite3.connect('grouppostlist.db')
        # SQL 쿼리 실행
        query = "SELECT * FROM groupedlist WHERE 할인타입그룹 = '%s'" % (self.entry_var)
        print(' 조회 쿼리 :', query)
        self.df = pd.read_sql_query(query, groupdb)

        print('입력 시작')
        print('[Debug : pswd]', self.submit_password)
        print('[Debug : entry_var]', self.entry_var, type(self.entry_var))
        #인터넷우체국 이용약관 동의 체크
        print('인터넷우체국 이용약관 동의 체크')
        check_btn_agreeall = self.driver.find_element(By.CLASS_NAME, 'check_all_label')
        #주의 driver 의 find_element_by_class_name은 위의 메소드로 변경됨
        print('찾음 :',check_btn_agreeall)
        check_btn_agreeall.click()

        #패스워드 입력
        time.sleep(0.5) #딜레이
        print('패스워드 입력')
        entry_post_password = self.driver.find_element(By.ID, 'guest_orderpw')
        print('찾음 :',entry_post_password)
        entry_post_password.send_keys(self.submit_password)

        #보내는 사람 입력
        time.sleep(0.5)  # 딜레이
        print('보내는 사람 입력')
        print('sender')
        print('sender :', self.df)
        sender = self.df.loc[0, '보내는사람'] #DF의 첫 레코드값을 이용
        print('DF 불러옴')
        entry_sender = self.driver.find_element(By.ID, 'tSndNm')
        print('찾음 :', entry_sender)
        entry_sender.send_keys(sender)

        # 주소입력
        print('주소 입력')
        btn_addr_search = self.driver.find_element(By.ID, 'SrchAddrBtn').click()
        # 새 창으로 컨트롤 이동
        self.driver.switch_to.window(self.driver.window_handles[1])
        print('주소검색창 전환 완료')
        # 로드될 때까지 대기
        self.driver.implicitly_wait(time_to_wait=10)
        # 주소 검색
        entry_addrsearch = self.driver.find_element(By.ID, 'keyword')
        entry_addrsearch.send_keys(self.var_sender_addr)
        entry_addrsearch_start = self.driver.find_element(By.ID, 'btnImgSrch').click()
        # 두 번째 창을 닫혀서 창의 수가 1개가 될때까지 대기
        WebDriverWait(self.driver, 9999).until(EC.number_of_windows_to_be(1))
        print('주소검색창 닫기 완료')
        # 원래 창으로 스위칭합니다.
        self.driver.switch_to.window(self.driver.window_handles[0])

        # 연락처(휴대전화)입력
        time.sleep(0.5)  # 딜레이
        print('연락처(휴대전화) 입력')
        entry_phoneno_1 = self.driver.find_element(By.ID, 'tSndHTel1')
        print('찾음 :', entry_phoneno_1)
        entry_phoneno_1.send_keys('') #010 자동입력

        entry_phoneno_2 = self.driver.find_element(By.ID, 'tSndHTel2')
        print('찾음 :', entry_phoneno_2)
        entry_phoneno_2.send_keys(self.var_phoneno_2)

        entry_phoneno_3 = self.driver.find_element(By.ID, 'tSndHTel3')
        print('찾음 :', entry_phoneno_3)
        entry_phoneno_3.send_keys(self.var_phoneno_3)

        # 연락처(일반전화)입력
        time.sleep(0.5) #딜레이  # 딜레이
        entry_telno_1 = self.driver.find_element(By.ID, 'tSndTel1')
        print('찾음 :', entry_telno_1)
        entry_telno_1.send_keys(self.var_telno_1)

        entry_telno_2 = self.driver.find_element(By.ID, 'tSndTel2')
        print('찾음 :', entry_telno_2)
        entry_telno_2.send_keys(self.var_telno_2)

        entry_telno_3 = self.driver.find_element(By.ID, 'tSndTel3')
        print('찾음 :', entry_telno_3)
        entry_telno_3.send_keys(self.var_telno_3)

        #배송정보수신여부 입력
        time.sleep(0.5) #딜레이
        if self.var_is_receiveinfo == True:
            entry_is_receiveinfo = self.driver.find_element(By.ID, 'recevYn4').click()
        else:
            pass

        # 이메일 입력
        entry_email = self.driver.find_element(By.ID, 'sendprsnemail')
        print('찾음 :', entry_email)
        entry_email.send_keys(self.var_entry_email)

        #보내는 등기 중량 입력
        gravity = self.df.loc[0, '우편중량']
        entry_gravity = self.driver.find_element(By.ID, 'tWeight1')
        print('찾음 :', entry_gravity)
        entry_gravity.send_keys(gravity)


        #xlsx 파일 생성
        # 라이브러리 호출
        sqldb = sqlite3.connect("grouppostlist.db")
        sqlcursor = sqldb.cursor()
        # sql db에서 df 로 변환
        query = 'SELECT 받는사람, 우편번호, 주소, 전화번호, 우편중량 FROM groupedlist WHERE "할인타입그룹" = %s' % (self.entry_var)
        print('쿼리 :', query)
        df = pd.read_sql(query, sqldb)
        # 불러와진 DF 확인
        print('[불러와진 DF]\n', df)

        # 우체국 양식에 맞게 변경
        #인덱스 칼럼 제거
        print('DF 가공 시작')
        df.rename(columns={'받는사람': '받는 분'}, inplace=True)
        df.rename(columns={'주소': '주소(시도+시군구+도로명+건물번호'}, inplace=True)
        # 현재 DB에 구현되어 있지 않은 칼럼 임시로 채워두기
        df.insert(3, '상세주소(동, 호수, 洞명칭, 아파트, 건물명 등)', '11111')
        df.insert(5, '일반전화(02-1234-5678)', '')
        df.insert(6, '등기번호(선납소포라벨만 입력가능)', '')
        #############################################
        df.rename(columns={'전화번호': '휴대전화(010-1234-5678)'}, inplace=True)
        df.rename(columns={'우편중량': '중량'}, inplace=True)
        df.rename(columns={'메모': '반려사유'}, inplace=True)
        # 출력할 DF 확인
        print('[xlsx 전달할 DF]\n', df)


        print('xlsxwrite 시작')
        # 작성 이전에 존재하는 xlsx 파일 지우기
        try:
            os.remove('df_output.xlsx')
            print('파일 삭제 완료')
        except:
            print('파일 없음')
        # xlsx 작성 시작 (참조 : https://minwook-shin.github.io/python-creating-excel-xlsx-files-using-xlsxwriter/)
        writer = pd.ExcelWriter('df_output.xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        # Close the Pandas Excel writer and output the Excel file.
        writer.close()
        # xlsx 파일 경로 담기
        inputfile = os.path.abspath('df_output.xlsx')
        print('절대경로 :', inputfile)


        # xlsx 파일 추가
        time.sleep(0.5) #딜레이
        btn_addr_file = self.driver.find_element(By.ID, 'btnUseAddrFile').click()
        # 새 창으로 컨트롤 이동
        self.driver.switch_to.window(self.driver.window_handles[1])
        print('파일선택창 전환 완료')
        # 로드될 때까지 대기
        self.driver.implicitly_wait(time_to_wait=10)
        # 파일선택 버튼 클릭
        print('파일선택 클릭')

        ### web엔티티에 파일 직접 입력시 인식이 안되므로 직접 파일 선택해야함
        #btn_submit_file = self.driver.find_element(By.ID, 'uploadFile')
        #btn_submit_file.send_keys(inputfile)

        # 두 번째 창을 닫아서 창의 수가 1개가 될때까지 대기
        WebDriverWait(self.driver, 9999).until(EC.number_of_windows_to_be(1))
        #추후 접수하면 됩니다.
        print('접수 끝')
        print('접수신청 버튼을 누르세요')
        # 첫 번째 창을 닫아서 창의 수가 0개가 될때까지 대기
        WebDriverWait(self.driver, 9999).until(EC.number_of_windows_to_be(0))
        print('창 닫기')








############################################################
# Dialog 구성 및 선택한 그룹 번호 entry_var 에 입력하는 클래스
class Msg_SelectGroup(QDialog):
    def __init__(self, groupcount, groupdb):
        super().__init__()
        # 윈도우 설정
        self.setWindowIcon(QIcon('title_big.png'))
        self.setWindowTitle('묶음그룹 선택')
        self.resize(300, 100)

        # 레이아웃 생성 및 연결
        self.layout = QVBoxLayout()

        # Label 추가
        self.desc_label = QLabel('사전접수를 진행할 묶음그룹을 선택하세요.')
        self.layout.addWidget(self.desc_label)


        # QPushButton 추가
        # 버튼 엔티티를 담을 list 추가
        for i in range(groupcount):
            radio_button = QRadioButton(f"{i + 1} 그룹", self)
            self.layout.addWidget(radio_button)
            radio_button.clicked.connect(self.setentryvar)
        #########################################################
        # 구분자 추가 (PySide6 에 구분자가 없어서 Frame 으로 우회생성)
        self.separator1 = QFrame(self)
        self.separator1.setFrameShape(QFrame.HLine)
        self.separator1.setFrameShadow(QFrame.Sunken)
        self.layout.addWidget(self.separator1)

        # Label & Line Edit 추가
        self.desc_password = QLabel('접수 시 비밀번호로 숫자 4자리를 입력해주세요.')
        self.layout.addWidget(self.desc_password)
        # desc_password 엔트리에 정수만 입력받게 하기
        self.entry_password = QLineEdit()
        self.entry_password_validator = QIntValidator(0, 9999)
        self.entry_password.setValidator(self.entry_password_validator)
        self.layout.addWidget(self.entry_password)

        # QCheckBox 추가
        self.is_receiveinfo = QCheckBox('배송정보를 수신합니다.')
        self.layout.addWidget(self.is_receiveinfo)

        #########################################################
        # 주소추가 (현재는 주소를 제대로 입력했을 시만 사용가능, 추후 여기에 Daum 주소검색 API 이용하여 유연성 조절필요)
        # Label & Line Edit 추가
        self.desc_sender_addr = QLabel('보내는 주소 입력')
        self.layout.addWidget(self.desc_sender_addr)
        self.entry_sender_addr = QLineEdit()
        self.layout.addWidget(self.entry_sender_addr)

        #########################################################
        # 구분자2 추가
        self.separator2 = QFrame(self)
        self.separator2.setFrameShape(QFrame.HLine)
        self.separator2.setFrameShadow(QFrame.Sunken)
        self.layout.addWidget(self.separator2)

        # Label 추가
        self.desc_callno = QLabel('접수 시 번호를 입력해 주세요.')
        self.layout.addWidget(self.desc_callno)

        # Label 추가
        self.desc_phoneno = QLabel('휴대폰 번호.')
        self.layout.addWidget(self.desc_phoneno)

        # 휴대폰번호 담기 위한 가로 Layout 추가
        self.phonenolayout = QHBoxLayout()
        self.phoneno_1 = QLineEdit()
        self.phoneno_1_validator = QIntValidator(0, 999)
        self.phoneno_1.setValidator(self.phoneno_1_validator)

        self.phoneno_slash1 = QLabel('-')
        self.phoneno_2 = QLineEdit()
        self.phoneno_2_validator = QIntValidator(0, 9999)
        self.phoneno_2.setValidator(self.phoneno_2_validator)

        self.phoneno_slash2 = QLabel('-')
        self.phoneno_3 = QLineEdit()
        self.phoneno_3_validator = QIntValidator(0, 9999)
        self.phoneno_3.setValidator(self.phoneno_3_validator)

        self.phonenolayout.addWidget(self.phoneno_1)
        self.phonenolayout.addWidget(self.phoneno_slash1)
        self.phonenolayout.addWidget(self.phoneno_2)
        self.phonenolayout.addWidget(self.phoneno_slash2)
        self.phonenolayout.addWidget(self.phoneno_3)

        # 전화번호 가로 레이아웃을 기존 레이아웃에 추가
        self.layout.addLayout(self.phonenolayout)
        #########################################################
        # Label 추가
        self.desc_telno = QLabel('전화번호.')
        self.layout.addWidget(self.desc_telno)

        # 전화번호 담기 위한 가로Layout 추가
        self.telnolayout = QHBoxLayout()
        self.telno_1 = QLineEdit()
        self.telno_1_validator = QIntValidator(0, 999)
        self.telno_1.setValidator(self.telno_1_validator)

        self.telno_slash1 = QLabel('-')
        self.telno_2 = QLineEdit()
        self.telno_2_validator = QIntValidator(0, 9999)
        self.telno_2.setValidator(self.telno_2_validator)

        self.telno_slash2 = QLabel('-')
        self.telno_3 = QLineEdit()
        self.telno_3_validator = QIntValidator(0, 9999)
        self.telno_3.setValidator(self.telno_3_validator)

        self.telnolayout.addWidget(self.telno_1)
        self.telnolayout.addWidget(self.telno_slash1)
        self.telnolayout.addWidget(self.telno_2)
        self.telnolayout.addWidget(self.telno_slash2)
        self.telnolayout.addWidget(self.telno_3)

        # 전화번호 가로 레이아웃을 기존 레이아웃에 추가
        self.layout.addLayout(self.telnolayout)

        #########################################################
        # 구분자3 추가
        self.separator3 = QFrame(self)
        self.separator3.setFrameShape(QFrame.HLine)
        self.separator3.setFrameShadow(QFrame.Sunken)
        self.layout.addWidget(self.separator3)

        # 이메일 담기 위한 가로Layout 추가
        self.email_layout = QHBoxLayout()
        self.desc_email = QLabel('이메일 :')
        self.entry_email = QLineEdit()
        self.email_layout.addWidget(self.desc_email)
        self.email_layout.addWidget(self.entry_email)
        # 이메일 레이아웃을 기존 레이아웃에 추가
        self.layout.addLayout(self.email_layout)

        #########################################################
        # 구분자4 추가
        self.separator4 = QFrame(self)
        self.separator4.setFrameShape(QFrame.HLine)
        self.separator4.setFrameShadow(QFrame.Sunken)
        self.layout.addWidget(self.separator4)

        #버튼을 담기 위한 Layout 추가
        self.btnlayout = QHBoxLayout()

        # 버튼 추가
        self.btn_proceed = QPushButton("진행")
        self.btnlayout.addWidget(self.btn_proceed)
        self.btn_closewin = QPushButton("취소")
        self.btnlayout.addWidget(self.btn_closewin)

        #버튼 레이아웃을 기존 레이아웃에 추가
        self.layout.addLayout(self.btnlayout)
        #########################################################
        #레이아웃 설정
        self.setLayout(self.layout)

        # 버튼 시그널 연결
        self.btn_proceed.clicked.connect(self.runprocess)
        self.btn_closewin.clicked.connect(self.closewindow)

        # 창 생성
        self.exec()

    #버튼 클릭 시 해당 그룹의 값을 entry_var 에 입력
    def setentryvar(self):
        # 선택된 버튼의 텍스트를 가져와서 entry_var에 할당
        # Sender : sender()는 PySide/PyQt 에서 이벤트를 발생시킨 객체를 반환하는 함수
        sender = self.sender()
        # 그룹 2는 SQL DB상에서 할인그룹 1이기에 -1 을 함
        WritePostData.entry_var = int(sender.text().split(' 그룹')[0]) - 1
        print('선택 :', WritePostData.entry_var)

    #창 닫기
    def closewindow(self):
        # Dialog 닫기
        self.close()

    #선택된 그룹 저장하고 자동화 수행 시작
    def runprocess(self):
        # 전달할 패스워드 설정
        WritePostData.submit_password = self.entry_password.text()
        WritePostData.var_is_receiveinfo = self.is_receiveinfo.isChecked()
        WritePostData.var_phoneno_1 = self.phoneno_1.text()
        WritePostData.var_phoneno_2 = self.phoneno_2.text()
        WritePostData.var_phoneno_3 = self.phoneno_3.text()
        WritePostData.var_telno_1 = self.telno_1.text()
        WritePostData.var_telno_2 = self.telno_2.text()
        WritePostData.var_telno_3 = self.telno_3.text()
        WritePostData.var_entry_email = self.entry_email.text()
        WritePostData.var_sender_addr = self.entry_sender_addr.text()


        print('[debug] pswd : ', WritePostData.submit_password)
        # 추후 패스워드 미입력시 경고창 팝업되게 수정필요

        #Dialog 닫기
        self.close()

################################################################