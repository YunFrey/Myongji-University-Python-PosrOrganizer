# PySide6 로드
from PySide6.QtWidgets import *
from PySide6.QtCore import QDate, Qt
# ui 로드용
from PySide6 import QtUiTools
from ui_loader import load_ui
# SQLite 로드
import sqlite3
# 파일 관리자
import os
# 시간 관련 라이브러리
from datetime import datetime
# Pandas 로드
import pandas as pd


class SubmitPostMenu(QMainWindow):
    ##### 사용하는 전역변수 선언 #####
    global listdb
    global cursor

    def __init__(self):
        super().__init__()
        load_ui('SubmitPostMenu.ui', self)
        self.show()

        # [버튼 시그널 생성]
        # 프로그램 종료
        self.action_closewindow.triggered.connect(self.closewindow)
        # 접수된 우편물 조회
        self.btn_loadpostlist.clicked.connect(self.loadpostlist)
        # 수정 버튼
        self.btn_modifypostcolumn.clicked.connect(self.modifypostsave)
        # 전체 선택 버튼
        self.btn_selectallrow.clicked.connect(self.selectallrow)

        # [QDateWidget 설정]
        self.datesel_start.setDate(QDate.currentDate())
        self.datesel_end.setDate(QDate.currentDate())
        self.datesel_start.setCalendarPopup(True)
        self.datesel_end.setCalendarPopup(True)

        # [QDateWidget 설정 2] -시작날짜가 끝 날짜와 겹치지 않게 날짜 변경시에 보정
        self.datesel_start.dateChanged.connect(lambda: self.leftdatefix())
        self.datesel_end.dateChanged.connect(lambda: self.rightdatefix())

        # [QTableWidget 레코드 선택 시]
        self.table_postlist.selectionModel().selectionChanged.connect(self.editrow_postinfo)

        # 이 창을 열었을 때 DB 조회 및 생성
        self.loadpostdb()
        # [우편레코드 제거버튼]
        self.btn_delpostcolumn.clicked.connect(self.deleterowselected)
        # [우편레코드 추가버튼]
        self.btn_addpostcolumn.clicked.connect(self.addpostcolumn)
        # [우편레코드 저장]
        self.btn_savepostlist.clicked.connect(self.savepostlist)
        # [사원번호 치고 엔터 시 실행 이벤트 연결]
        self.info_0_importerid.returnPressed.connect(self.getnamefromid)
        # [우편종류 설정]
        self.initiate_lengf()
        self.initiate_sorttype()

    def initiate_lengf(self):
        self.info_9_horizontal_format.addItem('cm')
        self.info_9_horizontal_format.addItem('m')
        self.info_10_vertical_format.addItem('cm')
        self.info_10_vertical_format.addItem('m')
        self.info_11_weight_format.addItem('g')
        self.info_11_weight_format.addItem('kg')
        self.info_12_height_format.addItem('cm')
        self.info_12_height_format.addItem('m')

    def initiate_sorttype(self):
        self.info_8_combo_sort.clear()
        self.info_8_combo_sort.addItem('우편')
        self.info_8_combo_sort.addItem('등기')
        self.info_8_combo_sort.addItem('익일특급')
        self.info_8_combo_sort.addItem('일반소포')
        self.info_8_combo_sort.addItem('등기소포')

    def __getidlist(self):  # private method
        if os.path.isfile("employee.db"):  # DB파일이 있을 경우
            # DB 불러오기
            id_db = sqlite3.connect("employee.db")
            id_cursor = id_db.cursor()
            input = str(self.info_0_importerid.text())
            query = 'SELECT ID FROM emplist'
            id_cursor.execute(query)
            output = id_cursor.fetchall()
            return output

        else:  # DB파일이 없을 경우
            self.statusBar().showMessage('에러 : employee.db 파일 없음')

    def getnamefromid(self):
        if os.path.isfile("employee.db"):  # DB파일이 있을 경우
            # DB 불러오기
            empdb = sqlite3.connect("employee.db")
            empcursor = empdb.cursor()
            input = str(self.info_0_importerid.text())
            query = 'SELECT NAME, DEPT FROM emplist WHERE ID = %s' % input
            try:
                empcursor.execute(query)
                output = empcursor.fetchone()
                self.info_1_importer.setText(output[0])
                self.info_2_depart.setText(output[1])
            except sqlite3.OperationalError as e:
                # 잘못된 입력이라고 표시
                msg_incorrect = QMessageBox()  # 메세지박스 생성
                msg_incorrect.information(self, "알림", "잘못된 형식입니다.")  # 메세지박스 설정
                output_incorrect = msg_incorrect.show()  # 메세지박스 실행
                if output_incorrect == QMessageBox.Ok:  # OK 누르면 닫기
                    msg_incorrect.close()
            except TypeError as e:
                # 없는 직원이라고 표시
                msg_notfound = QMessageBox()  # 메세지박스 생성
                msg_notfound.information(self, "알림", "해당 사번의 직원이 없습니다..")  # 메세지박스 설정
                output_notfound = msg_notfound.show()  # 메세지박스 실행
                if output_notfound == QMessageBox.Ok:  # OK 누르면 닫기
                    msg_notfound.close()
        else:  # DB파일이 없을 경우
            self.statusBar().showMessage('에러 : employee.db 파일 없음')

    def rightdatefix(self):
        if (int(self.datesel_start.date().toString('yyyyMMdd'))) > (int(self.datesel_end.date().toString('yyyyMMdd'))):
            self.datesel_start.setDate(self.datesel_end.date())

    def leftdatefix(self):
        if (int(self.datesel_end.date().toString('yyyyMMdd'))) < (int(self.datesel_start.date().toString('yyyyMMdd'))):
            self.datesel_end.setDate(self.datesel_start.date())

    def closewindow(self):  # 창 닫기
        print('SubmitPostMenu closed')
        self.close()

    def loadpostlist(self):
        # 시작날짜 변수에 담기
        str_datestart = self.datesel_start.date().toString("yyyyMMdd")
        print(str_datestart)
        # 끝나는날짜 변수에 남기
        str_dateend = self.datesel_end.date().toString("yyyyMMdd")
        print(str_dateend)
        # SQL DB 를 QTableWidget 에 표시
        # 오늘날짜 저장
        todaydat: str = datetime.today().strftime("%Y%m%d")
        # 결재된 우편물도 조회하는 옵션 켜져잇을 경우
        if self.info_14_isprocessed.isChecked() == True:
            optional = "AND 결재여부 = '1'"
        else:
            optional = "AND 결재여부 IS NULL"
        # 쿼리 완성
        query = 'SELECT id, 접수날짜, 사원번호, 접수자명, 부서명, 보내는사람, 받는사람, 주소, 우편번호, 제목, 수량, 종류, 긴급여부, 가로길이, 세로길이, 우편중량, 전화번호, 높이, 메모, 결재여부 FROM postlist WHERE "접수날짜" >= %s AND "접수날짜" <= %s %s' % (
        str_datestart, str_dateend, optional)
        print('쿼리 : ', query)
        #DB가 없을 때에는 실행 안함
        try:
            df = pd.read_sql(query, listdb)
        except pd.errors.DatabaseError:
            print('시스템 : DB 없음, 새로운')
        else:
            # 만약 조회결과가 없을 경우 아무것도 안나오고 오류 출력
            if df.empty == True:  # 데이터프레임이 비어있을 떄
                self.table_postlist.clear()  # QTableWidget 초기화
                # 오류 메세지창 팝업
                self.statusBar().showMessage('알림 : 해당 날짜에 조회된 내용이 없습니다.')
            else:  # 데이터프레임이 비지 않으면 다음 내용 시작
                # QTableWidget 에 DF 내용 쓰기
                self.table_postlist.clear()  # 부르기 전 초기화
                col = len(df.keys())
                self.table_postlist.setColumnCount(col)
                self.table_postlist.setHorizontalHeaderLabels(df.keys())
                row = len(df.index)
                self.table_postlist.setRowCount(row)
                for r in range(row):
                    for c in range(col):
                        item = QTableWidgetItem(str(df.iloc[r][c]))
                        self.table_postlist.setItem(r, c, item)
                self.table_postlist.resizeColumnsToContents()



    def loadpostdb(self):
        # 전역변수 불러오기
        global listdb
        global cursor
        print('DB 로드')
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

    def savepostlist(self):
        print('저장')
        listdb.commit()

    def getentryinfo(self):
        # 변수에 담기
        importerid = self.info_0_importerid.text()
        importer = self.info_1_importer.text()
        depart = self.info_2_depart.text()
        sender = self.info_3_sender.text()
        receiver = self.info_4_receiver.text()
        address = self.info_5_address.text()
        is_urgent = self.info_z_isurgent.isChecked()
        title = self.info_6_title.text()
        quantity = self.info_7_spin_quantity.value()
        sort = self.info_8_combo_sort.currentText()
        horizontal = self.info_9_horizontal.text()
        horizontal_f = self.info_9_horizontal_format.currentText()
        vertical = self.info_10_vertical.text()
        vertical_f = self.info_10_vertical_format.currentText()
        weight = self.info_11_weight.text()
        weight_f = self.info_11_weight_format.currentText()
        height = self.info_12_height.text()
        height_f = self.info_12_height_format.currentText()
        memo = self.info_13_memo.text()
        postno = self.info_14_postno.text()
        phonenum = self.info_15_phonenum.text()
        # 포맷변환(m -> cm, kg -> g)
        if horizontal_f == 'm':
            horizontal = int(horizontal) * 100
        if vertical_f == 'm':
            vertical = int(vertical) * 100
        if height_f == 'm':
            height = int(height) * 100
        if weight_f == 'kg':
            weight = int(weight) * 1000
        # 리스트 생성(우편물 한건당 정보를 담음)
        postline = list()
        postline.append(datetime.today().strftime("%Y%m%d"))
        postline.append(importerid)
        postline.append(importer)
        postline.append(depart)
        postline.append(sender)
        postline.append(receiver)
        postline.append(address)
        postline.append(postno)
        postline.append(title)
        postline.append(quantity)
        postline.append(sort)
        postline.append(is_urgent)
        postline.append(horizontal)
        postline.append(vertical)
        postline.append(weight)
        postline.append(phonenum)
        postline.append(height)
        postline.append(memo)
        print('[Postline 정보]\n', postline)
        # return
        return postline

    def addpostcolumn(self):
        #SQL 있으면 DB 연결
        if os.path.isfile("postlist.db"):  # DB파일이 있을 경우
            listdb = sqlite3.connect("postlist.db")
            cursor = listdb.cursor()
        else:  # DB파일이 없을 경우
            self.statusBar().showMessage('DB가 없으므로 DB파일을 생성합니다.')
            listdb = sqlite3.connect("postlist.db")
            cursor = listdb.cursor()
            cursor.execute(
                "CREATE TABLE postlist(id integer PRIMARY KEY AUTOINCREMENT, 접수날짜 text, 사원번호 text, 접수자명 text, 부서명 text, 보내는사람 text, 받는사람 text, 주소 text, 우편번호 text, 제목 text, 수량 integer, 종류 text, 긴급여부 text, 가로길이 real, 세로길이 real, 우편중량 real, 전화번호 text, 높이 real, 메모 text, 할인타입 text, 할인타입그룹 integer, 결재여부 , 반려여부 boolean, 비용 text)")
            listdb.commit()
        print('우편물 추가 실행')
        # entry 값 얻는 함수 실행(현재는 첫 엔트리만 체크)
        postline = self.getentryinfo()
        print('postline : ', postline)
        #아무것도 입력하지 않았을 때 오류 생성
        try:
            if postline[1] == '':
                raise Exception
        except Exception as e:  # 예외가 발생했을 때 실행됨
            msg_noentry = QMessageBox()  # 메세지박스 생성
            msg_noentry.warning(self, "오류", "추가할 우편물 정보를 입력하세요")  # 메세지박스 설정
            output_noentry = msg_noentry.show()  # 메세지박스 실행
            if output_noentry == QMessageBox.Ok:  # OK 누르면 닫기
                msg_noentry.close()
        else:
            #입력된 게 있을 경우
            # 긴급여부
            if postline[11] == '0':
                tempres = '0'
            else:
                tempres = '1'
            # 라인 추가
            print('추가 SQL 시작')
            cursor.execute(
                "INSERT INTO postlist(접수날짜, 사원번호, 접수자명, 부서명, 보내는사람, 받는사람, 주소, 우편번호, 제목, 수량, 종류, 긴급여부, 가로길이, 세로길이, 우편중량, 전화번호, 높이, 메모) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                postline)
            listdb.commit()

            # Table 업데이트
            self.loadpostlist()
            # 이건 메소드만 불러오기
            # QLineEdit 초기화
            self.clearallentry()
            # 초기화함수 실행
            self.initiate_lengf()
            self.initiate_sorttype()
        finally:
            listdb.close()



    def clearallentry(self):
        self.info_0_importerid.clear()
        self.info_1_importer.clear()
        self.info_2_depart.clear()
        self.info_3_sender.clear()
        self.info_4_receiver.clear()
        self.info_5_address.clear()
        self.info_z_isurgent.setChecked(False)
        self.info_6_title.clear()
        self.info_7_spin_quantity.clear()
        self.info_9_horizontal.clear()
        self.info_9_horizontal_format.clear()
        self.info_10_vertical.clear()
        self.info_10_vertical_format.clear()
        self.info_11_weight.clear()
        self.info_11_weight_format.clear()
        self.info_12_height.clear()
        self.info_12_height_format.clear()
        self.info_13_memo.clear()
        ### 우편번호와 전화번호 추가
        self.info_14_postno.clear()
        self.info_15_phonenum.clear()

    def editrow_postinfo(self):
        post_tofix = self.getselectedrowid_adv(self.table_postlist)
        print('디버그 :', post_tofix)
        # 칼럼 선택시 불러진 데이터값이 0 이 아닐 경우 엔트리 업데이트하기
        if (len(post_tofix) > 0):
            self.info_0_importerid.setText(post_tofix[2])
            self.info_1_importer.setText(post_tofix[3])
            self.info_2_depart.setText(post_tofix[4])
            self.info_3_sender.setText(post_tofix[5])
            self.info_4_receiver.setText(post_tofix[6])
            self.info_5_address.setText(post_tofix[7])
            self.info_14_postno.setText(post_tofix[8])
            self.info_6_title.setText(post_tofix[9])
            self.info_7_spin_quantity.setValue(int(post_tofix[10]))

            #CheckBox 는 str타입이 아닌 boolean 타입만 받음
            if post_tofix[12] == '0':
                tempres = 0
            else:
                tempres = 1

            # combobox 는 setText 가 안되니 작업 추가
            index = self.info_8_combo_sort.findText(post_tofix[11], Qt.MatchFixedString)
            self.info_8_combo_sort.setCurrentIndex(index)
            self.info_z_isurgent.setChecked(tempres)
            self.info_9_horizontal.setText(post_tofix[13])
            self.info_9_horizontal_format.setCurrentIndex(0)
            self.info_10_vertical.setText(post_tofix[14])
            self.info_9_horizontal_format.setCurrentIndex(0)
            self.info_11_weight.setText(post_tofix[15])
            self.info_9_horizontal_format.setCurrentIndex(0)
            self.info_11_weight.setText(post_tofix[15])
            self.info_15_phonenum.setText(post_tofix[16])
            self.info_12_height.setText(post_tofix[17])
            self.info_9_horizontal_format.setCurrentIndex(0)
            self.info_13_memo.setText(post_tofix[18])

        # 선택된 줄에 대한 ID 데이터를 읽는 함수

    def getselectedrowid(self, tablename):
        # 인텍스에 선택된 롤 집어넣기
        index = tablename.selectionModel().selectedRows()
        # 리스트 초기화
        templist = list()
        for row in index:  # index를 정렬하여 row 로 읽기
            # row 정보 중 0번칼럼 읽어서 텍스트로 출력
            id = tablename.item(row.row(), 0).text()
            # 선택된 모든 row 읽어서 리스트에 추가
            templist.append(id)
        return templist

    # 선택된 줄에 대한 모든 데이터를 읽는 함수
    def getselectedrowid_adv(self, tablename):
        # 인텍스에 선택된 롤 집어넣기
        index = tablename.selectionModel().selectedRows()
        # 리스트 초기화
        templist = list()
        for row in index:  # index를 정렬하여 row 로 읽기
            # 읽힌 모든 거 다 append 하기
            for i in range(tablename.columnCount()):
                tempvar = tablename.item(row.row(), i).text()
                templist.append(tempvar)
        return templist

    def deleterowselected(self):
        # 선택한 레코드가 없을 경우 오류 생성
        try:
            if len(self.table_postlist.selectionModel().selectedRows()) == 0:
                raise Exception
        except Exception as e:  # 예외가 발생했을 때 실행됨
            msg_noentry = QMessageBox()  # 메세지박스 생성
            msg_noentry.warning(self, "오류", "삭제할 우편물을 선택해 주세요")  # 메세지박스 설정
            output_noentry = msg_noentry.show()  # 메세지박스 실행
            if output_noentry == QMessageBox.Ok:  # OK 누르면 닫기
                msg_noentry.close()
        else:
            #선택한 레코드가 있을 경우
            indexlist = self.getselectedrowid(self.table_postlist)
            # SQL DELETE 문 실행
            # cursor.executemany("DELETE FROM postlist WHERE id = ?", indexlist)
            cursor.execute("DELETE FROM postlist WHERE id IN (%s)" % ",".join(indexlist))
            # 커밋
            listdb.commit()
            # 엔트리 초기화
            self.clearallentry()
            # 테이블 업데이트
            self.loadpostlist()



        ## 선택된 테이블 칼럼의 줄값정보를 가져오는 함수 (1개 or 마지막개수)

    def getrowid(self, tablename):
        try:
            idtofix = ''
            for row in self.table_postlist.selectionModel().selectedRows():
                idtofix = self.table_postlist.item(row.row(), 0).text()
        except UnboundLocalError as e:
            self.statusBar().showMessage('선택된 레코드가 없습니다.')

        return idtofix

    def modifypostsave(self):
        postline = self.getentryinfo()
        #수정할 우편물이 선택되지 않았을 경우 오류 생성
        try:
            if len(self.table_postlist.selectionModel().selectedRows()) == 0:
                raise Exception
        except Exception as e:  # 예외가 발생했을 때 실행됨
            msg_noentry = QMessageBox()  # 메세지박스 생성
            msg_noentry.warning(self, "오류", "수정할 우편물을 선택해 주세요")  # 메세지박스 설정
            output_noentry = msg_noentry.show()  # 메세지박스 실행
            if output_noentry == QMessageBox.Ok:  # OK 누르면 닫기
                msg_noentry.close()
        else:
            # 수정할 우편물이 선택되어 있을 경우
            # 선택한 라인 저장
            rowid = self.getrowid(self.table_postlist)
            # 라인 추가
            print('postline :', postline)
            print('rowid :', rowid)
            cursor.execute(
                "UPDATE postlist SET 접수날짜='%s', 사원번호='%s', 접수자명='%s', 부서명='%s', 보내는사람='%s', 받는사람='%s', 주소='%s', 우편번호='%s', 제목='%s', 수량='%s', 종류='%s', 긴급여부='%s', 가로길이='%s', 세로길이='%s', 우편중량='%s', 전화번호='%s', 높이='%s', 메모='%s' WHERE id = '%s'"
                % (datetime.today().strftime("%Y%m%d"), postline[1], postline[2], postline[3], postline[4], postline[5],
                   postline[6], postline[7], postline[8], postline[9], postline[10], postline[11], postline[12],
                   postline[13],
                   postline[14], postline[15], postline[16], postline[17], rowid))
            # 커밋
            listdb.commit()
            # Table 업데이트
            self.loadpostlist()
            # 이건 메소드만 불러오기
            # QLineEdit 초기화
            self.clearallentry()
            # 초기화함수 실행
            self.initiate_lengf()
            self.initiate_sorttype()
            print('저장')
            # 테이블 업데이트
            self.loadpostlist()
            # 상태바 업데이트
            self.statusBar().showMessage(str(rowid) + '번 라인 수정됨')



    # 전체선택 함수
    def selectallrow(self):
        self.table_postlist.selectAll()
