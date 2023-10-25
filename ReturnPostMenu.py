# Pyside6 로드
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from ui_loader import load_ui
# Pandas 로드
import pandas as pd
# 시간 관련 라이브러리
from datetime import datetime
# SQLite 3 로드
import sqlite3
# 파일 관리자
import os


class ReturnPostMenu(QMainWindow):

    def __init__(self, input): #res 는 객체 생성시 받은 Search ID값
        super().__init__()

        #SQLite 로드
        if os.path.isfile("returnpostlist.db"):  # DB파일이 있을 경우
            self.returndb = sqlite3.connect("returnpostlist.db")
            self.returncursor = self.returndb.cursor()
        else:  # DB파일이 없을 경우
            self.statusBar().showMessage('반려 DB가 없으므로 DB파일을 생성합니다.')
            self.returndb = sqlite3.connect("returnpostlist.db")
            self.returncursor = self.returndb.cursor()
            self.returncursor.execute(
                "CREATE TABLE postlist(id integer, 접수날짜 text, 사원번호 text, 접수자명 text, 부서명 text, 보내는사람 text, 받는사람 text, 주소 text, 우편번호 text, 제목 text, 수량 integer, 종류 text, 긴급여부 boolean, 가로길이 real, 세로길이 real, 우편중량 real, 전화번호 text, 높이 real, 메모 text, 할인타입 text, 할인타입그룹 integer, 결재여부 , 반려여부 boolean, 비용 text)")
            self.returncursor.commit()
        #SQLite postlist 로드
        if os.path.isfile("postlist.db"):  # DB파일이 있을 경우
            self.listdb = sqlite3.connect("postlist.db")
            self.cursor = self.listdb.cursor()
        else:  # DB파일이 없을 경우
            self.statusBar().showMessage('DB가 없으므로 DB파일을 생성합니다.')
            self.listdb = sqlite3.connect("postlist.db")
            self.cursor = self.listdb.cursor()
            self.cursor.execute(
                "CREATE TABLE postlist(id integer PRIMARY KEY AUTOINCREMENT, 접수날짜 text, 사원번호 text, 접수자명 text, 부서명 text, 보내는사람 text, 받는사람 text, 주소 text, 우편번호 text, 제목 text, 수량 integer, 종류 text, 긴급여부 boolean, 가로길이 real, 세로길이 real, 우편중량 real, 전화번호 text, 높이 real, 메모 text, 할인타입 text, 할인타입그룹 integer, 결재여부 boolean, 반려여부 boolean, 비용 text)")
            self.listdb.commit()

        # UI 로드
        load_ui('ReturnPostMenu.ui', self)
        # 윈도우 보이기
        self.show()

        # 버튼시그널
        self.btn_delpostcolumn.clicked.connect(self.deleterowselected)

        # 창 닫기
        self.action_closewindow.triggered.connect(self.closewindow)

        # 창 실행하자마자 사원번호 id를 input 으로 받아 반려우편물 조회
        self.loadpostlist(input)

        # [QTableWidget 레코드 선택 시]
        self.table_postlist.selectionModel().selectionChanged.connect(self.editrow_postinfo)

        # 수정 및 재접수 버튼 시그널
        self.btn_reuploadpost.clicked.connect(self.fixpostandresubmit)

        #첫 실행 시 콤보박스 초기화
        self.initiate_lengf()
        self.initiate_sorttype()

    #메인윈도우 X 눌러 종료 시 DB 정리
    def closeEvent(self, event):
        self.returndb.close()
        self.listdb.close()
        self.close()

    #반려우편물 조회
    def loadpostlist(self, input):
        # 반려우편울 조회조건(사원번호) 저장
        self.saveid = input
        #쿼리 완성
        query = 'SELECT id, 접수날짜, 사원번호, 접수자명, 부서명, 보내는사람, 받는사람, 주소, 우편번호, 제목, 수량, 종류, 긴급여부, 가로길이, 세로길이, 우편중량, 전화번호, 높이, 메모 FROM postlist WHERE 사원번호 = "%s"' % (self.saveid)
        print('쿼리 : ',query)
        df = pd.read_sql(query, self.returndb)
        # 메모 칼럼을 반려사유 칼럼으로 임시 변경
        df.rename(columns={'메모': '반려사유'})
        #df 정렬
        df = df.sort_values(by=['id'])
        #불러진 데이터 보기
        print('[데이터프레임]\n', df)

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

            # 상태창 업데이트
            self.statusBar().showMessage('추가완료')
            # return
            return postline

    def getrowid(self, tablename):
        try:
            idtofix = ''
            for row in self.table_postlist.selectionModel().selectedRows():
                idtofix = self.table_postlist.item(row.row(), 0).text()
        except UnboundLocalError as e:
            self.statusBar().showMessage('선택된 레코드가 없습니다.')

        return idtofix

        #선택된 줄에 대한 ID 데이터를 읽는 함수
    def getselectedrowid(self, tablename):
        #인텍스에 선택된 롤 집어넣기
        index = tablename.selectionModel().selectedRows()
        #리스트 초기화
        templist = list()
        for row in index: #index를 정렬하여 row 로 읽기
            #row 정보 중 0번칼럼 읽어서 텍스트로 출력
            id = tablename.item(row.row(), 0).text()
            #선택된 모든 row 읽어서 리스트에 추가
            templist.append(id)
        return templist


    #선택된 줄에 대한 모든 데이터를 읽는 함수
    def getselectedrowid_adv(self, tablename):
        #인텍스에 선택된 롤 집어넣기
        index = tablename.selectionModel().selectedRows()
        #리스트 초기화
        templist = list()
        for row in index: #index를 정렬하여 row 로 읽기
            #읽힌 모든 거 다 append 하기
            for i in range(tablename.columnCount()):
                tempvar = tablename.item(row.row(), i).text()
                templist.append(tempvar)
        return templist

    def deleterowselected(self):
        indexlist = self.getselectedrowid(self.table_postlist)
        # SQL DELETE 문 실행
        self.returncursor.execute("DELETE FROM postlist WHERE id IN (%s)"%",".join(indexlist))
        #커밋
        self.returndb.commit()
        # 엔트리 초기화
        self.clearallentry()
        # 테이블 업데이트
        self.loadpostlist(self.saveid)

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

    def editrow_postinfo(self):
        post_tofix = self.getselectedrowid_adv(self.table_postlist)
        print('디버그 :', post_tofix)
        #칼럼 선택시 불러진 데이터값이 0 이 아닐 경우 엔트리 업데이트하기
        if(len(post_tofix) > 0):
            self.info_0_importerid.setText(post_tofix[2])
            self.info_1_importer.setText(post_tofix[3])
            self.info_2_depart.setText(post_tofix[4])
            self.info_3_sender.setText(post_tofix[5])
            self.info_4_receiver.setText(post_tofix[6])
            self.info_5_address.setText(post_tofix[7])
            self.info_14_postno.setText(post_tofix[8])
            self.info_6_title.setText(post_tofix[9])
            self.info_7_spin_quantity.setValue(int(post_tofix[10]))

            # combobox 는 setText 가 안되니 작업 추가
            index = self.info_8_combo_sort.findText(post_tofix[11], Qt.MatchFixedString)
            self.info_8_combo_sort.setCurrentIndex(index)
            # 긴급여부가 False이면 False, True이면 True
            if post_tofix[12] == 'False':
                tempres = False
            else:
                tempres = True
            print('AAAA', post_tofix[12], tempres)
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

    #수정 및 재접수 버튼 누를 시 returnpost의 DB 수정 후 삭제한 다음 postlist.db 에 신규추가
    def fixpostandresubmit(self):

        if os.path.isfile("postlist.db"):  # DB파일이 있을 경우
            listdb = sqlite3.connect("postlist.db")
            cursor = listdb.cursor()
        else:  # DB파일이 없을 경우
            self.statusBar().showMessage('DB가 없으므로 DB파일을 생성합니다.')
            listdb = sqlite3.connect("postlist.db")
            cursor = listdb.cursor()
            cursor.execute(
                "CREATE TABLE postlist(id integer PRIMARY KEY AUTOINCREMENT, 접수날짜 text, 사원번호 text, 접수자명 text, 부서명 text, 보내는사람 text, 받는사람 text, 주소 text, 우편번호 text, 제목 text, 수량 integer, 종류 text, 긴급여부 boolean, 가로길이 real, 세로길이 real, 우편중량 real, 전화번호 text, 높이 real, 메모 text, 할인타입 text, 할인타입그룹 integer, 결재여부 , 반려여부 boolean, 비용 text)")
            listdb.commit()


        postline = self.getentryinfo()
        # 선택한 라인 저장
        rowid = self.getrowid(self.table_postlist)
        # 라인 추가
        print('postline :', postline)
        print('보내는 우편물 ID :', rowid)
        # 디버그 긴급표시 boolean > int 로
        if postline[12] == 'True':
            postline[12] = True
        else:
            postline[12] = False
        # 반려우편물의 업데이트
        self.returncursor.execute(
            "UPDATE postlist SET 접수날짜='%s', 사원번호='%s', 접수자명='%s', 부서명='%s', 보내는사람='%s', 받는사람='%s', 주소='%s', 우편번호='%s', 제목='%s', 수량='%s', 종류='%s', 긴급여부='%s', 가로길이='%s', 세로길이='%s', 우편중량='%s', 전화번호='%s', 높이='%s', 메모='%s' WHERE id = '%s'"
            % (datetime.today().strftime("%Y%m%d"), postline[1], postline[2], postline[3], postline[4], postline[5],
               postline[6], postline[7], postline[8], postline[9], postline[10], postline[11], postline[12], postline[13],
               postline[14], postline[15], postline[16], postline[17], rowid))
        # 반려우편물에서 우편물 id로 우편물 제거
        self.returncursor.execute("DELETE FROM postlist WHERE id IN (%s)" % rowid)
        self.returndb.commit()

        # 접수우편물 리스트에 우편물 추가
        print('추가할 우편데이터 : ', str(postline))
        postline[17] = '수정됨'
        # 라인 추가
        cursor.execute(
            "INSERT INTO postlist(접수날짜, 사원번호, 접수자명, 부서명, 보내는사람, 받는사람, 주소, 우편번호, 제목, 수량, 종류, 긴급여부, 가로길이, 세로길이, 우편중량, 전화번호, 높이, 메모) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            postline)
        # 커밋 및 정리
        listdb.commit()

        #Table 업데이트 시 조회 조건 콘솔에 표시
        print('조회 조건 :', self.saveid)
        # Table 업데이트
        self.loadpostlist(self.saveid)
        # 이건 메소드만 불러오기
        # QLineEdit 초기화
        self.clearallentry()
        # 초기화함수 실행
        self.initiate_lengf()
        self.initiate_sorttype()
        print('저장')
        # 테이블 업데이트
        self.loadpostlist(self.saveid)
        # 상태바 업데이트
        self.statusBar().showMessage(str(rowid) + '번 라인 수정됨')


    def closewindow(self): #윈도우 닫기
        self.returndb.close()
        self.listdb.close()
        self.close()

