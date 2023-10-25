# Pyside6 로드
from PySide6.QtWidgets import *
# ui 로드용
from ui_loader import load_ui
# SQLite 로드
import sqlite3
# Pandas 로드
import pandas as pd

class CheckGroupPost(QMainWindow):
    def __init__(self):
        super().__init__()
        #UI 로드
        load_ui('CheckGroupPost.ui', self)
        # 창 닫기 버튼 시그널
        self.action_closewindow.triggered.connect(self.closewindow)
        # 전체 선택 버튼 시그널
        self.btn_selectallrow.clicked.connect(self.selectallrow)
        # 창 보이기
        self.show()

        ################ SQLite DB 로드해서 DF로 옮기기############
        # SQLite DB 연결
        groupdb = sqlite3.connect('grouppostlist.db')
        # SQL 쿼리 실행
        query = "SELECT * FROM groupedlist"
        try:
            df = pd.read_sql_query(query, groupdb)
        except pd.errors.DatabaseError:
            self.statusBar().showMessage('DB에 레코드가 존재하지 않습니다.')
            # DB 연결 종료
            groupdb.close()
        else:
            # DB 연결 종료
            groupdb.close()
            print('[DF]\n', df)

            # QTableWidget 에 DF 내용 쓰기
            self.table_grouppostlist.clear()  # 부르기 전 초기화
            col = len(df.keys()) #DF의 키 길이를 칼럼수로 저장
            self.table_grouppostlist.setColumnCount(col)
            self.table_grouppostlist.setHorizontalHeaderLabels(df.keys())
            row = len(df.index) #DF의 리코드 길이를 행수로 저장
            self.table_grouppostlist.setRowCount(row)
            for r in range(row):
                for c in range(col):
                    item = QTableWidgetItem(str(df.iloc[r][c]))
                    self.table_grouppostlist.setItem(r, c, item)
            self.table_grouppostlist.resizeColumnsToContents()


    #QTableWidget 모든 행 선택
    def selectallrow(self):
        self.table_grouppostlist.selectAll()

    # 윈도우 닫기
    def closewindow(self): #윈도우 닫기
        self.close()