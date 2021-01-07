# QT
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5 import uic

# gspread
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# module
import pandas as pd
import datetime as dt
from dateutil.parser import parse
import re

scope = [
'https://spreadsheets.google.com/feeds',
'https://www.googleapis.com/auth/drive',
]
json_file_name = 'storage-299609-970c24909d3f.json'
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
gc = gspread.authorize(credentials)
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1nIPCJtoxa4_up-aV4DB-0KPOqmoyCCPfXi0GieZgwPo/edit#gid=1615035515'

##############################################

# 스프레스시트 문서 가져오기 
doc = gc.open_by_url(spreadsheet_url)

# 시트 선택하기
timeSheet = doc.worksheet('기록')
storageSheet = doc.worksheet('재고')

# UI파일 연결
form_class = uic.loadUiType("storage.ui")[0]

##############################################
# 전역변수
# timeToday
##############################################
# # 시간 측정용 임시
# timeTemp = pd.Timestamp.now()

# # 시간측정후 표시
# timeTemp = pd.Timestamp.now() - timeTemp
# print(timeTemp)
##############################################
# *.exec_() == 모달 (종속) 방식
# *.show() == 모달리스 (독립) 방식
##############################################
# loc == 인덱스로 찾기
# iloc == 행번호로 찾기
##############################################

# 리프레시 시간
refreshTime = 30000


class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        # 창 이름 설정
        self.setWindowTitle('재고 관리')
        
        # lambda 식을 이용하여 처리
        self.toMain.clicked.connect(lambda state, pageNum = 0 : self.setPage(state, pageNum))
        self.toSearch.clicked.connect(lambda state, pageNum = 1: self.setPage(state, pageNum))

        #################### 최근 내역 #########################
        self.stkPage.setCurrentIndex(0)

        # 로딩버튼이랑 연결(임시 테스트용)
        self.QBtn_road.clicked.connect(self.load_timeLine)

        # 오늘 날짜 저장 (0시를 기점으로)
        global timeToday
        timeToday = pd.Timestamp.now()
        timeToday = timeToday.replace(hour = 0, minute = 0, second = 0, microsecond = 0, nanosecond = 0)

        # 타임라인 로딩
        self.load_timeLine()

        # 타임라인 업데이트 설정
        self.reFresh = QTimer(self)
        self.reFresh.start(refreshTime)
        self.reFresh.timeout.connect(self.load_timeLine)
        
        ###################### 검색 ######################
        # Btn 연결
        self.QBtn_searchCls.clicked.connect(self.lineCls)
        
        # LineEdit 동작 연결
        self.QLine_search.textChanged.connect(self.findBar)
        self.QLine_search.returnPressed.connect(self.findItem)

        self.QList_search.itemDoubleClicked.connect(self.selItem)

    ################################################################
    def setPage(self, state, pageNum):
        self.stkPage.setCurrentIndex(pageNum)

    ####################### 최근 내역 ###############################
    def load_timeLine(self):
        # 일정 개수를 가져와서 후 처리
        timeLine = timeSheet.get('A1:C10000')

        # dataframe
        timeLine = pd.DataFrame(timeLine, columns = timeLine[0])
        timeLine = timeLine.reindex(timeLine.index.drop(0))
        # 바로 위에 과정을 통해서 인덱스의 시작은 1번이지만
        # 아래 for문에서는 위에서 '몇번째 인덱스' 순으로 사용을 해서 문제가 없음
    
        self.QTable_time.setRowCount(len(timeLine.index))
        for i, _ in enumerate(timeLine.index):
            if parse(timeLine.iloc[i, 0]) > timeToday :
                for j, _ in enumerate(timeLine.columns):
                    self.QTable_time.setItem(i, j, QTableWidgetItem(str(timeLine.iloc[i,j])))
                    # 가운데 정렬
                    self.QTable_time.item(i, j).setTextAlignment(Qt.AlignCenter)
            else:
                self.QTable_time.setRowCount(i)
                break

        # 칸 넓이 데이터에 맞게 수정
        self.QTable_time.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)


    #################################### 검색 #####################################
    def lineCls(self):
        self.QLine_search.clear()

    def listCls(self):
        self.QList_search.clear()


    def showItem(self, search):
        self.listCls()
        # 품목을 찾으면 정보 출력
        # 셀을 못찾으면 'CellNotFound(query)라는 오류를 내기에 try-except을 사용
        try:
            cell = storageSheet.find(search)

            # 항목을 찾았으면 항목의 정보와 1열의 그 항목이 어떤것인지 가져와서 출력
            # 차후 스프레드시트 내의 항목이 늘어나더라도 변경없이 사용가능
            itemHead = storageSheet.row_values(1)
            itemValue = storageSheet.row_values(cell.row)

            # 항목을 리스트 내에 출력
            for i, head in enumerate(itemHead):
                self.QList_search.addItem(head + ": " + itemValue[i])

            self.lineCls()
        
        # 셀을 못찾을경우 출력
        except:
            self.QList_search.addItem("검색 결과가 없습니다.")


    def findBar(self):
        barcode = self.QLine_search.text()
        
        # 8자리 혹은 13자리 숫자가 아니면 스킵 (EAN-13 바코드 규격을 따름)
        if barcode.isdigit() == True and (len(barcode) == 8 or len(barcode) == 13):
            pass
        else:
            return
        
        # 바코드 검색 후 결과 출력
        self.showItem(barcode)
        

    # 검색하고 목록을 띄움
    def findItem(self):
        self.listCls()
        
        # 검색창에 입력이 안되있을시 종료
        if len(self.QLine_search.text()) == 0:
            return

        searchCriteria = re.compile(self.QLine_search.text(), re.I)
        cellList = storageSheet.findall(searchCriteria)

        # 검색된것이 없으면 그냥 종료
        if len(cellList) == 0:
            self.QList_search.addItem("검색 결과가 없습니다.")
            return

        # 회사 혹은 품명이 일치할시 항목을 보여줌
        for i in cellList:
            if i.col == 2 or i.col == 3:
                self.QList_search.addItem("(" + storageSheet.cell(i.row, 2).value + ") " + storageSheet.cell(i.row, 3).value)

    # 검색후 더블 클릭해서 물건의 상세정보를 띄움
    def selItem(self):
        try:
            # 변수명이 너무 길어서 저장해서 사용
            itemName = self.QList_search.selectedItems()[0].text()
    
            # 제품 이름만 분리해서 저장
            namePos = re.search("\)\s", itemName).end()
            itemName = itemName[namePos:]
            
            # 상세 정보 출력
            self.showItem(itemName)
        
        except:
            pass


if __name__ == "__main__" :
    # QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv) 

    # WindowClass의 인스턴스 생성
    myWindow = WindowClass() 

    # 프로그램 화면을 보여주는 코드
    myWindow.show()

    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()





