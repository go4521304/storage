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


# self.close() 창 닫기
##########################################################################################

class editDialog(QDialog):
    def __init__(self):
        super().__init__()
        
        # ui 로딩
        uic.loadUi("edit.ui", self)
        
        # 창 타이틀 설정
        self.setWindowTitle("재고 추가 / 수정")

        # 바코드가 존재하는지 확인
        self.QLine_bar.textChanged.connect(self.checkBar)

        # 저장 버튼 연결
        self.QBtn_save.clicked.connect(self.save)

        # 덮어쓰기 작업할때 업데이트 할 장소 저장
        self.cell = None

        # 변경내역 확인 할때 사용
        self.itemValue = []

    # 종료시에 불림
    def closeEvent(self, event):
        value = [self.QLine_bar.text(), self.QLine_company.text(), self.QLine_model.text(), self.QLine_price.text(), self.QLine_quantity.text()]
        
        reply = None
        
        if self.cell == None:
            for i in value:
                if i != '':
                    reply = QMessageBox.question(self, '메시지', '아직 저장하지 않은 항목이 있습니다. \n 정말 닫으시겠습니까?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                    break
        
        else:
            for i, v in enumerate(self.itemValue):
                if v != value[i]:
                    reply = QMessageBox.question(self, '메시지', '아직 저장하지 않은 항목이 있습니다. \n 정말 닫으시겠습니까?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                    break
        
        if reply == QMessageBox.No:
            event.ignore()
            return

        # 사용한 것들 초기화 (변수, 텍스트 창)
        self.cell = None
        self.itemValue = []
        self.QLine_bar.setText('')
        self.QLine_company.setText('')
        self.QLine_model.setText('')
        self.QLine_price.setText('')
        self.QLine_quantity.setText('')

        # 종료
        event.accept()
    
    # 바코드가 이미 등록된 바코드인지 확인 - (이미 등록된 바코드면 등록된 정보를 불러옴)
    def checkBar(self):
        barcode = self.QLine_bar.text()
        
        # 8자리 혹은 13자리 숫자가 아니면 스킵 (EAN-13 바코드 규격을 따름)
        if barcode.isdigit() == True and (len(barcode) == 8 or len(barcode) == 13):
            pass
        else:
            self.cell = None
            self.itemValue = []
            return

        # 품목을 찾으면 정보 받아옴
        # 셀을 못찾으면 'CellNotFound(query)라는 오류를 내기에 try-except을 사용
        try:
            # 바코드가 존재하는거면 itemValue에 저장하고 화면에 내용을 띄움
            self.cell = storageSheet.find(barcode)
            self.itemValue = storageSheet.row_values(self.cell.row)
            
            # 화면에 띄우는 코드
            self.QLine_company.setText(self.itemValue[1])
            self.QLine_model.setText(self.itemValue[2])
            self.QLine_price.setText(self.itemValue[3])
            self.QLine_quantity.setText(self.itemValue[4])
            
        
        # 셀을 못찾을경우 self.cell, self.itemValue 값을 초기화
        except:
            self.cell = None
            self.itemValue = []

    def save(self):
        # 화면에 있는 정보들 저장
        value = [self.QLine_bar.text(), self.QLine_company.text(), self.QLine_model.text(), self.QLine_price.text(), self.QLine_quantity.text()]
        
        # 빈칸이 있을 경우 경고문을 띄움
        for i in value:
            if i == '':
                QMessageBox.warning(self, '경고', '비어있는 항목이 있습니다!', QMessageBox.Ok, QMessageBox.Ok)
                return

        # 저장작업
        # cell 이 None이 아니면(== 기존 품목 수정)
        if self.cell != None:
            for i, v in enumerate(self.itemValue):
                # 변경 사항이 있으면
                if v != value[i]:
                    storageSheet.update_cell(self.cell.row, i + 1, value[i])
        
        # 신규 저장
        else:
            storageSheet.append_row(value)

        # cell, itemValue 업데이트
        self.checkBar()

##########################################################################################

class editQuantity(QDialog):
    def __init__(self):
        super().__init__()
        
        # ui 로딩
        uic.loadUi("quantity.ui", self)
        
        # 창 타이틀 설정
        self.setWindowTitle("입/출고")

        # 수량 정보 저장 함수
        self.info = None

        # 버튼 연결
        self.QBtn_save.clicked.connect(self.save)

    def setInfo(self, info):
        self.info = info


    def closeEvent(self, event):
        # 혹시 모른 상황에 대비해서 사용한 변수 초기화
        self.info = None

        ##################################################################################
        # 여기에 시그널을 추가해서 종료시에 최근기록, 검색창을 다시 로딩하는 과정을 추가해야 함 #
        ##################################################################################
        
        event.accept()

    def save(self):
        try:
            # 숫자를 num에 저장
            num = (int)(self.QLine_quantity.text())

            # 출고 일때
            if self.QCombo_type.currentText() == '출고':
                self.info['value'] -= num
                
                # 만약 총 수량이 0보다 작으면 경고문 띄우고 다시 입력대기
                if self.info['value'] < 0:
                    QMessageBox.warning(self, '경고', '충분한 재고가 없습니다!', QMessageBox.Ok, QMessageBox.Ok)
                    self.info['value'] += num
                    return

            # 입고 일때
            elif self.QCombo_type.currentText() == '입고':
                self.info['value'] += num
            
            # '재고' 시트에 수량 업데이트
            storageSheet.update_cell(self.info['row'], self.info['col'], (str)(self.info['value']))
            
            # '기록' 시트에 넣을 값 추가
            # 시간 - 거래종류 - 수량 - 품명
            data = [dt.datetime.now().strftime(r"%Y-%m-%d %H:%M:%S"), self.QCombo_type.currentText(), num, self.info['name']]
            
            # '기록' 시트에 작업 내용 추가
            timeSheet.insert_row(data, 2)
            
            # 성공적으로 끝냈으면 종료 이벤트 호출
            self.close()
        
        except:
            # 숫자가 아니거나 실수, 빈칸 등등의 의도치 않은 상황일 때
            QMessageBox.warning(self, '경고', '제대로 입력해 주십시오!', QMessageBox.Ok, QMessageBox.Ok)
        

##########################################################################################


class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        # 창 이름 설정
        self.setWindowTitle('재고 관리')
        
        # lambda 식을 이용하여 처리
        self.QBtn_toMain.clicked.connect(lambda state, pageNum = 0 : self.setPage(state, pageNum))
        self.QBtn_toSearch.clicked.connect(lambda state, pageNum = 1: self.setPage(state, pageNum))

        #################### 최근 내역 #########################
        self.stkPage.setCurrentIndex(0)

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
        self.QBtn_search.clicked.connect(self.findItem)
        
        # LineEdit 동작 연결
        self.QLine_search.textChanged.connect(self.findBar)
        self.QLine_search.returnPressed.connect(self.findItem)

        self.QList_search.itemDoubleClicked.connect(self.selItem)
        ###################### 추가 / 수정 ######################
        # 추가 / 수정 class 연결
        self.editUi = editDialog()

        # 추가 / 수정 버튼 클릭시 Dialog 열기 (Modaless)
        self.QBtn_warehousing.clicked.connect(self.openEdit)

        ###################### 입 / 출고 ######################
        # 입 / 출고 class 연결
        self.quantityUi = editQuantity()
        
        # 입 / 출고 버튼 클릭시 Dialog 열기 (Modal)
        self.QBtn_quantity.clicked.connect(self.openQuantity)
        
        # 검색이 될 때까지 버튼 비활성화
        self.QBtn_quantity.setEnabled(False)

        # 수량 정보 저장
        self.quantityInfo = None

###################################################################################
    def setPage(self, state, pageNum):
        self.stkPage.setCurrentIndex(pageNum)

#################################### 최근 내역 #####################################
    def load_timeLine(self):
        # 일정 개수를 가져와서 후 처리
        timeLine = timeSheet.get('A1:D10000')

        # dataframe
        timeLine = pd.DataFrame(timeLine, columns = timeLine[0])
        timeLine = timeLine.reindex(timeLine.index.drop(0))
        # 바로 위에 과정을 통해서 인덱스의 시작은 1번이지만
        # 아래 for문에서는 위에서 '몇번째 인덱스' 순으로 사용을 해서 문제가 없음

        totalNum = 0    
        self.QTable_time.setRowCount(len(timeLine.index))
        for i, _ in enumerate(timeLine.index):
            if parse(timeLine.iloc[i, 0]) > timeToday :
                totalNum += int(timeLine .iloc[i]['수량'])
                for j, _ in enumerate(timeLine.columns):
                    self.QTable_time.setItem(i, j, QTableWidgetItem(str(timeLine.iloc[i,j])))
                    # 가운데 정렬
                    self.QTable_time.item(i, j).setTextAlignment(Qt.AlignCenter)
            else:
                self.QTable_time.setRowCount(i)
                self.QLabel_total.setText('합계 ' + str(i))
                self.QLabel_totalNum.setText('총 수량 ' + str(totalNum))
                break

        # 칸 넓이 데이터에 맞게 수정
        self.QTable_time.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.QTable_time.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.QTable_time.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.QTable_time.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)


    #################################### 검색 #####################################
    def lineCls(self):
        self.QLine_search.clear()

    def listCls(self):
        self.QList_search.clear()

    def cantFind(self):
        self.quantityInfo = None
        self.QBtn_quantity.setEnabled(False)
        self.QList_search.addItem("검색 결과가 없습니다.")


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
                if head == '품명':
                    self.quantityInfo = {'name': itemValue[i]}
                elif head == '수량':
                    self.quantityInfo['row'] = cell.row
                    self.quantityInfo['col'] = i + 1
                    self.quantityInfo['value'] = int(itemValue[i])

                # 단가와 현재 수량인경우 정규 표현식을 이용하여 천의 자리마다 끊어줌
                if head == '단가' or head == '수량':
                    itemValue[i] = re.sub(r'\B(?=(\d{3})+(?!\d))', ',', itemValue[i])

                self.QList_search.addItem(head + ": " + itemValue[i])
                self.QBtn_quantity.setEnabled(True)
        
        # 셀을 못찾을경우 출력
        except:
            self.cantFind()


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
            self.cantFind()
            return

        # 졍규 표현식을 사용 re.I (대소문자 상관 X) 옵션을 줌
        searchCriteria = re.compile(self.QLine_search.text(), re.I)
        cellList = storageSheet.findall(searchCriteria)

        # 검색된것이 없으면 그냥 종료
        if len(cellList) == 0:
            self.cantFind()
            return

        # 회사 혹은 품명이 일치할시 항목을 보여줌
        for i in cellList:
            if i.col == 2 or i.col == 3:
                self.QList_search.addItem("(" + storageSheet.cell(i.row, 2).value + ") " + storageSheet.cell(i.row, 3).value)
        
        # 보여준 항목이 없을시에 '검색결과 없음' 출력
        if (self.QList_search.count() == 0):
            self.cantFind()
        

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

    #################################### 추가 / 수정 #####################################
    def openEdit(self) :
        self.editUi.show()

    #################################### 입 / 출고 #####################################
    def openQuantity(self):
        if self.quantityInfo == None:
            return
        
        self.quantityUi.setInfo(self.quantityInfo)
        self.quantityUi.exec_()
        
    



if __name__ == "__main__" :
    # QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv) 

    # WindowClass의 인스턴스 생성
    myWindow = WindowClass() 

    # 프로그램 화면을 보여주는 코드
    myWindow.show()

    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()





