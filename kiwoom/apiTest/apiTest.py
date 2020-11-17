import sys

from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtWidgets import *
from PyQt5 import uic

form_class = uic.loadUiType("apiTest.ui")[0]  # ui 파일을 로드하여 form_class 생성


class MyWindow(QMainWindow, form_class):  # MyWindow 클래스 QMainWindow, form_class 클래스를 상속 받아 생성됨
    def __init__(self):  # MyWindow 클래스의 초기화 함수(생성자)
        super().__init__()  # 부모클래스 QMainWindow 클래스의 초기화 함수(생성자)를 호출
        self.setupUi(self)  # ui 파일 화면 출력

        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")  # 키움증권 Open API+의 ProgID를 사용하여 생성된 QAxWidget을 kiwoom 변수에 할당

        self.lineEditCode.setDisabled(True)  # 종목코드 입력란을 비활성화 상태로 변경
        self.btnSearch.setDisabled(True)  # 조회 버튼을 비활성화 상태로 변경
        self.pteLog.setDisabled(True)  # plainTextEdit를 비활성화 상태로 변경

        self.btnLogin.clicked.connect(
            self.btn_login)  # ui 파일을 생성할때 작성한 로그인 버튼의 objectName 으로 클릭 이벤트가 발생할 경우 btn_login 함수를 호출
        self.kiwoom.OnEventConnect.connect(self.event_connect)  # 키움 서버 접속 관련 이벤트가 발생할 경우 event_connect 함수 호출

    def btn_login(self): # Login 버튼 클릭 시 실행되는 함수
        ret = self.kiwoom.dynamicCall("CommConnect()")  # 키움 로그인 윈도우를 실행

    def event_connect(self, err_code):  # 키움 서버 접속 관련 이벤트가 발생할 경우 실행되는 함수
        if err_code == 0:   # err_code가 0이면 로그인 성공 그외 실패
            self.lineEditCode.setDisabled(False)    # 종목코드 입력란을 활성화 상태로 변경
            self.btnSearch.setDisabled(False)   # 조회 버튼을 활성화 상태로 변경
            self.pteLog.setDisabled(False)  # plainTextEdit를 활성화 상태로 변경
            self.pteLog.appendPlainText("로그인 성공")    # ui 파일을 생성할때 작성한 plainTextEdit의 objectName 으로 해당 plainTextEdit에 텍스트를 추가함

            account_num = self.kiwoom.dynamicCall("GetLoginInfo(QString)",["ACCNO"])  # 키움 dynamicCall 함수를 통해 GetLoginInfo 함수를 호출하여 계좌번호를 가져옴
            self.pteLog.appendPlainText("계좌번호: " + account_num.rstrip(';'))  # 키움은 전체 계좌를 반환하며 각 계좌 번호 끝에 세미콜론(;)이 붙어 있음으로 제거하여 plainTextEdit에 텍스트를 추가함

        else:
            self.lineEditCode.setDisabled(True)  # 종목코드 입력란을 비활성화 상태로 변경
            self.btnSearch.setDisabled(True)  # 조회 버튼을 비활성화 상태로 변경
            self.pteLog.setDisabled(True)  # plainTextEdit를 비활성화 상태로 변경
            self.pteLog.appendPlainText("로그인 실패")    # ui 파일을 생성할때 작성한 plainTextEdit의 objectName 으로 해당

        self.btnSearch.clicked.connect(self.btn_search)  # ui 파일을 생성할때 작성한 조회 버튼의 objectName 으로 클릭 이벤트가 발생할 경우 btn_search 함수를 호출
        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata)  # 키움 데이터 수신 관련 이벤트가 발생할 경우 receive_trdata 함수 호출

    def btn_search(self):  # 조회 버튼 클릭 시 실행되는 함수
        code = self.lineEditCode.text()  # ui 파일을 생성할때 작성한 종목코드 입력란의 objectName 으로 사용자가 입력한 종목코드의 텍스트를 가져옴
        self.pteLog.appendPlainText(
            "종목코드: " + code)  # ui 파일을 생성할때 작성한 plainTextEdit의 objectName 으로 해당 plainTextEdit에 텍스트를 추가함
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드",
                                code)  # 키움 dynamicCall 함수를 통해 SetInputValue 함수를 호출하여 종목코드를 셋팅함
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10001_req", "opt10001", 0,
                                "0101")  # 키움 dynamicCall 함수를 통해 CommRqData 함수를 호출하여 opt10001 API를 구분명 opt10001_req, 화면번호 0101으로 호출함

    def receive_trdata(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1,
                       msg2):  # 키움 데이터 수신 함수
        if rqname == "opt10001_req":  # 수신된 데이터 구분명이 opt10001_req 일 경우
            name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname,
                                           0, "종목명")  # 구분명 opt10001_req 의 종목명을 가져와서 name에 셋팅
            volume = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname,
                                             0, "거래량")  # 구분명 opt10001_req 의 거래량을 가져와서 volume에 셋팅

            self.pteLog.appendPlainText("종목명: " + name.strip())  # 종목명을 공백 제거해서 plainTextEdit에 텍스트를 추가함
            self.pteLog.appendPlainText("거래량: " + volume.strip())  # 거래량을 공백 제거해서 plainTextEdit에 텍스트를 추가함

# py 파일 실행시 제일 먼저 동작
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()  # MyWindow 클래스를 생성하여 myWondow 변수에 할당
    myWindow.show()  # MyWindow 클래스를 노출
    app.exec_()  # 메인 이벤트 루프에 진입 후 프로그램이 종료될 때까지 무한 루프 상태 대기