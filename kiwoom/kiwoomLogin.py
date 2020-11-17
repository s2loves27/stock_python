import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QAxContainer import *


class MyWindow(QMainWindow):  # MyWindow 클래스 QMainWindow 클래스를 상속 받아 생성됨
    def __init__(self):  # MyWindow 클래스의 초기화 함수(생성자)
        super().__init__()  # 부모클래스 QMainWindow 클래스의 초기화 함수(생성자)를 호출
        self.setWindowTitle("devshj")  # 윈도우 창 타이틀을 devshj로 변경
        self.setGeometry(300, 300, 300, 150)  # 윈도우 창 위치 및 크기를 지정 (x축 시작점, y축 시작점, 넓이, 높이)

        self.kiwoom = QAxWidget(
            "KHOPENAPI.KHOpenAPICtrl.1")  # 키움증권 Open API+의 ProgID를 사용하여 생성된 QAxWidget을 kiwoom 변수에 할당

        btn1 = QPushButton("Login", self)  # Login 텍스트를 포함한 버튼을 생성하여 btn1 변수에 할당
        btn1.move(20, 20)  # 해당 버튼의 위치를 지정 (x축 시작점, y축 시작점)
        btn1.clicked.connect(self.btn1_clicked)  # 해당 버튼의 클릭 이벤트를 btn1_clicked 함수와 연결함

        btn2 = QPushButton("Check state", self)  # Check state 텍스트를 포함한 버튼을 생성하여 btn2 변수에 할당
        btn2.move(20, 70)  # 해당 버튼의 위치를 지정 (x축 시작점, y축 시작점)
        btn2.clicked.connect(self.btn2_clicked)  # 해당 버튼의 클릭 이벤트를 btn2_clicked 함수와 연결함

        self.kiwoom.OnEventConnect.connect(self.event_connect)  # 통신 연결 상태가 바뀔 때 이벤트를 event_connect 함수와 연결함

    def btn1_clicked(self):  # Login 버튼 클릭 시 실행되는 함수
        ret = self.kiwoom.dynamicCall("CommConnect()")  # 키움 로그인 윈도우를 실행

    def btn2_clicked(self):  # Check state 버튼 클릭 시 실행되는 함수
        # GetConnectState() > 0:미연결, 1:연결완료
        if self.kiwoom.dynamicCall("GetConnectState()") == 0:  # 로그인 상태값이 0 이라면
            self.statusBar().showMessage("Not connected")  # PyQt5 statusBar에 Not connected 메시지를 노출함
        else:  # 로그인 상태값이 0 이 아니라면
            self.statusBar().showMessage("Connected")  # PyQt5 statusBar에 Connected 메시지를 노출함

    def event_connect(self, err_code):  # 통신 연결 상태가 바뀔 때 실행되는 함수
        if err_code == 0:  # 로그인이 성공하였다면
            self.statusBar().showMessage("Connected")  # PyQt5 statusBar에 Connected 메시지를 노출함
        else:
            self.statusBar().showMessage("Not connected")  # PyQt5 statusBar에 Not connected 메시지를 노출함

# py 파일 실행시 제일 먼저 동작
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()  # MyWindow 클래스를 생성하여 myWondow 변수에 할당
    myWindow.show()  # MyWindow 클래스를 노출
    app.exec_()  # 메인 이벤트 루프에 진입 후 프로그램이 종료될 때까지 무한 루프 상태 대기