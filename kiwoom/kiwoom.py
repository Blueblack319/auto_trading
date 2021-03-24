from PyQt5.QAxContainer import *
from PyQt5.QtCore import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print('Kiwoom 클래스 입니다.')

        ###### event loop 모음 ########
        self.login_event_loop = None
        ##############################

        self.get_ocx_instance()
        self.signal_login_commConnect()
        self.event_slots()

    def get_ocx_instance(self):
        # python이 키움증권 Open api에 대한 제어권을 가져올 수 있게 함.
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    # 이벤트를 모아두는 곳
    def event_slots(self):
        # 로그인 이벤트 처리
        self.OnEventConnect(self.login_slot)


    def login_slot(self, errCode):
        print(errCode)

        self.login_event_loop.exit()


    # 로그인 요청
    def signal_login_commConnect(self):
        # 다른 응용프로그램에 데이터 전송 가능케 함.
        self.dynamicCall("CommConnect()")

        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()