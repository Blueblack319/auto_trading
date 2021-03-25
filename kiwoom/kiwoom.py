from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print('Kiwoom 클래스 입니다.')

        ###### event loop 모음 ########
        self.login_event_loop = None
        ##############################

        ########### 변수 모음 ##########
        self.account_num = None
        ##############################

        self.get_ocx_instance()
        self.event_slots()

        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info()
        self.get_basic_info()



    def get_ocx_instance(self):
        # python이 키움증권 Open api에 대한 제어권을 가져올 수 있게 함.
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    # 이벤트를 모아두는 곳
    def event_slots(self):
        # 로그인 이벤트 처리
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.tr_data_slot)

    # 로그인 요청
    def signal_login_commConnect(self):
        # 다른 응용프로그램에 데이터 전송 가능케 함.
        self.dynamicCall("CommConnect()")

        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()


    def login_slot(self, errCode):
        print(errors(errCode))

        self.login_event_loop.exit()

    def get_account_info(self):
        account_list = self.dynamicCall('GetLoginInfo(String)', 'ACCNO')
        self.account_num = account_list.split(';')[0]
        print(f'나의 보유계좌번호 {self.account_num}')

    def detail_account_info(self):
        print('예수금 요청하는 부분')

        self.dynamicCall('SetInputValue(String, String)', '계좌번호', '8162873111')
        self.dynamicCall('SetInputValue(String, String)', '비밀번호', '0000')
        self.dynamicCall('SetInputValue(String, String)', '비밀번호입력매체구분', '00')
        self.dynamicCall('SetInputValue(String, String)', '조회구분', '2')
        self.dynamicCall('CommRqData(String, String, int, String)', '예수금상세현황요청', 'opw00001', '0', '2000')

    def get_basic_info(self):
        print('종목의 기본정보를 요청하는 부분')

        self.dynamicCall('SetInputValue(String, String)', '종목코드', '015760')
        self.dynamicCall('CommRqData(String, String, int, String)', '주식기본정보요청', 'opt10001', '0', '3000')

    def tr_data_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        TR 데이터를 받아오는 곳, 요청받는 곳
        :param sScrNo: 스크린번호
        :param sRQName: 내가 요청했을 때 지은 이름(사용자 구분명)
        :param sTrCode: TR코드(요청id)
        :param sRecordName: 사용안함
        :param sPrevNext: 다음 페이지가 있는지
        :return:
        '''

        if sRQName == '예수금상세현황요청':
            deposit = self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '예수금')
            print(f'예수금: {deposit}')
            usable_money = self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '주문가능금액')
            print(f'주문가능금액: {usable_money}')

        if sRQName == '주식기본정보요청':
            name = self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '종목명').strip()
            face_value = self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '액면가').strip()
            capital = self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '자본금').strip()
            market_capitalation = self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '시가총액').strip()
            roe = self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, 'ROE').strip()
            revenue = self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '매출액').strip()
            operating_income = self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '영업이익').strip()
            net_income = self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '당기순이익').strip()
            stock_price = self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '현재가').strip()
            dist_ratio = self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '유통비율').strip()


            print(f'종목명: {name}')
            print(f'액면가: {face_value}')
            print(f'자본금: {capital}')
            print(f'시가총액: {market_capitalation}')
            print(f'ROE: {roe}')
            print(f'매출액: {revenue}')
            print(f'영업이익: {operating_income}')
            print(f'당기순이익: {net_income}')
            print(f'현재가: {stock_price}')
            print(f'유통비율: {dist_ratio}')


