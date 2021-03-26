from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from PyQt5.QtTest import *
import os

from config.errorCode import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print('Kiwoom 클래스 입니다.')

        ###### event loop 모음 ########
        self.login_event_loop = None
        self.detail_account_info_event_loop = QEventLoop()
        self.calculator_event_loop = QEventLoop()
        self.basic_info_event_loop = None
        ##############################

        ###### 스크린번호 모음 ######
        self.screen_account_info = "2000"
        self.screen_stock_info = '4000'
        self.screen_realtime_stock = '5000' # 종목별로 할당할 스크린번호
        self.screen_trading_stock = '6000' # 매매용 스크린번호
        ##########################

        ########### 변수 모음 ##########
        self.account_num = None
        self.account_stock_dict = {}
        self.not_concluded_stock_dict = {}
        self.calcul_data = []
        self.portfolio_stock_dict = {}
        ##############################

        ###### 계좌관련 변수 모음 ######
        self.usable_money = 0
        self.usable_money_rate = 0.5
        #############################

        self.get_ocx_instance()
        self.event_slots()

        self.signal_login_commConnect() # 로그인 요청
        self.get_account_info() # 계좌번호 가져오기
        self.detail_account_info() # 예수금 가져오기
        self.get_basic_info() # 종목기본정보 가져오기
        self.account_eval_bal() # 계좌평가장고내역 가져오기
        self.not_concluded_stock() # 미체결요청 가져오기
        # self.calculate_stock() # 종목분석 실행하기
        
        self.read_file() # 분석된 파일 읽기
        self.set_screen_no() # 실시간용 스크린번호, 매매용 스크린번호 할당



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
        print('======예수금 요청하는 부분======')

        self.dynamicCall('SetInputValue(String, String)', '계좌번호', self.account_num)
        self.dynamicCall('SetInputValue(String, String)', '비밀번호', '0000')
        self.dynamicCall('SetInputValue(String, String)', '비밀번호입력매체구분', '00')
        self.dynamicCall('SetInputValue(String, String)', '조회구분', '2')
        self.dynamicCall('CommRqData(String, String, int, String)', '예수금상세현황요청', 'opw00001', '0', self.screen_account_info)

        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec_()

    def get_basic_info(self):
        print('======종목의 기본정보를 요청하는 부분======')

        self.dynamicCall('SetInputValue(String, String)', '종목코드', '015760')
        self.dynamicCall('CommRqData(String, String, int, String)', '주식기본정보요청', 'opt10001', '0', '3000')

        self.basic_info_event_loop = QEventLoop()
        self.basic_info_event_loop.exec_()

    def account_eval_bal(self, sPrevNext='0'):
        print('======계좌평가잔고내역 요청하는 부분======')

        self.dynamicCall('SetInputValue(String, String)', '계좌번호', self.account_num)
        self.dynamicCall('SetInputValue(String, String)', '비밀번호', '0000')
        self.dynamicCall('SetInputValue(String, String)', '비밀번호입력매체구분', '00')
        self.dynamicCall('SetInputValue(String, String)', '조회구분', '2')
        self.dynamicCall('CommRqData(String, String, int, String)', '계좌평가잔고내역요청', 'opw00018', sPrevNext, self.screen_account_info)

        if sPrevNext != '2':
            self.detail_account_info_event_loop.exec_()

    def not_concluded_stock(self):
        print('======미체결요청 요청하는 부분======')

        self.dynamicCall('SetInputValue(String, String)', '계좌번호', self.account_num)
        self.dynamicCall('SetInputValue(String, String)', '전체종목구분', '0')
        self.dynamicCall('SetInputValue(String, String)', '매매구분', '0')
        self.dynamicCall('SetInputValue(String, String)', '체결구분', '1')
        self.dynamicCall('CommRqData(String, String, int, String)', '미체결요청', 'opt10075', '0', self.screen_account_info)

        self.detail_account_info_event_loop.exec_()

    def get_code_list_by_market(self, market_code):
        '''
        종목코드 반환
        :param market_code:
        :return:
        '''
        code_list = self.dynamicCall('GetCodeListByMarket(QString)', market_code)
        code_list = code_list.split(';')[:-1]

        return code_list

    def calculate_stock(self):
        '''
        종목분석 실행용 함수
        :return:
        '''
        code_list = self.get_code_list_by_market('10')
        print(f'Number of KOSDAQ: {len(code_list)}')

        for idx, code in enumerate(code_list):
            self.dynamicCall('DisconnectRealData(QString)', self.screen_stock_info) # 스크린번호가 계속 쌓이니까 안전하게 귾어주기. but 덮어쓰니까 궅이 안써도 괜찮

            print(f'{idx+1} / {len(code_list)} : KOSDAQ Stock Code : {code} is updating...')

            self.day_kiwoom_db(code=code)



    def day_kiwoom_db(self, code=None, date=None, sPrevNext='0'):
        QTest.qWait(3600)

        self.dynamicCall('SetInputValue(QString, QString)', '종목코드', code)
        self.dynamicCall('SetInputValue(QString, QString)', '수정주가구분', '1')

        if date != None:
            self.dynamicCall('SetInputValue(QString, QString)', '기준일자', date)

        self.dynamicCall('CommRqData(QString, QString, int, QString)', '주식일봉차트조회요청', 'opt10081', sPrevNext, self.screen_stock_info)

        if sPrevNext != '2':
            self.calculator_event_loop.exec_()



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
            deposit = int(self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '예수금'))
            print(f'예수금: {deposit}')

            self.usable_money = deposit * self.usable_money_rate
            self.usable_money = self.usable_money / 4

            ok_deposit = int(self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '출금가능금액'))
            print(f'출금가능금액: {ok_deposit}')

            self.detail_account_info_event_loop.exit()

        if sRQName == '계좌평가잔고내역요청':
            total_purchase_amount = int(self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '총매입금액'))
            total_earning_rate = float(self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '총수익률(%)'))

            print(f'총매입금액: {total_purchase_amount}')
            print(f'총수익률(%): {total_earning_rate}')

            items = self.dynamicCall('GetRepeatCnt(QString, QString)', sTrCode, sRQName)

            for i in range(items):
                code = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '종목번호')
                code_name = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '종목명')
                stock_quantity = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '보유수량')
                buy_price = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '매입가')
                earning_rate = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '수익률(%)')
                current_price = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '현재가')
                total_chegual_price = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '매입금액')
                possible_quantity = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '매매가능수량')

                code = code.strip()

                if code in self.account_stock_dict:
                    pass
                else:
                    self.account_stock_dict[code] = {}

                code_name = code_name.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                earning_rate = float(earning_rate.strip())
                current_price = int(current_price.strip())
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = int(possible_quantity.strip())

                self.account_stock_dict[code].update({'종목명': code_name})
                self.account_stock_dict[code].update({'보유수량': stock_quantity})
                self.account_stock_dict[code].update({'매입가': buy_price})
                self.account_stock_dict[code].update({'수익률(%)': earning_rate})
                self.account_stock_dict[code].update({'현재가': current_price})
                self.account_stock_dict[code].update({'매입금액': total_chegual_price})
                self.account_stock_dict[code].update({'매매가능수량': possible_quantity})

            print(f'계좌에 가지고 있는 종목: {self.account_stock_dict}')
            print(f'종목개수: {len(self.account_stock_dict)}')
            if sPrevNext == "2":
                self.account_eval_bal(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()

        if sRQName == '미체결요청':
            items = self.dynamicCall('GetRepeatCnt(QString, QString)', sTrCode, sRQName)
            print(items)

            for i in range(items):
                code = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '종목코드')
                code_name = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '종목명')
                order_no = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '주문번호')
                order_status = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '주문상태')
                order_quantity = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '주문수량')
                order_price = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '주문가격')
                order_division = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '주문구분')
                not_quantity = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '미체결수량')
                ok_quantity = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '체결량')

                code = code.strip()
                code_name = code_name.strip()
                order_no = int(order_no.strip())
                order_status = order_status.strip()
                order_quantity = int(order_quantity.strip())
                order_price = int(order_price.strip())
                order_division = order_division.strip().lstrip('+').lstrip('-')
                not_quantity = int(not_quantity.strip())
                ok_quantity = int(ok_quantity.strip())

                if order_no not in self.not_concluded_stock_dict:
                    self.not_concluded_stock_dict.update({order_no: {}})

                ncsd = self.not_concluded_stock_dict

                ncsd[order_no].update({'종목코드': code})
                ncsd[order_no].update({'종목명': code_name})
                ncsd[order_no].update({'주문상태': order_status})
                ncsd[order_no].update({'주문수량': order_quantity})
                ncsd[order_no].update({'주문가격': order_price})
                ncsd[order_no].update({'주문구분': order_division})
                ncsd[order_no].update({'미체결수량': not_quantity})
                ncsd[order_no].update({'체결량': ok_quantity})

            self.detail_account_info_event_loop.exit()

        if sRQName == '주식일봉차트조회요청':
            code = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, 0, '종목코드')
            code = code.strip()

            cnt = self.dynamicCall('GetRepeatCnt(QString, QString)', sTrCode, sRQName)
            print(f'{code} 일봉데이터 요청')
            print(f'데이터일수 {cnt}')

            # data = self.dynamicCall('GetCommDataEx(QString, QString)', sTrCode, sRQName)
            # [['', '현재가', '거래량', '거래대금', '일자', '시가', '고가', '저가', ''], ['', '현재가','거래량', '거래대금', '일자', '시가', '고가', '저가', ''], ...]

            for i in range(cnt):
                data = []

                current_price = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '현재가')
                volume = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '거래량')
                trading_value = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '거래대금')
                date = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '일자')
                start_price = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '시가')
                high_price = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '고가')
                low_price = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '저가')

                data.append('')
                data.append(current_price.strip())
                data.append(volume.strip())
                data.append(trading_value.strip())
                data.append(date.strip())
                data.append(start_price.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())

                self.calcul_data.append(data.copy())


            if sPrevNext == '2':
                self.day_kiwoom_db(code=code, sPrevNext=sPrevNext)
            else:
                print(f'총 일수 {len(self.calcul_data)}')
                pass_success = False

                # 120일 이평선을 그릴만큼의 데이터가 있는치 확인
                if self.calcul_data == None or len(self.calcul_data) < 120:
                    pass_success = False
                else:
                    #120일 이상 되면

                    total_price = 0
                    for data in self.calcul_data[:120]:
                        total_price += int(data[1])

                    moving_average_price = total_price / 120

                    #오늘 주가 120이평선에 걸쳐있는거 확인
                    bottom_stock_price = False
                    check_price = None #오늘자 고가
                    if int(self.calcul_data[0][7]) <= moving_average_price and int(self.calcul_data[0][6]) >= moving_average_price:
                        print('오늘 주가 120이평선에 걸쳐있는 것 확인')
                        bottom_stock_price = True
                        check_price = int(self.calcul_data[0][6])

                    # 과거의 일봉들이 120일 이평선보다 밑에 있는지 확인
                    # 그렇게 확인을 하다가 일봉이 120일 이평선보다 위에 있으면 계산 진행
                    prev_price = None # 과거의 일봉 저가
                    if bottom_stock_price == True:

                        moving_average_price_prev = 0
                        price_top_moving = False

                        idx = 1
                        while True:

                            if len(self.calcul_data[idx:]) < 120:
                                print('120일치 없음')
                                break

                            total_price = 0
                            for value in self.calcul_data[idx: idx+120]:
                                total_price += int(value[1])
                            moving_average_price_prev = total_price / 120

                            if idx <= 20 and moving_average_price_prev <= int(self.calcul_data[idx][6]):
                                print('20일 동안 주가가 120일 이평선보다 같거나 위에 있으면 조건 통과 못함')
                                price_top_moving = False
                                break

                            elif idx > 20 and moving_average_price_prev < int(self.calcul_data[idx][7]):
                                print('120일 이평선 위에 있는 일봉 확인')
                                price_top_moving = True
                                prev_price = int(self.calcul_data[idx][7])
                                break

                            idx += 1
                        if price_top_moving == True:
                            if prev_price < check_price and moving_average_price_prev < moving_average_price:
                                print('포착된 이평선의 가격이 오늘자(최근일자) 이평선 가격보다낮은 것 확인됨')
                                print('포착된 부분의 일봉 저가가 오늘자 일봉의 고가보다 낮은지 확인됨')
                                pass_success = True

                if pass_success == True:
                    print('조건부 통과됨')

                    code_name = self.dynamicCall('GetMasterCodeName(QString)', code)

                    f = open('files/condition_stock.txt', 'a', encoding='utf8')
                    f.write(f'{code}\t{code_name}\t{str(self.calcul_data[0][1])}')
                    f.close()

                else:
                    print('조건부 통과 못함')

                self.calcul_data.clear()
                self.calculator_event_loop.exit()

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

            self.basic_info_event_loop.exit()

    def read_file(self):

        if os.path.exists('files/condition_stock.txt'):
            f = open('file/condition_stock.txt', 'r', 'utf8')
            lines = f.readlines()

            for line in lines:
                if line != '':
                    ls = line.split('\t')

                    stock_code = ls[0]
                    stock_name = ls[1]
                    stock_price = int(ls[2].split('\n')[0])
                    stock_price = abs(stock_price)

                    self.portfolio_stock_dict.update({stock_code: {'종목명': stock_name}, '현재가': stock_price})

            f.close()
            print(self.portfolio_stock_dict)

    def set_screen_no(self):

        all_stock_code = []

        # 계좌평가잔고내역에 있는 종목코드
        for code in self.account_stock_dict.keys():
            if code not in all_stock_code:
                all_stock_code.append(code)

        # 미체결에 있는 종목코드
        for order_no in self.not_concluded_stock_dict.keys():
            code = self.not_concluded_stock_dict[order_no]['종목코드']

            if code not in all_stock_code:
                all_stock_code.append(code)

        # 포트폴리오에 있는 종목코드
        for code in self.portfolio_stock_dict.keys():
            if code not in all_stock_code:
                all_stock_code.append(code)

        # 스크린번호 할당
        cnt = 0
        screen_realtime = int(self.screen_realtime_stock)
        screen_trading = int(self.screen_trading_stock)

        if cnt % 50 == 0:
            screen_realtime += 1
            screen_trading += 1
            self.screen_realtime_stock = str(screen_realtime)
            self.screen_trading_stock = str(screen_trading)

        for code in all_stock_code:
            if code in self.portfolio_stock_dict.keys():
                self.portfolio_stock_dict[code].update({'스크린번호': str(screen_realtime), '주문용 스크린번호': str(screen_trading)})
            else:
                self.portfolio_stock_dict.update({code: {'스크린번호': str(screen_realtime), '주문용 스크린번호': str(screen_trading)}})

            cnt +=1

        print(self.portfolio_stock_dict)








