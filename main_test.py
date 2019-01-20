# -*-coding: utf-8 -
import sys, os, re, datetime, copy, json
import xlwings as xw
import resource_rc

import util, kw_util

from PyQt5 import QtCore, QtWidgets
from PyQt5.QtCore import QObject, pyqtSlot, pyqtSignal, QUrl, QEvent
from PyQt5.QtCore import QStateMachine, QState, QTimer, QFinalState
from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget
from mainwindow_ui import Ui_MainWindow


###################################################################################################
# 사용자 정의 파라미터
###################################################################################################

AUTO_TRADING_OPERATION_TIME = [ [ [9, 1], [15, 19] ] ] #해당 시스템 동작 시간 설정 장시작시 급등하고 급락하여 매수 / 매도 시 손해 나는 것을 막기 위해 1분 유예 둠 (반드시 할것)
CONDITION_NAME = '스캘퍼_시가갭' #키움증권 HTS 에서 설정한 조건 검색 식 총이름
TOTAL_BUY_AMOUNT = 10000 #  매도 호가1, 2 총 수량이 TOTAL_BUY_AMOUNT 이상 안되면 매수금지  (슬리피지 최소화)

MAESU_BASE_UNIT = 10000 # 추가 매수 기본 단위
MAESU_LIMIT = 3 # 추가 매수 제한
MAESU_TOTAL_PRICE =         [ MAESU_BASE_UNIT * 1,  MAESU_BASE_UNIT * 1,    MAESU_BASE_UNIT * 2,    MAESU_BASE_UNIT * 4,    MAESU_BASE_UNIT * 8 ]
# 추가 매수 진행시 stoploss 및 stopplus 퍼센티지 변경 최대 6
STOP_PLUS_PER_MAESU_COUNT = [ 8,                    8,                      8,                      8,                      8                  ]
STOP_LOSS_PER_MAESU_COUNT = [ 40,                   40,                     40,                     40,                     40,                ]

EXCEPTION_LIST = [] # 장기 보유 종목 번호 리스트  ex) EXCEPTION_LIST = ['034220']
STOCK_POSSESION_COUNT = 20 + len(EXCEPTION_LIST)   # 보유 종목수 제한

###################################################################################################
###################################################################################################

TEST_MODE = False    # 주의 TEST_MODE 를 True 로 하면 1주 단위로 삼

# DAY_TRADING_END_TIME 시간에 모두 시장가로 팔아 버림  반드시 동시 호가 시간 이전으로 입력해야함
# auto_trading_operation_time 이전값을 잡아야 함
DAY_TRADING_ENABLE = False
DAY_TRADING_END_TIME = [15, 10]

TRADING_INFO_GETTING_TIME = [15, 35] # 트레이딩 정보를 저장하기 시작하는 시간

SLIPPAGE = 0.5 # 보통가로 거래하므로 매매 수수료만 적용
CHUMAE_TIME_LILMIT_HOURS  = 7  # 다음 추가 매수시 보내야될 시간 조건   장 운영 시간으로만 계산하므로 약 6.5 시간이 하루임
TIME_CUT_MAX_DAY = 10  # 추가 매수 안한지 ?일 지나면 타임컷 수행하도록 함

TR_TIME_LIMIT_MS = 3800 # 키움 증권에서 정의한 연속 TR 시 필요 딜레이

CHEGYEOL_INFO_FILE_PATH = "log" + os.path.sep +  "chegyeol.json"
JANGO_INFO_FILE_PATH =  "log" + os.path.sep + "jango.json"
CHEGYEOL_INFO_EXCEL_FILE_PATH = "log" + os.path.sep +  "chegyeol.xlsx"

class CloseEventEater(QObject):
    def eventFilter(self, obj, event):
        if( event.type() == QEvent.Close):
            test_make_jangoInfo()
            return True
        else:
            return super(CloseEventEater, self).eventFilter(obj, event)

class Stock(QObject):
    sigInitOk = pyqtSignal()
    sigConnected = pyqtSignal()
    sigDisconnected = pyqtSignal()
    sigTryConnect = pyqtSignal()
    sigGetConditionCplt = pyqtSignal()
    sigSelectCondition = pyqtSignal()
    sigWaitingTrade = pyqtSignal()
    sigRefreshCondition = pyqtSignal()

    sigStateStop = pyqtSignal()
    sigStockComplete = pyqtSignal()

    sigConditionOccur = pyqtSignal()
    sigRequestInfo = pyqtSignal()
    sigRequestEtcInfo = pyqtSignal()

    sigGetBasicInfo = pyqtSignal()
    sigGetEtcInfo = pyqtSignal()
    sigGet5minInfo = pyqtSignal()
    sigGetHogaInfo = pyqtSignal()
    sigTrWaitComplete = pyqtSignal()

    sigBuy = pyqtSignal()
    sigNoBuy = pyqtSignal()
    sigRequestRealHogaComplete = pyqtSignal()
    sigError = pyqtSignal()
    sigRequestJangoComplete = pyqtSignal()
    sigCalculateStoplossComplete = pyqtSignal()
    sigStartProcessBuy = pyqtSignal()
    sigStopProcessBuy = pyqtSignal()
    sigTerminating = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.fsm = QStateMachine()
        self.account_list = []
        self.timerSystem = QTimer()
        self.lineCmdText = ''

        self.todayTradedCodeList = []  # 금일 거래 되었던 종목

        self.yupjongInfo = {'코스피': {}, '코스닥': {}}  # { 'yupjong_code': { '현재가': 222, ...} }
        self.michegyeolInfo = {}
        self.jangoInfo = {}  # { 'jongmokCode': { '이익실현가': 222, ...}}
        self.jangoInfoFromFile = {}  # TR 잔고 정보 요청 조회로는 얻을 수 없는 데이터를 파일로 저장하고 첫 실행시 로드함
        self.chegyeolInfo = {}  # { '날짜' : [ [ '주문구분', '매도', '주문/체결시간', '체결가' , '체결수량', '미체결수량'] ] }
        self.conditionOccurList = []  # 조건 진입이 발생한 모든 리스트 저장하고 매수 결정에 사용되는 모든 정보를 저장함  [ {'종목코드': code, ...}]
        self.conditionRevemoList = []  # 조건 이탈이 발생한 모든 리스트 저장

        self.kospiCodeList = ()
        self.kosdaqCodeList = ()

        self.createState()
        self.createConnection()
        self.currentTime = datetime.datetime.now()

    def createState(self):
        # state defintion
        mainState = QState(self.fsm)
        stockCompleteState = QState(self.fsm)
        finalState = QFinalState(self.fsm)
        self.fsm.setInitialState(mainState)

        initState = QState(mainState)
        disconnectedState = QState(mainState)
        connectedState = QState(QtCore.QState.ParallelStates, mainState)

        systemState = QState(connectedState)

        initSystemState = QState(systemState)
        waitingTradeSystemState = QState(systemState)
        standbySystemState = QState(systemState)
        requestingJangoSystemState = QState(systemState)
        calculateStoplossSystemState = QState(systemState)
        terminatingSystemState = QState(systemState)

        # transition defition
        mainState.setInitialState(initState)
        mainState.addTransition(self.sigStateStop, finalState)
        mainState.addTransition(self.sigStockComplete, stockCompleteState)
        stockCompleteState.addTransition(self.sigStateStop, finalState)
        initState.addTransition(self.sigInitOk, disconnectedState)
        disconnectedState.addTransition(self.sigConnected, connectedState)
        disconnectedState.addTransition(self.sigTryConnect, disconnectedState)
        connectedState.addTransition(self.sigDisconnected, disconnectedState)

        systemState.setInitialState(initSystemState)
        initSystemState.addTransition(self.sigGetConditionCplt, requestingJangoSystemState)
        requestingJangoSystemState.addTransition(self.sigRequestJangoComplete, calculateStoplossSystemState)
        calculateStoplossSystemState.addTransition(self.sigCalculateStoplossComplete, waitingTradeSystemState)

        waitingTradeSystemState.addTransition(self.sigWaitingTrade, waitingTradeSystemState)
        waitingTradeSystemState.addTransition(self.sigSelectCondition, standbySystemState)

        standbySystemState.addTransition(self.sigRefreshCondition, initSystemState)
        standbySystemState.addTransition(self.sigTerminating, terminatingSystemState)

        # state entered slot connect
        mainState.entered.connect(self.mainStateEntered)
        stockCompleteState.entered.connect(self.stockCompleteStateEntered)
        initState.entered.connect(self.initStateEntered)
        disconnectedState.entered.connect(self.disconnectedStateEntered)
        connectedState.entered.connect(self.connectedStateEntered)

        systemState.entered.connect(self.systemStateEntered)
        initSystemState.entered.connect(self.initSystemStateEntered)
        waitingTradeSystemState.entered.connect(self.waitingTradeSystemStateEntered)
        requestingJangoSystemState.entered.connect(self.requestingJangoSystemStateEntered)
        calculateStoplossSystemState.entered.connect(self.calculateStoplossPlusStateEntered)
        standbySystemState.entered.connect(self.standbySystemStateEntered)
        terminatingSystemState.entered.connect(self.terminatingSystemStateEntered)

        # processBuy definition
        processBuyState = QState(connectedState)
        initProcessBuyState = QState(processBuyState)
        standbyProcessBuyState = QState(processBuyState)
        requestEtcInfoProcessBuyState = QState(processBuyState)
        request5minInfoProcessBuyState = QState(processBuyState)
        determineBuyProcessBuyState = QState(processBuyState)
        waitingTRlimitProcessBuyState = QState(processBuyState)

        processBuyState.setInitialState(initProcessBuyState)
        initProcessBuyState.addTransition(self.sigStartProcessBuy, standbyProcessBuyState)

        standbyProcessBuyState.addTransition(self.sigConditionOccur, standbyProcessBuyState)
        standbyProcessBuyState.addTransition(self.sigRequestEtcInfo, requestEtcInfoProcessBuyState)
        standbyProcessBuyState.addTransition(self.sigStopProcessBuy, initProcessBuyState)

        requestEtcInfoProcessBuyState.addTransition(self.sigGetEtcInfo, waitingTRlimitProcessBuyState)
        requestEtcInfoProcessBuyState.addTransition(self.sigRequestInfo, request5minInfoProcessBuyState)
        requestEtcInfoProcessBuyState.addTransition(self.sigError, standbyProcessBuyState)

        request5minInfoProcessBuyState.addTransition(self.sigGet5minInfo, determineBuyProcessBuyState)
        request5minInfoProcessBuyState.addTransition(self.sigError, standbyProcessBuyState)

        determineBuyProcessBuyState.addTransition(self.sigNoBuy, waitingTRlimitProcessBuyState)
        determineBuyProcessBuyState.addTransition(self.sigBuy, waitingTRlimitProcessBuyState)

        waitingTRlimitProcessBuyState.addTransition(self.sigTrWaitComplete, standbyProcessBuyState)

        processBuyState.entered.connect(self.processBuyStateEntered)
        initProcessBuyState.entered.connect(self.initProcessBuyStateEntered)
        standbyProcessBuyState.entered.connect(self.standbyProcessBuyStateEntered)
        requestEtcInfoProcessBuyState.entered.connect(self.requestEtcInfoProcessBuyStateEntered)
        #request5minInfoProcessBuyState.entered.connect(self.request5minInfoProcessBuyStateEntered)
        determineBuyProcessBuyState.entered.connect(self.determineBuyProcessBuyStateEntered)
        print('############# State End ################')
        waitingTRlimitProcessBuyState.entered.connect(self.waitingTRlimitProcessBuyStateEntered)

        # fsm start
        finalState.entered.connect(self.finalStateEntered)
        self.fsm.start()

        pass

    def createConnection(self):
        self.ocx.OnEventConnect[int].connect(self._OnEventConnect)
        self.ocx.OnReceiveMsg[str, str, str, str].connect(self._OnReceiveMsg)
        self.ocx.OnReceiveTrData[str, str, str, str, str,
                                    int, str, str, str].connect(self._OnReceiveTrData)
        self.ocx.OnReceiveRealData[str, str, str].connect(
            self._OnReceiveRealData)
        self.ocx.OnReceiveChejanData[str, int, str].connect(
            self._OnReceiveChejanData)
        self.ocx.OnReceiveConditionVer[int, str].connect(
            self._OnReceiveConditionVer)
        self.ocx.OnReceiveTrCondition[str, str, str, int, int].connect(
            self._OnReceiveTrCondition)
        self.ocx.OnReceiveRealCondition[str, str, str, str].connect(
            self._OnReceiveRealCondition)

        self.timerSystem.setInterval(1000)
        self.timerSystem.timeout.connect(self.onTimerSystemTimeout)

    def isTradeAvailable(self):
        # 시간을 확인해 거래 가능한지 여부 판단
        ret_vals = []
        current_time = self.currentTime.time()
        for start, stop in AUTO_TRADING_OPERATION_TIME:
            start_time = datetime.time(
                hour=start[0],
                minute=start[1])
            stop_time = datetime.time(
                hour=stop[0],
                minute=stop[1])
            if (current_time >= start_time and current_time <= stop_time):
                ret_vals.append(True)
            else:
                ret_vals.append(False)
                pass

        # 하나라도 True 였으면 거래 가능시간임
        if (ret_vals.count(True)):
            return True
        else:
            return False
        pass

    @pyqtSlot(str, str, str, int, str, int, int, str, str, result=int)
    def sendOrder(self, rQName, screenNo, accNo, orderType, code, qty, price, hogaGb, orgOrderNo):
        return self.ocx.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)", [rQName, screenNo, accNo, orderType, code, qty, price, hogaGb, orgOrderNo])

    # 체결잔고 데이터를 반환한다.
    @pyqtSlot(int, result=str)
    def getChejanData(self, fid):
        return self.ocx.dynamicCall("GetChejanData(int)", fid)

    # 서버에 저장된 사용자 조건식을 가져온다.
    @pyqtSlot(result=int)
    def getConditionLoad(self):
        return self.ocx.dynamicCall("GetConditionLoad()")

    @pyqtSlot()
    def mainStateEntered(self):
        pass

    @pyqtSlot()
    def stockCompleteStateEntered(self):
        print(util.whoami())
        self.sigStateStop.emit()
        pass

    @pyqtSlot()
    def initStateEntered(self):
        print(util.whoami())
        self.sigInitOk.emit()
        pass

    @pyqtSlot()
    def disconnectedStateEntered(self):
        # print(util.whoami())
        if( self.getConnectState() == 0 ):
            self.commConnect()
            QTimer.singleShot(90000, self.sigTryConnect)
            pass
        else:
            self.sigConnected.emit()

    @pyqtSlot()
    def connectedStateEntered(self):
        # get 계좌 정보

        account_cnt = self.getLoginInfo("ACCOUNT_CNT")
        acc_num = self.getLoginInfo("ACCNO")
        user_id = self.getLoginInfo("USER_ID")
        user_name = self.getLoginInfo("USER_NAME")
        keyboard_boan = self.getLoginInfo("KEY_BSECGB")
        firewall = self.getLoginInfo("FIREW_SECGB")
        print("account count: {}, acc_num: {}, user id: {}, " 
              "user_name: {}, keyboard_boan: {}, firewall: {}"
              .format(account_cnt, acc_num, user_id, user_name,
                      keyboard_boan, firewall))

        self.account_list = (acc_num.split(';')[:-1])

        print(util.whoami() + 'account list ' + str(self.account_list))

        # 코스피 , 코스닥 종목 코드 리스트 얻기
        result = self.getCodeListByMarket('0')
        self.kospiCodeList = tuple(result.split(';'))
        result = self.getCodeListByMarket('10')
        self.kosdaqCodeList = tuple(result.split(';'))
        pass

    @pyqtSlot()
    def systemStateEntered(self):
        pass

    @pyqtSlot()
    def initSystemStateEntered(self):
        # 체결정보 로드
        if( os.path.isfile(CHEGYEOL_INFO_FILE_PATH) == True ):
            with open(CHEGYEOL_INFO_FILE_PATH, 'r', encoding='utf8') as f:
                file_contents = f.read()
                self.chegyeolInfo = json.loads(file_contents)
            # 금일 체결 정보는 매수 금지 리스트로 추가함
            for trade_date, data_chunk in self.chegyeolInfo.items():
                if( datetime.datetime.strptime(trade_date, "%y%m%d").date() == self.currentTime.date() ):
                    for trade_info in data_chunk:
                        parse_str_list = [item.strip() for item in trade_info.split('|') ]
                        if( len(parse_str_list)  < 5):
                            continue
                        jongmok_code_index = kw_util.dict_jusik['체결정보'].index('종목코드')
                        jumun_gubun_index = kw_util.dict_jusik['체결정보'].index('주문구분')

                        jongmok_code = parse_str_list[jongmok_code_index]
                        jumun_gubun  = parse_str_list[jumun_gubun_index]

                        if( jumun_gubun == "-매도"):
                            self.todayTradedCodeList.append(jongmok_code)
                    break

        if( os.path.isfile(JANGO_INFO_FILE_PATH) == True ):
            with open(JANGO_INFO_FILE_PATH, 'r', encoding='utf8') as f:
                file_contents = f.read()
                self.jangoInfoFromFile = json.loads(file_contents)

        # get 조건 검색 리스트
        self.getConditionLoad()
        self.setRealReg(kw_util.sendRealRegUpjongScrNo, '1;101', kw_util.type_fidset['업종지수'], "0")
        self.setRealReg(kw_util.sendRealRegTradeStartScrNo, '', kw_util.type_fidset['장시작시간'], "0")
        self.timerSystem.start()
        pass

    @pyqtSlot()
    def waitingTradeSystemStateEntered(self):
        # 장시작 전에 조건이 시작하도록 함
        time_span = datetime.timedelta(minutes=40)
        expected_time = (self.currentTime + time_span).time()
        if (expected_time >= datetime.time(*AUTO_TRADING_OPERATION_TIME[0][0])):
            self.sigSelectCondition.emit()

            # 반환값 : 조건인덱스1^조건명1;조건인덱스2^조건명2;…;
            # result = '조건인덱스1^조건명1;조건인덱스2^조건명2;'
            result = self.getConditionNameList()
            searchPattern = r'(?P<index>[^\/:*?"<>|;]+)\^(?P<name>[^\/:*?"<>|;]+);'
            fileSearchObj = re.compile(searchPattern, re.IGNORECASE)
            findList = fileSearchObj.findall(result)

            tempDict = dict(findList)
            print(tempDict)

            condition_num = 0
            for number, condition in tempDict.items():
                if condition == CONDITION_NAME:
                    condition_num = int(number)
            print("select condition" + kw_util.sendConditionScreenNo, CONDITION_NAME)
            self.sendCondition(kw_util.sendConditionScreenNo, CONDITION_NAME, condition_num, 1)


        else:
            QTimer.singleShot(1000, self.sigWaitingTrade)

        pass

    @pyqtSlot()
    def requestingJangoSystemStateEntered(self):
        # print(util.whoami() )
        self.requestOpw00018(self.account_list[0])
        pass

    @pyqtSlot()
    def calculateStoplossPlusStateEntered(self):
        # print(util.whoami() )
        # 이곳으로 온 경우 이미 잔고 TR 은 요청한 상태임
        for jongmok_code in self.jangoInfo:
            self.makeEtcJangoInfo(jongmok_code)
        self.makeJangoInfoFile()
        self.sigCalculateStoplossComplete.emit()

    @pyqtSlot()
    def standbySystemStateEntered(self):
        print(util.whoami())
        # 프로그램 첫 시작시 TR 요청으로 인한 제한 시간  막기 위해 딜레이 줌
        QTimer.singleShot(TR_TIME_LIMIT_MS * 5, self.sigStartProcessBuy)
        pass

    @pyqtSlot()
    def terminatingSystemStateEntered(self):
        print(util.whoami())
        pass

    @pyqtSlot()
    def processBuyStateEntered(self):
        pass

    @pyqtSlot()
    def initProcessBuyStateEntered(self):
        print(util.whoami())
        pass

    @pyqtSlot()
    def standbyProcessBuyStateEntered(self):
        # print(util.whoami() )
        # 운영 시간이 아닌 경우 운영시간이 될때까지 지속적으로 확인
        if (self.isTradeAvailable() == False):
            print(util.whoami())
            QTimer.singleShot(10000, self.sigConditionOccur)
            return
        else:
            # 무한으로 시그널 발생 방지를 위해 딜레이 줌
            QTimer.singleShot(100, self.sigRequestEtcInfo)

    @pyqtSlot()
    def requestEtcInfoProcessBuyStateEntered(self):
        # 조건 발생 리스트 검색
        for jongmok_code in self.conditionRevemoList:
            self.removeConditionOccurList(jongmok_code)
        self.conditionRevemoList.clear()

        self.refreshRealRequest()

        jongmok_info = self.getConditionOccurList()

        if (jongmok_info):
            jongmok_code = jongmok_info['종목코드']
            jongmok_name = jongmok_info['종목명']
            if ('상한가' not in jongmok_info):
                self.requestOpt10001(jongmok_code)
            else:
                self.sigRequestInfo.emit()
            # print(util.whoami() , jongmok_name, jongmok_code )
            return
        else:
            self.sigError.emit()
        pass

    @pyqtSlot()
    def request5minInfoProcessBuyStateEntered(self):
        # print(util.whoami() )
        jongmok_info_dict = self.getConditionOccurList()

        if (not jongmok_info_dict):
            self.shuffleConditionOccurList()
            self.sigError.emit()
            return

        code = jongmok_info_dict['종목코드']
        # 아직 실시간 정보를 못받아온 상태라면
        # 체결 정보 받는데 시간 걸리므로 다른 종목 폴링
        # 혹은 집입했닥 이탈하면 데이터 삭제 하므로 실시간 정보가 없을수도 있다.
        if ('매도호가1' not in jongmok_info_dict or '등락율' not in jongmok_info_dict):
            self.shuffleConditionOccurList()
            if ('매도호가1' not in jongmok_info_dict):
                print('매도호가1 not in {0}'.format(code))
            else:
                print('등락율 not in {0}'.format(code))
            self.sigError.emit()
            return

        if (self.requestOpt10080(code) == False):
            self.sigError.emit()

        pass

    @pyqtSlot()
    def determineBuyProcessBuyStateEntered(self):
        print("########### determineBuyProcessBuyStateEntered ###############")
        jongmok_info_dict = []
        is_log_print_enable = False
        return_vals = []
        printLog = ''

        jongmok_info_dict = self.getConditionOccurList()
        print(jongmok_info_dict)
        #result = self.sendOrder("buy_" + jongmokCode, kw_util.sendOrderScreenNo,objKiwoom.account_list[0], kw_util.dict_order["신규매수"], jongmokCode, qty, 0, kw_util.dict_order["시장가"], "")


    # event
    # 통신 연결 상태 변경시 이벤트
    # nErrCode가 0이면 로그인 성공, 음수면 실패
    def _OnEventConnect(self, errCode):
        print(util.whoami() + '{}'.format(errCode))
        if errCode == 0:
            self.sigConnected.emit()
        else:
            self.sigDisconnected.emit()

    # 수신 메시지 이벤트
    def _OnReceiveMsg(self, scrNo, rQName, trCode, msg):
        # print(util.whoami() + 'sScrNo: {}, sRQName: {}, sTrCode: {}, sMsg: {}'
        # .format(scrNo, rQName, trCode, msg))

        # [107066] 매수주문이 완료되었습니다.
        # [107048] 매도주문이 완료되었습니다
        # [571489] 장이 열리지않는 날입니다
        # [100000] 조회가 완료되었습니다
        printData =  'sScrNo: {}, sRQName: {}, sTrCode: {}, sMsg: {}'.format(scrNo, rQName, trCode, msg)

        # buy 하다가 오류 난경우 강제로 buy signal 생성
        if( 'buy' in rQName and '107066' not in msg ):
            QTimer.singleShot(TR_TIME_LIMIT_MS,  self.sigBuy )

        print(printData)
        util.save_log(printData, "시스템메시지", "log")
        pass


    @pyqtSlot()
    def waitingTRlimitProcessBuyStateEntered(self):
        # print(util.whoami())
        # TR 은 개당 3.7 초 제한
        # 5연속 조회시 17초 대기 해야함
        # print(util.whoami() )
        QTimer.singleShot(TR_TIME_LIMIT_MS,  self.sigTrWaitComplete)
        pass

    @pyqtSlot()
    def finalStateEntered(self):
        print(util.whoami())
        self.makeJangoInfoFile()
        util.save_log('', subject= '', folder='log')
        util.save_log('', subject= '', folder='log')
        util.save_log('', subject= '', folder='log')
        util.save_log('', subject= '', folder='log')
        util.save_log('', subject= '', folder='log')
        util.save_log('', subject= '', folder='log')
        util.save_log('', subject= '', folder='log')
        util.save_log('', subject= '', folder='log')
        util.save_log('', subject= '', folder='log')
        import subprocess
        subprocess.call(["shutdown", "-s", "-t", "500"])
        sys.exit()
        pass

    def makeEtcJangoInfo(self, jongmok_code, priority='server'):

        if (jongmok_code not in self.jangoInfo):
            return
        current_jango = {}

        if (priority == 'server'):
            current_jango = self.jangoInfo[jongmok_code]
            maeip_price = current_jango['매입가']

            if ('매수횟수' not in current_jango):
                if (jongmok_code in self.jangoInfoFromFile):
                    current_jango['매수횟수'] = self.jangoInfoFromFile[jongmok_code].get('매수횟수', 1)
                else:
                    current_jango['매수횟수'] = 1
                    pass

            maesu_count = current_jango['매수횟수']
            # 손절가 다시 계산
            stop_loss_value = STOP_LOSS_PER_MAESU_COUNT[maesu_count - 1]
            stop_plus_value = STOP_PLUS_PER_MAESU_COUNT[maesu_count - 1]

            current_jango['손절가'] = round(maeip_price * (1 - (stop_loss_value - SLIPPAGE) / 100), 2)
            current_jango['이익실현가'] = round(maeip_price * (1 + (stop_plus_value + SLIPPAGE) / 100), 2)

            if ('주문/체결시간' not in current_jango):
                if (jongmok_code in self.jangoInfoFromFile):
                    current_jango['주문/체결시간'] = self.jangoInfoFromFile[jongmok_code].get('주문/체결시간', [])
                else:
                    current_jango['주문/체결시간'] = []

            if ('체결가/체결시간' not in current_jango):
                if (jongmok_code in self.jangoInfoFromFile):
                    current_jango['체결가/체결시간'] = self.jangoInfoFromFile[jongmok_code].get('체결가/체결시간', [])
                else:
                    current_jango['체결가/체결시간'] = []

            if ('최근매수가' not in current_jango):
                if (jongmok_code in self.jangoInfoFromFile):
                    current_jango['최근매수가'] = self.jangoInfoFromFile[jongmok_code].get('최근매수가', [])
                else:
                    current_jango['최근매수가'] = []
        else:

            if (jongmok_code in self.jangoInfoFromFile):
                current_jango = self.jangoInfoFromFile[jongmok_code]
            else:
                current_jango = self.jangoInfo[jongmok_code]

        self.jangoInfo[jongmok_code].update(current_jango)
        pass

    @pyqtSlot()
    def makeJangoInfoFile(self):
        print(util.whoami())
        remove_keys = ['매도호가1', '매도호가2', '매도호가수량1', '매도호가수량2', '매도호가총잔량',
                       '매수호가1', '매수호가2', '매수호가수량1', '매수호가수량2', '매수호가수량3', '매수호가수량4', '매수호가총잔량',
                       '현재가', '호가시간', '세금', '전일종가', '현재가', '종목번호', '수익율', '수익', '잔고', '매도중', '시가', '고가', '저가', '장구분',
                       '거래량', '등락율', '전일대비', '기준가', '상한가', '하한가', '5분봉타임컷기준']
        temp = copy.deepcopy(self.jangoInfo)
        # 불필요 필드 제거
        for jongmok_code, contents in temp.items():
            for key in remove_keys:
                if (key in contents):
                    del contents[key]

        with open(JANGO_INFO_FILE_PATH, 'w', encoding='utf8') as f:
            f.write(json.dumps(temp, ensure_ascii=False, indent=2, sort_keys=True))
        pass

        # 조건검색 조건명 리스트를 받아온다.
        # 조건명 리스트(인덱스^조건명)
        # 조건명 리스트를 구분(“;”)하여 받아온다

    @pyqtSlot(result=str)
    def getConditionNameList(self):
        return self.ocx.dynamicCall("GetConditionNameList()")

    @pyqtSlot(str, str, int, int)
    def sendCondition(self, scrNo, conditionName, index, search):
        self.ocx.dynamicCall("SendCondition(QString,QString, int, int)", scrNo, conditionName, index, search)

    @pyqtSlot(str, str, int)
    def sendConditionStop(self, scrNo, conditionName, index):
        self.ocx.dynamicCall("SendConditionStop(QString, QString, int)", scrNo, conditionName, index)

    @pyqtSlot(str, bool, int, int, str, str)
    def commKwRqData(self, arrCode, next, codeCount, typeFlag, rQName, screenNo):
        self.ocx.dynamicCall("CommKwRqData(QString, QBoolean, int, int, QString, QString)", arrCode, next, codeCount,
                             typeFlag, rQName, screenNo)

    # 실시간 시세 이벤트
    def _OnReceiveRealData(self, jongmokCode, realType, realData):
        # print(util.whoami() + 'jongmokCode: {}, {}, realType: {}'
        #         .format(jongmokCode, self.getMasterCodeName(jongmokCode),  realType))

        # 장전에도 주식 호가 잔량 값이 올수 있으므로 유의해야함
        if (realType == "주식호가잔량"):
            # print(util.whoami() + 'jongmokCode: {}, realType: {}, realData: {}'
            #     .format(jongmokCode, realType, realData))

            self.makeHogaJanRyangInfo(jongmokCode)

            # 주식 체결로는 사고 팔기에는 반응이 너무 느림
        elif (realType == "주식체결"):
            # print(util.whoami() + 'jongmokCode: {}, realType: {}, realData: {}'
            #     .format(jongmokCode, realType, realData))
            self.makeBasicInfo(jongmokCode)

            # WARNING: 장중에 급등으로 거래 정지 되어 동시 호가진행되는 경우에 대비하여 체결가 정보 발생했을때만 stoploss 진행함.
            self.processStopLoss(jongmokCode)
            pass

        elif (realType == "주식시세"):
            # 장종료 후에 나옴
            # print(util.whoami() + 'jongmokCode: {}, realType: {}, realData: {}'
            #     .format(jongmokCode, realType, realData))
            pass

        elif (realType == "업종지수"):
            # print(util.whoami() + 'jongmokCode: {}, realType: {}, realData: {}'
            #     .format(jongmokCode, realType, realData))
            result = ''
            for col_name in kw_util.dict_jusik['실시간-업종지수']:
                result = self.getCommRealData(jongmokCode, kw_util.name_fid[col_name])
                if (jongmokCode == '001'):
                    self.yupjongInfo['코스피'][col_name] = result.strip()
                elif (jongmokCode == '100'):
                    self.yupjongInfo['코스닥'][col_name] = result.strip()
            pass

        elif (realType == '장시작시간'):
            # TODO: 장시작 30분전부터 실시간 정보가 올라오는데 이를 토대로 가변적으로 장시작시간을 가늠할수 있도록 기능 추가 필요
            # 장운영구분(0:장시작전, 2:장종료전, 3:장시작, 4,8:장종료, 9:장마감)
            # 동시호가 시간에 매수 주문
            result = self.getCommRealData(realType, kw_util.name_fid['장운영구분'])
            if (result == '2'):
                self.sigTerminating.emit()
            elif (result == '4'):  # 장종료 후 5분뒤에 프로그램 종료 하게 함
                QTimer.singleShot(300000, self.sigStockComplete)

            # print(util.whoami() + 'jongmokCode: {}, realType: {}, realData: {}'
            #     .format(jongmokCode, realType, realData))

            print(util.whoami() + 'jongmokCode: {}, realType: {}, realData: {}'
                  .format(jongmokCode, realType, realData))
            pass

    def calculateSuik(self, jongmok_code, current_price):
        current_jango = self.jangoInfo[jongmok_code]
        maeip_price = abs(int(current_jango['매입가']))
        boyou_suryang = int(current_jango['보유수량'])

        suik_price = round((current_price - maeip_price) * boyou_suryang, 2)
        current_jango['수익'] = suik_price
        current_jango['수익율'] = round(((current_price - maeip_price) / maeip_price) * 100, 2)
        pass

    # 실시간 호가 잔량 정보
    def makeHogaJanRyangInfo(self, jongmokCode):
        # 주식 호가 잔량 정보 요청
        result = None
        for col_name in kw_util.dict_jusik['실시간-주식호가잔량']:
            result = self.getCommRealData(jongmokCode, kw_util.name_fid[col_name])

            if (jongmokCode in self.jangoInfo):
                self.jangoInfo[jongmokCode][col_name] = result.strip()
            if (jongmokCode in self.getCodeListConditionOccurList()):
                self.setHogaConditionOccurList(jongmokCode, col_name, result.strip())
        pass

        # 실시간 체결(기본) 정보

    def makeBasicInfo(self, jongmokCode):
        # 주식 호가 잔량 정보 요청
        result = None
        for col_name in kw_util.dict_jusik['실시간-주식체결']:
            result = self.getCommRealData(jongmokCode, kw_util.name_fid[col_name])

            if (jongmokCode in self.jangoInfo):
                self.jangoInfo[jongmokCode][col_name] = result.strip()
            if (jongmokCode in self.getCodeListConditionOccurList()):
                self.setHogaConditionOccurList(jongmokCode, col_name, result.strip())
        pass

    def processStopLoss(self, jongmokCode):
        jongmokName = self.getMasterCodeName(jongmokCode)
        if (self.isTradeAvailable() == False):
            return

        # 예외 처리 리스트이면 종료
        if (jongmokCode in EXCEPTION_LIST):
            return

        # 잔고에 없는 종목이면 종료
        if (jongmokCode not in self.jangoInfo):
            return
        current_jango = self.jangoInfo[jongmokCode]

        if ('손절가' not in current_jango or '매수호가1' not in current_jango or '매매가능수량' not in current_jango):
            return

        jangosuryang = int(current_jango['매매가능수량'])
        stop_loss = 0
        ########################################################################################
        # 업종 이평가를 기준으로 stop loss 값 조정
        # twenty_avr = 0
        # five_avr = 0
        # if( '20봉평균' in self.yupjongInfo['코스피'] and
        #     '5봉평균' in self.yupjongInfo['코스피'] and
        #     '20봉평균' in self.yupjongInfo['코스닥'] and
        #     '5봉평균' in self.yupjongInfo['코스닥']
        # ):
        #     if( jongmokCode in self.kospiCodeList):
        #         twenty_avr = abs(float(self.yupjongInfo['코스피']['20봉평균']))
        #         five_avr = abs(float(self.yupjongInfo['코스피']['5봉평균']))
        #     else:
        #         twenty_avr = abs(float(self.yupjongInfo['코스닥']['20봉평균']))
        #         five_avr = abs(float(self.yupjongInfo['코스닥']['5봉평균']))

        #     stop_loss = int(current_jango['손절가'])
        # else:
        #     stop_loss = int(current_jango['손절가'])
        stop_loss = int(current_jango['손절가'])
        stop_plus = int(current_jango['이익실현가'])
        maeipga = int(current_jango['매입가'])

        # 호가 정보는 문자열로 기준가 대비 + , - 값이 붙어 나옴
        maesuHoga1 = abs(int(current_jango['매수호가1']))
        maesuHogaAmount1 = int(current_jango['매수호가수량1'])
        maesuHoga2 = abs(int(current_jango['매수호가2']))
        maesuHogaAmount2 = int(current_jango['매수호가수량2'])
        #    print( util.whoami() +  maeuoga1 + " " + maesuHogaAmount1 + " " + maesuHoga2 + " " + maesuHogaAmount2 )
        totalAmount = maesuHoga1 * maesuHogaAmount1 + maesuHoga2 * maesuHogaAmount2
        # print( util.whoami() + jongmokName + " " + str(sum))

        isSell = False
        printData = jongmokCode + ' {0:20} '.format(jongmokName)

        ########################################################################################
        # time cut 적용
        base_time_str = ''
        last_chegyeol_time_str = ''
        # if( '5분봉타임컷기준' in current_jango ):
        #     base_time_str =  current_jango['5분봉타임컷기준'][2]
        #     base_time = datetime.datetime.strptime(base_time_str, '%Y%m%d%H%M%S')
        #     last_chegyeol_time_str = current_jango['주문/체결시간'][-1]
        #     maeip_time = datetime.datetime.strptime(last_chegyeol_time_str, '%Y%m%d%H%M%S')

        #     if( maeip_time < base_time ):
        #         stop_loss = 99999999

        #########################################################################################
        # day trading 용
        if (DAY_TRADING_ENABLE == True):
            # day trading 주식 거래 시간 종료가 가까운 경우 모든 종목 매도
            time_span = datetime.timedelta(minutes=10)
            dst_time = datetime.datetime.combine(datetime.date.today(),
                                                 datetime.time(*DAY_TRADING_END_TIME)) + time_span

            current_time = datetime.datetime.now()
            if (datetime.time(*DAY_TRADING_END_TIME) < current_time.time() and dst_time > current_time):
                # 0 으로 넣고 로그 남기면서 매도 처리하게 함
                stop_loss = 0
                pass

        # 손절 / 익절 계산
        # 정리나, 손절의 경우 시장가로 팔고 익절의 경우 보통가로 팜
        isSijanga = False
        maedo_type = ''
        if (stop_loss == 0):
            maedo_type = "(당일정리)"
            printData += maedo_type
            isSijanga = True
            isSell = True
        elif (stop_loss == 99999999):
            maedo_type = "(타임컷임)"
            printData += maedo_type
            isSijanga = True
            isSell = True
        elif (stop_loss >= maesuHoga1):
            maedo_type = "(손절이다)"
            printData += maedo_type
            isSijanga = True
            isSell = True
        elif (stop_plus < maesuHoga1):
            if (totalAmount >= TOTAL_BUY_AMOUNT):
                maedo_type = "(익절이다)"
                printData += maedo_type
                isSell = True
            else:
                maedo_type = "(익절미달)"
                printData += maedo_type
                isSell = True

        printData += ' 손절가: {0:7}/'.format(str(stop_loss)) + \
                     ' 이익실현가: {0:7}/'.format(str(stop_plus)) + \
                     ' 매입가: {0:7}/'.format(str(maeipga)) + \
                     ' 잔고수량: {0:7}'.format(str(jangosuryang)) + \
                     ' 타임컷 기준 시간: {0:7}'.format(base_time_str) + \
                     ' 최근 주문/체결시간: {0:7}'.format(last_chegyeol_time_str) + \
                     ' 매수호가1 {0:7}/'.format(str(maesuHoga1)) + \
                     ' 매수호가수량1 {0:7}/'.format(str(maesuHogaAmount1)) + \
                     ' 매수호가2 {0:7}/'.format(str(maesuHoga2)) + \
                     ' 매수호가수량2 {0:7}/'.format(str(maesuHogaAmount2))

        if (isSell == True):
            # processStop 의 경우 체결될때마다 호출되므로 중복 주문이 나가지 않게 함
            if ('매도중' not in current_jango):
                current_jango['매도중'] = maedo_type
                if (isSijanga == True):
                    result = self.sendOrder("sell_" + jongmokCode, kw_util.sendOrderScreenNo, objKiwoom.account_list[0],
                                            kw_util.dict_order["신규매도"],
                                            jongmokCode, jangosuryang, 0, kw_util.dict_order["시장가"], "")
                else:
                    result = self.sendOrder("sell_" + jongmokCode, kw_util.sendOrderScreenNo, objKiwoom.account_list[0],
                                            kw_util.dict_order["신규매도"],
                                            jongmokCode, jangosuryang, maesuHoga1, kw_util.dict_order["지정가"], "")

                util.save_log(printData, '매도', 'log')
                print("S " + jongmokCode + ' ' + str(result), sep="")
            pass
        pass

    # 체결데이터를 받은 시점을 알려준다.
    # sGubun – 0:주문체결통보, 1:잔고통보, 3:특이신호
    # sFidList – 데이터 구분은 ‘;’ 이다.
    '''
    _OnReceiveChejanData gubun: 1, itemCnt: 27, fidList: 9201;9001;917;916;302;10;930;931;932;933;945;946;950;951;27;28;307;8019;957;958;918;990;991;992;993;959;924
    {'종목코드': 'A010050', '당일실현손익률(유가)': '0.00', '대출일': '00000000', '당일실현손익률(신용)': '0.00', '(최우선)매수호가': '+805', '당일순매수수량': '5', '총매입가': '4043', 
    '당일총매도손일': '0', '만기일': '00000000', '신용금액': '0', '당일실현손익(신용)': '0', '현재가': '+806', '기준가': '802', '계좌번호': ', '보유수량': '5', 
    '예수금': '0', '주문가능수량': '5', '종목명': '우리종금                                ', '손익율': '0.00', '당일실현손익(유가)': '0', '담보대출수량': '0', '924': '0', 
    '매입단가': '809', '신용구분': '00', '매도/매수구분': '2', '(최우선)매도호가': '+806', '신용이자': '0'}
    '''

    # 로컬에 사용자조건식 저장 성공여부 응답 이벤트
    # 0:(실패) 1:(성공)
    def _OnReceiveConditionVer(self, ret, msg):
        print(util.whoami() + 'ret: {}, msg: {}'
              .format(ret, msg))
        if ret == 1:
            self.sigGetConditionCplt.emit()

    # 조건검색 조회응답으로 종목리스트를 구분자(“;”)로 붙어서 받는 시점.
    # LPCTSTR sScrNo : 종목코드
    # LPCTSTR strCodeList : 종목리스트(“;”로 구분)
    # LPCTSTR strConditionName : 조건명
    # int nIndex : 조건명 인덱스
    # int nNext : 연속조회(2:연속조회, 0:연속조회없음)
    def _OnReceiveTrCondition(self, scrNo, codeList, conditionName, index, next):
        # print(util.whoami() + 'scrNo: {}, codeList: {}, conditionName: {} '
        # 'index: {}, next: {}'
        # .format(scrNo, codeList, conditionName, index, next ))
        codes = codeList.split(';')[:-1]
        # 마지막 split 결과 None 이므로 삭제
        for code in codes:
            print('condition occur list add code: {} '.format(code) + self.getMasterCodeName(code))
            self.addConditionOccurList(code)

    # 편입, 이탈 종목이 실시간으로 들어옵니다.
    # strCode : 종목코드
    # strType : 편입(“I”), 이탈(“D”)
    # strConditionName : 조건명
    # strConditionIndex : 조건명 인덱스
    def _OnReceiveRealCondition(self, code, type, conditionName, conditionIndex):
        print(util.whoami() + 'code: {}, type: {}, conditionName: {}, conditionIndex: {}'
              .format(code, type, conditionName, conditionIndex))
        if type == 'I':
            self.addConditionOccurList(code)  # 조건 발생한 경우 해당 내용 list 에 추가
        else:
            self.conditionRevemoList.append(code)
            pass

    def addConditionOccurList(self, jongmok_code):
        # 발생시간, 종목코드,  종목명
        jongmok_name = self.getMasterCodeName(jongmok_code)
        ret_vals = []

        # 중복 제거
        for item_dict in self.conditionOccurList:
            if (jongmok_code == item_dict['종목코드']):
                ret_vals.append(True)

        if (ret_vals.count(True)):
            pass
        else:
            self.conditionOccurList.append({'종목명': jongmok_name, '종목코드': jongmok_code})
            self.sigConditionOccur.emit()
        pass

    def removeConditionOccurList(self, jongmok_code):
        for item_dict in self.conditionOccurList:
            if (item_dict['종목코드'] == jongmok_code):
                self.conditionOccurList.remove(item_dict)
                break
        pass

    def getConditionOccurList(self):
        if (len(self.conditionOccurList)):
            return self.conditionOccurList[0]
        else:
            return None
        pass

    def getCodeListConditionOccurList(self):
        items = []
        for item_dict in self.conditionOccurList:
            items.append(item_dict['종목코드'])
        return items

    def setHogaConditionOccurList(self, jongmok_code, col_name, value):
        for index, item_dict in enumerate(self.conditionOccurList):
            if (item_dict['종목코드'] == jongmok_code):
                item_dict[col_name] = value

    # 다음 codition list 를 감시 하기 위해 종목 섞기
    def shuffleConditionOccurList(self):
        jongmok_info_dict = self.getConditionOccurList()
        jongmok_code = jongmok_info_dict['종목코드']
        self.removeConditionOccurList(jongmok_code)
        self.conditionOccurList.append(jongmok_info_dict)

    # 실시간  주식 정보 요청 요청리스트 갱신
    # WARNING: 실시간 요청도 TR 처럼 초당 횟수 제한이 있으므로 잘 사용해야함
    def refreshRealRequest(self):
        # 버그로 모두 지우고 새로 등록하게 함
        # print(util.whoami() )
        self.setRealRemove("ALL", "ALL")
        codeList = []

        for code in self.jangoInfo.keys():
            if (code not in codeList):
                codeList.append(code)

        condition_list = self.getCodeListConditionOccurList()
        for code in condition_list:
            if (code not in codeList):
                codeList.append(code)

        if (len(codeList) == 0):
            # 종목 미보유로 실시간 체결 요청 할게 없는 경우 코스닥 코스피 실시간 체결가가 올라오지 않으므로 임시로 하나 등록
            codeList.append('044180')
        else:
            for code in codeList:
                if (code not in EXCEPTION_LIST):
                    self.addConditionOccurList(code)

        # 실시간 호가 정보 요청 "0" 은 이전거 제외 하고 새로 요청
        if (len(codeList)):
            #  WARNING: 주식 시세 실시간은 리턴되지 않음!
            #    tmp = self.setRealReg(kw_util.sendRealRegSiseSrcNo, ';'.join(codeList), kw_util.type_fidset['주식시세'], "0")
            tmp = self.setRealReg(kw_util.sendRealRegHogaScrNo, ';'.join(codeList), kw_util.type_fidset['주식호가잔량'], "0")
            tmp = self.setRealReg(kw_util.sendRealRegChegyeolScrNo, ';'.join(codeList), kw_util.type_fidset['주식체결'],
                                  "0")
            tmp = self.setRealReg(kw_util.sendRealRegUpjongScrNo, '001;101', kw_util.type_fidset['업종지수'], "0")

    @pyqtSlot(str, str, str, str,  result=int)
    def setRealReg(self, screenNo, codeList, fidList, optType):
        return self.ocx.dynamicCall("SetRealReg(QString, QString, QString, QString)", screenNo, codeList, fidList, optType)

    @pyqtSlot(str, str)
    def setRealRemove(self, scrNo, delCode):
        self.ocx.dynamicCall("SetRealRemove(QString, QString)", scrNo, delCode)

    def make_excel(self, file_path, data_dict):
        result = False
        result = os.path.isfile(file_path)
        if (result):
            # excel open
            wb = xw.Book(file_path)
            sheet_names = [sheet.name for sheet in wb.sheets]
            insert_sheet_names = []
            # print(sheet_names)
            for key, value in data_dict.items():
                # sheet name 이 존재 안하면 sheet add
                sheet_name = key[0:4]
                if (sheet_name not in sheet_names):
                    if (sheet_name not in insert_sheet_names):
                        insert_sheet_names.append(sheet_name)

            for insert_sheet in insert_sheet_names:
                wb.sheets.add(name=insert_sheet)
            # sheet name 은 YYMM 형식
            sheet_names = [sheet.name for sheet in wb.sheets]
            all_items = []

            for sheet_name in sheet_names:
                # key 값이 match 되는것을 찾음
                for sorted_key in sorted(data_dict):
                    input_data_sheet_name = sorted_key[0:4]
                    if (input_data_sheet_name == sheet_name):
                        all_items.append([sorted_key, '', '', '', '', '', '', '', '', '', '-' * 128])
                        for line in data_dict[sorted_key]:
                            items = [item.strip() for item in line.split('|')]
                            items.insert(0, '')
                            all_items.append(items)

                wb.sheets[sheet_name].activate()
                xw.Range('A1').value = all_items
                all_items.clear()

            # save
            wb.save()
            wb.app.quit()

    # method
    # 로그인
    # 0 - 성공, 음수값은 실패
    # 단순 API 호출이 되었느냐 안되었느냐만 확인 가능
    @pyqtSlot(result=int)
    def commConnect(self):
        return self.ocx.dynamicCall("CommConnect()")

    # 로그인 상태 확인
    # 0:미연결, 1:연결완료, 그외는 에러
    @pyqtSlot(result=int)
    def getConnectState(self):
        return self.ocx.dynamicCall("GetConnectState()")

    # 로그 아웃
    @pyqtSlot()
    def commTerminate(self):
        self.ocx.dynamicCall("CommTerminate()")

    # 로그인한 사용자 정보를 반환한다.
    # “ACCOUNT_CNT” – 전체 계좌 개수를 반환한다.
    # "ACCNO" – 전체 계좌를 반환한다. 계좌별 구분은 ‘;’이다.
    # “USER_ID” - 사용자 ID를 반환한다.
    # “USER_NAME” – 사용자명을 반환한다.
    # “KEY_BSECGB” – 키보드보안 해지여부. 0:정상, 1:해지
    # “FIREW_SECGB” – 방화벽 설정 여부. 0:미설정, 1:설정, 2:해지
    @pyqtSlot(str, result=str)
    def getLoginInfo(self, tag):
        return self.ocx.dynamicCall("GetLoginInfo(QString)", [tag])

    # Tran 입력 값을 서버통신 전에 입력값일 저장한다.
    @pyqtSlot(str, str)
    def setInputValue(self, id, value):
        self.ocx.dynamicCall("SetInputValue(QString, QString)", id, value)

    @pyqtSlot(str, result=str)
    def getCodeListByMarket(self, sMarket):
        return self.ocx.dynamicCall("GetCodeListByMarket(QString)", sMarket)

    # 통신 데이터를 송신한다.
    # 0이면 정상
    # OP_ERR_SISE_OVERFLOW – 과도한 시세조회로 인한 통신불가
    # OP_ERR_RQ_STRUCT_FAIL – 입력 구조체 생성 실패
    # OP_ERR_RQ_STRING_FAIL – 요청전문 작성 실패
    # OP_ERR_NONE – 정상처리
    @pyqtSlot(str, str, int, str, result=int)
    def commRqData(self, rQName, trCode, prevNext, screenNo):
        return self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)", rQName, trCode, prevNext, screenNo)

    # 수신 받은 데이터의 반복 개수를 반환한다.
    @pyqtSlot(str, str, result=int)
    def getRepeatCnt(self, trCode, recordName):
        return self.ocx.dynamicCall("GetRepeatCnt(QString, QString)", trCode, recordName)

    # Tran 데이터, 실시간 데이터, 체결잔고 데이터를 반환한다.
    # 1. Tran 데이터
    # 2. 실시간 데이터
    # 3. 체결 데이터
    # 1. Tran 데이터
    # sJongmokCode : Tran명
    # sRealType : 사용안함
    # sFieldName : 레코드명
    # nIndex : 반복인덱스
    # sInnerFieldName: 아이템명
    # 2. 실시간 데이터
    # sJongmokCode : Key Code
    # sRealType : Real Type
    # sFieldName : Item Index (FID)
    # nIndex : 사용안함
    # sInnerFieldName:사용안함
    # 3. 체결 데이터
    # sJongmokCode : 체결구분
    # sRealType : “-1”
    # sFieldName : 사용안함
    # nIndex : ItemIndex
    # sInnerFieldName:사용안함
    @pyqtSlot(str, str, str, int, str, result=str)
    def commGetData(self, jongmokCode, realType, fieldName, index, innerFieldName):
        return self.ocx.dynamicCall("CommGetData(QString, QString, QString, int, QString)", jongmokCode, realType,
                                    fieldName, index, innerFieldName).strip()

    # strRealType – 실시간 구분
    # nFid – 실시간 아이템
    # Ex) 현재가출력 - openApi.GetCommRealData(“주식시세”, 10);
    # 참고)실시간 현재가는 주식시세, 주식체결 등 다른 실시간타입(RealType)으로도 수신가능
    @pyqtSlot(str, int, result=str)
    def getCommRealData(self, realType, fid):
        return self.ocx.dynamicCall("GetCommRealData(QString, int)", realType, fid).strip()
    # Tran 수신시 이벤트
    def _OnReceiveTrData(   self, scrNo, rQName, trCode, recordName,
                            prevNext, dataLength, errorCode, message,
                            splmMsg):
        # print(util.whoami() + 'sScrNo: {}, rQName: {}, trCode: {}, prevNext {}'
        # .format(scrNo, rQName, trCode, prevNext))

        # rQName 은 계좌번호임
        if ( trCode == 'opw00018' ):
            if( self.makeOpw00018Info(rQName) ):
                self.sigRequestJangoComplete.emit()
            else:
                self.sigError.emit()
            pass
        elif( trCode =='opt10081'):
            if( self.makeOpt10081Info(rQName) ):
                # 잔고 정보를 뒤져서 손절가 책정이 되었는지 확인
                # ret_vals = []
                # for jangoInfo in self.jangoInfo.values():
                #     if( '손절가' not in jangoInfo.keys() ):
                #         ret_vals.append(False)
                # if( ret_vals.count(False) == 0 ):
                #     self.printStockInfo()
                #     self.sigCalculateStoplossComplete.emit()
                pass
            else:
                self.sigError.emit()

        #주식 기본 정보 요청 rQName 은 개별 종목 코드임
        elif( trCode == "opt10001"):
            if( self.makeOpt10001Info(rQName) ):
                self.sigGetEtcInfo.emit()
            else:
                self.sigError.emit()
            pass

        # 주식 분봉 정보 요청 rQName 개별 종목 코드
        elif( trCode == "opt10080"):
            if( self.makeOpt10080Info(rQName) ) :
                print('TEST',self.sigGet5minInfo)
                self.sigGet5minInfo.emit()
            else:
                self.sigError.emit()
            pass

        # 업종 분봉 rQName 업종 코드
        elif( trCode == "opt20005"):
            if( self.makeOpt20005Info(rQName) ) :
                self.sigGetEtcInfo.emit()
            else:
                self.sigError.emit()
            pass

    @pyqtSlot()
    def onTimerSystemTimeout(self):
        # print(".", end='')
        self.currentTime = datetime.datetime.now()
        if (self.getConnectState() != 1):
            util.save_log("Disconnected!", "시스템", folder="log")
            self.sigDisconnected.emit()
        else:
            if (datetime.time(*TRADING_INFO_GETTING_TIME) <= self.currentTime.time()):
                self.timerSystem.stop()
                util.save_log("Stock Trade Terminate!\n\n\n\n\n", "시스템", folder="log")
                pass
            else:
                pass
        pass

    # 주식 잔고정보 요청
    def requestOpw00018(self, account_num):
        self.setInputValue('계좌번호', account_num)
        self.setInputValue('비밀번호', '') #  사용안함(공백)
        self.setInputValue('비빌번호입력매체구분', '00')
        self.setInputValue('조회구분', '1')

        ret = self.commRqData(account_num, "opw00018", 0, kw_util.sendAccountInfoScreenNo)
        errorString = None
        if( ret != 0 ):
            errorString =   account_num + " commRqData() " + kw_util.parseErrorCode(str(ret))
            print(util.whoami() + errorString )
            util.save_log(errorString, util.whoami(), folder = "log" )
            return False
        return True

    # 주식 잔고 정보 #rQName 의 경우 계좌 번호로 넘겨줌
    def makeOpw00018Info(self, rQName):
        data_cnt = self.getRepeatCnt('opw00018', rQName)
        for cnt in range(data_cnt):
            info_dict = {}
            for item_name in kw_util.dict_jusik['TR:계좌평가잔고내역요청']:
                result = self.getCommData("opw00018", rQName, cnt, item_name)
                # 없는 컬럼은 pass
                if (len(result) == 0):
                    continue
                if (item_name == '종목명'):
                    info_dict[item_name] = result.strip()
                elif (item_name == '종목번호'):
                    info_dict[item_name] = result[1:-1].strip()
                elif (item_name == '수익률(%)'):
                    info_dict[item_name] = int(result) / 100
                else:
                    info_dict[item_name] = int(result)

            jongmokCode = info_dict['종목번호']

            if (jongmokCode not in self.jangoInfo.keys()):
                self.jangoInfo[jongmokCode] = info_dict
            else:
                # 기존에 가지고 있는 종목이면 update
                self.jangoInfo[jongmokCode].update(info_dict)

        # print(self.jangoInfo)
        return True

        # 주식 1일봉 요청

    def requestOpt10081(self, jongmokCode):
        # print(util.cur_time_msec() )
        datetime_str = datetime.datetime.now().strftime('%Y%m%d')
        self.setInputValue("종목코드", jongmokCode)
        self.setInputValue("기준일자", datetime_str)
        self.setInputValue('수정주가구분', '1')
        ret = self.commRqData(jongmokCode, "opt10081", 0, kw_util.sendGibonScreenNo)
        errorString = None
        if (ret != 0):
            errorString = jongmokCode + " commRqData() " + kw_util.parseErrorCode(str(ret))
            print(util.whoami() + errorString)
            util.save_log(errorString, util.whoami(), folder="log")
            return False
        return True

    # 주식 일봉 차트 조회
    def makeOpt10081Info(self, rQName):
        return True

    # 주식 분봉 tr 요청
    def requestOpt10080(self, jongmokCode):
        # 분봉 tr 요청의 경우 너무 많은 데이터를 요청하므로 한개씩 수행
        self.setInputValue("종목코드", jongmokCode)
        self.setInputValue("틱범위", "5:5분")
        self.setInputValue("수정주가구분", "1")
        # rQName 을 데이터로 외부에서 사용
        ret = self.commRqData(jongmokCode, "opt10080", 0, kw_util.send5minScreenNo)
        errorString = None
        if (ret != 0):
            errorString = jongmokCode + " commRqData() " + kw_util.parseErrorCode(str(ret))
            print(util.whoami() + errorString)
            util.save_log(errorString, util.whoami(), folder="log")
            return False
        return True

    # 분봉 데이터 생성
    def makeOpt10080Info(self, rQName):
        jongmok_info_dict = self.getConditionOccurList()
        if (jongmok_info_dict):
            pass
        else:
            return False
        repeatCnt = self.getRepeatCnt("opt10080", rQName)

        total_current_price_list = []

        for i in range(min(repeatCnt, 800)):
            line = []
            for item_name in kw_util.dict_jusik['TR:분봉']:
                result = self.getCommData("opt10080", rQName, i, item_name)
                if (item_name == "현재가"):
                    total_current_price_list.append(abs(int(result)))

                line.append(result.strip())
            key_value = '5분 {0}봉전'.format(i)
            jongmok_info_dict[key_value] = line

        for i in range(0, 200):
            fivebong_sum, twentybong_sum, sixtybong_sum, twohundred_sum = 0, 0, 0, 0

            twohundred_sum = sum(total_current_price_list[i:200 + i])
            jongmok_info_dict['200봉{}평균'.format(i)] = int(twohundred_sum / 200)
        jongmok_code = jongmok_info_dict['종목코드']
        if (jongmok_code in self.jangoInfo):
            time_cut_5min = '5분 {0}봉전'.format(TIME_CUT_MAX_DAY * 78)
            self.jangoInfo[jongmok_code]['5분봉타임컷기준'] = jongmok_info_dict[time_cut_5min]

        # RSI 14 calculate
        rsi_up_sum = 0
        rsi_down_sum = 0
        index_current_price = kw_util.dict_jusik['TR:분봉'].index('현재가')

        for i in range(14, -1, -1):
            key_value = '5분 {0}봉전'.format(i)
            if (i != 14):
                key_value = '5분 {0}봉전'.format(i + 1)
                prev_fivemin_close = abs(int(jongmok_info_dict[key_value][index_current_price]))
                key_value = '5분 {0}봉전'.format(i)
                fivemin_close = abs(int(jongmok_info_dict[key_value][index_current_price]))
                if (prev_fivemin_close < fivemin_close):
                    rsi_up_sum += fivemin_close - prev_fivemin_close
                elif (prev_fivemin_close > fivemin_close):
                    rsi_down_sum += prev_fivemin_close - fivemin_close
            pass

        rsi_up_avg = rsi_up_sum / 14
        rsi_down_avg = rsi_down_sum / 14
        if (rsi_up_avg != 0 and rsi_down_avg != 0):
            rsi_value = round(rsi_up_avg / (rsi_up_avg + rsi_down_avg) * 100, 1)
        else:
            rsi_value = 100
        jongmok_info_dict['RSI14'] = str(rsi_value)
        # print(util.whoami(), self.getMasterCodeName(jongmok_code), jongmok_code,  'rsi_value: ',  rsi_value)
        return True

    # 업종 분봉 tr 요청
    def requestOpt20005(self, yupjong_code):
        self.setInputValue("업종코드", yupjong_code)
        self.setInputValue("틱범위", "5:5분")
        self.setInputValue("수정주가구분", "1")
        ret = 0
        if (yupjong_code == '001'):
            ret = self.commRqData(yupjong_code, "opt20005", 0, kw_util.sendReqYupjongKospiScreenNo)
        else:
            ret = self.commRqData(yupjong_code, "opt20005", 0, kw_util.sendReqYupjongKosdaqScreenNo)

        errorString = None
        if (ret != 0):
            errorString = yupjong_code + " commRqData() " + kw_util.parseErrorCode(str(ret))
            print(util.whoami() + errorString)
            util.save_log(errorString, util.whoami(), folder="log")
            return False
        return True

    # 업종 분봉 데이터 생성
    def makeOpt20005Info(self, rQName):
        if (rQName == '001'):
            yupjong_info_dict = self.yupjongInfo['코스피']
        elif (rQName == '101'):
            yupjong_info_dict = self.yupjongInfo['코스닥']
        else:
            return

        repeatCnt = self.getRepeatCnt("opt20005", rQName)

        fivebong_sum = 0
        twentybong_sum = 0
        for i in range(min(repeatCnt, 20)):
            line = []
            for item_name in kw_util.dict_jusik['TR:업종분봉']:
                result = self.getCommData("opt20005", rQName, i, item_name)
                if (item_name == "현재가"):
                    current_price = abs(int(result)) / 100
                    if (i < 5):
                        fivebong_sum += current_price
                    twentybong_sum += current_price
                    line.append(str(current_price))
                else:
                    line.append(result.strip())
            key_value = '5분 {0}봉전'.format(i)
            yupjong_info_dict[key_value] = line

        yupjong_info_dict['20봉평균'] = str(round(twentybong_sum / 20, 2))
        yupjong_info_dict['5봉평균'] = str(round(fivebong_sum / 5, 2))
        return True

    # 주식 기본 정보 요청
    def requestOpt10001(self, jongmokCode):
        # print(util.cur_time_msec() )
        self.setInputValue("종목코드", jongmokCode)
        ret = self.commRqData(jongmokCode, "opt10001", 0, kw_util.sendGibonScreenNo)
        errorString = None
        if (ret != 0):
            errorString = jongmokCode + " commRqData() " + kw_util.parseErrorCode(str(ret))
            print(util.whoami() + errorString)
            util.save_log(errorString, util.whoami(), folder="log")
            return False
        return True

    # 주식 기본 차트 조회 ( multi data 아님 )
    def makeOpt10001Info(self, rQName):
        jongmok_code = rQName
        jongmok_info_dict = self.getConditionOccurList()
        if (jongmok_info_dict):
            pass
        else:
            return False

        for item_name in kw_util.dict_jusik['TR:기본정보']:
            result = self.getCommData("opt10001", rQName, 0, item_name)
            if (jongmok_code in self.jangoInfo):
                self.jangoInfo[jongmok_code][item_name] = result.strip()
            jongmok_info_dict[item_name] = result.strip()
        return True


    @pyqtSlot(str, str, int, str, result=str)
    def getCommData(self, trCode, recordName, index, itemName):
        return self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)",
        trCode, recordName, index, itemName)

    @pyqtSlot(str, str, result=str)
    def getCommDataEx(self, trCode, recordName):
        return self.ocx.dynamicCall("GetCommDataEx(QString, QString)", trCode, recordName)

    def _OnReceiveChejanData(self, gubun, itemCnt, fidList):
        # print(util.whoami() + 'gubun: {}, itemCnt: {}, fidList: {}'
        #         .format(gubun, itemCnt, fidList))
        if (gubun == "1"):  # 잔고 정보
            # 잔고 정보에서는 매도/매수 구분이 되지 않음

            jongmok_code = self.getChejanData(kw_util.name_fid['종목코드'])[1:]
            boyou_suryang = int(self.getChejanData(kw_util.name_fid['보유수량']))
            jumun_ganeung_suryang = int(self.getChejanData(kw_util.name_fid['주문가능수량']))
            maeip_danga = int(self.getChejanData(kw_util.name_fid['매입단가']))
            jongmok_name = self.getChejanData(kw_util.name_fid['종목명']).strip()
            current_price = abs(int(self.getChejanData(kw_util.name_fid['현재가'])))

            # 미체결 수량이 있는 경우 잔고 정보 저장하지 않도록 함
            if (jongmok_code in self.michegyeolInfo):
                if (self.michegyeolInfo[jongmok_code]['미체결수량']):
                    return
                    # 미체결 수량이 없으므로 정보 삭제
            del (self.michegyeolInfo[jongmok_code])
            if (boyou_suryang == 0):
                # 보유 수량이 0 인 경우 매도 수행
                if (jongmok_code not in self.todayTradedCodeList):
                    self.todayTradedCodeList.append(jongmok_code)
                self.jangoInfo.pop(jongmok_code)
                self.removeConditionOccurList(jongmok_code)
            else:
                # 보유 수량이 늘었다는 것은 매수수행했다는 소리임
                self.sigBuy.emit()

                # 아래 잔고 정보의 경우 TR:계좌평가잔고내역요청 필드와 일치하게 만들어야 함
                current_jango = {}
                current_jango['보유수량'] = boyou_suryang
                current_jango['매매가능수량'] = jumun_ganeung_suryang  # TR 잔고에서 매매가능 수량 이란 이름으로 사용되므로
                current_jango['매입가'] = maeip_danga
                current_jango['종목번호'] = jongmok_code
                current_jango['종목명'] = jongmok_name.strip()
                chegyeol_info = util.cur_date_time('%Y%m%d%H%M%S') + ":" + str(current_price)

                if (jongmok_code not in self.jangoInfo):
                    current_jango['주문/체결시간'] = [util.cur_date_time('%Y%m%d%H%M%S')]
                    current_jango['체결가/체결시간'] = [chegyeol_info]
                    current_jango['최근매수가'] = [current_price]
                    current_jango['매수횟수'] = 1

                    self.jangoInfo[jongmok_code] = current_jango

                else:
                    chegyeol_time_list = self.jangoInfo[jongmok_code]['주문/체결시간']
                    chegyeol_time_list.append(util.cur_date_time('%Y%m%d%H%M%S'))
                    current_jango['주문/체결시간'] = chegyeol_time_list

                    last_chegyeol_info = self.jangoInfo[jongmok_code]['체결가/체결시간'][-1]
                    if (int(last_chegyeol_info.split(':')[1]) != current_price):
                        chegyeol_info_list = self.jangoInfo[jongmok_code]['체결가/체결시간']
                        chegyeol_info_list.append(chegyeol_info)
                        current_jango['체결가/체결시간'] = chegyeol_info_list

                    price_list = self.jangoInfo[jongmok_code]['최근매수가']
                    last_price = price_list[-1]
                    if (last_price != current_price):
                        # 매수가 나눠져서 진행 중이므로 자료 매수횟수 업데이트 안함
                        price_list.append(current_price)
                    current_jango['최근매수가'] = price_list

                    chumae_count = self.jangoInfo[jongmok_code]['매수횟수']
                    if (last_price != current_price):
                        current_jango['매수횟수'] = chumae_count + 1
                    else:
                        current_jango['매수횟수'] = chumae_count

                    self.jangoInfo[jongmok_code].update(current_jango)

            self.makeEtcJangoInfo(jongmok_code)
            self.makeJangoInfoFile()
            pass

        elif (gubun == "0"):
            jumun_sangtae = self.getChejanData(kw_util.name_fid['주문상태'])
            jongmok_code = self.getChejanData(kw_util.name_fid['종목코드'])[1:]
            michegyeol_suryang = int(self.getChejanData(kw_util.name_fid['미체결수량']))
            # 주문 상태
            # 매수 시 접수(gubun-0) - 체결(gubun-0) - 잔고(gubun-1)
            # 매도 시 접수(gubun-0) - 잔고(gubun-1) - 체결(gubun-0) - 잔고(gubun-1)   순임
            # 미체결 수량 정보를 입력하여 잔고 정보 처리시 미체결 수량 있는 경우에 대한 처리를 하도록 함
            if (jongmok_code not in self.michegyeolInfo):
                self.michegyeolInfo[jongmok_code] = {}
            self.michegyeolInfo[jongmok_code]['미체결수량'] = michegyeol_suryang

            if (jumun_sangtae == "체결"):
                self.makeChegyeolInfo(jongmok_code, fidList)
                self.makeChegyeolInfoFile()
                pass

            pass

    @pyqtSlot(str, result=str)
    def getMasterCodeName(self, strCode):
        return self.ocx.dynamicCall("GetMasterCodeName(QString)", strCode)

if __name__ == "__main__":

    @pyqtSlot()
    def test_make_jangoInfo():
        objKiwoom.makeJangoInfoFile()
        pass

    myApp = QtWidgets.QApplication(sys.argv)
    objKiwoom = Stock()

    sys.exit(myApp.exec_())

