# -*- coding: utf-8 -*-
'''    [화면번호]
        화면번호는 서버에 시세조회나 주문을 요청할때 이 요청을 구별하기 위한 키값으로 이해하시면 됩니다.
        0000(혹은 0)을 제외한 임의의 숫자를 사용하시면 되는데 갯수가 200개로 한정되어 있기 때문에 이 갯수를 넘지 않도록 관리하셔야 합니다.
        만약 사용하는 화면번호가 200개를 넘는 경우 조회결과나 주문결과에 다른 데이터가 섞이거나 원하지 않는 결과를 나타날 수 있습니다.
'''
sendConditionScreenNo = "001"

# 실시간 체결 화면번호
# 실시간 잔고의 경우 체결시 콜백함수를 이용해야 하며 실시간 잔고로는 잔고 조회가 안됨  
sendRealRegHogaScrNo = "002"
sendRealRegUpjongScrNo = '003'
sendRealRegChegyeolScrNo = '004'
sendRealRegTradeStartScrNo = '005'
sendRealRegSiseSrcNo = "006"


sendOrderScreenNo = "101"

sendOrderETFScreenNo = "201"
sendOrderETFPairScreenNo = "202"

sendReqYupjongKospiScreenNo = "102"
sendReqYupjongKosdaqScreenNo = "103"

sendGibonScreenNo = "104"
send5minScreenNo = "105"
sendHogaScreenNo = "106"
sendAccountInfoScreenNo = '107'



dict_order = {
    "신규매수" : 1,
    "신규매도" : 2,
    "매수취소" : 3,
    "매도취소" : 4,
    "매수정정" : 5,
    "매도정정" : 6,
    "지정가"   : "00",
    "시장가"   : "03",
    "조건부지정가" : "05",
    "최유리지정가" : "06",
    "최우선지정가" : "07",
    "지정가IOC"    : "10",
    "시장가IOC"    : "13",
    "최유리IOC"    : "16",
    "시장가FOK"    : "23",
    "최유리FOK"    : "26",
    "장전시간외종가": "61",
    "시간외단일가매매": "62",
    "장후시간외종가xxx": "81"
}

# 저장시 필요한 리스트만 나열한 것임
dict_jusik = {
    # 체결 정보는 파일에 저장됨
    "체결정보": (
        '수익율', '수익', '매수횟수', '주문타입', # 원래 없는 멤버
        '종목코드', '주문구분', '체결가', '체결량', '주문/체결시간', '종목명'),

    'TR:계좌평가잔고내역요청': (
        '종목명', '종목번호', #'평가손익', '수익률(%)', --> 실시간 잔고 기능 사용시 계속 잔고 요청해야하므로 삭제
        '매입가', '전일종가', '보유수량',
        '매매가능수량', '현재가' # '세금'
    ),
    "TR:업종분봉": (
        "현재가", "거래량", "체결시간"
    ),
    "TR:분봉": (
        "현재가", "거래량", "체결시간", "시가", "고가", "저가"
    ),
    "TR:일봉": (
        "현재가", "거래량", "시가", "고가", "저가", '일자'
    ),
    "TR:기본정보": (
        "상한가", "하한가", "기준가"
    ),

    "실시간-주식체결":(
        # "체결시간",
        "현재가",
        "전일대비",
        "등락율",
        # "(최우선)매도호가",
        # "(최우선)매수호가",
        # "거래량",
        # "누적거래량",
        # # "누적거래대금",
        "시가",
        "고가",
        "저가"
        #"시가총액(억)",
        #"장구분"
    ),
    "실시간-주식호가잔량": (
        '호가시간',
        '매도호가1',
        '매도호가수량1',
        '매수호가1',
        '매수호가수량1',
        '매도호가2',
        '매도호가수량2',
        '매수호가2',
        '매수호가수량2',
        # '매수호가수량3',
        # '매수호가수량4',
        #'누적거래량'
        # '예상체결가',
        # '예상체결수량',
        # '누적거래량',
    ),
    '실시간-업종지수': (
        '체결시간',
        '현재가',
        '등락율',
        '전일대비기호',
    ),
    '실시간-장시작시간':(
        '장운영구문',
        '체결시간',
        '장시작예상잔여시간'
    )
}
# fid 는 다 넣을 필요 없음
type_fidset = {
    "주식시세": "10;11;12;27;28;13;14;16;17;18;25;26;29;30;31;32;311;567;568",
    "주식체결": "20;10;11;12;27;28;15;13;14;16;17;18;25;26;29;30;31;32;311;290;691;567;568",
    "주식호가잔량":"21;41;61;81;51;71;91",
    '업종지수': "20;10;11;12;15;13;14;16;17;18;25;26",
    '장시작시간': "215;20;214"
}
name_fid = {
    '호가시간': 21,
    '매도호가1': 41,
    '매도호가수량1': 61,
    '매수호가1': 51,
    '매수호가수량1': 71,
    '매도호가2': 42,
    '매도호가수량2': 62,
    '매수호가2': 52,
    '매수호가수량2': 72,
    '매도호가3': 43,
    '매도호가수량3': 63,
    '매수호가3': 53,
    '매수호가수량3': 73,
    '매도호가4': 44,
    '매도호가수량4': 64,
    '매수호가4': 54,
    '매수호가수량4': 74,
    '매도호가5': 45,
    '매도호가수량5': 65,
    '매수호가5': 55,
    '매수호가수량5': 75,
    '매도호가총잔량': 121,
    '매수호가총잔량': 125,
    '누적거래량': 13,

    "계좌번호": 9201,
    "주문번호": 9203,
    "관리자사번": 9205,
    "종목코드": 9001,
    "주문업무분류": 912,  # (JJ:주식주문, FJ:선물옵션, JG:주식잔고, FG:선물옵션잔고)
    "주문상태": 913,
    "종목명": 302,
    "주문수량": 900,
    "주문가격": 901,
    "미체결수량": 902,
    "체결누계금액": 903,
    "원주문번호": 904,
    "주문구분": 905, # (+현금내수,-현금매도…)
    "매매구분": 906,  # (보통,시장가…)
    "매도매수구분": 907, #(1:매도,2:매수)
    "주문/체결시간": 908,
    "체결번호": 909,
    "체결가": 910,
    "체결량": 911,
    "현재가": 10,
    "(최우선)매도호가": 27,
    "(최우선)매수호가": 28,
    "단위체결가": 914,
    "단위체결량": 915,
    "당일매매수수료": 938,
    "당일매매세금": 939,
    "거부사유": 919,
    "화면번호": 920,
    "921": 921,
    "922": 922,
    "923" : 923,
    "924" : 924,
    "신용구분": 917,
    "대출일": 916,
    "보유수량": 930,
    "매입단가": 931,
    "총매입가": 932,
    "주문가능수량": 933,
    "당일순매수수량": 945,
    "매도/매수구분":  946,
    "당일총매도손실": 950,
    "예수금": 951,
    "담보대출수량": 959,
    "기준가": 307,
    "손익율": 8019,
    "신용금액": 957,
    "신용이자": 958,
    "만기일": 918,
    "당일실현손익(유가)": 990,
    "당일실현손익률(유가)": 991,
    "당일실현손익(신용)": 992,
    "당일실현손익률(신용)":  993,
    "파생상품거래단위": 397,
    "상한가":   305,
    "하한가": 306,
    "기준가": 307,

    "등락율": 12,
    "체결시간": 20,
    "전일대비기호":   25,
    '전일대비': 11,
    '거래량': 15,
    '누적거래대금': 14,
    '시가': 16,
    '고가': 17,
    '저가': 18,
    '장구분': 290,
    '장운영구분': 215
}

def parseErrorCode(code):
    """에러코드 메시지

        :param code: 에러 코드
        :type code: str
        :return: 에러코드 메시지를 반환

        ::

            kw_util.parseErrorCode("00310") # 모의투자 조회가 완료되었습니다
    """
    code = str(code)
    ht = {
        "0" : "정상처리",
        "-100" : "사용자정보교환에 실패하였습니다. 잠시후 다시 시작하여 주십시오.",
        "-101" : "서버 접속 실패",
        "-102" : "버전처리가 실패하였습니다.",
        "-103" : "개인방화벽 실패",
        "-104" : "메모리 보호 실패",
        "-105" : "함수 입력값 오류",
        "-106" : "통신 연결 종료",

        "-200" : "시세조회 과부하",
        "-201" : "REQUEST_INPUT_st Failed",
        "-202" : "요청 전문 작성 실패",
        "-203" : "데이터 없음",
        "-204" : "조회가능한 종목수 초과, 한번에 조회 가능한 종목 개수는 최대 100종목",
        "-205" : "데이터 수신 실패",
        "-206" : "실시간 해제 오류",

        "-300" : "주문 입력값 오류",
        "-301" : "계좌비밀번호를 입력하십시오.",
        "-302" : "타인계좌는 사용할 수 없습니다.",
        "-303" : "주문가격이 20억원을 초과합니다.",
        "-304" : "주문가격은 50억원을 초과할 수 없습니다.",
        "-305" : "주문수량이 총발행주수의 1%를 초과합니다.",
        "-306" : "주문수량은 총발행주수의 3%를 초과할 수 없습니다.",
        "-307" : "주문전송 실패",
        "-308" : "주문전송 과부하",
        "-309" : "주문수량 300계약 초과",
        "-310" : "계좌 정보 없음",
        "-500" : "종목코드 없음"
    }
    return ht[code] + " (%s)" % code if code in ht else code