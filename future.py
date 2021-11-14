'''
실시간으로 신규 매수주문 넣기, 테스트 주문형태 포함
'''

import win32com.client
import pythoncom
import time
from config.errCode import *
from config.accountCalculator import *

'''
사용될 변수 모아놓는 클래스
'''
class Object:

    ##### 오브젝트 모음 #####
    # 우리가 요청하는 TR요청 설정 변수
    XAQuery_o3105 = None # 종목정보 요청
    XAQuery_CIDBQ01500 = None # 해외선물 미결제 정보내역 요청
    XAQuery_CIDBQ03000 = None # 해외선물 예수금/잔고현황
    XAQuery_CIDBT00100 = None # 해외선물 신규주문
    XAQuery_CIDBT01000 = None # 해외선물 취소주문
    XAQuery_CIDBT00900 = None # 해외선물 정정주문

    # 실시간으로 수신되는 설정 변수
    XAReal_OVH = None  # 종목 실시간 호가 정보 받기
    XAReal_OVC = None  # 종목 실시간 틱봉 정보 받기
    XARealOrder_TC1 = None # 주문 접수 데이터 받기
    XARealOrder_TC2 = None # 주문 응답 데이터 받기
    XARealOrder_TC3 = None # 주문 체결 데이터 받기


    #############################

    ##### 함수를 할당한 변수 #####
    tr_signal_o3105 = None # 종목정보tr요청 함수를 할당
    tr_signal_CIDBQ01500 = None # 미결제 함수를 할당
    tr_signal_CIDBQ03000 = None # 예수금/잔고현황 함수를 할당
    order_buy_CIDBT00100 = None # 해외선물 신규주문
    order_cancel_CIDBT01000 = None # 해외선물 취소주문 
    
    #############################

    ##### 기타 변수 모음 #####
    TR처리완료 = False # TR요청완료 기다리기
    로그인완료 = False # 로그인완료인지 확인하기
    해외선물_계좌번호 = None # 해외선물 계좌번호만 할당
    종목정보_딕셔너리 = {} # 종목의 상세 정보 데이터
    미결제_딕셔너리 = {} # 체결 후 결제되지 않은 상품
    예수금_딕셔너리 = {} # 예수금 상세정보 데이터
    실시간호가_딕셔너리 = {} # 실시간으로 변하는 호가 데이터
    실시간체결_딕셔너리 = {} # 실시간으로 변하는 틱봉 데이터

    주문접수_딕셔너리 = {} # 신규매수, 매도, 정정, 취소 주문 접수 데이터를 일시적으로 담아 놓는다.
    주문응답_딕셔너리 = {} # 정정, 취소 주문이 정상적으로 응답된 데이터를 일시적으로 담아 놓는다.
    주문체결_딕셔너리 = {} # 신규매수, 매도, 정정 주문이 체결되면 데이터를 일시적으로 담아 놓는다.

    #############################

    ##### 기타설정 #####
    매수 = False # 매수 주문 추가 방지
    매도 = False # 매도 주문 추가 방지
    취소 = False # 취소 주문 추가 방지
    정정 = False # 정정 주문 추가 방지
    #############################


'''
로그인 및 연결 상태에 대한 정보를 반환받게 정의해주는 이벤트
'''
class XASessionEvent(Object):

    # 로그인 처리 결과를 받는 이벤트 핸들러
    def OnLogin(self, szCode, szMsg):
        print("★★★ 로그인 %s, %s" % (szCode, szMsg))

        if szCode == "0000":
            Object.로그인완료 = True
        else:
            Object.로그인완료 = False

    # 연결 끊김 결과를 받는 이벤트 핸들러
    def OnDisconnect(self):
        print("★★★ 연결 끊김")

'''
단일 요청으로 원하는 데이터를 반환받게 정의해주는 이벤트
'''
class XAQueryEvent(Object):

    def OnReceiveData(self, szTrCode):

        if szTrCode == "o3105":
            print("★★★ 해외선물 종목정보 결과반환")

            종목코드 = self.GetFieldData("o3105OutBlock", "Symbol", 0)
            if 종목코드 != "":
                종목명 = self.GetFieldData("o3105OutBlock", "SymbolNm", 0)
                종목배치수신일 = self.GetFieldData("o3105OutBlock", "ApplDate", 0)
                기초상품코드 = self.GetFieldData("o3105OutBlock", "BscGdsCd", 0)
                기초상품명 = self.GetFieldData("o3105OutBlock", "BscGdsNm", 0)
                거래소코드 = self.GetFieldData("o3105OutBlock", "ExchCd", 0)
                거래소명 = self.GetFieldData("o3105OutBlock", "ExchNm", 0)
                정산구분코드 = self.GetFieldData("o3105OutBlock", "EcCd", 0)
                기준통화코드 = self.GetFieldData("o3105OutBlock", "CrncyCd", 0)
                진법구분코드 = self.GetFieldData("o3105OutBlock", "NotaCd", 0)
                호가단위간격 = self.GetFieldData("o3105OutBlock", "UntPrc", 0)
                최소가격변동금액 = self.GetFieldData("o3105OutBlock", "MnChgAmt", 0)
                가격조정계수 = self.GetFieldData("o3105OutBlock", "RgltFctr", 0)
                계약당금액 = self.GetFieldData("o3105OutBlock", "CtrtPrAmt", 0)
                상장개월수 = self.GetFieldData("o3105OutBlock", "LstngMCnt", 0)
                상품구분코드 = self.GetFieldData("o3105OutBlock", "GdsCd", 0)
                시장구분코드 = self.GetFieldData("o3105OutBlock", "MrktCd", 0)
                Emini구분코드 = self.GetFieldData("o3105OutBlock", "EminiCd", 0)
                상장년 = self.GetFieldData("o3105OutBlock", "LstngYr", 0)
                상장월 = self.GetFieldData("o3105OutBlock", "LstngM", 0)
                월물순서 = self.GetFieldData("o3105OutBlock", "SeqNo", 0)
                상장일자 = self.GetFieldData("o3105OutBlock", "LstngDt", 0)
                만기일자 = self.GetFieldData("o3105OutBlock", "MtrtDt", 0)
                최종거래일 = self.GetFieldData("o3105OutBlock", "FnlDlDt", 0)
                # 최초인도통지일자 = self.GetFieldData("o3105OutBlock", "FstTrsfrDt", 0)
                정산가격 = self.GetFieldData("o3105OutBlock", "EcPrc", 0)
                거래시작일자_한국 = self.GetFieldData("o3105OutBlock", "DlDt", 0)
                거래시작시간_한국 = self.GetFieldData("o3105OutBlock", "DlStrtTm", 0)
                거래종료시간_한국 = self.GetFieldData("o3105OutBlock", "DlEndTm", 0)
                거래시작일자_현지 = self.GetFieldData("o3105OutBlock", "OvsStrDay", 0)
                거래시작시간_현지 = self.GetFieldData("o3105OutBlock", "OvsStrTm", 0)
                거래종료일자_현지 = self.GetFieldData("o3105OutBlock", "OvsEndDay", 0)
                거래종료시간_현지 = self.GetFieldData("o3105OutBlock", "OvsEndTm", 0)
                # 거래가능구분코드 = self.GetFieldData("o3105OutBlock", "DlPsblCd", 0)
                # 증거금징수구분코드 = self.GetFieldData("o3105OutBlock", "MgnCltCd", 0)
                개시증거금 = self.GetFieldData("o3105OutBlock", "OpngMgn", 0)
                유지증거금 = self.GetFieldData("o3105OutBlock", "MntncMgn", 0)
                # 개시증거금율 = self.GetFieldData("o3105OutBlock", "OpngMgnR", 0)
                # 유지증거금율 = self.GetFieldData("o3105OutBlock", "MntncMgnR", 0)
                # 유효소수점자리수 = self.GetFieldData("o3105OutBlock", "DotGb", 0)
                시차 = self.GetFieldData("o3105OutBlock", "TimeDiff", 0)
                현지체결일자 = self.GetFieldData("o3105OutBlock", "OvsDate", 0)
                한국체결일자 = self.GetFieldData("o3105OutBlock", "KorDate", 0)
                한국체결시간 = self.GetFieldData("o3105OutBlock", "TrdTm", 0)
                한국체결시각 = self.GetFieldData("o3105OutBlock", "RcvTm", 0)
                체결가격 = self.GetFieldData("o3105OutBlock", "TrdP", 0)
                체결수량 = self.GetFieldData("o3105OutBlock", "TrdQ", 0)
                누적거래량 = self.GetFieldData("o3105OutBlock", "TotQ", 0)
                # 체결거래대금 = self.GetFieldData("o3105OutBlock", "TrdAmt", 0)
                # 누적거래대금 = self.GetFieldData("o3105OutBlock", "TotAmt", 0)
                # 시가 = self.GetFieldData("o3105OutBlock", "OpenP", 0)
                # 고가 = self.GetFieldData("o3105OutBlock", "HighP", 0)
                # 저가 = self.GetFieldData("o3105OutBlock", "LowP", 0)
                # 전일종가 = self.GetFieldData("o3105OutBlock", "CloseP", 0)
                # 전일대비 = self.GetFieldData("o3105OutBlock", "YdiffP", 0)
                # 전일대비구분 = self.GetFieldData("o3105OutBlock", "YdiffSign", 0)
                # 체결구분 = self.GetFieldData("o3105OutBlock", "Cgubun", 0)
                # 등락율 = self.GetFieldData("o3105OutBlock", "Diff", 0)

                if 종목코드 not in Object.종목정보_딕셔너리.keys():
                    Object.종목정보_딕셔너리.update({종목코드:{}})

                Object.종목정보_딕셔너리[종목코드].update({"종목코드":종목코드})
                Object.종목정보_딕셔너리[종목코드].update({"종목명": 종목명})
                Object.종목정보_딕셔너리[종목코드].update({"종목배치수신일": 종목배치수신일})
                Object.종목정보_딕셔너리[종목코드].update({"기초상품코드": 기초상품코드})
                Object.종목정보_딕셔너리[종목코드].update({"기초상품명": 기초상품명})
                Object.종목정보_딕셔너리[종목코드].update({"거래소코드": 거래소코드})
                Object.종목정보_딕셔너리[종목코드].update({"거래소명": 거래소명})
                Object.종목정보_딕셔너리[종목코드].update({"정산구분코드": 정산구분코드})
                Object.종목정보_딕셔너리[종목코드].update({"기준통화코드": 기준통화코드})
                Object.종목정보_딕셔너리[종목코드].update({"진법구분코드": 진법구분코드})
                Object.종목정보_딕셔너리[종목코드].update({"호가단위간격": float(호가단위간격)})
                Object.종목정보_딕셔너리[종목코드].update({"최소가격변동금액": float(최소가격변동금액)})
                Object.종목정보_딕셔너리[종목코드].update({"가격조정계수": float(가격조정계수)})
                Object.종목정보_딕셔너리[종목코드].update({"계약당금액": float(계약당금액)})
                Object.종목정보_딕셔너리[종목코드].update({"상장개월수": int(상장개월수)})
                Object.종목정보_딕셔너리[종목코드].update({"상품구분코드": 상품구분코드})
                Object.종목정보_딕셔너리[종목코드].update({"시장구분코드": 시장구분코드})
                Object.종목정보_딕셔너리[종목코드].update({"Emini구분코드": Emini구분코드})
                Object.종목정보_딕셔너리[종목코드].update({"상장년": 상장년})
                Object.종목정보_딕셔너리[종목코드].update({"상장월": 상장월})
                Object.종목정보_딕셔너리[종목코드].update({"월물순서": int(월물순서)})
                Object.종목정보_딕셔너리[종목코드].update({"상장일자": 상장일자})
                Object.종목정보_딕셔너리[종목코드].update({"만기일자": 만기일자})
                Object.종목정보_딕셔너리[종목코드].update({"최종거래일": 최종거래일})
                # Object.종목정보_딕셔너리[종목코드].update({"최초인도통지일자": 최초인도통지일자})
                Object.종목정보_딕셔너리[종목코드].update({"정산가격": float(정산가격)})
                Object.종목정보_딕셔너리[종목코드].update({"거래시작일자_한국": 거래시작일자_한국})
                Object.종목정보_딕셔너리[종목코드].update({"거래시작시간_한국": 거래시작시간_한국})
                Object.종목정보_딕셔너리[종목코드].update({"거래종료시간_한국": 거래종료시간_한국})
                Object.종목정보_딕셔너리[종목코드].update({"거래시작일자_현지": 거래시작일자_현지})
                Object.종목정보_딕셔너리[종목코드].update({"거래시작시간_현지": 거래시작시간_현지})
                Object.종목정보_딕셔너리[종목코드].update({"거래종료일자_현지": 거래종료일자_현지})
                Object.종목정보_딕셔너리[종목코드].update({"거래종료시간_현지": 거래종료시간_현지})
                # Object.종목정보_딕셔너리[종목코드].update({"거래가능구분코드": 거래가능구분코드})
                # Object.종목정보_딕셔너리[종목코드].update({"증거금징수구분코드": 증거금징수구분코드})
                Object.종목정보_딕셔너리[종목코드].update({"개시증거금": float(개시증거금)})
                Object.종목정보_딕셔너리[종목코드].update({"유지증거금": float(유지증거금)})
                # Object.종목정보_딕셔너리[종목코드].update({"개시증거금율": float(개시증거금율)})
                # Object.종목정보_딕셔너리[종목코드].update({"유지증거금율": float(유지증거금율)})
                # Object.종목정보_딕셔너리[종목코드].update({"유효소수점자리수": int(유효소수점자리수)})
                Object.종목정보_딕셔너리[종목코드].update({"시차": int(시차)})
                Object.종목정보_딕셔너리[종목코드].update({"현지체결일자": 현지체결일자})
                Object.종목정보_딕셔너리[종목코드].update({"한국체결일자": 한국체결일자})
                Object.종목정보_딕셔너리[종목코드].update({"한국체결시간": 한국체결시간})
                Object.종목정보_딕셔너리[종목코드].update({"한국체결시각": 한국체결시각})
                Object.종목정보_딕셔너리[종목코드].update({"체결가격": float(체결가격)})
                Object.종목정보_딕셔너리[종목코드].update({"체결수량": int(체결수량)})
                Object.종목정보_딕셔너리[종목코드].update({"누적거래량": int(누적거래량)})
                # Object.종목정보_딕셔너리[종목코드].update({"체결거래대금": float(체결거래대금)})
                # Object.종목정보_딕셔너리[종목코드].update({"누적거래대금": float(누적거래대금)})
                # Object.종목정보_딕셔너리[종목코드].update({"시가": float(시가)})
                # Object.종목정보_딕셔너리[종목코드].update({"고가": float(고가)})
                # Object.종목정보_딕셔너리[종목코드].update({"저가": float(저가)})
                # Object.종목정보_딕셔너리[종목코드].update({"전일종가": float(전일종가)})
                # Object.종목정보_딕셔너리[종목코드].update({"전일대비": float(전일대비)})
                # Object.종목정보_딕셔너리[종목코드].update({"전일대비구분": 전일대비구분})
                # Object.종목정보_딕셔너리[종목코드].update({"체결구분": 체결구분})
                # Object.종목정보_딕셔너리[종목코드].update({"등락율": float(등락율)})

                print("\n===== 종목정보 ========"
                                        "\n%s"
                                        "\n%s"
                                        "\n======================"
                                        % (종목코드, Object.종목정보_딕셔너리[종목코드]))

            Object.TR처리완료 = True

        elif szTrCode == "CIDBQ01500":
            print("★★★ 미결제잔고내역 조회")

            occurs_count = self.GetBlockCount("CIDBQ01500OutBlock2")
            for i in range(occurs_count):
                기준일자 = self.GetFieldData("CIDBQ01500OutBlock2", "BaseDt", i) # 20200320
                예수금 = self.GetFieldData("CIDBQ01500OutBlock2", "Dps", i) # 26727 <- 26726.75 : 올림되서 나옴
                청산손익금액 = self.GetFieldData("CIDBQ01500OutBlock2", "LpnlAmt", i)
                선물만기전청산손익금액 = self.GetFieldData("CIDBQ01500OutBlock2", "FutsDueBfLpnlAmt", i)
                선물만기전수수료 = self.GetFieldData("CIDBQ01500OutBlock2", "FutsDueBfCmsn", i)
                위탁증거금액 = self.GetFieldData("CIDBQ01500OutBlock2", "CsgnMgn", i)
                유지증거금 = self.GetFieldData("CIDBQ01500OutBlock2", "MaintMgn", i)
                # 신용한도금액 = self.GetFieldData("CIDBQ01500OutBlock2", "CtlmtAmt", i)
                추가증거금액 = self.GetFieldData("CIDBQ01500OutBlock2", "AddMgn", i)
                마진콜율 = self.GetFieldData("CIDBQ01500OutBlock2", "MgnclRat", i)
                주문가능금액 = self.GetFieldData("CIDBQ01500OutBlock2", "OrdAbleAmt", i)
                인출가능금액 = self.GetFieldData("CIDBQ01500OutBlock2", "WthdwAbleAmt", i)
                계좌번호 = self.GetFieldData("CIDBQ01500OutBlock2", "AcntNo", i)
                종목코드값 = self.GetFieldData("CIDBQ01500OutBlock2", "IsuCodeVal", i)
                종목명 = self.GetFieldData("CIDBQ01500OutBlock2", "IsuNm", i)
                통화코드값 = self.GetFieldData("CIDBQ01500OutBlock2", "CrcyCodeVal", i)
                해외파생상품코드 = self.GetFieldData("CIDBQ01500OutBlock2", "OvrsDrvtPrdtCode", i)
                해외파생옵션구분코드 = self.GetFieldData("CIDBQ01500OutBlock2", "OvrsDrvtOptTpCode", i)
                만기일자 = self.GetFieldData("CIDBQ01500OutBlock2", "DueDt", i)
                # 해외파생행사가격 = self.GetFieldData("CIDBQ01500OutBlock2", "OvrsDrvtXrcPrc", i)
                매매구분코드 = self.GetFieldData("CIDBQ01500OutBlock2", "BnsTpCode", i)
                공통코드명 = self.GetFieldData("CIDBQ01500OutBlock2", "CmnCodeNm", i)
                # 구분코드명 = self.GetFieldData("CIDBQ01500OutBlock2", "TpCodeNm", i)
                잔고수량 = self.GetFieldData("CIDBQ01500OutBlock2", "BalQty", i)
                매입가격 = self.GetFieldData("CIDBQ01500OutBlock2", "PchsPrc", i)
                해외파생현재가 = self.GetFieldData("CIDBQ01500OutBlock2", "OvrsDrvtNowPrc", i)
                해외선물평가손익금액 = self.GetFieldData("CIDBQ01500OutBlock2", "AbrdFutsEvalPnlAmt", i)
                위탁수수료 = self.GetFieldData("CIDBQ01500OutBlock2", "CsgnCmsn", i)
                # 포지션번호 = self.GetFieldData("CIDBQ01500OutBlock2", "PosNo", i)
                # 거래소비용1수수료금액 = self.GetFieldData("CIDBQ01500OutBlock2", "EufOneCmsnAmt", i)
                # 거래소비용2수수료금액 = self.GetFieldData("CIDBQ01500OutBlock2", "EufTwoCmsnAmt", i)

                if 종목코드값 not in Object.미결제_딕셔너리.keys():
                    Object.미결제_딕셔너리.update({종목코드값:{}})

                Object.미결제_딕셔너리[종목코드값].update({"기준일자": 기준일자})
                Object.미결제_딕셔너리[종목코드값].update({"예수금": int(예수금)})
                Object.미결제_딕셔너리[종목코드값].update({"청산손익금액": float(청산손익금액)})
                Object.미결제_딕셔너리[종목코드값].update({"선물만기전청산손익금액": float(선물만기전청산손익금액)})
                Object.미결제_딕셔너리[종목코드값].update({"선물만기전수수료": float(선물만기전수수료)})
                Object.미결제_딕셔너리[종목코드값].update({"위탁증거금액": int(위탁증거금액)})
                Object.미결제_딕셔너리[종목코드값].update({"유지증거금": int(유지증거금)})
                # Object.미결제_딕셔너리[종목코드값].update({"신용한도금액": int(신용한도금액)})
                Object.미결제_딕셔너리[종목코드값].update({"추가증거금액": int(추가증거금액)})
                Object.미결제_딕셔너리[종목코드값].update({"마진콜율": float(마진콜율)})
                Object.미결제_딕셔너리[종목코드값].update({"주문가능금액": int(주문가능금액)})
                Object.미결제_딕셔너리[종목코드값].update({"인출가능금액": int(인출가능금액)})
                Object.미결제_딕셔너리[종목코드값].update({"계좌번호": 계좌번호})
                Object.미결제_딕셔너리[종목코드값].update({"종목코드값": 종목코드값})
                Object.미결제_딕셔너리[종목코드값].update({"종목명": 종목명})
                Object.미결제_딕셔너리[종목코드값].update({"통화코드값": 통화코드값})
                Object.미결제_딕셔너리[종목코드값].update({"해외파생상품코드": 해외파생상품코드})
                Object.미결제_딕셔너리[종목코드값].update({"해외파생옵션구분코드": 해외파생옵션구분코드})
                Object.미결제_딕셔너리[종목코드값].update({"만기일자": 만기일자})
                # Object.미결제_딕셔너리[종목코드값].update({"해외파생행사가격": 해외파생행사가격})
                Object.미결제_딕셔너리[종목코드값].update({"매매구분코드": 매매구분코드})
                Object.미결제_딕셔너리[종목코드값].update({"공통코드명": 공통코드명})
                # Object.미결제_딕셔너리[종목코드값].update({"구분코드명": 구분코드명})
                Object.미결제_딕셔너리[종목코드값].update({"잔고수량": int(잔고수량)})
                Object.미결제_딕셔너리[종목코드값].update({"매입가격": float(매입가격)})
                Object.미결제_딕셔너리[종목코드값].update({"해외파생현재가": float(해외파생현재가)})
                Object.미결제_딕셔너리[종목코드값].update({"해외선물평가손익금액": float(해외선물평가손익금액)})
                Object.미결제_딕셔너리[종목코드값].update({"위탁수수료": float(위탁수수료)})
                # Object.미결제_딕셔너리[종목코드값].update({"포지션번호": 포지션번호})
                # Object.미결제_딕셔너리[종목코드값].update({"거래소비용1수수료금액": 거래소비용1수수료금액})
                # Object.미결제_딕셔너리[종목코드값].update({"거래소비용2수수료금액": 거래소비용2수수료금액})

                print("\n====== 미결제 ======="
                                        "\n%s"
                                        "\n%s"
                                        "\n======================"
                                        % (종목코드값, Object.미결제_딕셔너리[종목코드값]))

            # 데이터가 더 존재하면 다시 조회한다.
            if Object.XAQuery_CIDBQ01500.IsNext:
                Object.tr_signal_CIDBQ01500(IsNext=True)
            else:
                Object.TR처리완료 = True

        elif szTrCode == "CIDBQ03000":
            # 통화마다 보유한 예수금 조회
            print("★★★ 해외선물 예수금/잔고현황")

            occurs_count = self.GetBlockCount("CIDBQ03000OutBlock2")
            for i in range(occurs_count):
                계좌번호 = self.GetFieldData("CIDBQ03000OutBlock2", "AcntNo", i)
                거래일자 = self.GetFieldData("CIDBQ03000OutBlock2", "TrdDt", i)
                통화대상코드 = self.GetFieldData("CIDBQ03000OutBlock2", "CrcyObjCode", i)
                해외선물예수금 = self.GetFieldData("CIDBQ03000OutBlock2", "OvrsFutsDps", i)
                # 고객입출금금액 = self.GetFieldData("CIDBQ03000OutBlock2", "CustmMnyioAmt", i)
                해외선물청산손익금액 = self.GetFieldData("CIDBQ03000OutBlock2", "AbrdFutsLqdtPnlAmt", i)  # 실시간 계산
                해외선물수수료금액 = self.GetFieldData("CIDBQ03000OutBlock2", "AbrdFutsCmsnAmt", i)  # 실시간 계산
                # 가환전예수금 = self.GetFieldData("CIDBQ03000OutBlock2", "PrexchDps", i)
                평가자산금액 = self.GetFieldData("CIDBQ03000OutBlock2", "EvalAssetAmt", i)
                해외선물위탁증거금액 = self.GetFieldData("CIDBQ03000OutBlock2", "AbrdFutsCsgnMgn", i) # 실시간 계산
                # 해외선물추가증거금액 = self.GetFieldData("CIDBQ03000OutBlock2", "AbrdFutsAddMgn", i)
                # 해외선물인출가능금액 = self.GetFieldData("CIDBQ03000OutBlock2", "AbrdFutsWthdwAbleAmt", i)
                해외선물주문가능금액 = self.GetFieldData("CIDBQ03000OutBlock2", "AbrdFutsOrdAbleAmt", i) # 실시간 계산, but 환율정보를 받을 수 없기에 수치가 부정확
                해외선물평가손익금액 = self.GetFieldData("CIDBQ03000OutBlock2", "AbrdFutsEvalPnlAmt", i) # 실시간 계산
                # 최종결제손익금액 = self.GetFieldData("CIDBQ03000OutBlock2", "LastSettPnlAmt", i)
                # 해외옵션결제금액 = self.GetFieldData("CIDBQ03000OutBlock2", "OvrsOptSettAmt", i)
                # 해외옵션잔고평가금액 = self.GetFieldData("CIDBQ03000OutBlock2", "OvrsOptBalEvalAmt", i)

                if 통화대상코드 not in Object.예수금_딕셔너리.keys():
                    Object.예수금_딕셔너리.update({통화대상코드 : {}})

                Object.예수금_딕셔너리[통화대상코드].update({"계좌번호": 계좌번호})
                Object.예수금_딕셔너리[통화대상코드].update({"거래일자": 거래일자})
                Object.예수금_딕셔너리[통화대상코드].update({"통화대상코드": 통화대상코드})
                Object.예수금_딕셔너리[통화대상코드].update({"해외선물예수금": float(해외선물예수금)})
                # Object.예수금_딕셔너리[통화대상코드].update({"고객입출금금액": 고객입출금금액})
                Object.예수금_딕셔너리[통화대상코드].update({"해외선물청산손익금액": float(해외선물청산손익금액)})
                Object.예수금_딕셔너리[통화대상코드].update({"해외선물수수료금액": float(해외선물수수료금액)})
                # Object.예수금_딕셔너리[통화대상코드].update({"가환전예수금": float(가환전예수금)})
                Object.예수금_딕셔너리[통화대상코드].update({"평가자산금액": 평가자산금액})
                Object.예수금_딕셔너리[통화대상코드].update({"해외선물위탁증거금액": float(해외선물위탁증거금액)})
                # Object.예수금_딕셔너리[통화대상코드].update({"해외선물추가증거금액": float(해외선물추가증거금액)})
                # Object.예수금_딕셔너리[통화대상코드].update({"해외선물인출가능금액": 해외선물인출가능금액})
                Object.예수금_딕셔너리[통화대상코드].update({"해외선물주문가능금액": float(해외선물주문가능금액)})
                Object.예수금_딕셔너리[통화대상코드].update({"해외선물평가손익금액": float(해외선물평가손익금액)})
                # Object.예수금_딕셔너리[통화대상코드].update({"최종결제손익금액": 최종결제손익금액})
                # Object.예수금_딕셔너리[통화대상코드].update({"해외옵션결제금액": 해외옵션결제금액})
                # Object.예수금_딕셔너리[통화대상코드].update({"해외옵션잔고평가금액": 해외옵션잔고평가금액})
                Object.예수금_딕셔너리[통화대상코드].update({"해외선물유지증거금액": 0.0})  # 임시로 만듬

                print("\n===== 예수금 ========"
                                        "\n%s"
                                        "\n%s"
                                        "\n======================"
                                        % (통화대상코드, Object.예수금_딕셔너리[통화대상코드]))

            if Object.XAQuery_CIDBQ03000.IsNext:
                Object.tr_signal_CIDBQ03000(IsNext=True)
            else:
                Object.TR처리완료 = True

    def OnReceiveMessage(self, systemError, messageCode, message):
        '''
        02662 - 모의투자 정정 가능한 수량을 초과하였습니다.
        02258 - 모의투자 원주문번호를 잘못 입력하셨습니다.
        01231 - 주문가격은 현재가보다 커야합니다.
        03576 - 모의투자 증거금 한도 초과로 주문 불가합니다.
        :param systemError:
        :param messageCode:
        :param message:
        :return:
        '''

        if messageCode != "00000":
            if Object.취소 == True:
                Object.취소 = False

            if len(Object.주문접수_딕셔너리) == 0 and Object.매수 == True:
                Object.매수 = False

            if Object.정정 == True:
                Object.정정 = False

            elif Object.정정 == False:
                Object.매도 = False


        print("★★★ systemError: %s, messageCode: %s, message: %s" % (systemError, messageCode, message))

'''
실시간으로 데이터 받게 정의해주는 이벤트
'''
class XARealEvent(Object):

    def OnReceiveRealData(self, trCode):

        if trCode == "OVH":

            종목코드 = self.GetFieldData("OutBlock", "symbol")
            호가시간 = self.GetFieldData("OutBlock", "hotime")
            매도호가1 = self.GetFieldData("OutBlock", "offerho1")
            매수호가1 = self.GetFieldData("OutBlock", "bidho1")
            매도호가잔량1 = self.GetFieldData("OutBlock", "offerrem1")
            매수호가잔량1 = self.GetFieldData("OutBlock", "bidrem1")
            매도호가건수1 = self.GetFieldData("OutBlock", "offerno1")
            매수호가건수1 = self.GetFieldData("OutBlock", "bidno1")
            매도호가2 = self.GetFieldData("OutBlock", "offerho2")
            매수호가2 = self.GetFieldData("OutBlock", "bidho2")
            매도호가잔량2 = self.GetFieldData("OutBlock", "offerrem2")
            매수호가잔량2 = self.GetFieldData("OutBlock", "bidrem2")
            매도호가건수2 = self.GetFieldData("OutBlock", "offerno2")
            매수호가건수2 = self.GetFieldData("OutBlock", "bidno2")
            매도호가3 = self.GetFieldData("OutBlock", "offerho3")
            매수호가3 = self.GetFieldData("OutBlock", "bidho3")
            매도호가잔량3 = self.GetFieldData("OutBlock", "offerrem3")
            매수호가잔량3 = self.GetFieldData("OutBlock", "bidrem3")
            매도호가건수3 = self.GetFieldData("OutBlock", "offerno3")
            매수호가건수3 = self.GetFieldData("OutBlock", "bidno3")
            매도호가4 = self.GetFieldData("OutBlock", "offerho4")
            매수호가4 = self.GetFieldData("OutBlock", "bidho4")
            매도호가잔량4 = self.GetFieldData("OutBlock", "offerrem4")
            매수호가잔량4 = self.GetFieldData("OutBlock", "bidrem4")
            매도호가건수4 = self.GetFieldData("OutBlock", "offerno4")
            매수호가건수4 = self.GetFieldData("OutBlock", "bidno4")
            매도호가5 = self.GetFieldData("OutBlock", "offerho5")
            매수호가5 = self.GetFieldData("OutBlock", "bidho5")
            매도호가잔량5 = self.GetFieldData("OutBlock", "offerrem5")
            매수호가잔량5 = self.GetFieldData("OutBlock", "bidrem5")
            매도호가건수5 = self.GetFieldData("OutBlock", "offerno5")
            매수호가건수5 = self.GetFieldData("OutBlock", "bidno5")
            매도호가총건수 = self.GetFieldData("OutBlock", "totoffercnt")
            매수호가총건수 = self.GetFieldData("OutBlock", "totbidcnt")
            매도호가총수량 = self.GetFieldData("OutBlock", "totbidcnt")
            매수호가총수량 = self.GetFieldData("OutBlock", "totbidrem")

            if 종목코드 not in Object.실시간호가_딕셔너리.keys():
                Object.실시간호가_딕셔너리.update({종목코드: {}})

            Object.실시간호가_딕셔너리[종목코드].update({"종목코드": 종목코드})
            Object.실시간호가_딕셔너리[종목코드].update({"호가시간": 호가시간})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가1": float(매도호가1)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가1": float(매수호가1)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가잔량1": int(매도호가잔량1)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가잔량1": int(매수호가잔량1)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가건수1": int(매도호가건수1)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가건수1": int(매수호가건수1)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가2": float(매도호가2)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가2": float(매수호가2)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가잔량2": int(매도호가잔량2)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가잔량2": int(매수호가잔량2)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가건수2": int(매도호가건수2)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가건수2": int(매수호가건수2)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가3": float(매도호가3)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가3": float(매수호가3)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가잔량3": int(매도호가잔량3)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가잔량3": int(매수호가잔량3)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가건수3": int(매도호가건수3)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가건수3": int(매수호가건수3)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가4": float(매도호가4)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가4": float(매수호가4)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가잔량4": int(매도호가잔량4)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가잔량4": int(매수호가잔량4)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가건수4": int(매도호가건수4)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가건수4": int(매수호가건수4)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가5": float(매도호가5)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가5": float(매수호가5)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가잔량5": int(매도호가잔량5)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가잔량5": int(매수호가잔량5)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가건수5": int(매도호가건수5)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가건수5": int(매수호가건수5)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가총건수": int(매도호가총건수)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가총건수": int(매수호가총건수)})
            Object.실시간호가_딕셔너리[종목코드].update({"매도호가총수량": int(매도호가총수량)})
            Object.실시간호가_딕셔너리[종목코드].update({"매수호가총수량": int(매수호가총수량)})

        elif trCode == "OVC":

            종목코드 = self.GetFieldData("OutBlock", "symbol")
            체결일자_현지 = self.GetFieldData("OutBlock", "ovsdate")
            체결일자_한국 = self.GetFieldData("OutBlock", "kordate")
            체결시간_현지 = self.GetFieldData("OutBlock", "trdtm")
            체결시간_한국 = self.GetFieldData("OutBlock", "kortm")
            체결가격 = self.GetFieldData("OutBlock", "curpr")
            전일대비 = self.GetFieldData("OutBlock", "ydiffpr")
            전일대비기호 = self.GetFieldData("OutBlock", "ydiffSign")
            시가 = self.GetFieldData("OutBlock", "open")
            고가 = self.GetFieldData("OutBlock", "high")
            저가 = self.GetFieldData("OutBlock", "low")
            등락율 = self.GetFieldData("OutBlock", "chgrate")
            건별체결수량 = self.GetFieldData("OutBlock", "trdq")
            누적체결수량 = self.GetFieldData("OutBlock", "totq")
            체결구분 = self.GetFieldData("OutBlock", "cgubun")
            매도누적체결수량 = self.GetFieldData("OutBlock", "mdvolume")
            매수누적체결수량 = self.GetFieldData("OutBlock", "msvolume")
            장마감일 = self.GetFieldData("OutBlock", "ovsmkend")

            if 종목코드 not in Object.실시간체결_딕셔너리.keys():
                Object.실시간체결_딕셔너리.update({종목코드 : {}})

            Object.실시간체결_딕셔너리[종목코드].update({"종목코드": 종목코드})
            Object.실시간체결_딕셔너리[종목코드].update({"체결일자_현지": 체결일자_현지})
            Object.실시간체결_딕셔너리[종목코드].update({"체결일자_한국": 체결일자_한국})
            Object.실시간체결_딕셔너리[종목코드].update({"체결시간_현지": 체결시간_현지})
            Object.실시간체결_딕셔너리[종목코드].update({"체결시간_한국": 체결시간_한국})
            Object.실시간체결_딕셔너리[종목코드].update({"체결가격": float(체결가격)})
            Object.실시간체결_딕셔너리[종목코드].update({"전일대비": float(전일대비)})
            Object.실시간체결_딕셔너리[종목코드].update({"전일대비기호": 전일대비기호})
            Object.실시간체결_딕셔너리[종목코드].update({"시가": float(시가)})
            Object.실시간체결_딕셔너리[종목코드].update({"고가": float(고가)})
            Object.실시간체결_딕셔너리[종목코드].update({"저가": float(저가)})
            Object.실시간체결_딕셔너리[종목코드].update({"등락율": float(등락율)})
            Object.실시간체결_딕셔너리[종목코드].update({"건별체결수량": int(건별체결수량)})
            Object.실시간체결_딕셔너리[종목코드].update({"누적체결수량": int(누적체결수량)})
            Object.실시간체결_딕셔너리[종목코드].update({"체결구분": 체결구분})
            Object.실시간체결_딕셔너리[종목코드].update({"매도누적체결수량": 매도누적체결수량})
            Object.실시간체결_딕셔너리[종목코드].update({"매수누적체결수량": 매수누적체결수량})
            Object.실시간체결_딕셔너리[종목코드].update({"장마감일": 장마감일})

            print(Object.실시간체결_딕셔너리)

            if 종목코드 in Object.실시간호가_딕셔너리:
                self.condition(종목코드=종목코드, 체결일자_한국=체결일자_한국)

    '''
    매수/매도 조건 계산
    '''
    '''
    매수/매도 조건 계산
    '''
    def condition(self, 종목코드=None, 체결일자_한국=None):

        호_딕 = Object.실시간호가_딕셔너리[종목코드]
        종_딕 = Object.종목정보_딕셔너리[종목코드]

        매도호가1 = 호_딕["매도호가1"]
        매수호가1 = 호_딕["매수호가1"]
        호가단위간격 = 종_딕["호가단위간격"]

        # 주문완료인지 확인
        if 종목코드 in self.미결제_딕셔너리:

            잔고수량 = self.미결제_딕셔너리[종목코드]["잔고수량"]
            매입가격 = self.미결제_딕셔너리[종목코드]["매입가격"]
            가격 = 매입가격 + (호가단위간격 * 3)

            if Object.매도 == False:
                Object.매도 = True

                Object.order_buy_CIDBT00100(
                    레코드갯수=1,
                    주문일자=체결일자_한국,
                    지점코드="",
                    계좌번호=Object.해외선물_계좌번호,
                    비밀번호='0000',
                    종목코드값=종목코드,
                    선물주문구분코드="1",
                    매매구분코드="1",
                    해외선물주문유형코드="2",
                    통화코드="",
                    해외파생주문가격=가격,
                    조건주문가격=0,
                    주문수량=잔고수량,
                    상품코드="",
                    만기년월="",
                    거래소코드=""
                )

        elif 종목코드 not in self.미결제_딕셔너리:

            if Object.매수 == False:

                interval = (매도호가1 - 매수호가1)
                compareTo = (호가단위간격 * 4)  # 4
                if interval >= compareTo:
                    Object.매수 = True

                    Object.order_buy_CIDBT00100(
                        레코드갯수=1,
                        주문일자=체결일자_한국,
                        지점코드="",
                        계좌번호=Object.해외선물_계좌번호,
                        비밀번호='0000',
                        종목코드값=종목코드,
                        선물주문구분코드="1",
                        매매구분코드="2",
                        해외선물주문유형코드="2",
                        통화코드="",
                        해외파생주문가격=매수호가1,
                        조건주문가격=0,
                        주문수량=1,
                        상품코드="",
                        만기년월="",
                        거래소코드=""
                    )

        # 주문 취소와 관련된 부분.
        주문접수번호_리스트 = list(Object.주문접수_딕셔너리)
        for 주문번호 in 주문접수번호_리스트:
            접_딕 = Object.주문접수_딕셔너리[주문번호]
            정정취소유형 = 접_딕['정정취소유형']
            매도매수유형 = 접_딕['매도매수유형']
            주문가격 = 접_딕["주문가격"]
            주문수량 = 접_딕["주문수량"]

            # api에서 주문 취소를 할때 주문 번호를 10자리를 맞춰줘야 한다.
            번호 = 주문번호
            while True:
                order_line = len(번호)
                if order_line == 10:
                    break
                else :
                    번호 = '0' + 번호

            if 매도매수유형 == "1" and len(Object.미결제_딕셔너리) == 0:
                # 딕셔너리에 키값이 많이 쌓이므로 불필요한 것은 바로바로 삭제해준다
                del Object.주문접수_딕셔너리[주문번호]

            # 지금은 단방향 거래만 생각하고 있다./ 주문후 체결이 안되고, 내가 원하는 방향과 반대로 움직일때, 주문 취소를 해야한다.
            if 접_딕["종목코드"] == 종목코드 and 매도매수유형 == "2" and 정정취소유형 not in ["2", "3"] \
                and Object.매도 == False and Object.매수 == True and Object.취소 == False and Object.정정 == False:

                check = 매수호가1 - 주문가격
                compare = 호가단위간격 * 6
                reverse_check = 주문가격 - 매수호가1
                reverse_compare = 호가단위간격 * 6

                if (check > compare or reverse_check > reverse_compare):
                    Object.취소 = True

                    Object.order_cancel_CIDBT01000(
                        레코드갯수=1,
                        주문일자=체결일자_한국,
                        지점번호="",
                        계좌번호=Object.해외선물_계좌번호,
                        비밀번호="0000",
                        종목코드값=종목코드,
                        해외선물원주문번호=번호,
                        선물주문구분코드="3",
                        상품구분코드="",
                        거래소코드=""
                    )

            # 정정하는 부분
            elif 접_딕["종목코드"] == 종목코드 and 매도매수유형 == "1" and 정정취소유형 != "3" \
                and Object.매도 == True and Object.매수 == False and Object.정정 == False and Object.취소 == False:

                check = 주문가격 - 매도호가1
                compare = 호가단위간격 * 6
                reverse_check = 매도호가1 - 주문가격
                reverse_compare = 호가단위간격 * 6

                if (check > compare or reverse_check > reverse_compare):
                    Object.정정 = True

                    Object.order_cancel_CIDBT00900(
                        레코드갯수=1,
                        주문일자=체결일자_한국,
                        등록지점번호="",
                        계좌번호=Object.해외선물_계좌번호,
                        비밀번호="0000",
                        해외선물원주문번호=번호,
                        종목코드값=종목코드,
                        선물주문구분코드="2",
                        매매구분코드=매도매수유형,
                        선물주문유형코드="2",
                        통화코드값="",
                        해외파생주문가격=매도호가1,
                        조건주문가격=0,
                        주문수량=주문수량,
                        해외파생상품코드="",
                        만기년월="",
                        거래소코드=""
                    )


'''
실시간으로 주문 정보 정의해주는 이벤트
'''
class XARealOrderEvent(Object):
    # 실시간으로 체결되면 정보를 받아오는 부분
    def OnReceiveRealData(self, trCode):

        if trCode == "TC1":
            라인일련번호 = self.GetFieldData("OutBlock", "lineseq")
            key = self.GetFieldData("OutBlock", "key")
            조직자ID = self.GetFieldData("OutBlock", "user") # 내 아이디
            서비스ID = self.GetFieldData("OutBlock", "svc_id") # 거래소에서 던져줌
            주문일자 = self.GetFieldData("OutBlock", "ordr_dt")
            지점번호 = self.GetFieldData("OutBlock", "brn_cd")
            주문번호 = self.GetFieldData("OutBlock", "ordr_no")
            원주문번호 = self.GetFieldData("OutBlock", "orgn_ordr_no")
            모주문번호 = self.GetFieldData("OutBlock", "mthr_ordr_no")
            계좌번호 = self.GetFieldData("OutBlock", "ac_no")
            종목코드 = self.GetFieldData("OutBlock", "is_cd")
            매도매수유형 = self.GetFieldData("OutBlock", "s_b_ccd") #1:매도, 2:매수
            정정취소유형 = self.GetFieldData("OutBlock", "ordr_ccd") #1:신규, 2:정정, 3:취소
            주문유형코드 = self.GetFieldData("OutBlock", "ordr_typ_cd") #1:시장가, 2:지정가, 3:Stop Market, 4: Stop limit
            주문기간코드 = self.GetFieldData("OutBlock", "ordr_typ_prd_ccd") #01:일반, 02:Average, 03:Spread
            주문적용시작일자 = self.GetFieldData("OutBlock", "ordr_aplc_strt_dt")
            주문적용종료일자 = self.GetFieldData("OutBlock", "ordr_aplc_end_dt")
            주문가격 = self.GetFieldData("OutBlock", "ordr_prc")
            주문조건가격 = self.GetFieldData("OutBlock", "cndt_ordr_prc")
            주문수량 = self.GetFieldData("OutBlock", "ordr_q")
            주문시간 = self.GetFieldData("OutBlock", "ordr_tm")
            사용자ID = self.GetFieldData("OutBlock", "userid")
            만기행사유무 = self.GetFieldData("OutBlock", "xrc_rsv_tcp_code")

            if 주문번호 not in Object.주문접수_딕셔너리.keys():
                Object.주문접수_딕셔너리.update({주문번호 : {}})

            Object.주문접수_딕셔너리[주문번호].update({"라인일련번호": 라인일련번호})
            Object.주문접수_딕셔너리[주문번호].update({"key": key})
            Object.주문접수_딕셔너리[주문번호].update({"조직자ID": 조직자ID})
            Object.주문접수_딕셔너리[주문번호].update({"서비스ID": 서비스ID})
            Object.주문접수_딕셔너리[주문번호].update({"주문일자": 주문일자})
            Object.주문접수_딕셔너리[주문번호].update({"지점번호": 지점번호})
            Object.주문접수_딕셔너리[주문번호].update({"주문번호": 주문번호})
            Object.주문접수_딕셔너리[주문번호].update({"원주문번호": 원주문번호})
            Object.주문접수_딕셔너리[주문번호].update({"모주문번호": 모주문번호})
            Object.주문접수_딕셔너리[주문번호].update({"계좌번호": 계좌번호})
            Object.주문접수_딕셔너리[주문번호].update({"종목코드": 종목코드})
            Object.주문접수_딕셔너리[주문번호].update({"매도매수유형": 매도매수유형})
            Object.주문접수_딕셔너리[주문번호].update({"정정취소유형": 정정취소유형})
            Object.주문접수_딕셔너리[주문번호].update({"주문유형코드": 주문유형코드})
            Object.주문접수_딕셔너리[주문번호].update({"주문기간코드": 주문기간코드})
            Object.주문접수_딕셔너리[주문번호].update({"주문적용시작일자": 주문적용시작일자})
            Object.주문접수_딕셔너리[주문번호].update({"주문적용종료일자": 주문적용종료일자})
            Object.주문접수_딕셔너리[주문번호].update({"주문가격": float(주문가격)})
            Object.주문접수_딕셔너리[주문번호].update({"주문조건가격": float(주문조건가격)})
            Object.주문접수_딕셔너리[주문번호].update({"주문수량": int(주문수량)})
            Object.주문접수_딕셔너리[주문번호].update({"주문시간": 주문시간})
            Object.주문접수_딕셔너리[주문번호].update({"사용자ID": 사용자ID})
            Object.주문접수_딕셔너리[주문번호].update({"만기행사유무": 만기행사유무})

            print("\n===== 주문접수 ========"
                                    "\n%s"
                                    "\n%s"
                                    "\n%s"
                                    "\n======================"
                                    % (Object.매수, 주문번호, Object.주문접수_딕셔너리[주문번호]))

        elif trCode == "TC2":
            라인일련번호 = self.GetFieldData("OutBlock", "lineseq")
            key = self.GetFieldData("OutBlock", "key")
            조직자ID = self.GetFieldData("OutBlock", "user")
            서비스ID = self.GetFieldData("OutBlock", "svc_id")
            주문일자 = self.GetFieldData("OutBlock", "ordr_dt")
            지점번호 = self.GetFieldData("OutBlock", "brn_cd")
            주문번호 = self.GetFieldData("OutBlock", "ordr_no")
            원주문번호 = self.GetFieldData("OutBlock", "orgn_ordr_no")
            모주문번호 = self.GetFieldData("OutBlock", "mthr_ordr_no")
            계좌번호 = self.GetFieldData("OutBlock", "ac_no")
            종목코드 = self.GetFieldData("OutBlock", "is_cd")
            매도매수유형 = self.GetFieldData("OutBlock", "s_b_ccd")
            정정취소유형 = self.GetFieldData("OutBlock", "ordr_ccd")
            주문유형코드 = self.GetFieldData("OutBlock", "ordr_typ_cd")
            주문기간코드 = self.GetFieldData("OutBlock", "ordr_typ_prd_ccd")
            주문적용시작일자 = self.GetFieldData("OutBlock", "ordr_aplc_strt_dt")
            주문적용종료일자 = self.GetFieldData("OutBlock", "ordr_aplc_end_dt")
            주문가격 = self.GetFieldData("OutBlock", "ordr_prc")
            주문조건가격 = self.GetFieldData("OutBlock", "cndt_ordr_prc")
            주문수량 = self.GetFieldData("OutBlock", "ordr_q")
            주문시간 = self.GetFieldData("OutBlock", "ordr_tm")
            호가확인수량 = self.GetFieldData("OutBlock", "cnfr_q")
            호가거부사유코드 = self.GetFieldData("OutBlock", "rfsl_cd")
            호가거부사유코드명 = self.GetFieldData("OutBlock", "text")
            사용자ID = self.GetFieldData("OutBlock", "user_id")

            if 주문번호 not in Object.주문응답_딕셔너리.keys():
                Object.주문응답_딕셔너리.update({주문번호 : {}})

            Object.주문응답_딕셔너리[주문번호].update({"라인일련번호": 라인일련번호})
            Object.주문응답_딕셔너리[주문번호].update({"key": key})
            Object.주문응답_딕셔너리[주문번호].update({"조직자ID": 조직자ID})
            Object.주문응답_딕셔너리[주문번호].update({"서비스ID": 서비스ID})
            Object.주문응답_딕셔너리[주문번호].update({"주문일자": 주문일자})
            Object.주문응답_딕셔너리[주문번호].update({"지점번호": 지점번호})
            Object.주문응답_딕셔너리[주문번호].update({"주문번호": 주문번호})
            Object.주문응답_딕셔너리[주문번호].update({"원주문번호": 원주문번호})
            Object.주문응답_딕셔너리[주문번호].update({"모주문번호": 모주문번호})
            Object.주문응답_딕셔너리[주문번호].update({"계좌번호": 계좌번호})
            Object.주문응답_딕셔너리[주문번호].update({"종목코드": 종목코드})
            Object.주문응답_딕셔너리[주문번호].update({"매도매수유형": 매도매수유형})
            Object.주문응답_딕셔너리[주문번호].update({"정정취소유형": 정정취소유형})
            Object.주문응답_딕셔너리[주문번호].update({"주문유형코드": 주문유형코드})
            Object.주문응답_딕셔너리[주문번호].update({"주문기간코드": 주문기간코드})
            Object.주문응답_딕셔너리[주문번호].update({"주문적용시작일자": 주문적용시작일자})
            Object.주문응답_딕셔너리[주문번호].update({"주문적용종료일자": 주문적용종료일자})
            Object.주문응답_딕셔너리[주문번호].update({"주문가격": float(주문가격)})
            Object.주문응답_딕셔너리[주문번호].update({"주문조건가격": float(주문조건가격)})
            Object.주문응답_딕셔너리[주문번호].update({"주문수량": int(주문수량)})
            Object.주문응답_딕셔너리[주문번호].update({"주문시간": 주문시간})
            Object.주문응답_딕셔너리[주문번호].update({"호가확인수량": int(호가확인수량)})
            Object.주문응답_딕셔너리[주문번호].update({"호가거부사유코드": 호가거부사유코드})
            Object.주문응답_딕셔너리[주문번호].update({"호가거부사유코드명": 호가거부사유코드명})
            Object.주문응답_딕셔너리[주문번호].update({"사용자ID": 사용자ID})

            print("\n===== 주문응답 ========"
                                    "\n%s"
                                    "\n%s"
                                    "\n======================"
                                    % (주문번호, Object.주문응답_딕셔너리[주문번호]))

            ## 오류 처리 부분
            if 원주문번호 in Object.주문접수_딕셔너리:
                del Object.주문접수_딕셔너리[원주문번호]
            del Object.주문응답_딕셔너리[주문번호]

            if 정정취소유형 == "2":
                Object.정정 = False

            elif 정정취소유형 == "3":
                if 주문번호 in Object.주문접수_딕셔너리:
                    del Object.주문접수_딕셔너리[주문번호]

                Object.매수 = False
                Object.취소 = False

        elif trCode == "TC3":
            라인일련번호 = self.GetFieldData("OutBlock", "lineseq")
            key = self.GetFieldData("OutBlock", "key")
            조직자ID = self.GetFieldData("OutBlock", "user")
            서비스ID = self.GetFieldData("OutBlock", "svc_id")
            주문일자 = self.GetFieldData("OutBlock", "ordr_dt")
            지점번호 = self.GetFieldData("OutBlock", "brn_cd")
            주문번호 = self.GetFieldData("OutBlock", "ordr_no")
            원주문번호 = self.GetFieldData("OutBlock", "orgn_ordr_no")
            모주문번호 = self.GetFieldData("OutBlock", "mthr_ordr_no")
            계좌번호 = self.GetFieldData("OutBlock", "ac_no")
            종목코드 = self.GetFieldData("OutBlock", "is_cd")
            매도매수유형 = self.GetFieldData("OutBlock", "s_b_ccd")
            정정취소유형 = self.GetFieldData("OutBlock", "ordr_ccd")
            체결수량 = self.GetFieldData("OutBlock", "ccls_q")
            체결가격 = self.GetFieldData("OutBlock", "ccls_prc")
            체결번호 = self.GetFieldData("OutBlock", "ccls_no")
            체결시간 = self.GetFieldData("OutBlock", "ccls_tm")
            매입평균단가 = self.GetFieldData("OutBlock", "avg_byng_uprc")
            매입금액 = self.GetFieldData("OutBlock", "byug_amt")
            청산손익 = self.GetFieldData("OutBlock", "clr_pl_amt")
            위탁수수료 = self.GetFieldData("OutBlock", "ent_fee")  #2개 체결시 2배로 나옴
            # FCM수수료 = self.GetFieldData("OutBlock", "fcm_fee")
            사용자ID = self.GetFieldData("OutBlock", "userid")
            # 현재가격 = self.GetFieldData("OutBlock", "now_prc")
            통화코드 = self.GetFieldData("OutBlock", "crncy_cd")
            만기일자 = self.GetFieldData("OutBlock", "mtrt_dt")

            if 주문번호 not in Object.주문체결_딕셔너리.keys():
                Object.주문체결_딕셔너리.update({주문번호 : {}})

            Object.주문체결_딕셔너리[주문번호].update({"라인일련번호": 라인일련번호})
            Object.주문체결_딕셔너리[주문번호].update({"key": key})
            Object.주문체결_딕셔너리[주문번호].update({"조직자ID": 조직자ID})
            Object.주문체결_딕셔너리[주문번호].update({"서비스ID": 서비스ID})
            Object.주문체결_딕셔너리[주문번호].update({"주문일자": 주문일자})
            Object.주문체결_딕셔너리[주문번호].update({"지점번호": 지점번호})
            Object.주문체결_딕셔너리[주문번호].update({"주문번호": 주문번호})
            Object.주문체결_딕셔너리[주문번호].update({"원주문번호": 원주문번호})
            Object.주문체결_딕셔너리[주문번호].update({"모주문번호": 모주문번호})
            Object.주문체결_딕셔너리[주문번호].update({"계좌번호": 계좌번호})
            Object.주문체결_딕셔너리[주문번호].update({"종목코드": 종목코드})
            Object.주문체결_딕셔너리[주문번호].update({"매도매수유형": 매도매수유형})
            Object.주문체결_딕셔너리[주문번호].update({"정정취소유형": 정정취소유형})
            Object.주문체결_딕셔너리[주문번호].update({"체결수량": int(체결수량)})
            Object.주문체결_딕셔너리[주문번호].update({"체결가격": float(체결가격)})
            Object.주문체결_딕셔너리[주문번호].update({"체결번호": 체결번호})
            Object.주문체결_딕셔너리[주문번호].update({"체결시간": 체결시간})
            Object.주문체결_딕셔너리[주문번호].update({"매입평균단가": float(매입평균단가)})
            Object.주문체결_딕셔너리[주문번호].update({"매입금액": float(매입금액)})
            Object.주문체결_딕셔너리[주문번호].update({"청산손익": float(청산손익)})
            Object.주문체결_딕셔너리[주문번호].update({"위탁수수료": float(위탁수수료)})
            # Object.주문체결_딕셔너리[주문번호].update({"FCM수수료": float(FCM수수료)})
            Object.주문체결_딕셔너리[주문번호].update({"사용자ID": 사용자ID})
            # Object.주문체결_딕셔너리[주문번호].update({"현재가격": 현재가격})
            Object.주문체결_딕셔너리[주문번호].update({"통화코드": 통화코드})
            Object.주문체결_딕셔너리[주문번호].update({"만기일자": 만기일자})

            print("\n===== 주문체결 ========"
                                    "\n%s"
                                    "\n%s"
                                    "\n%s"
                                    "\n======================"
                                    % (주문번호, 체결수량, Object.주문체결_딕셔너리[주문번호]))

            # 주문을 체결한 후 tr요청해서 정보를 받아서 미결제를 다시 True or False로 바꿔주기에는 tr요청이 1초 1회라서 너무 느리다.
            # 따라서 체결과 동시에 미결제_딕셔너리 정보에 값을 넣어준다.
            미_딕 = 미결제_업데이트(
                종목코드 = 종목코드,
                주문번호 = 주문번호,
                미결제_딕셔너리 = Object.미결제_딕셔너리,
                종목정보_딕셔너리 = Object.종목정보_딕셔너리,
                실시간체결_딕셔너리 = Object.실시간체결_딕셔너리,
                주문체결_딕셔너리 = Object.주문체결_딕셔너리,
            )

            if 매도매수유형 == "1":

                print("\n====== 매도 체결 후 미결제 ======="
                                        "\n%s"
                                        "\n%s"
                                        "\n======================"
                                        % (종목코드, Object.미결제_딕셔너리[종목코드]['잔고수량']))

                if Object.미결제_딕셔너리[종목코드]['잔고수량'] <= 0:
                    del Object.미결제_딕셔너리[종목코드]

                if 주문번호 in Object.주문접수_딕셔너리:
                    del Object.주문접수_딕셔너리[주문번호]

                if 정정취소유형 == "1":
                    Object.매도 = False
                if 정정취소유형 == "2":
                    Object.정정 = False
                    Object.매도 = False

            elif 매도매수유형 == "2":

                print("\n====== 매수 체결 후 미결제 ======="
                        "\n%s"
                        "\n%s"
                        "\n======================"
                        % (종목코드, 미_딕))

                if 주문번호 in Object.주문접수_딕셔너리:
                    del Object.주문접수_딕셔너리[주문번호]

                Object.매수 = False

            del Object.주문체결_딕셔너리[주문번호]

class XingApi_Class(Object):

    def __init__(self):

        ##### 함수모음 #####
        Object.tr_signal_o3105 = self.tr_signal_o3105
        Object.tr_signal_CIDBQ01500 = self.tr_signal_CIDBQ01500
        Object.tr_signal_CIDBQ03000 = self.tr_signal_CIDBQ03000
        Object.order_buy_CIDBT00100 = self.order_buy_CIDBT00100
        Object.order_cancel_CIDBT01000 = self.order_cancel_CIDBT01000
        Object.order_cancel_CIDBT00900 = self.order_cancel_CIDBT00900

        ####################

        ##### XASession COM 객체를 생성한다. ("API이벤트이름", 콜백클래스) #####
        self.XASession_object = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEvent)
        ####################

        ##### 로그인 상태에서 Xing 서버와 접속 중이면 끊음. #####
        self.login_status_check()
        ####################

        ##### Xing 실서버, 모의서버 구분해서 연결하기 ("hts. 실서버, demo. 모의서버", "포트넘버") #####
        self.server_connect()
        ####################

        ##### 로그인 시도하기 ("아이디", "비밀번호", "공인인증 비밀번호", "서버타입(사용안함)", "발생한에러표시여부(무시)") #####
        self.login_connect_signal()
        ####################

        ##### 계좌번호 리스트 받기 #####
        self.get_account_info()
        ####################

        ##### 종목정보 #####
        Object.XAQuery_o3105 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvent)
        Object.XAQuery_o3105.ResFileName = "C:\\eBEST\\xingAPI\\Res\\o3105.res"
        ####################

        ##### 미결제 #####
        Object.XAQuery_CIDBQ01500 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvent)
        Object.XAQuery_CIDBQ01500.ResFileName = "C:\\eBEST\\xingAPI\\Res\\CIDBQ01500.res"
        ####################

        ##### 예수금/잔고현황 #####
        Object.XAQuery_CIDBQ03000 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvent)
        Object.XAQuery_CIDBQ03000.ResFileName = "C:\\eBEST\\xingAPI\\Res\\CIDBQ03000.res"
        ####################

        self.tr_signal_CIDBQ01500()
        time.sleep(1.1)
        self.tr_signal_CIDBQ03000()
        time.sleep(1.1)
        self.tr_signal_o3105("HCEIM21")

        ##### XA_DataSet의 XAReal COM 객체를 생성한다. ("API이벤트이름", 콜백클래스) #####
        Object.XAReal_OVH = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", XARealEvent)
        Object.XAReal_OVH.ResFileName = "C:\\eBEST\\xingAPI\\Res\\OVH.res"
        Object.XAReal_OVC = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", XARealEvent)
        Object.XAReal_OVC.ResFileName = "C:\\eBEST\\xingAPI\\Res\\OVC.res"
        ####################

        time.sleep(1.1)
        self.set_real_signal(symbol="HCEIM21")

        #####해외선물 TR 신규주문 #####
        Object.XAQuery_CIDBT00100 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvent)
        Object.XAQuery_CIDBT00100.ResFileName = "C:\\eBEST\\xingAPI\\Res\\CIDBT00100.res"
        ####################

        #####해외선물 TR 취소주문 #####
        Object.XAQuery_CIDBT01000 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvent)
        Object.XAQuery_CIDBT01000.ResFileName = "C:\\eBEST\\xingAPI\\Res\\CIDBT01000.res"
        ####################

        ##### 해외선물 TR 정정주문 #####
        Object.XAQuery_CIDBT00900 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvent)
        Object.XAQuery_CIDBT00900.ResFileName = "C:\\eBEST\\xingAPI\\Res\\CIDBT00900.res"
        ####################


        ##### 주문접수 실시간 COM 객체 실행 #####
        Object.XARealOrder_TC1 = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", XARealOrderEvent)
        Object.XARealOrder_TC1.ResFileName = "C:\\eBEST\\xingAPI\\Res\\TC1.res"
        Object.XARealOrder_TC1.AdviseRealData()
        ####################

        ##### 주문응답 실시간 COM 객체 실행 #####
        Object.XARealOrder_TC2 = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", XARealOrderEvent)
        Object.XARealOrder_TC2.ResFileName = "C:\\eBEST\\xingAPI\\Res\\TC2.res"
        Object.XARealOrder_TC2.AdviseRealData()
        ####################

        ##### 주문체결 실시간 COM 객체 실행 #####
        Object.XARealOrder_TC3 = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", XARealOrderEvent)
        Object.XARealOrder_TC3.ResFileName = "C:\\eBEST\\xingAPI\\Res\\TC3.res"
        Object.XARealOrder_TC3.AdviseRealData()
        ####################


        ##### COM 쓰레드 동작에서 메세지 큐에 들어온 데이터를 펌프한다. #####
        while True:
            pythoncom.PumpWaitingMessages()
        ####################


    # 로그인 상태 함수
    def login_status_check(self):
        print("★★★ login_status_check() 로그인 상태 함수")

        if Object.로그인완료 == True:
            self.XASession_object.DisconnectServer()
            Object.로그인완료 = False

    # 서버접속 확인 함수
    def server_connect(self):
        print("★★★ 서버접속 확인 함수")
        
        if self.XASession_object.ConnectServer("demo.ebestsec.co.kr", 20001) == True:
            print("★★★ 서버에 연결 됨")
        else:
            nErrCode = XASessionEvent.GetLastError()
            strErrMsg = XASessionEvent.GetErrorMessage(nErrCode)
            print(strErrMsg)

    # 로그인 시도 함수
    def login_connect_signal(self):
        print("★★★ 로그인 시도 함수")
        
        if self.XASession_object.Login(ID, PW, "", 0, 0) == True:
            print("★★★ 로그인 성공")

        while Object.로그인완료 == False:
            # COM 스레드에 메시지 루프가 필요할 때 현재 스레드에 대한 모든 대기 메시지를 체크합니다.
            pythoncom.PumpWaitingMessages()
            # print(self.로그인완료)

    # 계좌정보 가져오기 함수
    def get_account_info(self):
        print("★★★ 계좌정보 가져오기 함수")

        계좌수 = self.XASession_object.GetAccountListCount()
        for i in range(계좌수):
            계좌번호 = self.XASession_object.GetAccountList(i)
            if "55555" in 계좌번호:
                Object.해외선물_계좌번호 = 계좌번호

        print("★★★ 해외선물 계좌번호 %s" % Object.해외선물_계좌번호)

    # TR요청 시그널 함수
    def tr_signal_o3105(self, symbol=None):
        print("★★★ tr_signal_o3105() 해외선물 종목정보 TR요청 %s " % symbol)

        Object.XAQuery_o3105.SetFieldData("o3105InBlock", "symbol", 0, symbol)

        error = Object.XAQuery_o3105.Request(False) # 연속 조회일 경우만 True
        if error < 0:
            print("★★★ 에러코드 %s, 에러내용 %s" % (error, 에러코드(error)))

        Object.TR처리완료 = False
        while Object.TR처리완료 == False:
            # COM 스레드에 메시지 루프가 필요할 때 현재 스레드에 대한 모든 대기 메시지를 체크합니다.
            pythoncom.PumpWaitingMessages()


    def tr_signal_CIDBQ01500(self, IsNext=False):
        print("★★★ tr_signal_CIDBQ01500() 해외선물 미결제 잔고내역 TR요청")

        Object.XAQuery_CIDBQ01500.SetFieldData("CIDBQ01500InBlock1", "RecCnt", 0, 1)
        Object.XAQuery_CIDBQ01500.SetFieldData("CIDBQ01500InBlock1", "AcntTpCode", 0, "1")
        Object.XAQuery_CIDBQ01500.SetFieldData("CIDBQ01500InBlock1", "AcntNo", 0, Object.해외선물_계좌번호)
        Object.XAQuery_CIDBQ01500.SetFieldData("CIDBQ01500InBlock1", "FcmAcntNo", 0, "")
        Object.XAQuery_CIDBQ01500.SetFieldData("CIDBQ01500InBlock1", "Pwd", 0, "0000")
        Object.XAQuery_CIDBQ01500.SetFieldData("CIDBQ01500InBlock1", "QryDt", 0, "")
        Object.XAQuery_CIDBQ01500.SetFieldData("CIDBQ01500InBlock1", "BalTpCode", 0, "1")

        error = Object.XAQuery_CIDBQ01500.Request(IsNext) # 연속 조회일 경우만 True
        if error < 0:
            print("★★★ 에러코드 %s, 에러내용 %s" % (error, 에러코드(error)))

        Object.TR처리완료 = False
        while Object.TR처리완료 == False:
            # COM 스레드에 메시지 루프가 필요할 때 현재 스레드에 대한 모든 대기 메시지를 체크합니다.

            pythoncom.PumpWaitingMessages()


    def tr_signal_CIDBQ03000(self, IsNext=False):
        print("★★★ tr_signal_CIDBQ03000() 해외선물 예수금/잔고현황")

        Object.XAQuery_CIDBQ03000.SetFieldData("CIDBQ03000InBlock1", "RecCnt", 0, 1)
        Object.XAQuery_CIDBQ03000.SetFieldData("CIDBQ03000InBlock1", "AcntTpCode", 0, "1")
        Object.XAQuery_CIDBQ03000.SetFieldData("CIDBQ03000InBlock1", "AcntNo", 0, Object.해외선물_계좌번호)
        Object.XAQuery_CIDBQ03000.SetFieldData("CIDBQ03000InBlock1", "AcntPwd", 0, "0000")
        Object.XAQuery_CIDBQ03000.SetFieldData("CIDBQ03000InBlock1", "TrdDt", 0, "")

        error = Object.XAQuery_CIDBQ03000.Request(IsNext) # 연속 조회일 경우만 True
        if error < 0:
            print("★★★ 에러코드 %s, 에러내용 %s" % (error, 에러코드(error)))

        Object.TR처리완료 = False
        while Object.TR처리완료 == False:
            # COM 스레드에 메시지 루프가 필요할 때 현재 스레드에 대한 모든 대기 메시지를 체크합니다.
            pythoncom.PumpWaitingMessages()


    # 해외선물 체결정보 실시간
    def set_real_signal(self, symbol=None):
        print("★★★ set_real_signal() 해외선물 호가/체결정보 실시간요청 %s" % symbol)

        Object.XAReal_OVH.SetFieldData("InBlock", "symbol", symbol)
        Object.XAReal_OVH.AdviseRealData()
        Object.XAReal_OVC.SetFieldData("InBlock", "symbol", symbol)
        Object.XAReal_OVC.AdviseRealData()


    def order_buy_CIDBT00100(self, 레코드갯수=1, 주문일자=None, 지점코드=None, 계좌번호=None, 비밀번호=None, 종목코드값=None, 선물주문구분코드=None,
                            매매구분코드=None, 해외선물주문유형코드=None, 통화코드=None, 해외파생주문가격=0, 조건주문가격=0, 주문수량=0, 상품코드=None, 만기년월=None, 거래소코드=None):

        '''
        :param 레코드갯수: 1
        :param 주문일자: "YYYYMMDD"
        :param 지점코드: SPACE
        :param 계좌번호: "234235023"
        :param 비밀번호: "0000"
        :param 종목코드값: "GCJ20"
        :param 선물주문구분코드: "1:신규"
        :param 매매구분코드: "1:매도, 2:매수"
        :param 해외선물주문유형코드: "1:시장가, 2:지정가"
        :param 통화코드: SPACE
        :param 해외파생주문가격: "주문가격"
        :param 조건주문가격: 0
        :param 주문수량: 1
        :param 상품코드: SPACE
        :param 만기년월: SPACE
        :param 거래소코드: SPACE
        :return:
        '''

        print("★★★ order_buy_CIDBT00100() 해외선물 신규주문")

        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "RecCnt", 0, 레코드갯수)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "OrdDt", 0, 주문일자)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "BrnCode", 0, 지점코드)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "AcntNo", 0, 계좌번호)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "Pwd", 0, 비밀번호)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "IsuCodeVal", 0, 종목코드값)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "FutsOrdTpCode", 0, 선물주문구분코드)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "BnsTpCode", 0, 매매구분코드)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "AbrdFutsOrdPtnCode", 0, 해외선물주문유형코드)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "CrcyCode", 0, 통화코드)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "OvrsDrvtOrdPrc", 0, 해외파생주문가격)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "CndiOrdPrc", 0, 조건주문가격)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "OrdQty", 0, 주문수량)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "PrdtCode", 0, 상품코드)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "DueYymm", 0, 만기년월)
        Object.XAQuery_CIDBT00100.SetFieldData("CIDBT00100InBlock1", "ExchCode", 0, 거래소코드)

        error = Object.XAQuery_CIDBT00100.Request(False) # 연속 조회일 경우만 True
        if error < 0:
            print("order_buy_CIDBT00100 에러코드 %s, 에러내용 %s" % (error, 에러코드(error)))
            Object.매수 = False


    def order_cancel_CIDBT01000(self, 레코드갯수=1, 주문일자=None, 지점번호=None, 계좌번호=None, 비밀번호=None, 종목코드값=None, 해외선물원주문번호=None,
                             선물주문구분코드=None, 상품구분코드=None, 거래소코드=None):
        '''

        :param 레코드갯수:
        :param 주문일자: "YYYYMMDD"
        :param 지점번호: SPACE
        :param 계좌번호:
        :param 비밀번호:
        :param 종목코드값:
        :param 해외선물원주문번호: 0000012345 10자리
        :param 선물주문구분코드: 3
        :param 상품구분코드: SPACE
        :param 거래소코드: SPACE
        :return:
        '''

        print("★★★ order_buy_CIDBT01000() 해외선물 취소주문")

        Object.XAQuery_CIDBT01000.SetFieldData("CIDBT01000InBlock1", "RecCnt", 0, 레코드갯수)
        Object.XAQuery_CIDBT01000.SetFieldData("CIDBT01000InBlock1", "OrdDt", 0, 주문일자)
        Object.XAQuery_CIDBT01000.SetFieldData("CIDBT01000InBlock1", "BrnNo", 0, 지점번호)
        Object.XAQuery_CIDBT01000.SetFieldData("CIDBT01000InBlock1", "AcntNo", 0, 계좌번호)
        Object.XAQuery_CIDBT01000.SetFieldData("CIDBT01000InBlock1", "Pwd", 0, 비밀번호)
        Object.XAQuery_CIDBT01000.SetFieldData("CIDBT01000InBlock1", "IsuCodeVal", 0, 종목코드값)
        Object.XAQuery_CIDBT01000.SetFieldData("CIDBT01000InBlock1", "OvrsFutsOrgOrdNo", 0, 해외선물원주문번호)
        Object.XAQuery_CIDBT01000.SetFieldData("CIDBT01000InBlock1", "FutsOrdTpCode", 0, 선물주문구분코드)
        Object.XAQuery_CIDBT01000.SetFieldData("CIDBT01000InBlock1", "PrdtTpCode", 0, 상품구분코드)
        Object.XAQuery_CIDBT01000.SetFieldData("CIDBT01000InBlock1", "ExchCode", 0, 거래소코드)

        error = Object.XAQuery_CIDBT01000.Request(False)  # 연속 조회일 경우만 True
        if error < 0:
            print("order_buy_CIDBT01000 에러코드 %s, 에러내용 %s" % (error, 에러코드(error)))
            Object.취소 = False


    def order_cancel_CIDBT00900(self, 레코드갯수=1, 주문일자=None, 등록지점번호=None, 계좌번호=None, 비밀번호=None, 해외선물원주문번호=None, 종목코드값=None,
                             선물주문구분코드=None, 매매구분코드=None, 선물주문유형코드=None, 통화코드값=None, 해외파생주문가격=None, 조건주문가격=None, 주문수량=None,
                                해외파생상품코드=None, 만기년월=None, 거래소코드=None):
        '''

        :param 레코드갯수:
        :param 주문일자: YYYYMMDD
        :param 등록지점번호: SPACE
        :param 계좌번호:
        :param 비밀번호:
        :param 해외선물원주문번호:
        :param 종목코드값:
        :param 선물주문구분코드: 2:정정
        :param 매매구분코드: 1:매도, 2:매수
        :param 선물주문유형코드: 2:지정가
        :param 통화코드값: SPACE
        :param 해외파생주문가격:
        :param 조건주문가격:
        :param 주문수량:
        :param 해외파생상품코드: SPACE
        :param 만기년월: SPACE
        :param 거래소코드: SPACE
        :return:
        '''

        print("★★★ order_buy_CIDBT00900() 해외선물 정정주문")

        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "RecCnt", 0, 레코드갯수)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "OrdDt", 0, 주문일자)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "RegBrnNo", 0, 등록지점번호)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "AcntNo", 0, 계좌번호)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "Pwd", 0, 비밀번호)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "OvrsFutsOrgOrdNo", 0, 해외선물원주문번호)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "IsuCodeVal", 0, 종목코드값)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "FutsOrdTpCode", 0, 선물주문구분코드)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "BnsTpCode", 0, 매매구분코드)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "FutsOrdPtnCode", 0, 선물주문유형코드)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "CrcyCodeVal", 0, 통화코드값)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "OvrsDrvtOrdPrc", 0, 해외파생주문가격)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "CndiOrdPrc", 0, 조건주문가격)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "OrdQty", 0, 주문수량)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "OvrsDrvtPrdtCode", 0, 해외파생상품코드)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "DueYymm", 0, 만기년월)
        Object.XAQuery_CIDBT00900.SetFieldData("CIDBT00900InBlock1", "ExchCode", 0, 거래소코드)

        error = Object.XAQuery_CIDBT00900.Request(False)  # 연속 조회일 경우만 True
        if error < 0:
            print("order_buy_CIDBT00900 에러코드 %s, 에러내용 %s" % (error, 에러코드(error)))
            Object.정정 = False

if __name__ == "__main__":
    XingApi_Class()