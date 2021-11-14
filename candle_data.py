'''
1분봉 요청하기
'''
import win32com.client
import pythoncom
import time
from config.errCode import *

'''
사용될 변수 모아놓는 클래스
'''
class Object:

    ##### 오브젝트 모음 #####
    # 우리가 요청하는 TR요청 설정 변수
    XAQuery_o3105 = None # 종목정보 요청
    XAQuery_CIDBQ01500 = None # 해외선물 미결제 정보내역 요청
    XAQuery_CIDBQ03000 = None # 해외선물 예수금/잔고현황
    XAQuery_o3103 = None # 해외선물 분봉 조회
    #############################

    ##### 함수를 할당한 변수 #####
    tr_signal_o3105 = None # 종목정보tr요청 함수를 할당
    tr_signal_CIDBQ01500 = None # 미결제 함수를 할당
    tr_signal_CIDBQ03000 = None # 예수금/잔고현황 함수를 할당
    tr_signal_o3103 = None # 해외선물 분봉 조회
    #############################

    ##### 기타 변수 모음 #####
    TR처리완료 = False # TR요청완료 기다리기
    로그인완료 = False # 로그인완료인지 확인하기
    해외선물_계좌번호 = None # 해외선물 계좌번호만 할당
    종목정보_딕셔너리 = {} # 종목의 상세 정보 데이터
    미결제_딕셔너리 = {} # 체결 후 결제되지 않은 상품
    예수금_딕셔너리 = {} # 예수금 상세정보 데이터
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

        elif szTrCode == "o3103":
            print("★★★ 해외선물 분봉 조회 결과반환")

            종목코드 = self.GetFieldData("o3103OutBlock", "shcode", 0)
            시차 = self.GetFieldData("o3103OutBlock", "timediff", 0)
            조회건수 = self.GetFieldData("o3103OutBlock", "readcnt", 0)
            연속일자 = self.GetFieldData("o3103OutBlock", "cts_date", 0)
            연속시간 = self.GetFieldData("o3103OutBlock", "cts_time", 0)

            occurs_count = self.GetBlockCount("o3103OutBlock1")
            print("===== 종목코드: %s, 분봉 갯수 : %s =====" % (종목코드, occurs_count))


            for i in range(occurs_count):
                날짜 = self.GetFieldData("o3103OutBlock1", "date", i)
                현지시간 = self.GetFieldData("o3103OutBlock1", "time", i)
                시가 = self.GetFieldData("o3103OutBlock1", "open", i)
                고가 = self.GetFieldData("o3103OutBlock1", "high", i)
                저가 = self.GetFieldData("o3103OutBlock1", "low", i)
                종가 = self.GetFieldData("o3103OutBlock1", "close", i)
                거래량 = self.GetFieldData("o3103OutBlock1", "volume", i)

            if Object.XAQuery_o3103.IsNext:
                print("이어서 조회하려는 날짜: %s, 현지시간: %s" % (날짜, 현지시간))
                Object.tr_signal_o3103(symbol=종목코드, cts_date=날짜, cts_time=현지시간, IsNext=True)
            else:
                Object.TR처리완료 = True


    def OnReceiveMessage(self, systemError, messageCode, message):
        '''
        -38
        02662 - 모의투자 정정 가능한 수량을 초과하였습니다.
        02258 - 모의투자 원주문번호를 잘못 입력하셨습니다.
        01231 - 주문가격은 현재가보다 커야합니다.
        03576 - 모의투자 증거금 한도 초과로 주문 불가합니다.
        :param systemError:
        :param messageCode:
        :param message:
        :return:
        '''

        print("★★★ systemError: %s, messageCode: %s, message: %s" % (systemError, messageCode, message))

class XingApi_Class(Object):

    def __init__(self):

        ##### 함수모음 #####
        Object.tr_signal_o3105 = self.tr_signal_o3105
        Object.tr_signal_CIDBQ01500 = self.tr_signal_CIDBQ01500
        Object.tr_signal_CIDBQ03000 = self.tr_signal_CIDBQ03000
        Object.tr_signal_o3103 = self.tr_signal_o3103
        ####################

        ##### XASession COM 객체를 생성한다. ("API이벤트이름", 콜백클래스) #####
        self.XASession_object = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEvent)
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

        ##### 분봉 요청 #####
        Object.XAQuery_o3103 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvent)
        Object.XAQuery_o3103.ResFileName = "C:\\eBEST\\xingAPI\\Res\\o3103.res"
        ####################

        self.tr_signal_CIDBQ01500()
        time.sleep(1.1)
        self.tr_signal_CIDBQ03000()
        time.sleep(1.1)
        self.tr_signal_o3105("HCEIM21")
        time.sleep(1.1)
        self.tr_signal_o3103("HCEIM21")

    # 서버접속 확인 함수
    def server_connect(self):
        print("★★★ 서버접속 확인 함수")
        
        if self.XASession_object.ConnectServer("demo.ebestsec.co.kr", 20001) == True:
            print("★★★ 서버에 연결 됨")
        else:
            nErrCode = self.XASession_object.GetLastError()
            strErrMsg = self.XASession_object.GetErrorMessage(nErrCode)
            print(strErrMsg)

    # 로그인 시도 함수
    def login_connect_signal(self):
        print("★★★ 로그인 시도 함수")
        
        if self.XASession_object.Login("uj02030", "0998bnm", "", 0, 0) == True:
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

    def tr_signal_o3103(self, symbol=None, cts_date=None, cts_time=None, IsNext=False):
        print("★★★ tr_signal_o3103() 해외선물차트 분봉 조회")

        time.sleep(1.1)

        Object.XAQuery_o3103.SetFieldData("o3103InBlock", "shcode", 0, symbol)
        Object.XAQuery_o3103.SetFieldData("o3103InBlock", "ncnt", 0, 1) # N분주기
        Object.XAQuery_o3103.SetFieldData("o3103InBlock", "readcnt", 0, 100)
        Object.XAQuery_o3103.SetFieldData("o3103InBlock", "cts_date", 0, cts_date)
        Object.XAQuery_o3103.SetFieldData("o3103InBlock", "cts_time", 0, cts_time)

        error = Object.XAQuery_o3103.Request(IsNext) # 연속 조회일 경우만 True (다음조회)버튼 누르기
        if error < 0:
            print("★★★ 에러코드 %s, 에러내용 %s" % (error, 에러코드(error)))

        Object.TR처리완료 = False
        while Object.TR처리완료 == False:
            # COM 스레드에 메시지 루프가 필요할 때 현재 스레드에 대한 모든 대기 메시지를 체크합니다.
            pythoncom.PumpWaitingMessages()
            
if __name__ == "__main__":
    XingApi_Class()