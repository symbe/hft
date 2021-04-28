#!/usr/bin/env python
# coding: utf-8

# In[ ]:

# 확장모듈 중 PyQt5 설치 필요
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QAxContainer import *
from tkinter import messagebox

strDemoId = '20181373'
strDemoPwd = 'eugenegn1!@'
strPwd = 'eugene14!!'
g_DStockType = []  # 국내 주야간구분
g_DStockKey = []  # 국내 종목 Key
g_DStockCode = []  # 국내 종목코드
g_DStockReal = []  # 국내 리얼코드
g_GStockExnm = []  # 해외 거래소코드
g_GStockCode = []  # 해외 종목코드
g_GStockReal = []  # 해외 리얼코드
# 종목코드 버퍼 생성
g_nBufRow = 1
g_nBufCol = 21
g_listBuf = [[0 for col in range(g_nBufCol)] for row in range(g_nBufRow)]

# 국내 계좌 07879010

class RSI():

    def __init__(self):

        self.over = False

    def calculate(self, date):

        AU = (sum( item for item in date if item > 0) / 14)
        AD = (-sum( item for item in date if item < 0) / 14)

        RSI = ( AU / (AU + AD) ) * 100

        if RSI > 72:
            self.over = True
        elif RSI < 23:
            self.over = True

    def reset(self):
        self.over = False


# 원하는 종목 인덱스 찾음.
def GetIndexByCode(strCode):
    global g_listBuf
    for i in range(g_nBufRow):
        if g_listBuf[i][0] == "":
            g_listBuf[i][0] = strCode
            return i
        if g_listBuf[i][0] == strCode:
            return i
    return -1


class Quote():

    def __init__(self):

        self.now = 0
        self.bid_price = 0
        self.ask_price = 0

        self.bid_size = 0
        self.ask_size = 0
        self.bid_change = 0
        self.ask_change = 0

        self.bidv = 0
        self.askv = 0

        self.traded = False
        self.level_ct = 0
        self.time = 0

    def reset(self):
        self.traded = True
        self.level_ct += 1
        print(self.level_ct)

    def max(self):
        self.traded = True

    def update(self, data):

        self.now = data[0]
        self.bidv = data[5]
        self.askv = data[6]
        self.bid_change = data[7]
        self.ask_change = data[8]
        self.time = data[9]

        if (
                self.bid_price != data[1]
                and self.ask_price != data[2]

        ):

            self.bid_price = data[1]
            self.ask_price = data[2]

            print("Level Change:", self.bid_price, self.ask_price)
            self.traded = False


class MyWindow(QMainWindow):

    # 생성자
    def __init__(self):
        global g_Api, g_listBuf
        super().__init__()
        self.setWindowTitle("API")
        self.setGeometry(300, 300, 400, 200)
        # API 로드
        self.api = QAxWidget("OPENAPIX.Ctrl.1")

        # 접속버튼
        btn1 = QPushButton("Connect", self)
        btn1.move(20, 20)
        btn1.clicked.connect(self.btn1_clicked)

        # 로그인 버튼
        btn2 = QPushButton("Login", self)
        btn2.move(20, 70)
        btn2.clicked.connect(self.btn2_clicked)

        # 계좌리스트 콤보
        self.combo2 = QComboBox(self)
        self.combo2.move(20, 110)
        self.combo2.resize(100, 20)

        # 국내 시세조회 버튼
        btn3 = QPushButton("국내매매실행", self)
        btn3.move(140, 20)
        btn3.clicked.connect(self.btn3_clicked)

        # 국내 실시간 손익
        btn4 = QPushButton("실시간손익", self)
        btn4.move(250, 20)
        btn4.clicked.connect(self.btn4_clicked)

        # 매수
        btn5 = QPushButton("매수 1계약", self)
        btn5.move(140, 110)
        btn5.clicked.connect(self.btn5_clicked)

        # 매도
        btn6 = QPushButton("매도 1계약", self)
        btn6.move(250, 110)
        btn6.clicked.connect(self.btn6_clicked)

        lbl0 = QLabel("종목 : ", self)
        lbl0.move(140, 62)
        lbl0.resize(40, 15)
        self.combo0 = QComboBox(self)
        self.combo0.move(180, 60)
        self.combo0.resize(80, 20)
        self.combo1 = QComboBox(self)
        self.combo1.move(270, 60)
        self.combo1.resize(80, 20)

        global quote
        global rsi
        rsi = RSI()
        quote = Quote()

        # 로그인 이벤트 선언.
        self.api.LoginOK.connect(self.event_loginok)
        # 조회 응답 이벤트 선언
        self.api.ReceiveData.connect(self.event_receivedata)
        # 실시간 응답 이벤트 선언
        self.api.ReceiveReal.connect(self.event_receivereal)

        # 버퍼 초기화
        for i in range(g_nBufRow):
            for j in range(g_nBufCol):
                g_listBuf[i][j] = ""

    # 소멸자
    def __del__(self):
        # 체널매니저 종료
        self.api.dynamicCall("DoClose()")

    # 종료버튼 눌렀을 시 이벤트 처리
    def closeEvent(self, event):
        self.api.dynamicCall("DoClose()")
        super(QMainWindow, self).closeEvent(event)

    # 처음 컨넥션 함수 실행.
    def btn1_clicked(self):
        ret = self.api.dynamicCall("DoConnect()")
        if ret == 1:
            QMessageBox.question(self, '연결성공', "성공", QMessageBox.Ok)
        else:
            QMessageBox.question(self, '연결실패', "실패!!!", QMessageBox.Ok)

    # Login 버튼을 눌렀을 시 처리
    def btn2_clicked(self):
        # 접속자타입 (0:HTS고객, 1:Front, 2:모의HTS고객, 3:모의Front), 아이디, 패스워드, 인증서비밀번호
        ret = self.api.dynamicCall("DoLogin(int, QString, QString, QString)", 0, strDemoId, strDemoPwd, strPwd)
        # if ret == 1:
        #    QMessageBox.question(self, '로그인성공', "성공", QMessageBox.Ok)
        # else:
        #    QMessageBox.question(self, '로그인실패', "실패!!!", QMessageBox.Ok)
        return

    # 국내 시세 조회.
    def btn3_clicked(self):
        # 조회서버 (0:국내시세, 1:해외시세, -1:모든 주문/계좌관련), 전문코드, 입력스트링, 다음조회키, 사용자정의문자열
        # 현재가조회
        nSel = self.combo0.currentIndex()

        # 현재가 실시간.
        ret = self.api.dynamicCall("DoRegistReal(QString, QString)", "KA", g_DStockReal[nSel])
        # 호가 실시간.
        ret = self.api.dynamicCall("DoRegistReal(QString, QString)", "KB", g_DStockReal[nSel])


    # 국내 실시간 손익현황 조회
    def btn4_clicked(self):
        # 조회서버 (0:국내시세, 1:해외시세, -1:모든 주문/계좌관련), 전문코드, 입력스트링, 다음조회키, 사용자정의문자열

        #미결제 수량
        strr = "200000099999909"
        ret = self.api.dynamicCall("DoRequestData(int, QString, QString, QString, QString)", -1, "DSBB010030", strr, "", "")

        nSel = self.combo0.currentIndex()
        str = '0001' + g_DStockKey[nSel] + g_DStockCode[nSel].ljust(12, ' ')
        ret = self.api.dynamicCall("DoRequestData(int, QString, QString, QString, QString)", 0, "PIBO1001", str, "", "TEST1")

        # 손익 조회
        strInput = "200000099999930"
        ret = self.api.dynamicCall("DoRequestData(int, QString, QString, QString, QString)", -1, "DSBB010040", strInput, "", "")


    # 매수 1계약
    def btn5_clicked(self):

        pw = self.api.dynamicCall("GetEncodeText(Qstring)", "6644")
        now = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO1001", 1, 0, 5)  # 현재가
        Qs = ""
        str1 = ["2", "07879010", pw, "167Q3000", "1", "1", "2", "1", now, Qs]
        # 주야간 구분, 계좌번호, 비밀번호, 종목, 매매구분, 가격조건, 체결조건, 주문수량, 주문가격, 사용자정의 문자열
        self.api.dynamicCall(
            "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
            str1)

    # 매도 1계약
    def btn6_clicked(self):

        pw = self.api.dynamicCall("GetEncodeText(Qstring)", "6644")
        now = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO1001", 1, 0, 5)  # 현재가
        Qs = ""
        str2 = ["2", "07879010", pw, "167Q3000", "2", "1", "2", "1", now, Qs]
        # 주야간 구분, 계좌번호, 비밀번호, 종목, 매매구분, 가격조건, 체결조건, 주문수량, 주문가격, 사용자정의 문자열
        self.api.dynamicCall(
            "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
            str2)

    # 위에 선언된 로그인 이벤트 응답에 대한 처리
    def event_loginok(self, err_code):
        strResultMsg = 'event_loginok. ErrorCode : %s' % (err_code)
        print(strResultMsg)
        if err_code == '00000':
            # 정상이라면 관리계좌와 종목 정보 조회
            strInput = strDemoId.ljust(8, ' ') + 'H'
            strResultMsg = 'Request CSBA000101 : %s' % (strInput)
            print(strResultMsg)
            # 관리계좌 요청
            ret = self.api.dynamicCall("DoRequestData(int, QString, QString, QString, QString)", -1, "CSBA000101",
                                       strInput, "", "")
        return


    # 조회 응답 이벤트 선언
    def event_receivedata(self, DataSeq, TrCode, ErrCode, quote):

        global stop

        # 응답값이 에러가 있는지 확인
        if ErrCode != "" and ErrCode != "U91001" and ErrCode != "U49004" and ErrCode != "U46537" and ErrCode != "U31051" and ErrCode != "U75001" and ErrCode != "U53102" and ErrCode != "U00000" and ErrCode != "G00000" and ErrCode != "G10152" and ErrCode != "G50128" and ErrCode != "G46537":
            strResultMsg = '%s 전문 조회 에러. C:\\Champion\\Data\\ErrorCode.ini에서 코드(%s)를 확인하십시오.' % (TrCode, ErrCode)
            print(strResultMsg)
            return

        if TrCode == "PIBO1001":
            # 현재가
            strValue = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO1001", 1, 0, 5)

        # RSI 1분봉
        if TrCode == "PIBO2102":

            try:
                a1 = float(self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 0, 6))

            except:
                return

            a2 = float(self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 1, 6))
            a3 = float(self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 2, 6))
            a4 = float(self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 3, 6))
            a5 = float(self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 4, 6))
            a6 = float(self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 5, 6))
            a7 = float(self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 6, 6))
            a8 = float(self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 7, 6))
            a9 = float(self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 8, 6))
            a10 = float(self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 9, 6))
            a11 = float(self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 10, 6))
            a12 = float(self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 11, 6))
            a13 = float(self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 12, 6))
            a14 = float(self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 13, 6))

        #국내 손익현황
        if TrCode == "DSBB010040":

            a = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "DSBB010040", 1, 0, 8) #순손익
            b = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "DSBB010040", 1, 0, 7) #수수료

            print("Total:", a, "fees:", b)

        #현재 미결제 수량 제한
        if TrCode == "DSBB130603":
            position = self.api.dynamicCall("GetDataAt(QString, int, int, int)", TrCode, 1, 0, 3) #포지션
            quantity = self.api.dynamicCall("GetDataAt(QString, int, int, int)", TrCode, 1, 0, 5) #수량

        if TrCode == "DSBB010030":

            nRow = int(self.api.dynamicCall("GetDataRowCount(QString, int)", TrCode, 1))

            for i in range(0, nRow):

                pss = self.api.dynamicCall("GetDataAt(QString, int, int, int)", TrCode, 1, i, 6) #포지션
                qtt = self.api.dynamicCall("GetDataAt(QString, int, int, int)", TrCode, 1, i, 10) #수량
                prs = self.api.dynamicCall("GetDataAt(QString, int, int, int)", TrCode, 1, i, 14)  #평단

                print("Position:", pss, "/ Count:", qtt, "/ price:", prs)

        #미체결 주문 확인
        if TrCode == "DSBB010010":

            nRow = int(self.api.dynamicCall("GetDataRowCount(QString, int)", TrCode, 1))

            for i in range(0, nRow):

                num = self.api.dynamicCall("GetDataAt(QString, int, int, int)", TrCode, 1, i, 4) #주문번호
                ps = self.api.dynamicCall("GetDataAt(QString, int, int, int)", TrCode, 1, i, 10) #매매구분
                qt = self.api.dynamicCall("GetDataAt(QString, int, int, int)", TrCode, 1, i, 14) #주문수량
                pr = self.api.dynamicCall("GetDataAt(QString, int, int, int)", TrCode, 1, i, 18) #주문가격
                now = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO1001", 1, 0, 5) #현재가

                stop = [[num, ps, qt, pr, now] for row in range(nRow)]

                try:
                    a = int(num)
                    b = float(ps)
                    c = float(qt)
                    d = float(pr)
                    e = float(now)

                except:
                    return

                else:
                    pass

                #미체결 익절 주문 손절주문로 정정(매수 포지션)

                if (stop[i][1] == "1"
                and (float(stop[i][3]) + 0.1) < float(stop[i][4])
                ):

                    print("BUY STOP LOSS /", "BID:", stop[i][3], "/ NOW:", stop[i][4])

                    pw = self.api.dynamicCall("GetEncodeText(Qstring)", "6644")
                    try:
                        nm = float(stop[i][0])
                    except:
                        return

                    nw = float(stop[i][4]) + 0.01

                    # 정정주문
                    str1 = ["10", "2", "07879010", pw, "1", "1", "2", nw, num, ""]
                    self.api.dynamicCall(
                        "DoRequestDomChangeOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
                        str1)

                    print("Stop Loss: Buy at", nw)


                #미체결 익절 주문 손절주문으로 정정(매도 포지션)
                elif (stop[i][1] == "2"
                and (float(stop[i][3]) - 0.1) > float(stop[i][4])
                ):


                    print("SELL STOP LOSS /", "ASK:", stop[i][3], "/ NOW:", stop[i][4])
                    pw = self.api.dynamicCall("GetEncodeText(Qstring)", "6644")


                    nw = float(stop[i][4]) - 0.01

                    # 정정주문
                    str3 = ["10", "2", "07879010", pw, "1", "1", "2", nw, num, ""]
                    self.api.dynamicCall(
                        "DoRequestDomChangeOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)", str3)

                    print("Stop Loss: Sell at", nw)

            return

        if TrCode == "CSBA000101":
            # 관리계좌 응답처리.
            strResultMsg = 'Receive CSBA000101 Error : %s' % (ErrCode)
            print(strResultMsg)
            # CSBA000101 전문의 응답 0번 레이어 총 Row 갯수
            nRow = int(self.api.dynamicCall("GetDataRowCount(QString, int)", TrCode, 0))
            # CSBA000101 전문의 응답 0번 레이어 총 Col 갯수
            nCol = int(self.api.dynamicCall("GetDataColCount(QString, int)", TrCode, 0))

            # 총 관리계좌 수랑.
            nCount = int(self.api.dynamicCall("GetDataAt(QString, int, int, int)", TrCode, 0, 0, 0))
            nRow = int(self.api.dynamicCall("GetDataRowCount(QString, int)", TrCode, 1))
            # 실제로 nCount와 nRow는 값이 같다. 둘중 뭘 써도 된다.
            print('Count / Row : ' + str(nCount) + ' / ' + str(nRow))
            # CSBA000101 전문의 응답 1번 레이어 총 Col 갯수
            nCol = int(self.api.dynamicCall("GetDataColCount(QString, int)", TrCode, 1))
            for i in range(0, nRow):
                strtemp = ''
                strGrpNo = ''
                for j in range(0, nCol):
                    # CSBA000101 전문의 총 Col 갯수
                    strtemp = strtemp + self.api.dynamicCall("GetDataAt(QString, int, int, int)", TrCode, 1, i, j)
                    strtemp = strtemp + ","
                    # 테스트 프로그램에서는 계좌만 콤보에 넣을 것이므로.. 그룹명이 없는 계좌들만 넣는다.
                    if j == 0:
                        strGrpNo = self.api.dynamicCall("GetDataAt(QString, int, int, int)", TrCode, 1, i, j).strip()
                    if j == 1 and strGrpNo == '':
                        strAccnt = self.api.dynamicCall("GetDataAt(QString, int, int, int)", TrCode, 1, i, j).strip()
                        self.combo2.addItem(strAccnt)
                # print(strtemp)
            # 종목 파일 읽기
            strPath = "C:\\Champion\\data\\master"
            strDomeFile = strPath + "\\sfile03.dat"
            print(strDomeFile)
            fp = open(strDomeFile, 'r')
            for line in fp:
                strTemp = line.replace("\'", "")
                arr = strTemp.split(',')
                # 주야간구분
                g_DStockType.append(arr[5])
                # 국내 종목 Key
                g_DStockKey.append(arr[7])
                # 국내 종목코드
                g_DStockCode.append(arr[8])
                # 국내 리얼요청코드
                g_DStockReal.append(arr[22])
                # 콤보에 추가
                self.combo0.addItem(arr[8])
            fp.close()
        return

    # 실시간 응답 이벤트 선언
    def event_receivereal(self, RealCode, RealKey, UserArea):

        global g_listBuf
        global data
        global quote
        global rsi

        if RealCode == "KA":
            # 종목코드
            strCode = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 1)
            nRow = GetIndexByCode(strCode)
            if nRow != -1:
                # 현재가
                g_listBuf[nRow][1] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 11)
                # 매수최우선
                g_listBuf[nRow][2] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 25)
                # 매도최우선
                g_listBuf[nRow][3] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 23)


        if RealCode == "KB":

            # 종목코드
            strCode = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 1)
            nRow = GetIndexByCode(strCode)
            if nRow != -1:
                # 총매수수량
                g_listBuf[nRow][4] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 8)
                # 총매도수량
                g_listBuf[nRow][5] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 4)
                # 매수1
                g_listBuf[nRow][6] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 22)
                # 매수2
                g_listBuf[nRow][7] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 38)
                # 매수3
                g_listBuf[nRow][8] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 54)
                # 매도1
                g_listBuf[nRow][9] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 14)
                # 매도2
                g_listBuf[nRow][10] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 30)
                # 매도3
                g_listBuf[nRow][11] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 46)
                # 총매수 직전대비
                g_listBuf[nRow][12] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 11)
                # 총매도 직전대비
                g_listBuf[nRow][13] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 7)
                # 시간
                g_listBuf[nRow][14] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 3)

                # 매수4
                g_listBuf[nRow][15] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 70)
                # 매수5
                g_listBuf[nRow][16] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 86)
                # 매도4
                g_listBuf[nRow][17] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 62)
                # 매도5
                g_listBuf[nRow][18] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 78)

                # 매수1호가
                g_listBuf[nRow][19] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 21)
                # 매도1호가
                g_listBuf[nRow][20] = self.api.dynamicCall("GetRealAt(QString, QString, int)", RealCode, RealKey, 13)

        # 미체결 주문 확인 및 손절 주문
        nSel = self.combo0.currentIndex()
        str = '0001' + g_DStockKey[nSel] + g_DStockCode[nSel].ljust(12, ' ')
        ret = self.api.dynamicCall("DoRequestData(int, QString, QString, QString, QString)", 0, "PIBO1001", str, "",
                                   "TEST1")

        strInput = "200000099999930"
        ret = self.api.dynamicCall("DoRequestData(int, QString, QString, QString, QString)", -1, "DSBB010010", strInput,
                                   "", "")

        #1 신규 주문 및 익절 주문
        for j in range(g_nBufCol):

            try:
                a = float(g_listBuf[0][2])
            except:
                return
            else:
                pass

            try:
                b = float(g_listBuf[0][6])
            except:
                return
            else:
                pass

            bidp = float(g_listBuf[0][19])
            askp= float(g_listBuf[0][20])


            #매수평균의도가격
            bidV = (bidp * float(g_listBuf[0][6]) + (bidp - 0.01) * float(g_listBuf[0][7]) + (bidp - 0.02) * float(g_listBuf[0][8]) + (bidp - 0.03) * float(g_listBuf[0][15]) + (bidp - 0.04) * float(g_listBuf[0][16])) / (float(g_listBuf[0][6]) + float(g_listBuf[0][7]) + float(g_listBuf[0][8]) + float(g_listBuf[0][15]) + float(g_listBuf[0][16]))

            # 매도평균의도가격
            askV = (askp * float(g_listBuf[0][9]) + (askp + 0.01) * float(g_listBuf[0][10]) + (askp + 0.02) * float(g_listBuf[0][11]) + (askp + 0.03) * float(g_listBuf[0][17]) + (askp + 0.04) * float(g_listBuf[0][18])) / (float(g_listBuf[0][9]) + float(g_listBuf[0][10]) + float(g_listBuf[0][11]) + float(g_listBuf[0][17]) + float(g_listBuf[0][18]))

            data = [0 for col in range(12)]
            data[0] = g_listBuf[0][1] #현재가
            data[1] = g_listBuf[0][2] #매수최우선
            data[2] = g_listBuf[0][3] #매도최우선
            data[3] = g_listBuf[0][4] #총매수
            data[4] = g_listBuf[0][5] #총매도

            data[5] = bidV #매수평균의도가격
            data[6] = askV #매도평균의도가격
            data[7] = float(g_listBuf[0][6]) + float(g_listBuf[0][7]) + float(g_listBuf[0][8]) + float(g_listBuf[0][15]) #매수1~3
            data[8] = float(g_listBuf[0][9]) + float(g_listBuf[0][10]) + float(g_listBuf[0][11]) + float(g_listBuf[0][17])  # 매도1~3

            data[9] = g_listBuf[0][12] #총매수 직전대비
            data[10] = g_listBuf[0][13] #총매도 직전대비
            data[11] = g_listBuf[0][14] #시간

            quote.update(data)
            rsi.reset()

            #미결제수량 초과 시
            strr = "H" + "20181373"
            ret = self.api.dynamicCall("DoRequestData(int, QString, QString, QString, QString)", -1, "DSBB130603", strr, "", "")
            quantity = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "DSBB130603", 1, 0, 5)  # 수량

            try:
                Q = float(quantity)
            except:
                return
            else:
                pass

            if Q > 12:
                quote.max()
                return

            if quote.traded:
                return

            VWAP = (float(quote.bidv) + float(quote.askv)) * 0.5


            # RSI 과매도 or 과매수 시 멈춤, 23~72
            nSel = self.combo0.currentIndex()
            strInput = g_DStockKey[nSel] + g_DStockCode[nSel].ljust(12,' ') + '00000000000000000099999999235959010000030'
            ret = self.api.dynamicCall("DoRequestData(int, QString, QString, QString, QString)", 0, "PIBO2102", strInput, "", "")

            a1 = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 0, 6)

            try:
                float(a1)
            except:
                return

            a2 = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 1, 6)
            a3 = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 2, 6)
            a4 = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 3, 6)
            a5 = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 4, 6)
            a6 = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 5, 6)
            a7 = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 6, 6)
            a8 = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 7, 6)
            a9 = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 8, 6)
            a10 = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 9, 6)
            a11 = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 10, 6)
            a12 = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 11, 6)
            a13 = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 12, 6)
            a14 = self.api.dynamicCall("GetDataAt(QString, int, int, int)", "PIBO2102", 1, 13, 6)

            date = [float(a1), float(a2), float(a3), float(a4), float(a5), float(a6), float(a7), float(a8), float(a9), float(a10), float(a11), float(a12), float(a13), float(a14)]
            rsi.calculate(date)

            if rsi.over:
                return

            #총매수 압도적 우위
            if (float(data[3]) > float(data[4]) * 3
            and quote.now == quote.ask_price):

                bidpoint = float(data[0])
                targetsell = float(data[0]) + 0.02
                pw = self.api.dynamicCall("GetEncodeText(Qstring)", "6644")
                # 매도호가에서 매수
                str1 = ["2", "07879010", pw, "167Q3000", "1", "1", "1", "2", bidpoint, ""]
                self.api.dynamicCall(
                    "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)", str1)

                # 익절
                str2 = ["2", "07879010", pw, "167Q3000", "2", "1", "1", "2", targetsell, ""]
                self.api.dynamicCall(
                    "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
                    str2)

                print("BUY at: ", quote.now, "/ time: ", data[11], "Strong BID")
                quote.reset()

            elif float(data[4]) * 1.5 < float(data[3]) < float(data[4]) * 3:


                #매수1~3 > 매도1~3 받쳐줄때
                if ( float(data[7]) > float(data[8]) * 2
                and quote.now == quote.ask_price ):

                    bidpoint = float(data[0])
                    targetsell = float(data[0]) + 0.02
                    pw = self.api.dynamicCall("GetEncodeText(Qstring)", "6644")

                    # 매도호가에서 매수
                    str1 =["2", "07879010", pw, "167Q3000", "1", "1", "1", "2", bidpoint, ""]
                    self.api.dynamicCall("DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)", str1)

                    # 익절
                    str2 = ["2", "07879010", pw, "167Q3000", "2", "1", "1", "2", targetsell, ""]
                    self.api.dynamicCall("DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)", str2)

                    print("BUY at: ", quote.now, "/ time: ", data[11], "Weak BID, and bid3 > ask3" )
                    quote.reset()

                #매도1~3 >>> 매수1~3, 매수막힌 후 소폭 조정
                elif (float(data[7]) * 3 < float(data[8])
                and quote.now == quote.bid_price ):

                    askpoint = float(data[0])
                    targetbuy = float(data[0]) - 0.02
                    pw = self.api.dynamicCall("GetEncodeText(Qstring)", "6644")

                    # 매수호가에서 매도
                    str3 = ["2", "07879010", pw, "167Q3000", "2", "1", "1", "2", askpoint, ""]
                    self.api.dynamicCall(
                        "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
                        str3)

                    str4 = ["2", "07879010", pw, "167Q3000", "1", "1", "1", "2", targetbuy, ""]
                    self.api.dynamicCall(
                        "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
                        str4)

                    print("SELL at: ", quote.now, "/ time: ", data[11], "Weak BID, BUT ask3 > bid3")
                    quote.reset()

            #총매수, 총매도 의미 없을 때
            elif float(data[4]) * 0.66 < float(data[3]) < float(data[4]) * 1.5:

                # 매수의도 > 매도
                if (VWAP > float(quote.now) + 0.008
                and quote.now == quote.ask_price):
                    bidpoint = float(data[0])
                    targetsell = float(data[0]) + 0.02
                    pw = self.api.dynamicCall("GetEncodeText(Qstring)", "6644")

                    # 매도호가에서 매수
                    str1 = ["2", "07879010", pw, "167Q3000", "1", "1", "1", "2", bidpoint, ""]
                    self.api.dynamicCall(
                        "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
                        str1)

                    str2 = ["2", "07879010", pw, "167Q3000", "2", "1", "1", "2", targetsell, ""]
                    self.api.dynamicCall(
                        "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
                        str2)


                    print("BUY at: ", quote.now, "/ time: ", data[11], "VWAP:", VWAP)
                    quote.reset()

                # 매도의도 > 매수의도
                if (VWAP < float(quote.now) - 0.008
                and quote.now == quote.bid_price):

                    askpoint = float(data[0])
                    targetbuy = float(data[0]) - 0.02
                    pw = self.api.dynamicCall("GetEncodeText(Qstring)", "6644")

                    # 매수호가에서 매도
                    str3 = ["2", "07879010", pw, "167Q3000", "2", "1", "1", "2", askpoint, ""]
                    self.api.dynamicCall(
                        "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
                        str3)

                    str4 = ["2", "07879010", pw, "167Q3000", "1", "1", "1", "2", targetbuy, ""]
                    self.api.dynamicCall(
                        "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
                        str4)

                    print("Sell at: ", quote.now, "/ time: ", data[11], "VWAP:", VWAP)
                    quote.reset()


            #총매도 소폭 우위
            elif float(data[3]) * 1.5 < float(data[4]) < float(data[3]) * 3:

                # 매도1~3 > 매수1~3 받쳐줄때
                if (float(data[8]) > float(data[7]) * 2
                and quote.now == quote.bid_price ):

                    askpoint = float(data[0])
                    targetbuy = float(data[0]) - 0.02
                    pw = self.api.dynamicCall("GetEncodeText(Qstring)", "6644")

                    # 매수호가에서 매도
                    str1 = ["2", "07879010", pw, "167Q3000", "2", "1", "1", "2", askpoint, ""]
                    self.api.dynamicCall(
                    "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
                    str1)

                    # 익절
                    str2 = ["2", "07879010", pw, "167Q3000", "1", "1", "1", "2", targetbuy, ""]
                    self.api.dynamicCall(
                        "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
                        str2)

                    print("SELL at: ", quote.now, "/ time: ", data[11], "Weak ASK, and ask3 > bid3")
                    quote.reset()


                # 매수1~3 >>> 매도1~3, 매수막힌 후 소폭 조정
                elif ( float(data[8]) * 3 < float(data[7])
                and quote.now == quote.ask_price ):

                    bidpoint = float(data[0])
                    targetsell = float(data[0]) + 0.02
                    pw = self.api.dynamicCall("GetEncodeText(Qstring)", "6644")

                    # 매도호가에서 매수
                    str3 = ["2", "07879010", pw, "167Q3000", "1", "1", "1", "2", bidpoint, ""]
                    self.api.dynamicCall(
                        "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
                        str3)

                    # 익절
                    str4 = ["2", "07879010", pw, "167Q3000", "2", "1", "1", "2", targetsell, ""]
                    self.api.dynamicCall("DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
                        str4)

                    print("BUY at: ", quote.now, "/ time: ", data[11], "Weak ASK, BUT bid3 > ask3")
                    quote.reset()

            #총매도 압도적 우위
            if (float(data[4]) > float(data[3]) * 3
            and quote.now == quote.bid_price):

                askpoint = float(data[0])
                targetbuy = float(data[0]) - 0.02
                pw = self.api.dynamicCall("GetEncodeText(Qstring)", "6644")

                # 매수호가에서 매도
                str1 = ["2", "07879010", pw, "167Q3000", "2", "1", "1", "2", askpoint, ""]
                self.api.dynamicCall(
                    "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
                    str1)

                # 익절
                str2 = ["2", "07879010", pw, "167Q3000", "1", "1", "1", "2", targetbuy, ""]
                self.api.dynamicCall(
                    "DoRequestDomNewOrder(Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring, Qstring)",
                    str2)

                print("SELL at: ", quote.now, "/ time: ", data[11], "Strong ASK")
                quote.reset()

            return

if __name__ == "__main__":
    # import win32com.shell.shell as shell
    import os, sys

    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
    sys.exit()

