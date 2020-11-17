import sys

import pandas
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import *
import win32com.client
from pandas import Series, DataFrame
import locale

from daishin import balance
from daishin import load_10


class Form(QtWidgets.QDialog):
    def __init__(self, parent=None):
        QtWidgets.QDialog.__init__(self, parent)


        # plus 상태 체크
        if balance.InitPlusCheck() == False:
            exit()

        self.ui = uic.loadUi("hoga_1.ui", self)
        self.ui.show()
        self.objMst = load_10.CpRPCurrentPrice()
        self.item = load_10.stockPricedData()

        # 일자별
        self.objWeek = load_10.CpWeekList()
        self.rpWeek = DataFrame()  # 일자별 데이터프레임

        # 시간대별
        self.rpStockBid = DataFrame()
        self.objStockBid = load_10.CpStockBid()
        self.todayIndex = 0

        # 6033 잔고 object
        self.obj6033 = balance.Cp6033()
        self.jangoData = {}
        self.balance = {}

        self.isSB = False
        self.objCur = {}

        # 현재가 정보
        self.curDatas = {}
        self.objRPCur = balance.CpRPCurrentPrice()

        # 실시간 주문 체결
        self.objConclusion = balance.CpPBConclusion()

        self.setCode("000660")

        # 잔고 요청
        self.requestJango()
        self.displyBlanceStock()

    @pyqtSlot()
    def slot_codeupdate(self):
        code = self.ui.editCode.toPlainText()
        self.setCode(code)

    def slot_codechanged(self):
        code = self.ui.editCode.toPlainText()
        self.setCode(code)

    def monitorPriceChange(self):
        self.displyHoga()
        self.updateWeek()
        self.updateStockBid()

    def monitorBlanceChange(self):
        self.updateBlance()

    def monitorOfferbidChange(self):
        self.displyHoga()

    def updateBlance(self):
        self.ui.label_b_acc.setText(str(self.balance['계좌명']))
        self.ui.label_b_price.setText(str(self.balance['평가금액']))
        self.ui.label_b_profit.setText(str(round(self.balance['수익률'],4) * 100))

    def displyBlanceStock(self):
        rowcnt = len(self.jangoData)
        if rowcnt == 0:
            return
        self.ui.tableBlance.setRowCount(rowcnt)

        nRow = 0

        for index, row in self.jangoData.items():
            # 행 내에 표시할 데이터 - 컬럼순
            datas = [row['종목명'], row['대비'], '-', row['현재가'], row['잔고수량'], row['장부가'], row['손익단가'], row['평가금액'], row['평가손익'], row['수익률']]
            for col in range(len(datas)):
                val = ''
                ###
                val = str(col)

                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.ui.tableBlance.setItem(nRow, col, item)
            nRow += 1
        self.tableBlance.resizeColumnsToContents()

    def setCode(self, code):
        if len(code) < 6:
            return

        print(code)
        if not (code[0] == "A"):
            code = "A" + code

        name = load_10.g_objCodeMgr.CodeToName(code)
        if len(name) == 0:
            print("종목코드 확인")
            return

        self.ui.label_name.setText(name)

        if (self.objMst.Request(code, self.item, self) == False):
            return
        self.displyHoga()

        # 일자별
        self.ui.tableWeek.clearContents()
        if (self.objWeek.Request(code, self) == True):
            print(self.rpWeek)
            self.displyWeek()

        # 시간대별
        self.ui.tableStockBid.clearContents()
        if (self.objStockBid.Request(code, self) == True):
            self.displyStockBid()

    # 10차 호가 UI 채우기
    def displyHoga(self):
        self.ui.label_offer10.setText(format(self.item.offer[9], ','))
        self.ui.label_offer9.setText(format(self.item.offer[8], ','))
        self.ui.label_offer8.setText(format(self.item.offer[7], ','))
        self.ui.label_offer7.setText(format(self.item.offer[6], ','))
        self.ui.label_offer6.setText(format(self.item.offer[5], ','))
        self.ui.label_offer5.setText(format(self.item.offer[4], ','))
        self.ui.label_offer4.setText(format(self.item.offer[3], ','))
        self.ui.label_offer3.setText(format(self.item.offer[2], ','))
        self.ui.label_offer2.setText(format(self.item.offer[1], ','))
        self.ui.label_offer1.setText(format(self.item.offer[0], ','))

        self.ui.label_offer_v10.setText(format(self.item.offervol[9], ','))
        self.ui.label_offer_v9.setText(format(self.item.offervol[8], ','))
        self.ui.label_offer_v8.setText(format(self.item.offervol[7], ','))
        self.ui.label_offer_v7.setText(format(self.item.offervol[6], ','))
        self.ui.label_offer_v6.setText(format(self.item.offervol[5], ','))
        self.ui.label_offer_v5.setText(format(self.item.offervol[4], ','))
        self.ui.label_offer_v4.setText(format(self.item.offervol[3], ','))
        self.ui.label_offer_v3.setText(format(self.item.offervol[2], ','))
        self.ui.label_offer_v2.setText(format(self.item.offervol[1], ','))
        self.ui.label_offer_v1.setText(format(self.item.offervol[0], ','))

        self.ui.label_bid10.setText(format(self.item.bid[9], ','))
        self.ui.label_bid9.setText(format(self.item.bid[8], ','))
        self.ui.label_bid8.setText(format(self.item.bid[7], ','))
        self.ui.label_bid7.setText(format(self.item.bid[6], ','))
        self.ui.label_bid6.setText(format(self.item.bid[5], ','))
        self.ui.label_bid5.setText(format(self.item.bid[4], ','))
        self.ui.label_bid4.setText(format(self.item.bid[3], ','))
        self.ui.label_bid3.setText(format(self.item.bid[2], ','))
        self.ui.label_bid2.setText(format(self.item.bid[1], ','))
        self.ui.label_bid1.setText(format(self.item.bid[0], ','))

        self.ui.label_bid_v10.setText(format(self.item.bidvol[9], ','))
        self.ui.label_bid_v9.setText(format(self.item.bidvol[8], ','))
        self.ui.label_bid_v8.setText(format(self.item.bidvol[7], ','))
        self.ui.label_bid_v7.setText(format(self.item.bidvol[6], ','))
        self.ui.label_bid_v6.setText(format(self.item.bidvol[5], ','))
        self.ui.label_bid_v5.setText(format(self.item.bidvol[4], ','))
        self.ui.label_bid_v4.setText(format(self.item.bidvol[3], ','))
        self.ui.label_bid_v3.setText(format(self.item.bidvol[2], ','))
        self.ui.label_bid_v2.setText(format(self.item.bidvol[1], ','))
        self.ui.label_bid_v1.setText(format(self.item.bidvol[0], ','))

        cur = self.item.cur
        diff = self.item.diff
        diffp = self.item.diffp
        if (self.item.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            cur = self.item.expcur
            diff = self.item.expdiff
            diffp = self.item.expdiffp

        strcur = format(cur, ',')
        if (self.item.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            strcur = "*" + strcur

        curcolor = self.item.getCurColor()
        self.ui.label_cur.setStyleSheet(curcolor)
        self.ui.label_cur.setText(strcur)
        strdiff = str(diff) + "  " + format(diffp, '.2f')
        strdiff += "%"
        self.ui.label_diff.setText(strdiff)
        self.ui.label_diff.setStyleSheet(curcolor)

        self.ui.label_totoffer.setText(format(self.item.totOffer, ','))
        self.ui.label_totbid.setText(format(self.item.totBid, ','))



    # 일자별 리스트 UI 채우기
    def displyWeek(self):
        rowcnt = len(self.rpWeek.index)
        if rowcnt == 0:
            return
        self.ui.tableWeek.setRowCount(rowcnt)

        nRow = 0

        for index, row in self.rpWeek.iterrows():
            datas = [index, row['close'], row['diff'], row['diffp'], row['vol'], row['open'], row['high'], row['low'],
                     row['for_v'], row['for_d'], row['for_p']]
            for col in range(len(datas)):
                val = ''
                if (col == 0):  # 일자
                    # 20170929 ==> 2017/09/29
                    yyyy = int(datas[col] / 10000)
                    mm = int(datas[col] - (yyyy * 10000))
                    dd = mm % 100
                    mm = mm / 100
                    val = '%04d/%02d/%02d' % (yyyy, mm, dd)
                elif (col == 3 or col == 10):  # 대비율
                    val = locale.format('%.2f', datas[col], 1)
                    val += "%"

                else:
                    val = locale.format('%d', datas[col], 1)

                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.ui.tableWeek.setItem(nRow, col, item)

            if (nRow == 0):
                self.todayIndex = index
            nRow += 1

            self.tableWeek.resizeColumnsToContents()
        return

    # 일자별 리스트 UI 채우기 - 오늘 날짜 업데이트
    def updateWeek(self):
        rowcnt = len(self.rpWeek.index)
        if rowcnt == 0:
            return

        # 오늘 날짜 데이터 업데이트
        self.rpWeek.set_value(self.todayIndex, 'close', self.item.cur)
        self.rpWeek.set_value(self.todayIndex, 'open', self.item.open)
        self.rpWeek.set_value(self.todayIndex, 'high', self.item.high)
        self.rpWeek.set_value(self.todayIndex, 'low', self.item.low)
        self.rpWeek.set_value(self.todayIndex, 'vol', self.item.vol)
        self.rpWeek.set_value(self.todayIndex, 'diff', self.item.diff)
        self.rpWeek.set_value(self.todayIndex, 'diffp', self.item.diffp)

        datas = [self.todayIndex, self.item.cur, self.item.diff, self.item.diffp, self.item.vol,
                 self.item.open, self.item.high, self.item.low]
        for col in range(len(datas)):
            val = ''
            if (col == 0):  # 일자
                # 20170929 ==> 2017/09/29
                yyyy = int(datas[col] / 10000)
                mm = int(datas[col] - (yyyy * 10000))
                dd = mm % 100
                mm = mm / 100
                val = '%04d/%02d/%02d' % (yyyy, mm, dd)
            elif (col == 3):  # 대비율
                val = locale.format('%.2f', datas[col], 1)
                val += "%"

            else:
                val = locale.format('%d', datas[col], 1)

            item = QTableWidgetItem(val)
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.ui.tableWeek.setItem(0, col, item)

        return

    # 시간대별 리스트 UI  채우기
    def displyStockBid(self):
        rowcnt = len(self.rpStockBid.index)
        if rowcnt == 0:
            return
        self.ui.tableStockBid.setRowCount(rowcnt)

        nRow = 0

        for index, row in self.rpStockBid.iterrows():
            # 행 내에 표시할 데이터 - 컬럼 순
            datas = [row['time'], row['cur'], row['diff'], row['offer'], row['bid'], row['vol'], row['tvol'],
                     row['tvol'], row['volstr']]
            market = row['market']
            for col in range(len(datas)):
                val = ''
                if col == 0:  # 시각
                    # 155925 ==> 15:59:25
                    hh = int(datas[col] / 10000)
                    mm = int(datas[col] - (hh * 10000))
                    ss = mm % 100
                    mm = mm / 100
                    val = '%02d:%02d:%02d' % (hh, mm, ss)
                elif col == 6:  # 체결매도
                    market = row['flag']
                    if (market == "체결매도"):
                        val = locale.format('%d', datas[col], 1)
                elif col == 7:  # 체결매수
                    market = row['flag']
                    if (market == "체결매수"):
                        val = locale.format('%d', datas[col], 1)
                elif col == 8:  # 체결강도
                    val = locale.format('%.2f', datas[col], 1)
                elif col == 1:  # 현재가
                    val = locale.format('%d', datas[col], 1)
                    if (market == "예상체결"):
                        val = '*' + val
                else:  # 기타
                    val = locale.format('%d', datas[col], 1)
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.ui.tableStockBid.setItem(nRow, col, item)
            nRow += 1

        self.tableStockBid.resizeColumnsToContents()
        return

    def updateStockBid(self):
        rowcnt = len(self.rpStockBid.index)
        if rowcnt == 0:
            return
        if (self.item.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            return

        buyvol = sellvol = 0
        if self.item.volFlag == ord('1'):
            buyvol = self.item.tvol
        if self.item.volFlag == ord('2'):
            sellvol = self.item.tvol
        line = DataFrame({"time": self.item.time,
                          "cur": self.item.cur,
                          "diff": self.item.diff,
                          "offer": self.item.offer[0],
                          "bid": self.item.bid[0],
                          "vol": self.item.vol,
                          "tvol": buyvol,
                          "tvol": sellvol,
                          "volstr": self.item.volstr},
                         index=[0])

        self.rpStockBid = pandas.concat([line, self.rpStockBid.ix[:]]).reset_index(drop=True)

        # 행 내에 표시할 데이터 - 컬럼 순
        datas = [self.item.time, self.item.cur, self.item.diff, self.item.offer[0], self.item.bid[0],
                 self.item.vol, sellvol, buyvol, self.item.volstr]
        self.ui.tableStockBid.insertRow(0)
        for col in range(len(datas)):
            val = ''
            if col == 0:  # 시각
                # 155925 ==> 15:59:25
                hh = int(datas[col] / 10000)
                mm = int(datas[col] - (hh * 10000))
                ss = mm % 100
                mm = mm / 100
                val = '%02d:%02d:%02d' % (hh, mm, ss)
            elif col == 6:  # 체결매도
                val = locale.format('%d', datas[col], 1)
            elif col == 7:  # 체결매수
                val = locale.format('%d', datas[col], 1)
            elif col == 8:  # 체결강도
                val = locale.format('%.2f', datas[col], 1)
            else:  # 기타
                val = locale.format('%d', datas[col], 1)

            item = QTableWidgetItem(val)
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.ui.tableStockBid.setItem(0, col, item)

        return
    def StopSubscribe(self):
        if self.isSB:
            for key, obj in self.objCur.items():
                obj.Unsubscribe()
            self.objCur = {}

        self.isSB = False
        self.objConclusion.Unsubscribe()


    def requestJango(self):
        self.StopSubscribe();

        # 주식 잔고 통신
        if self.obj6033.requestJango(self) == False:
            return

        # 잔고 현재가 통신
        codes = set()
        for code, value in self.jangoData.items():
            codes.add(code)

        objMarkeyeye = balance.CpMarketEye()
        codelist = list(codes)
        if (objMarkeyeye.Request(codelist, self) == False):
            exit()

        # 실시간 현재가  요청
        cnt = len(codelist)
        for i in range(cnt):
            code = codelist[i]
            self.objCur[code] = balance.CpPBStockCur()
            self.objCur[code].Subscribe(code, self)
        self.isSB = True

        # 실시간 주문 체결 요청
        self.objConclusion.Subscribe('', self)

    # 실시간 주문 체결 처리 로직
    def updateJangoCont(self, pbCont):
        # 주문 체결에서 들어온 신용 구분 값 ==> 잔고 구분값으로 치환
        dicBorrow = {
            '현금': ord(' '),
            '유통융자': ord('Y'),
            '자기융자': ord('Y'),
            '주식담보대출': ord('B'),
            '채권담보대출': ord('B'),
            '매입담보대출': ord('M'),
            '플러스론': ord('P'),
            '자기대용융자': ord('I'),
            '유통대용융자': ord('I'),
            '기타': ord('Z')
        }

        # 잔고 리스트 map 의 key 값
        # key = (pbCont['종목코드'], dicBorrow[pbCont['현금신용']], pbCont['대출일'])
        # key = pbCont['종목코드']
        code = pbCont['종목코드']

        # 접수, 거부, 확인 등은 매도 가능 수량만 업데이트 한다.
        if pbCont['체결플래그'] == '접수' or pbCont['체결플래그'] == '거부' or pbCont['체결플래그'] == '확인':
            if (code not in self.jangoData):
                return
            self.jangoData[code]['매도가능'] = pbCont['매도가능수량']
            return

        if (pbCont['체결플래그'] == '체결'):
            if (code not in self.jangoData):  # 신규 잔고 추가
                if (pbCont['체결기준잔고수량'] == 0):
                    return
                print('신규 잔고 추가', code)
                # 신규 잔고 추가
                item = {}
                item['종목코드'] = pbCont['종목코드']
                item['종목명'] = pbCont['종목명']
                item['현금신용'] = dicBorrow[pbCont['현금신용']]
                item['대출일'] = pbCont['대출일']
                item['잔고수량'] = pbCont['체결기준잔고수량']
                item['매도가능'] = pbCont['매도가능수량']
                item['장부가'] = pbCont['장부가']
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']

                print('신규 현재가 요청', code)
                self.objRPCur.Request(code, self)
                self.objCur[code] = balance.CpPBStockCur()
                self.objCur[code].Subscribe(code, self)

                item['현재가'] = self.curDatas[code]['cur']
                item['대비'] = self.curDatas[code]['diff']
                item['거래량'] = self.curDatas[code]['vol']

                self.jangoData[code] = item

            else:
                # 기존 잔고 업데이트
                item = self.jangoData[code]
                item['종목코드'] = pbCont['종목코드']
                item['종목명'] = pbCont['종목명']
                item['현금신용'] = dicBorrow[pbCont['현금신용']]
                item['대출일'] = pbCont['대출일']
                item['잔고수량'] = pbCont['체결기준잔고수량']
                item['매도가능'] = pbCont['매도가능수량']
                item['장부가'] = pbCont['장부가']
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']

                # 잔고 수량이 0 이면 잔고 제거
                if item['잔고수량'] == 0:
                    del self.jangoData[code]
                    self.objCur[code].Unsubscribe()
                    del self.objCur[code]

        return

    # 실시간 현재가 처리 로직
    def updateJangoCurPBData(self, curData):
        code = curData['code']
        self.curDatas[code] = curData
        self.upjangoCurData(code)

    def upjangoCurData(self, code):
        # 잔고에 동일 종목을 찾아 업데이트 하자 - 현재가/대비/거래량/평가금액/평가손익
        curData = self.curDatas[code]
        item = self.jangoData[code]
        item['현재가'] = curData['cur']
        item['대비'] = curData['diff']
        item['거래량'] = curData['vol']


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    w = Form()
    sys.exit(app.exec())