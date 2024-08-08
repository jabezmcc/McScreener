import sys, os, time, csv
from datetime import datetime, date, timezone, timedelta
from PyQt5 import QtCore
from PyQt5.QtWidgets import QApplication, QMessageBox
from PyQt5.uic import loadUiType
import pandas as pd
import numpy as np
import yahooquery as yq
from pytz import timezone
local_tz = timezone('America/New_York')
import pickle, re
import requests
import random
from bs4 import BeautifulSoup
from platform import system
import xlsxwriter

version = "0.1"

Ui_MainWindow, QMainWindow = loadUiType('McScreener_main.ui') 
Ui_downloading_hist, Qdownloading_hist = loadUiType('Downloading_hist.ui')
Ui_continuing_hist, Qcontinuing_hist = loadUiType('Continuing_hist.ui')
Ui_fund_downloading, Qfund_downloading = loadUiType('Fundamental_downloading.ui')
Ui_aboutMcScreener, QaboutMcScreener = loadUiType('aboutMcScreener.ui')

user_agents = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36'
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36'
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'
    'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.1 Safari/605.1.15'
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 13_1) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.1 Safari/605.1.15'
]

class Main(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(Main,self).__init__()
        self.setupUi(self)
        self.actionQuit.triggered.connect(self.quit)
        self.actionAbout.triggered.connect(self.openabout)
        self.actionDocumentation.triggered.connect(self.openDocs)
        # read in previous data
        if os.path.isfile('goodlist.csv'):
            self.goodlist = self.read_goodlist()
            self.stocksSaved_label.setText(str(len(self.goodlist))+" stocks from most recent performance screen saved for fundamental analysis.")
            self.nogoodlist = False
        else:
            msg2 = QMessageBox()
            msg2.setText('No performance-screened stocks found, please run a screen on the performance screen tab.')
            msg2.exec() 
            self.lastPerf.setText('No previous screened history found.')
            self.nogoodlist = True
        if os.path.isfile('pricehistdict.p'):
            self.pricehistdict = pickle.load(open("pricehistdict.p","rb"))
            self.set_lastHist()
            self.nohist = False
        else:
            msg = QMessageBox()
            msg.setText('No previous history found, please go to performance screen tab and download some price history.')
            msg.exec()
            self.perfHistLabel.setText('No previous history found.')
            self.nohist = True
        if os.path.isfile('funddata.p'):
            self.funddata = pickle.load(open("funddata.p","rb"))
            self.set_lastFund()
        else:
            msg3 = QMessageBox()
            msg3.setText('No fundamental data found.')
            msg3.exec()            
            self.lastFund.setText('No fundamental data found.')
        if os.path.isfile('pricehistdict_paused.p'):
            pause_stamp = os.path.getmtime('pricehistdict_paused.p')
            stampstr = datetime.fromtimestamp(pause_stamp).strftime('%m-%d-%Y %H:%M:%S')
            self.partialStatusLabel.setText('Found a paused historical download created '+stampstr)
            self.continuePausedButton.setStyleSheet('QPushButton:enabled { background-color #ffffff; color: #000000 }')
        else:
            self.partialStatusLabel.setText('No paused historical download found.')
            self.continuePausedButton.setStyleSheet('QPushButton:!enabled { background-color #ffffff; color: #bebebe }')
            self.continuePausedButton.setEnabled(False)
        self.downloadNewFund.clicked.connect(self.downloadnewfund)
        self.runFundScreen.clicked.connect(self.runfundscreen)
        self.resetCriteria.clicked.connect(self.resetcriteria)
        self.openResult.clicked.connect(self.openresult)
        self.downloadNewHist.clicked.connect(self.downloadnewhist)

        self.continuePausedButton.clicked.connect(self.continuepaused)
        self.runPerfScreen.clicked.connect(self.runperfscreen)
        self.udFail_label.setText('')
        self.trendFail_label.setText('')
        self.annretFail_label.setText('')
        self.stddevFail_label.setText('')
        if self.nohist or self.nogoodlist:
            self.tabWidget.setCurrentWidget(self.perfTab)
        else:
            self.tabWidget.setCurrentWidget(self.fundTab)
            
    def read_goodlist(self):
        gl = []
        tickfile = "goodlist.csv"
        with open(tickfile) as csvfile:
            contents = csv.reader(csvfile,dialect='excel')
            for row in contents:
                gl.append(row[0])
        return gl

    def set_lastHist(self):
        phdstamp = os.path.getmtime('pricehistdict.p')
        phddatestr = datetime.fromtimestamp(phdstamp).strftime('%m/%d/%Y')
        if not self.nogoodlist:
            glstamp = os.path.getmtime('goodlist.csv')
            self.gldatestr = datetime.fromtimestamp(glstamp).strftime('%m/%d/%Y')
        else:
            self.gldatestr = "--/--/----"
        nticks = str(len(self.pricehistdict))
        lastdate = self.pricehistdict[list(self.pricehistdict.keys())[0]].index[-1][1]
        if isinstance(lastdate, datetime): lastdate = lastdate.date()
        firstdate = self.pricehistdict[list(self.pricehistdict.keys())[0]].index[0][1]
        if isinstance(firstdate, datetime): firstdate = firstdate.date()
        nyears = str(int(round((lastdate-firstdate).days/365.2425)))
        if self.nogoodlist:
            self.lastPerf.setText('No performance screen results so far') 
        else:
            self.lastPerf.setText('Last performance screen on '+self.gldatestr+' yielded '+str(len(self.goodlist))+' possible stocks.')
        self.perfHistLabel.setText('Last historical download on '+phddatestr+' yielded '+str(nticks)+' possible stocks with '+str(nyears)+' years of data.') 

    def set_lastFund(self):
        dlstamp=os.path.getmtime('funddata.p')
        datestr = datetime.fromtimestamp(dlstamp).strftime('%m/%d/%Y')
        self.lastFund.setText('Last fundamental download '+datestr)

    def downloadnewfund(self):
        if os.path.isfile('fund_logfile.txt'):
            os.replace('fund_logfile.txt','fund_logfile.bak')
        fundlogfile = open('fund_logfile.txt','w')
        newfunddata={}
        ndl = 0
        self.msg = Fundamental_downloading()
        self.msg.funddl_progress.reset()
        self.msg.status_msg_label.setText('Downloading data...')        
        self.msg.OK_pushButton.setEnabled(False)
        self.msg.show() 
        nticks = len(self.goodlist)       
        for i in range(nticks):
            tick = self.goodlist[i]
            self.msg.funddl_progress.setValue(int(i*100/nticks))
            newfunddata[tick] = get_fund_data(tick)
            if  newfunddata[tick][1] != '': 
                self.msg.status_msg_label.setText('Downloaded fundamental data for '+tick)
                QApplication.processEvents()
                now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                fundlogfile.write(now+' Downloaded fundamental data for '+tick+'\n')
                ndl += 1
            else:
                self.msg.status_msg_label.setText('Unable to retrieve data for '+tick)
                QApplication.processEvents()
                now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                fundlogfile.write(now+' Unable to retrieve data for '+tick+'\n')
            time.sleep(0.1)
            if self.msg.cancelflag == True:
                now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                fundlogfile.write(now+' Fundamental download cancelled.')
                fundlogfile.close()
                self.msg.status_msg_label.setText('Fundamental download cancelled.')
                self.msg.OK_pushButton.setEnabled(True)
                return
        fundlogfile.close()
        self.msg.status_msg_label.setText('Retrieved data for '+str(ndl)+' stocks')
        QApplication.processEvents()
        self.funddata = newfunddata
        pickle.dump(self.funddata, open("funddata.p", "wb"))  
        self.set_lastFund()
        flagfile = open('needfundscreen','w')
        flagfile.write('no')
        flagfile.close()
        self.msg.OK_pushButton.setEnabled(True)

    def runfundscreen(self):
        flagfile = open('needfundscreen','r')
        flag = flagfile.read().strip()
        flagfile.close()
        if flag == 'yes':
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText('A new performance screen has been created.\n Please download new fundamental data first.')
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec()
            return
        norm = (float(self.PE_weight.text())+float(self.PB_weight.text()) \
               +float(self.PMpct_weight.text())+ float(self.EPS5yrpct_weight.text()) \
               +float(self.SG5yrpct_weight.text())+float(self.CR_weight.text())
               +float(self.DE_weight.text())+ float(self.Price_weight.text()))/100.
        #print(norm)
        #print(funddata)
        result_count=0
        self.result_list = [['Ticker','Company Name','P/E','P/B','Prft mrgn %','5 yr EPS grwth %', '5 yr sls grwth %', 'Curr. ratio', 'Debt/equity', 'Price','Score']]
        totnum = len(self.funddata)
        for tickdata in self.funddata.values():
            score = 0
            #print(tickdata)
            if float(tickdata[2]) > float(self.PE_min.text()) and \
               float(tickdata[2]) < float(self.PE_max.text()):
                score = score + float(self.PE_weight.text())
            if float(tickdata[3]) > float(self.PB_min.text()) and \
               float(tickdata[3]) < float(self.PB_max.text()):
                score = score + float(self.PB_weight.text())
            if float(tickdata[4].strip('%')) > float(self.PMpct_min.text()) and \
               float(tickdata[4].strip('%')) < float(self.PMpct_max.text()):
                score = score + float(self.PMpct_weight.text())
            if float(tickdata[5].strip('%')) > float(self.EPS5yrpct_min.text()) and \
               float(tickdata[5].strip('%')) < float(self.EPS5yrpct_max.text()):
                score = score + float(self.EPS5yrpct_weight.text())                     
            if float(tickdata[6].strip('%')) > float(self.SG5yrpct_min.text()) and \
               float(tickdata[6].strip('%')) < float(self.SG5yrpct_max.text()):
                score = score + float(self.SG5yrpct_weight.text())   
            if float(tickdata[7]) > float(self.CR_min.text()) and \
               float(tickdata[7]) < float(self.CR_max.text()):
                score = score + float(self.CR_weight.text())  
            if float(tickdata[8]) > float(self.DE_min.text()) and \
               float(tickdata[8]) < float(self.DE_max.text()):
                score = score + float(self.DE_weight.text()) 
            if float(tickdata[9]) > float(self.Price_min.text()) and \
               float(tickdata[9]) < float(self.Price_max.text()):
                score = score + float(self.Price_weight.text()) 
            score = score/norm
            score_cutoff = float(self.scoreCutoff_lineEdit.text())
            keep = score >= score_cutoff
            bank = re.match('.*[Bb]an[Cck].*',tickdata[1]) or re.match('.*Financ.*',tickdata[1])
            test = keep
            if self.banBanks_checkbox.isChecked():
                test = keep and (not bank)
            if test:
                result_count+=1
                #tickdata.append(score)
                #print(tickdata)
                self.result_list.append([tickdata[0],tickdata[1],float(tickdata[2]),float(tickdata[3]), \
                                   float(tickdata[4].strip('%')),float(tickdata[5].strip('%')),float(tickdata[6].strip('%')), \
                                   float(tickdata[7]),float(tickdata[8]),float(tickdata[9]),score])
        #for r in self.result_list:
        #    print(r)
        self.fundResult.setText('Result count: '+str(result_count)+' out of '+str(totnum)+' stocks')

    def resetcriteria(self):
        self.set_defaults(self)  

    def set_defaults(self):
        criteria={}
        criteria['PE'] = [0, 25.0, 1.0]
        criteria['PB'] = [0, 3.0, 1.0]
        criteria['PMpct'] = [5.0, 1000.0, 1.0]
        criteria['EPS5yrpct'] = [5.0, 1000.0, 1.0]
        criteria['SG5yrpct'] = [5.0, 1000.0, 1.0]
        criteria['CR'] = [2.0, 1000.0, 1.0]
        criteria['DE'] = [0, 0.5, 1.0]
        criteria['Price'] = [3.0, 2000.0, 0.3]
        
        self.PE_min.setText(str(criteria['PE'][0]))
        self.PE_max.setText(str(criteria['PE'][1]))
        self.PE_weight.setText(str(criteria['PE'][2]))
        self.PB_min.setText(str(criteria['PB'][0]))
        self.PB_max.setText(str(criteria['PB'][1]))
        self.PB_weight.setText(str(criteria['PB'][2]))
        self.PMpct_min.setText(str(criteria['PMpct'][0]))
        self.PMpct_max.setText(str(criteria['PMpct'][1]))
        self.PMpct_weight.setText(str(criteria['PMpct'][2]))
        self.EPS5yrpct_min.setText(str(criteria['EPS5yrpct'][0]))
        self.EPS5yrpct_max.setText(str(criteria['EPS5yrpct'][1]))
        self.EPS5yrpct_weight.setText(str(criteria['EPS5yrpct'][2]))
        self.SG5yrpct_min.setText(str(criteria['SG5yrpct'][0]))
        self.SG5yrpct_max.setText(str(criteria['SG5yrpct'][1]))
        self.SG5yrpct_weight.setText(str(criteria['SG5yrpct'][2]))
        self.CR_min.setText(str(criteria['CR'][0]))
        self.CR_max.setText(str(criteria['CR'][1]))
        self.CR_weight.setText(str(criteria['CR'][2]))    
        self.DE_min.setText(str(criteria['DE'][0]))
        self.DE_max.setText(str(criteria['DE'][1]))
        self.DE_weight.setText(str(criteria['DE'][2]))
        self.Price_min.setText(str(criteria['Price'][0]))
        self.Price_max.setText(str(criteria['Price'][1]))
        self.Price_weight.setText(str(criteria['Price'][2]))
        self.scoreCutoff_lineEdit.setText('80')
        self.fundResult.setText('Results found: 0')

    def openresult(self):
        try:
            dummy = self.result_list[0]
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setText('Please run a fundamental screen first') 
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec()
            return
        wb = xlsxwriter.Workbook('ScreenResults.xlsx')
        ws = wb.add_worksheet()
        ws.set_column(1,1,30)
        ws.set_column(4,4,10)
        ws.set_column(5,5,14)
        ws.set_column(6,6,14)
        ws.set_column(7,7,9)
        ws.set_column(8,8,10)
        rownum = 0
        for row in self.result_list:
            ws.write_row(rownum,0,row)
            rownum +=1
        wb.close()
        err = 0
        if system() == 'Linux':
            err = os.system('xdg-open ScreenResults.xlsx')       
        elif system() == 'Windows':
            err = os.system('start ScreenResults.xlsx')
        else:
            err = os.system('open ScreenResults.xlsx') 
        if err:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setText('Unable to find a program to open xlsx file.') 
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec()           

    def downloadnewhist(self):
        self.dl = Downloading_hist()
        while self.dl.running == True:
            QApplication.processEvents()
        if self.dl.done: 
            self.pricehistdict = self.dl.pricehistdict
        now = datetime.now().strftime('%m-%d-%Y')
        self.perfHistLabel.setText('Last historical download on '+now+' yielded '+ \
             str(len(self.pricehistdict))+' possible stocks with '+ \
             self.dl.years_to_download.text()+' years of data.') 
        
    def continuepaused(self):
        self.contpaused = Continuing_hist()
        while self.contpaused.running == True:
            QApplication.processEvents()
        if not self.contpaused.pause and not self.contpaused.cancel: 
            self.pricehistdict = self.contpaused.pricehistdict
        now = datetime.now().strftime('%m-%d-%Y')
        self.perfHistLabel.setText('Last historical download on '+now+' yielded '+ \
             str(len(self.pricehistdict))+' possible stocks with '+ \
             str(self.contpaused.numyears)+' years of data.') 

    def runperfscreen(self):
        if os.path.isfile('perf_logfile.txt'):
            os.replace('perf_logfile.txt', 'perf_logfile.bak')
        perflogfile = open('perf_logfile.txt','w')
        lastdate = self.pricehistdict[list(self.pricehistdict.keys())[0]].index[-1][1]
        firstdate = self.pricehistdict[list(self.pricehistdict.keys())[0]].index[0][1]
        # Now compare price data to S&P 500
        # First get S&P 500 data
        yqtick = yq.Ticker('^GSPC')
        SnPdf = yqtick.history(start=firstdate,end=lastdate,interval='1mo')
        #yftick =yf.Ticker('^GSPC')
        #SnPdf = yftick.history(start=firstdate,end=lastdate,interval='1mo')
        nups_SnP,ndowns_SnP,trend_SnP,ann_ret_SnP,normstd_SnP = get_stats(SnPdf)
        #print('SnP500',nups_SnP,ndowns_SnP,trend_SnP,ann_ret_SnP,normstd_SnP)   
        # Now iterate over the saved data
        now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
        perflogfile.write(now+' Starting to screen vs S&P500...\n')    
        ticklist = list(self.pricehistdict.keys())    
        self.goodlist = []
        ud_fail = 0
        trend_fail = 0
        ar_fail = 0
        sd_fail = 0
        for tick in ticklist:
            nups,ndowns,trend,ann_ret,normstd = get_stats(self.pricehistdict[tick])
            #print(tick,nups,ndowns,trend,ann_ret,normstd)
            keep_ud = True
            updown_factor = float(self.updown_lineEdit.text())             
            if self.upDown_checkbox.isChecked() and (nups < updown_factor*ndowns):
                now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                perflogfile.write(now+' '+tick+' failed nup > ndown test\n')
                keep_ud = False
                ud_fail +=1
            keep_trend = True
            SnPtrendfactor = float(self.trend_lineEdit.text()) 
            if self.trend_checkBox.isChecked() and (trend < SnPtrendfactor*trend_SnP):
                now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                perflogfile.write(now+' '+tick+' failed trend test, trend='+str(trend)+' vs. '+str(trend_SnP)+'\n')
                keep_trend = False
                trend_fail +=1
            keep_ar = True
            SnPannretfactor = float(self.annRet_lineEdit.text())
            try:
                if self.annRet_checkBox.isChecked() and (ann_ret < SnPannretfactor*ann_ret_SnP):
                    now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                    perflogfile.write(now +' '+tick+' failed annual return test\n')
                    keep_ar = False
                    ar_fail +=1
            except:
                now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                perflogfile.write(now +' '+tick+' encountered error calculating annual return\n')
                keep_ar = False
                ar_fail +=1
            keep_sd = True
            SnPstdfactor = float(self.stddev_lineEdit.text())
            if self.stddev_checkBox.isChecked() and (normstd > SnPstdfactor*normstd_SnP):
                keep_sd = False
                sd_fail +=1
            if (keep_ud and keep_trend and keep_ar and keep_sd):
                self.goodlist.append(tick)
                now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                perflogfile.write(now+' '+tick+' included in good list\n')  
        self.udFail_label.setText(str(ud_fail)+' stocks failed the up/down test.')
        self.trendFail_label.setText(str(trend_fail)+' stocks failed the trend test.')
        self.annretFail_label.setText(str(ar_fail)+' stocks failed the ann. ret. test.')
        self.stddevFail_label.setText(str(sd_fail)+' stocks failed the std. dev. test.')
        self.stocksSaved_label.setText(str(len(self.goodlist))+" stocks from most recent performance screen saved for fundamental analysis.")
        today = datetime.now().strftime('%m-%d-%Y')
        self.lastPerf.setText('Last performance screen on '+today+' yielded '+str(len(self.goodlist))+' possible stocks.')
        with open('goodlist.csv', 'w', newline = '') as csvfile:
            writer = csv.writer(csvfile, dialect = 'excel')
            for tick in self.goodlist:
                writer.writerow([tick])   
        perflogfile.close()
        flagfile = open('needfundscreen','w')
        flagfile.write('yes')
        flagfile.close()

    def openDocs(self):
        try:
            if system() == 'Linux':
                err = os.system('xdg-open documentation.pdf')       
            elif system() == 'Windows':
                err = os.system('start documentation.pdf')
            else:
                err = os.system('open documentation.pdf') 
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setText('Unable to open documentation file.') 
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec()  

    def openabout(self):
        self.aboutwin = AboutMcScreener()  

    def quit(self):
        self.close()


class Downloading_hist(Qdownloading_hist, Ui_downloading_hist):
    def __init__(self):
        super(Downloading_hist,self).__init__()
        self.setupUi(self)
        self.running = True
        self.done = False
        self.allticklist = get_tickers()
        #self.allticklist = get_tickers()[:30] #for testing
        numticks = len(self.allticklist)    
        self.foundTickers_label.setText('Found '+str(numticks)+' tickers in the NYSE and NASDAQ.')
        self.download_status_label.setText('') 
        self.noButton.clicked.connect(self.quit) 
        self.proceedButton.clicked.connect(self.download_hist_data)
        self.okButton.setEnabled(False)
        self.pauseButton.setEnabled(False)
        self.cancelButton.setEnabled(False)
        self.pause = False
        self.pauseButton.clicked.connect(self.pause_download)
        self.cancel = False

        self.cancelButton.clicked.connect(self.cancelbutton)
        self.okButton.clicked.connect(self.quit) 
        if os.path.isfile('hist_logfile.txt'):
            os.replace('hist_logfile.txt', 'hist_logfile.bak')
        self.histlogfile = open('hist_logfile.txt','w')
        self.show()
        
    def download_hist_data(self):
        # Create a dictionary of dataframes, keyed by ticker, that has desired years of price history.  
        # Reject all stocks with less data than needed.
        self.okButton.setEnabled(False)
        self.noButton.setEnabled(False)
        self.proceedButton.setEnabled(False)
        self.pauseButton.setEnabled(True)
        self.cancelButton.setEnabled(True)
        numyears = int(self.years_to_download.text())
        lastday = datetime.now()-timedelta(days=1)
        lastday = lastday.date()
        startday = lastday - timedelta(days=365.25*numyears+31)
        self.pricehistdict = {} 
        self.newhistdl_progress.reset()
        self.blacklist = []
        #max = 20 #for testing
        max = len(self.allticklist)      
        for i in range(0,max):
            self.newhistdl_progress.setValue(int((i+1)*100/max))
            if self.cancel:
                break
            if self.pause:
                pickle.dump(self.blacklist,open("blacklist.p", "wb"))
                pickle.dump([self.pricehistdict,numyears], open("pricehistdict_paused.p", "wb"))
                now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                self.histlogfile.write(now+' Paused price history download.\n')
                pause_stamp = os.path.getmtime('pricehistdict_paused.p')
                stampstr = datetime.fromtimestamp(pause_stamp).strftime('%m-%d-%Y %H:%M:%S')
                main.partialStatusLabel.setText('Found a paused historical download created '+stampstr)
                main.continuePausedButton.setEnabled(True)
                self.okButton.setEnabled(True) 
                break 
            tick = self.allticklist[i]
            self.download_status_label.setText('Attempting to download '+str(numyears)+' yrs of data for '+tick+'.....('+str(i)+'/'+str(max)+')')
            QApplication.processEvents()
            yqtick = yq.Ticker(tick)
            #yftick = yf.Ticker(tick)
            try:
                pricedf = yqtick.history(start=startday,end=lastday,interval='1mo')
                #pricedf = yftick.history(start=startday,end=lastday,interval='1mo')
                pricedf = pricedf.dropna(axis=0)
                dt1 = pricedf.index[0][1] 
                if isinstance(dt1,datetime): dt1 = dt1.date()
                dt2 = startday+timedelta(days=30)                
                if(dt1 <= dt2):
                    self.pricehistdict[tick] = pricedf
                    now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                    self.histlogfile.write(now+' Kept prices for '+tick+'\n')
                    self.download_status_label.setText('Found sufficient data for '+tick)
                else:
                    self.download_status_label.setText('Insufficient data for '+tick)
                    self.blacklist.append[tick]
                    now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                    self.histlogfile.write(now+' Not enough data for '+tick+', first date is '+dt1.strftime('%Y-%m-%d')+'\n')
                QApplication.processEvents()
            except Exception as e: 
                self.download_status_label.setText('Failed to download data for '+tick+' (see logfile.txt).')
                self.blacklist.append(tick)
                now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                self.histlogfile.write(now+' Failed to download data for '+tick+'\n'+str(e)+'\n')
                time.sleep(1.0)
                QApplication.processEvents()
            time.sleep(1.0)
        if self.cancel: 
            self.histlogfile.close()
            self.running = False
            self.close()
        if self.pause:
            msg = QMessageBox()
            msg.setText('Paused historical download and saved for continuing later')
            msg.exec()
            self.histlogfile.close()
            self.running = False
            self.close()
        self.download_status_label.setText('Kept historical data for '+str(len(self.pricehistdict))+' tickers out of '+str(max)+',')
        self.ticklist=list(self.pricehistdict.keys())
        if not self.cancel and not self.pause:
            pickle.dump(self.pricehistdict, open("pricehistdict.p", "wb")) 
            now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
            self.histlogfile.write(now+' Finished download of '+str(len(self.pricehistdict))+' stock prices.\n')
            if os.path.isfile('pricehistdict_paused.p'):
                os.replace('pricehistdict_paused.p', 'pricehistdict_paused.p.bak')
            main.partialStatusLabel.setText('No paused historical download found.')
            main.continuePausedButton.setStyleSheet('QPushButton:!enabled { background-color #ffffff; color: #bebebe }')
            main.continuePausedButton.setEnabled(False)
            main.perfHistLabel.setText('Last historical download on '+now+' yielded '+str(len(self.pricehistdict))+' possible stocks with '+str(numyears)+' years of data.')
        self.okButton.setEnabled(True) 
        self.cancelButton.setEnabled(False)
        self.pauseButton.setEnabled(False)
        self.histlogfile.close()
        self.running = False
        self.done = True
 
    def quit(self):
        self.running = False
        self.close()
    
    def cancelbutton(self):
        self.cancel = True
        
    def pause_download(self):
        self.pause = True
 
class Continuing_hist(Qcontinuing_hist, Ui_continuing_hist):
    def __init__(self):
        super(Continuing_hist,self).__init__()
        self.setupUi(self)
        self.running = True
        self.allticklist = get_tickers()
        #self.allticklist = get_tickers()[:30]  #for testing
        try:
            self.blacklist = pickle.load(open("blacklist.p","rb"))
        except:
            self.blacklist = []
        numticks = len(self.allticklist)
        self.cont_foundTickersLabel.setText('Found '+str(numticks)+' tickers in the NYSE and NASDAQ.')    
        self.cont_download_status_label.setText('')
        self.cont_noButton.clicked.connect(self.quit) 
        self.cont_proceedButton.clicked.connect(self.cont_download_hist_data)
        self.pause = False
        self.cont_pauseButton.clicked.connect(self.cont_pause_download)
        self.cancel = False
        self.cont_cancelButton.clicked.connect(self.cancelbutton)
        self.cont_okButton.clicked.connect(self.quit) 
        self.cont_okButton.setEnabled(False)
        self.cont_pauseButton.setEnabled(False)
        self.cont_cancelButton.setEnabled(False)
        self.histlogfile = open('hist_logfile.txt','a')
        paused_data = pickle.load(open("pricehistdict_paused.p","rb"))
        self.pricehistdict = paused_data[0]
        self.numyears =paused_data[1]
        self.requiringLabel.setText('Requiring at least '+str(self.numyears)+' years of data.')  
        self.doneticks = list(self.pricehistdict.keys())
        print('done ticks =',self.doneticks)
        print('blacklist =',self.blacklist)
        self.ticks_toskip = list(set(self.doneticks) | set(self.blacklist))
        self.ticks_todo = list(set(self.allticklist) - set(self.ticks_toskip))
        print('ticks_toskip =',self.ticks_toskip)
        print('ticks_todo =',self.ticks_todo)
        self.nticks_toskip = len(self.ticks_toskip)
        percent = int(self.nticks_toskip/(self.nticks_toskip+len(self.ticks_todo))*100) 
        self.conthistdl_progress.setValue(percent)
        self.tickersRemainingLabel.setText(str(len(self.ticks_todo))+' tickers remaining to download. Do you want to proceed?')
        self.show()

    def cont_download_hist_data(self):
        self.cont_noButton.setEnabled(False)
        self.cont_proceedButton.setEnabled(False)
        self.cont_pauseButton.setEnabled(True)
        self.cont_cancelButton.setEnabled(True)
        lastday = self.pricehistdict[self.doneticks[0]].index[-1][1]
        if isinstance(lastday,datetime): lastday = lastday.date()
        startday = self.pricehistdict[self.doneticks[0]].index[0][1]
        if isinstance(startday,datetime): startday = startday.date()
        max = len(self.ticks_todo)    
        for i in range(0,max):
            if self.cancel:
                break
            if self.pause:
                pickle.dump([self.pricehistdict,self.numyears], open("pricehistdict_paused.p", "wb"))
                pickle.dump(self.blacklist, open("blacklist.p", "wb"))
                now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                self.histlogfile.write(now+' Paused price history download.\n')
                pause_stamp = os.path.getmtime('pricehistdict_paused.p')
                stampstr = datetime.fromtimestamp(pause_stamp).strftime('%m-%d-%Y %H:%M:%S')
                main.partialStatusLabel.setText('Found a paused historical download created '+stampstr)
                main.continuePausedButton.setEnabled(True)
                self.cont_okButton.setEnabled(True) 
                break 
            numerator = i+1+self.nticks_toskip
            denom = self.nticks_toskip+len(self.ticks_todo)
            percent =int(numerator/denom*100)
            self.conthistdl_progress.setValue(percent)
            QApplication.processEvents()
            tick = self.ticks_todo[i]
            self.cont_download_status_label.setText('Attempting to download '+str(self.numyears)+' yrs of data for '+tick+'...('+str(i)+'/'+str(max)+')')
            yqtick = yq.Ticker(tick)
            #yftick = yf.Ticker(tick)
            try:
                pricedf = yqtick.history(start=startday,end=lastday,interval='1mo')
                pricedf = pricedf.dropna(axis=0)
                dt1 = pricedf.index[0][1]
                if isinstance(dt1,datetime): dt1 = dt1.date()
                dt2 = startday+timedelta(days=30)             
                if(dt1 <= dt2):
                    self.pricehistdict[tick] = pricedf
                    self.download_status_label.setText('Found sufficient data for '+tick)
                    now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                    self.histlogfile.write(now+' Kept prices for '+tick+'\n')
                else:
                    now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                    self.download_status_label.setText('Insufficient data for '+tick)
                    self.blacklist.append(tick)
                    self.histlogfile.write(now+' Not enough data for '+tick+', first date is '+dt1.strftime('%Y-%m-%d')+'\n')
                QApplication.processEvents()
            except Exception as e: 
                self.cont_download_status_label.setText('Failed to download data for '+tick+' (see logfile.txt).')
                self.blacklist.append(tick)
                now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
                self.histlogfile.write(now+' Failed to download data for '+tick+'...'+str(e)+'\n')
                QApplication.processEvents()
                time.sleep(1.0)
            time.sleep(1.0)
        if self.cancel: 
            self.histlogfile.close()
            self.running = False
            self.close()
        if self.pause:
            msg = QMessageBox()
            msg.setText('Paused historical download and saved for continuing later')
            msg.exec()
            self.histlogfile.close()
            self.running = False
            self.close()
        self.cont_download_status_label.setText('Kept historical data for '+str(len(self.pricehistdict))+' tickers out of '+str(max)+',')
        self.ticklist=list(self.pricehistdict.keys())
        if not self.cancel and not self.pause:
            pickle.dump(self.pricehistdict, open("pricehistdict.p", "wb"))  
            now = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
            self.histlogfile.write(now+' Finished download of '+str(len(self.pricehistdict))+' stock prices.\n')
            if os.path.isfile('pricehistdict_paused.p'):
                os.replace('pricehistdict_paused.p', 'pricehistdict_paused.p.bak')
            self.tickersRemainingLabel.setText('Finished download.')
            main.partialStatusLabel.setText('No paused historical download found.')
            main.continuePausedButton.setStyleSheet('QPushButton:!enabled { background-color #ffffff; color: #bebebe }')
            main.continuePausedButton.setEnabled(False)
            main.perfHistLabel.setText('Last historical download on '+now+' yielded '+str(len(self.pricehistdict))+' possible stocks with '+str(self.numyears)+' years of data.')
        self.cont_okButton.setEnabled(True)
        self.cont_noButton.setEnabled(False)
        self.cont_proceedButton.setEnabled(False)
        self.cont_pauseButton.setEnabled(False)
        self.cont_cancelButton.setEnabled(False)
        self.histlogfile.close()
        self.running = False
        
    def quit(self):
        self.running = False
        self.close()
    
    def cancelbutton(self):
        self.cancel = True
        
    def cont_pause_download(self):
        self.pause = True       
 
class Fundamental_downloading(Qfund_downloading, Ui_fund_downloading):
    def __init__(self):
        super(Fundamental_downloading,self).__init__()
        self.setupUi(self) 
        self.status_msg_label.setText('')    
        self.cancelflag = False
        self.OK_pushButton.clicked.connect(self.closeme)
        self.Cancel_pushButton.clicked.connect(self.cancel)

    def closeme(self):
        self.close()
        
    def cancel(self):
        self.cancelflag = True
        
class AboutMcScreener(QaboutMcScreener, Ui_aboutMcScreener):
    def __init__(self):
        super(AboutMcScreener,self).__init__()
        self.setupUi(self)
        self.versionLabel.setText('Version '+version)
        self.OKButt.clicked.connect(self.closeme)
        self.licenseButt.clicked.connect(self.show_license)
        self.show()
        
    def show_license(self):
        try:
            if system() == 'Linux':
                err = os.system('xdg-open license.txt')       
            elif system() == 'Windows':
                err = os.system('start license.txt')
            else:
                err = os.system('open license.txt') 
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setText('Unable to open license file.') 
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec()  
            
    def closeme(self):
        self.close()
        
def get_stats(tickdf):
    pricedate = [x[1] for x in list(tickdf.index.values)]
    closeprice = list(tickdf['adjclose'])
    yrlyrise = []
    ndowns = 0
    nups= 0
    iyear = 0
    for i in range(len(pricedate)):
        if (not (i % 12) and (i>0)):
            yrlyrise.append((closeprice[i] - closeprice[i-12])/closeprice[i])
            if (yrlyrise[iyear] <= 0):
                ndowns+=1
            else:
                nups+=1
            iyear+=1
    logcloseprice = np.log(closeprice)
    datenum = []
    for x in pricedate:
        if isinstance(x,datetime): x = x.date()
        datenum.append(float((x-date(1970,1,1)).days))
    p1, resid, rank, singvals, rcond = np.polyfit(datenum, logcloseprice, 1,full='True')
    trend = p1[0]
    ann_ret = (closeprice[-1]/closeprice[0])**(365.25/(datenum[-1] -datenum[0]))-1
    normstd = np.sqrt(resid[0]/(np.mean(closeprice)*(len(closeprice)-1)))
    return nups, ndowns, trend, ann_ret, normstd

def get_fund_data(tick):
    url = 'https://finviz.com/quote.ashx?t='+tick
    try:
        headers = {'User-Agent': random.choice(user_agents)}
        doc = requests.get(url,headers=headers).text
        soup = BeautifulSoup(doc,"html.parser")
        txt=soup.get_text()
        m = re.search(r'[A-Z]{2,}\s+\-\s+(.*)\s+Stock',txt)
        if(m):
            title = m.group(1)
        else:
            title=''
        m = re.search(r'P/E(-*\d+\.\d+)',txt)
        if(m):
            PE = m.group(1)
        else:
            PE='0'
        m = re.search(r'P/B(-*\d+\.\d+)',txt)
        if(m):
            PB = m.group(1)
        else:
            PB='0'
        m = re.search(r'Profit Margin(-*\d+\.\d+)',txt)
        if(m):
            ProfitMargin = str(m.group(1))+'%'
        else:
            ProfitMargin='0'
        m = re.search(r'EPS past 5Y(-*\d+\.\d+)',txt)
        if(m):
            EPSgrowth = str(m.group(1))+'%'
        else:
            EPSgrowth = '0'
        m = re.search(r'Current Ratio(-*\d+\.\d+)',txt)
        if(m):
            CurrentRatio = m.group(1)
        else:
            CurrentRatio='0'
        m = re.search(r'Debt/Eq(-*\d+\.\d+)',txt)
        if(m):
            DebtToEquity = '%.2f' % (float(m.group(1)))
        else:
            DebtToEquity = '0'
        m = re.search(r'Sales past 5Y(-*\d+\.\d+)',txt)
        if(m):
            CashFlowGrowth = str(m.group(1))+'%'
        else:
            CashFlowGrowth = '0'
        m = re.search(r'^Price(-*\d+\.\d+)',txt,re.MULTILINE)
        if(m):
            Price = m.group(1)
        else:
            Price='0'
        return [tick,title,PE,PB,ProfitMargin,EPSgrowth,CashFlowGrowth,CurrentRatio,DebtToEquity,Price]   
    except:
        return [tick,'','0', '0', '0', '0', '0', '0', '0', '0']    

def get_tickers():
    # see https://www.nasdaqtrader.com/trader.aspx?id=symboldirdefs#other for details of data source
    # First download all NASDAQ tickers
    ndfile  = 'ftp://ftp.nasdaqtrader.com/symboldirectory/nasdaqlisted.txt'
    try:
        nasdaq_ticks =  pd.read_csv(ndfile,delimiter="|")
    except:
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText('Unable to download NASDAQ tickers from '+ndfile)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec()
        return []
    # Keep only domestic stocks, clean out anything global
    nasdaq_ticks = nasdaq_ticks.loc[nasdaq_ticks['Market Category']=='S']
    # Remove "test issues"
    nasdaq_ticks = nasdaq_ticks.loc[nasdaq_ticks['Test Issue']=='N']
    # Keep only normnal financial status
    nasdaq_ticks = nasdaq_ticks.loc[nasdaq_ticks['Financial Status']=='N']
    # Clear out ETFs
    nasdaq_ticks = nasdaq_ticks.loc[nasdaq_ticks['ETF']=='N']
    # Clear out any ticker with a period or $ in its name
    nasdaq_ticks = nasdaq_ticks.loc[nasdaq_ticks['Symbol'].str.contains('.',regex=False)==False]
    nasdaq_ticks = nasdaq_ticks.loc[nasdaq_ticks['Symbol'].str.contains('$',regex=False)==False]
    allticklist = nasdaq_ticks['Symbol'].to_list()
    numNASDAQ = len(allticklist)
    # Now add the NYSE tickers
    otherfile  = 'ftp://ftp.nasdaqtrader.com/symboldirectory/otherlisted.txt'
    try:
        nyse_ticks = pd.read_csv(otherfile,delimiter="|")
    except:
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText('Unable to download NYSE tickers from '+otherfile)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec()
        return []
    # Keep only NYSE issues
    nyse_ticks = nyse_ticks.loc[nyse_ticks['Exchange']=='N']
    # Remove ETFs
    nyse_ticks = nyse_ticks.loc[nyse_ticks['ETF']=='N']
    # Remove "test issues"
    nyse_ticks = nyse_ticks.loc[nyse_ticks['Test Issue']=='N']
    # Clear out any ticker with a period or $ in its name
    nyse_ticks = nyse_ticks.loc[nyse_ticks['ACT Symbol'].str.contains('.',regex=False)==False]
    nyse_ticks = nyse_ticks.loc[nyse_ticks['ACT Symbol'].str.contains('$',regex=False)==False]
    allticklist = allticklist + nyse_ticks['ACT Symbol'].to_list()
    return allticklist




if __name__=="__main__":
    app = QApplication(sys.argv)
    main = Main()
    main.set_defaults()
    main.show()
    sys.exit(app.exec())
