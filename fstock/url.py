import urllib2, urllib,lxml.html
import os, sys
import re
import html2csv
import fileinput

class FYClass:
    def __init__(self):
        self.F = ['netSalesOps', 'otherOpsIncome', 
                  'stockChange', 'rawMaterial', 
                  'purchaseGoods', 'employeeCost',
                  'depreciation', 'otherExpenditure',
                  'totalExpenditure', 'profitBeforOtherIncomeIntrestExceptional',
                  'otherIncome', 'profitBeforeIntrestExceptional',
                  'interest', 'profileafterIntrestbeforeExceptional',
                  'exceptional', 'profitBeforeTax',
                  'tax', 'profitAfterTax',
                  'extra']
        self.FY = {}
        self.FS = ['Net Sales/Income from Operations', 
                   'Other Operating Income',
                   'Increase/Decrease in Stock in trade and work in progress',
                   'Consumption of Raw Materials',
                   'Purchase of traded goods',
                   'Employees Cost',
                   'Depreciation',
                   'Other Expenditure',
                   'Total Expenditure',
                   'Profit from Operations before Other Income, Interest & Exceptional Items',
                   'Other Income',
                   'Profit before Interest & Exceptional Items',
                   'Interest',
                   'Profit after Interest but before Exceptional Items',
                   'Exceptional items',
                   'Profit(+)/Loss(-) from Ordinary Activities before tax',
                   'Tax Expense',
                   'Net Profit(+)/Loss(-) from Ordinary Activities after tax',
                   'Extraordinary Items',
                   'Net Profit (+) / Loss (-) for the period'
                   'Dividend (%)'
                   ]
    def getYS(self,keys,k):
        vs = ";"
        for i in keys:
            if k in self.FY[i]:
                vs = vs  + str(self.FY[i][k])+';'
            else:
                vs = vs + ';'
        return vs
        
    def FYPrint(self,fileName):
        if(os.path.isfile(fileName+'_FY.csv')):
            os.remove(fileName+'_FY.csv')
        fd = open(fileName+'_FY.csv','a')
        keys = sorted(self.FY.keys())
        ks = ""
        for i in keys:
            ks =ks+";"+i 
        fd.write("Description"+ks+"\n")
        for i in range(0,20):
            fd.write(self.FS[i]+self.getYS(keys, self.FS[i])+"\n")
        
        
        fd.close()


    
    

class NSEResults:
    def __init__(self,stockName):
        self.pwd = os.path.dirname(__file__)
        self.stockName = stockName
        self.linkDir = self.pwd+'\\link\\'
        self.htmlFile = self.pwd+'\\'+stockName+'.html'
        self.FYList = {}
        self.FY = FYClass()
        #self.linkOut = self.linkDir+stockName+'.txt'
        #if not os.path.exists(self.linkDir):
        #    os.makedirs(self.linkDir)
        
        self.resultsY = {}
        self.resultsQ = {}
        self.resultsH = {}   
        self.resultsO = {}
            
        self.getLinkWeb()
        self.getResultLink()
        self.getResultY()
        self.FY.FYPrint(self.stockName)
    def getVal(self,vs):
        ns = vs.split('"')[3]
        ns = ns.replace(",", "")
        return ns
    def mStr(self,str):
        return re.escape(str)
    def readResultY(self,csvFile):
        start = False
        dataFile = open(csvFile,'r')
        y = csvFile.split('.')[0].split('_')[0]
        self.FY.FY[y] = {}
        for line in dataFile:
            if(re.match('^\"Description(.*)',line,re.M|re.I)):
                start = True
            elif(re.match('^\"Segment\ Reporting(.*)',line,re.M|re.I)):
                start = False
            if(start ):
                if(re.match('^\"'+ str(self.mStr(self.FY.FS[0])) +'(.*)',line,re.M|re.I) or re.match('^\"Net\ Income\ from\ sales(.*)',line,re.M|re.I) ):
                    self.FY.FY[y][self.FY.FS[0]] = float(self.getVal(line))
                for i in range(1,20):
                    if(re.match('^\"'+ str(self.mStr(self.FY.FS[i])) +'(.*)',line,re.M|re.I)  ):
                        val = self.getVal(line)
                        if(val.isdigit()):
                            self.FY.FY[y][self.FY.FS[i]] = float(val)
                        else:
                            self.FY.FY[y][self.FY.FS[i]] = val
                
                    
        dataFile.close()
        
            
            
        
                    
    def storeLinks(self,link):      
        if(re.match('(.*)[0-9][0-9]\-APR\-(.*)\-MAR\-(.*)',link,re.M|re.I)):
            y= link.split('-')[2][:-2]
            self.resultsY[y] = 'http://www.nseindia.com'+link

        elif(re.match('(.*)[0-9][0-9][0-9][0-9][0-9][0-9]\-[A-Z][A-Z][A-Z]\-[0-9][0-9][0-9][0-9]Q',link,re.M|re.I)):
            q = link.split('-')[4][0:6]
            self.resultsQ[q] = 'http://www.nseindia.com'+link
        elif(re.match('(.*)[0-9][0-9][0-9][0-9][0-9][0-9]\-[A-Z][A-Z][A-Z]\-[0-9][0-9][0-9][0-9]H',link,re.M|re.I)):
            h =  link.split('-')[4][0:5]
            self.resultsH[h] = 'http://www.nseindia.com'+link
        else:
            self.resultsO = 'http://www.nseindia.com'+link
    def getResultY(self):
        vbs=self.pwd+'\\getResult.vbs'
        for k,y in enumerate(self.resultsY):
            print(self.resultsY[y])
            os.system('cscript.exe  %s %s %s %s'%(vbs,'\"'+self.resultsY[y]+'\"',y,self.stockName))
            h2c = y+"_"+self.stockName+".html"
            if(os.path.isfile(y+"_"+self.stockName+".csv")):
                os.remove(y+"_"+self.stockName+".csv")
            html2csv.html2csv_by_PP(h2c)
            if(os.path.isfile(h2c)):
                os.remove(h2c)
                
            #for line in fileinput.FileInput(y+"_"+self.stockName+".csv", inplace=1):
            #    if(re.match('^\ (.*)',line,re.M|re.I)):
                    
                
            #self.readResultY(y+"_"+self.stockName+".csv")
            #if(os.path.isfile(y+"_"+self.stockName+".csv")):
            #    os.remove(y+"_"+self.stockName+".csv")
            
            
    def getResultLink(self): 
        if not os.path.isfile(self.htmlFile) :
            print("html file not exist\n")
            exit()
        fd = open(self.htmlFile,'r')
        
        for line in fd:
            word = line.split()
            for w in word:
                if(re.match('href\=\/marketinfo(.*)',w,re.M|re.I) ):
                    r = w.split('=',1)[1]
                    l = r.split('>',1)[0]
                    self.storeLinks(l)
                    
        fd.close()
        if(os.path.isfile(self.htmlFile)):
            os.remove(self.htmlFile)
    def getLinkWeb(self):
        vbs=self.pwd+'\\getLink.vbs'
        os.system('cscript.exe  %s %s'%(vbs,self.stockName))
        
if __name__ == '__main__':

    NSEResults('WIPRO')
    #NSEResults('TCS')
    #FY = FYClass()
    #print(FY.FY)
