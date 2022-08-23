import requests as req
import settings, json
from openpyxl.styles import Font

#openpyxl Font
yesFont = Font(color='00FF00', bold='True')
noFont = Font(color='FF0000')

#CoinMarketCap request
url = 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest'
headers = {
    "X-CMC_PRO_API_KEY": settings.key,
    "Accepts" : 'application/json'
}

class get():
    def __init__(self, volmin, volmax) -> None:
        self.volmin = volmin
        self.volmax = volmax
        self.parameters = {
            'convert':'USD',
            'volume_24h_min':'',
            'volume_24h_max':'',
            'percent_change_24h_min':'-20',
            'percent_change_24h_max':'20',
        }
        self.parameters['volume_24h_max'] = str(volmax * 1000000)
        self.parameters['volume_24h_min'] = str(volmin * 1000000)
        self.session = req.Session()
        self.session.headers.update(headers)
        self.data = self.session.get(url, params=self.parameters)
        self.getdata = json.loads(self.data.text)

    def name(self):
        namelist = []
        try:
            i = 0
            while True:
                fdv = self.getdata['data'][i]['quote']['USD']['fully_diluted_market_cap']
                mc = self.getdata['data'][i]['quote']['USD']['market_cap']
                if fdv is not None and fdv != 0 and mc != 0:
                    NameResult = self.getdata['data'][i]['name'] #dictionary
                    namelist.append(NameResult)
                i += 1
        except:
            pass
        return namelist

    def tvl(self):
        tvllist = []
        try:
            i = 0
            while True:
                fdv = self.getdata['data'][i]['quote']['USD']['fully_diluted_market_cap']
                mc = self.getdata['data'][i]['quote']['USD']['market_cap']
                if fdv is not None and fdv != 0 and mc != 0:
                    TVLResult = self.getdata['data'][i]['quote']['USD']['volume_24h'] #dictionary
                    tvllist.append(TVLResult)
                i += 1
        except:
            pass
        return tvllist
    def change(self):
        changelist = []
        try:
            i = 0
            while True:
                fdv = self.getdata['data'][i]['quote']['USD']['fully_diluted_market_cap']
                mc = self.getdata['data'][i]['quote']['USD']['market_cap']
                if fdv is not None and fdv != 0 and mc != 0:
                    changeResult = self.getdata['data'][i]['quote']['USD']['percent_change_24h'] #dictionary
                    changelist.append(changeResult)
                i += 1
        except:
            pass
        return changelist

    def FDV(self):
        fdvlist = []
        try:
            i = 0
            while True:
                fdv = self.getdata['data'][i]['quote']['USD']['fully_diluted_market_cap']
                mc = self.getdata['data'][i]['quote']['USD']['market_cap']
                if fdv is not None and fdv != 0 and mc != 0:
                    fdvlist.append(fdv)
                i += 1
        except:
            pass
        return fdvlist
    
    def MC(self):
        mclist = []
        try:
            i = 0
            while True:
                fdv = self.getdata['data'][i]['quote']['USD']['fully_diluted_market_cap']
                mc = self.getdata['data'][i]['quote']['USD']['market_cap']
                if fdv is not None and fdv != 0 and mc != 0:
                    mclist.append(mc)
                i += 1
        except:
            pass
        return mclist

def handle(volmin, volmax, fillExcel):
    i = 2
    cur = get(volmin, volmax)
    for name in cur.name():
        fillExcel[f'A{i}'] = name
        i += 1
    i = 2
    for tvl in cur.tvl():
        fillExcel[f'B{i}'] = tvl
        i += 1
    i = 2
    for change in cur.change():
        fillExcel[f'C{i}'] = change
        i += 1
    i = 2
    for fdv in cur.FDV():
        fillExcel[f'D{i}'] = fdv
        i += 1
    i = 2
    for mc in cur.MC():
        fillExcel[f'E{i}'] = mc
        i += 1
    return i - 2

def longTerm(judge, total):
    for t in range(2, total + 2):
        judge[f'F{t}'] = (judge[f'E{t}'].value) * 100 / (judge[f'D{t}'].value)
        if judge[f'F{t}'].value > 40:
            judge[f'F{t}'].font = yesFont
        else:
            judge[f'F{t}'].font = noFont

def underrated(judge, total):
    for t in range(2, total + 2):
        judge[f'G{t}'] = (judge[f'E{t}'].value) / (judge[f'B{t}'].value)
        if judge[f'G{t}'].value < 1:
            judge[f'G{t}'].font = yesFont
        else:
            judge[f'G{t}'].font = noFont
