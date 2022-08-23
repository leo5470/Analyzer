import requests as req
import settings, json

url = 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest'
headers = {
    "X-CMC_PRO_API_KEY": settings.key,
    "Accepts" : 'application/json'
}

def name(volmin, volmax):
    parameters = {
        'convert':'USD',
        'volume_24h_min':'',
        'volume_24h_max':'',
        'percent_change_24h_min':'-20',
        'percent_change_24h_max':'20',
    }
    parameters['volume_24h_max'] = str(volmax * 1000000)
    parameters['volume_24h_min'] = str(volmin * 1000000)

    session = req.Session()
    session.headers.update(headers)
    data = session.get(url, params=parameters)
    get = json.loads(data.text)
    namelist = []
    try:
        i = 0
        while True:
            fdv = get['data'][i]['quote']['USD']['fully_diluted_market_cap']
            mc = get['data'][i]['quote']['USD']['market_cap']
            if fdv is not None and fdv != 0 and mc != 0:
                NameResult = get['data'][i]['name'] #dictionary
                namelist.append(NameResult)
            i += 1
    except:
        pass
    return namelist

def tvl(volmin, volmax):
    parameters = {
        'convert':'USD',
        'volume_24h_min':'',
        'volume_24h_max':'',
        'percent_change_24h_min':'-20',
        'percent_change_24h_max':'20',
    }
    parameters['volume_24h_max'] = str(volmax * 1000000)
    parameters['volume_24h_min'] = str(volmin * 1000000)

    session = req.Session()
    session.headers.update(headers)
    data = session.get(url, params=parameters)
    get = json.loads(data.text)
    tvllist = []
    try:
        i = 0
        while True:
            fdv = get['data'][i]['quote']['USD']['fully_diluted_market_cap']
            mc = get['data'][i]['quote']['USD']['market_cap']
            if fdv is not None and fdv != 0 and mc != 0:
                TVLResult = get['data'][i]['quote']['USD']['volume_24h'] #dictionary
                tvllist.append(TVLResult)
            i += 1
    except:
        pass
    return tvllist

def change(volmin, volmax):
    parameters = {
        'convert':'USD',
        'volume_24h_min':'',
        'volume_24h_max':'',
        'percent_change_24h_min':'-20',
        'percent_change_24h_max':'20',
    }
    parameters['volume_24h_max'] = str(volmax * 1000000)
    parameters['volume_24h_min'] = str(volmin * 1000000)

    session = req.Session()
    session.headers.update(headers)
    data = session.get(url, params=parameters)
    get = json.loads(data.text)
    changelist = []
    try:
        i = 0
        while True:
            fdv = get['data'][i]['quote']['USD']['fully_diluted_market_cap']
            mc = get['data'][i]['quote']['USD']['market_cap']
            if fdv is not None and fdv != 0 and mc != 0:
                changeResult = get['data'][i]['quote']['USD']['percent_change_24h'] #dictionary
                changelist.append(changeResult)
            i += 1
    except:
        pass
    return changelist
    
def fdv(volmin, volmax):
    parameters = {
        'convert':'USD',
        'volume_24h_min':'',
        'volume_24h_max':'',
        'percent_change_24h_min':'-20',
        'percent_change_24h_max':'20',
    }
    parameters['volume_24h_max'] = str(volmax * 1000000)
    parameters['volume_24h_min'] = str(volmin * 1000000)

    session = req.Session()
    session.headers.update(headers)
    data = session.get(url, params=parameters)
    get = json.loads(data.text)
    fdvlist = []
    try:
        i = 0
        while True:
            fdv = get['data'][i]['quote']['USD']['fully_diluted_market_cap']
            mc = get['data'][i]['quote']['USD']['market_cap']
            if fdv is not None and fdv != 0 and mc != 0:
                fdvlist.append(fdv)
            i += 1
    except:
        pass
    return fdvlist

def mc(volmin, volmax):
    parameters = {
        'convert':'USD',
        'volume_24h_min':'',
        'volume_24h_max':'',
        'percent_change_24h_min':'-20',
        'percent_change_24h_max':'20',
    }
    parameters['volume_24h_max'] = str(volmax * 1000000)
    parameters['volume_24h_min'] = str(volmin * 1000000)

    session = req.Session()
    session.headers.update(headers)
    data = session.get(url, params=parameters)
    get = json.loads(data.text)
    mclist = []
    try:
        i = 0
        while True:
            fdv = get['data'][i]['quote']['USD']['fully_diluted_market_cap']
            mc = get['data'][i]['quote']['USD']['market_cap']
            if fdv is not None and fdv != 0 and mc != 0:
                mclist.append(mc)
            i += 1
    except:
        pass
    return mclist
