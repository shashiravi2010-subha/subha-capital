#!/usr/bin/env python3
"""
SUBHA CAPITAL — Market Data Fetcher v7
4-Phase Pre-Market Intelligence System:
  8:45 AM → Gift Nifty + Pre-market sentiment
  9:00 AM → Pre-open + Sector scan  
  9:30 AM → ORB + Stock selection (ROCKERS)
  9:45 AM → FINNIFTY ORB + Options signal
"""

import os,json,math,pyotp,requests,pytz,time,openpyxl
from openpyxl.styles import PatternFill,Font,Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime,timedelta

CLIENT_ID   = os.environ["ANGEL_CLIENT_ID"]
API_KEY     = os.environ["ANGEL_API_KEY"]
TOTP_SECRET = os.environ["ANGEL_TOTP_SECRET"]
PIN         = os.environ["ANGEL_PIN"]
CAPITAL     = 100000
RISK_PCT    = 1.0
IST         = pytz.timezone("Asia/Kolkata")
BASE        = "https://apiconnect.angelone.in"

# ── NIFTY 100 UNIVERSE ────────────────────────────────────────────────
STOCKS = {
    "HDFCBANK":   {"token":"1333",  "sec":"BANKING","slp":0.9},
    "ICICIBANK":  {"token":"4963",  "sec":"BANKING","slp":0.85},
    "SBIN":       {"token":"3045",  "sec":"BANKING","slp":1.0},
    "AXISBANK":   {"token":"5900",  "sec":"BANKING","slp":1.0},
    "KOTAKBANK":  {"token":"1922",  "sec":"BANKING","slp":0.85},
    "BANKBARODA": {"token":"4668",  "sec":"BANKING","slp":1.2},
    "INDUSINDBK": {"token":"5258",  "sec":"BANKING","slp":1.1},
    "CANBK":      {"token":"10794", "sec":"BANKING","slp":1.2},
    "BAJFINANCE": {"token":"317",   "sec":"FINANCE","slp":0.85},
    "BAJAJFINSV": {"token":"16675", "sec":"FINANCE","slp":0.90},
    "HDFCLIFE":   {"token":"467",   "sec":"FINANCE","slp":0.85},
    "SBILIFE":    {"token":"21808", "sec":"FINANCE","slp":0.85},
    "ICICIPRULI": {"token":"18529", "sec":"FINANCE","slp":0.85},
    "CHOLAFIN":   {"token":"685",   "sec":"FINANCE","slp":1.0},
    "MUTHOOTFIN": {"token":"13923", "sec":"FINANCE","slp":1.0},
    "PFC":        {"token":"14299", "sec":"FINANCE","slp":1.0},
    "RECLTD":     {"token":"21614", "sec":"FINANCE","slp":1.0},
    "SHRIRAMFIN": {"token":"4306",  "sec":"FINANCE","slp":1.0},
    "INFY":       {"token":"1594",  "sec":"IT","slp":0.80},
    "TCS":        {"token":"11536", "sec":"IT","slp":0.75},
    "WIPRO":      {"token":"3787",  "sec":"IT","slp":0.90},
    "HCLTECH":    {"token":"7229",  "sec":"IT","slp":0.85},
    "TECHM":      {"token":"13538", "sec":"IT","slp":1.0},
    "LTIM":       {"token":"17818", "sec":"IT","slp":0.90},
    "MPHASIS":    {"token":"4503",  "sec":"IT","slp":1.0},
    "PERSISTENT": {"token":"18365", "sec":"IT","slp":1.0},
    "COFORGE":    {"token":"11543", "sec":"IT","slp":1.0},
    "OFSS":       {"token":"10738", "sec":"IT","slp":0.85},
    "MARUTI":     {"token":"10999", "sec":"AUTO","slp":0.85},
    "MM":         {"token":"2031",  "sec":"AUTO","slp":0.80},
    "TATAMOTORS": {"token":"3456",  "sec":"AUTO","slp":1.2},
    "BAJAJ-AUTO": {"token":"16669", "sec":"AUTO","slp":0.90},
    "HEROMOTOCO": {"token":"1348",  "sec":"AUTO","slp":0.90},
    "EICHERMOT":  {"token":"910",   "sec":"AUTO","slp":0.90},
    "TVSMOTOR":   {"token":"3518",  "sec":"AUTO","slp":1.0},
    "ASHOKLEY":   {"token":"212",   "sec":"AUTO","slp":1.2},
    "SUNPHARMA":  {"token":"3351",  "sec":"PHARMA","slp":0.75},
    "DRREDDY":    {"token":"881",   "sec":"PHARMA","slp":0.80},
    "CIPLA":      {"token":"694",   "sec":"PHARMA","slp":0.85},
    "DIVISLAB":   {"token":"10940", "sec":"PHARMA","slp":0.85},
    "AUROPHARMA": {"token":"275",   "sec":"PHARMA","slp":1.0},
    "ALKEM":      {"token":"13634", "sec":"PHARMA","slp":0.90},
    "BIOCON":     {"token":"524",   "sec":"PHARMA","slp":1.1},
    "UPL":        {"token":"11287", "sec":"PHARMA","slp":1.1},
    "TATASTEEL":  {"token":"3499",  "sec":"METAL","slp":1.0},
    "JSWSTEEL":   {"token":"11723", "sec":"METAL","slp":1.0},
    "HINDALCO":   {"token":"1363",  "sec":"METAL","slp":1.1},
    "COALINDIA":  {"token":"20374", "sec":"METAL","slp":1.0},
    "VEDL":       {"token":"3063",  "sec":"METAL","slp":1.2},
    "NMDC":       {"token":"15332", "sec":"METAL","slp":1.1},
    "HINDUNILVR": {"token":"1394",  "sec":"FMCG","slp":0.70},
    "ITC":        {"token":"1660",  "sec":"FMCG","slp":0.75},
    "BRITANNIA":  {"token":"547",   "sec":"FMCG","slp":0.80},
    "NESTLEIND":  {"token":"17963", "sec":"FMCG","slp":0.70},
    "DABUR":      {"token":"772",   "sec":"FMCG","slp":0.80},
    "GODREJCP":   {"token":"10099", "sec":"FMCG","slp":0.85},
    "MARICO":     {"token":"4067",  "sec":"FMCG","slp":0.85},
    "COLPAL":     {"token":"1367",  "sec":"FMCG","slp":0.80},
    "RELIANCE":   {"token":"2885",  "sec":"ENERGY","slp":0.80},
    "ONGC":       {"token":"2475",  "sec":"ENERGY","slp":0.90},
    "BPCL":       {"token":"526",   "sec":"ENERGY","slp":1.0},
    "NTPC":       {"token":"11630", "sec":"ENERGY","slp":0.90},
    "POWERGRID":  {"token":"14977", "sec":"ENERGY","slp":0.85},
    "TATAPOWER":  {"token":"3426",  "sec":"ENERGY","slp":1.1},
    "IOC":        {"token":"1624",  "sec":"ENERGY","slp":1.0},
    "DLF":        {"token":"14732", "sec":"REALTY","slp":1.0},
    "GODREJPROP": {"token":"9742",  "sec":"REALTY","slp":1.0},
    "OBEROIRLTY": {"token":"20316", "sec":"REALTY","slp":1.0},
    "LT":         {"token":"11483", "sec":"INFRA","slp":0.85},
    "ADANIPORTS": {"token":"15083", "sec":"INFRA","slp":1.0},
    "SIEMENS":    {"token":"3150",  "sec":"INFRA","slp":0.90},
    "ABB":        {"token":"13",    "sec":"INFRA","slp":0.90},
    "BEL":        {"token":"383",   "sec":"INFRA","slp":1.0},
    "HAL":        {"token":"10455", "sec":"INFRA","slp":1.0},
    "BHEL":       {"token":"438",   "sec":"INFRA","slp":1.2},
    "TITAN":      {"token":"3506",  "sec":"CONSUMER","slp":0.85},
    "ASIANPAINT": {"token":"236",   "sec":"CONSUMER","slp":0.80},
    "PIDILITIND": {"token":"2664",  "sec":"CONSUMER","slp":0.80},
    "HAVELLS":    {"token":"10350", "sec":"CONSUMER","slp":0.90},
    "VOLTAS":     {"token":"3718",  "sec":"CONSUMER","slp":1.0},
    "DMART":      {"token":"14413", "sec":"CONSUMER","slp":0.85},
    "TRENT":      {"token":"3721",  "sec":"CONSUMER","slp":0.90},
    "ZOMATO":     {"token":"5097",  "sec":"CONSUMER","slp":1.2},
    "NYKAA":      {"token":"21431", "sec":"CONSUMER","slp":1.2},
    "IRCTC":      {"token":"13611", "sec":"CONSUMER","slp":1.0},
    "ULTRACEMCO": {"token":"11532", "sec":"CEMENT","slp":0.85},
    "GRASIM":     {"token":"1232",  "sec":"CEMENT","slp":0.90},
    "AMBUJACEM":  {"token":"1270",  "sec":"CEMENT","slp":1.0},
    "SHREECEM":   {"token":"3103",  "sec":"CEMENT","slp":0.85},
    "BHARTIARTL": {"token":"10604", "sec":"TELECOM","slp":0.85},
    "INDUSTOWER": {"token":"17491", "sec":"TELECOM","slp":1.0},
}

INDICES = {"NIFTY50":"26000","FINNIFTY":"26037"}

SECTOR_STOCKS = {
    "BANKING":  ["HDFCBANK","ICICIBANK","SBIN","AXISBANK","KOTAKBANK","INDUSINDBK"],
    "FINANCE":  ["BAJFINANCE","BAJAJFINSV","CHOLAFIN","PFC","RECLTD","MUTHOOTFIN"],
    "IT":       ["INFY","TCS","WIPRO","HCLTECH","TECHM","LTIM"],
    "AUTO":     ["MARUTI","MM","TATAMOTORS","BAJAJ-AUTO","HEROMOTOCO"],
    "PHARMA":   ["SUNPHARMA","DRREDDY","CIPLA","DIVISLAB","AUROPHARMA"],
    "METAL":    ["TATASTEEL","JSWSTEEL","HINDALCO","COALINDIA","VEDL"],
    "FMCG":     ["HINDUNILVR","ITC","BRITANNIA","NESTLEIND","DABUR"],
    "ENERGY":   ["RELIANCE","ONGC","BPCL","NTPC","POWERGRID"],
    "REALTY":   ["DLF","GODREJPROP","OBEROIRLTY"],
    "INFRA":    ["LT","ADANIPORTS","BEL","HAL","BHEL"],
    "CONSUMER": ["TITAN","ASIANPAINT","DMART","TRENT","ZOMATO"],
    "CEMENT":   ["ULTRACEMCO","GRASIM","AMBUJACEM"],
    "TELECOM":  ["BHARTIARTL","INDUSTOWER"],
}

# ── TIME PHASE DETECTION ─────────────────────────────────────────────
def get_phase(ist):
    h,m = ist.hour, ist.minute
    if h<9 or (h==8 and m>=45):
        if h==8 and m>=45: return "PREMARKET_845"
    if h==9 and m<15: return "PREMARKET_900"
    if h==9 and m<30: return "MARKET_OPEN_915"
    if h==9 and m<45: return "ORB_930"
    if h>=9 and h<15: return "ORB_945_PLUS"
    return "CLOSED"

def get_phase_label(phase):
    labels = {
        "PREMARKET_845":  "8:45 AM — Pre-Market Brief",
        "PREMARKET_900":  "9:00 AM — Pre-Open Analysis",
        "MARKET_OPEN_915":"9:15 AM — Market Open Watch",
        "ORB_930":        "9:30 AM — ORB Confirmation",
        "ORB_945_PLUS":   "9:45 AM — FINNIFTY ORB Signal",
        "CLOSED":         "Market Closed — EOD Data",
    }
    return labels.get(phase,"—")

# ── GIFT NIFTY / SGX DATA ────────────────────────────────────────────
def fetch_gift_nifty():
    """Fetch Gift Nifty / SGX Nifty from multiple sources"""
    gift = {"ltp":0,"chg":0,"gap":0,"gap_pct":0,"bias":"—","source":"—"}
    hdrs = {
        "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept":"application/json, text/plain, */*",
        "Accept-Language":"en-US,en;q=0.9",
        "Referer":"https://www.google.com",
    }
    sources = [
        ("https://query1.finance.yahoo.com/v8/finance/chart/%5ENSEI?interval=1d&range=1d","yahoo1"),
        ("https://query2.finance.yahoo.com/v8/finance/chart/%5ENSEI?interval=1d&range=1d","yahoo2"),
        ("https://query1.finance.yahoo.com/v8/finance/chart/NIFTYBEES.NS?interval=1d&range=1d","niftybees"),
        ("https://query2.finance.yahoo.com/v8/finance/chart/%5ECNXNIFTY?interval=1d&range=1d","cnx"),
    ]
    try:
        for url, src in sources:
            try:
                r = requests.get(url, headers=hdrs, timeout=10)
                if r.status_code == 200:
                    d = r.json()
                    result = d.get("chart",{}).get("result",[])
                    if result:
                        meta = result[0].get("meta",{})
                        p = float(meta.get("regularMarketPrice",0) or 0)
                        prev = float(meta.get("previousClose",0) or meta.get("chartPreviousClose",0) or 0)
                        if p > 0:
                            chg = round((p-prev)/prev*100,2) if prev>0 else 0
                            gift["ltp"] = round(p,2)
                            gift["chg"] = chg
                            gift["source"] = src
                            print(f"Gift Nifty OK: {p} {chg}% from {src}")
                            break
            except Exception as e:
                print(f"Gift Nifty {src} error: {e}")
                continue

        if gift["ltp"] > 0:
            gift["gap"] = round(gift["chg"] * gift["ltp"] / 100, 0)
            gift["gap_pct"] = gift["chg"]
            if gift["chg"] > 0.5:    gift["bias"] = "GAP UP ▲ BULLISH"
            elif gift["chg"] > 0.2:  gift["bias"] = "SLIGHT GAP UP"
            elif gift["chg"] < -0.5: gift["bias"] = "GAP DOWN ▼ BEARISH"
            elif gift["chg"] < -0.2: gift["bias"] = "SLIGHT GAP DOWN"
            else:                     gift["bias"] = "FLAT — WAIT"
        else:
            print("Gift Nifty: All sources failed")
    except Exception as e:
        print(f"Gift Nifty fetch error: {e}")
    return gift

def fetch_vix():
    """Fetch India VIX from multiple sources"""
    vix={"ltp":0,"chg":0,"prev":0,"status":"—","source":"—"}
    hdrs_common={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                 "Accept":"application/json, text/plain, */*",
                 "Accept-Language":"en-US,en;q=0.9"}
    try:
        # Method 1: Yahoo Finance ^INDIAVIX
        try:
            s=requests.Session()
            s.headers.update(hdrs_common)
            # Get cookies first
            s.get("https://finance.yahoo.com",timeout=6)
            r=s.get("https://query1.finance.yahoo.com/v8/finance/chart/%5EINDIAVIX?interval=1d&range=2d",
                    timeout=8)
            if r.ok:
                d=r.json()
                result=d.get("chart",{}).get("result",[])
                if result:
                    meta=result[0].get("meta",{})
                    ltp=float(meta.get("regularMarketPrice",0))
                    prev=float(meta.get("previousClose",0))
                    if ltp>0:
                        vix["ltp"]=round(ltp,2)
                        vix["prev"]=round(prev,2)
                        vix["chg"]=round(ltp-prev,2) if prev>0 else 0
                        vix["source"]="Yahoo Finance"
                        vix["status"]="<13 Safe" if ltp<13 else "13-16 Caution" if ltp<16 else ">16 HIGH RISK"
                        print(f"VIX from Yahoo: {ltp}")
                        return vix
        except Exception as e:
            print(f"VIX Yahoo error: {e}")

        # Method 2: NSE India with session
        try:
            s=requests.Session()
            s.headers.update({**hdrs_common,"Referer":"https://www.nseindia.com/"})
            # Get session cookies
            s.get("https://www.nseindia.com",timeout=8)
            time.sleep(1)
            r=s.get("https://www.nseindia.com/api/allIndices",timeout=8)
            if r.ok:
                data=r.json()
                for idx in data.get("data",[]):
                    name=str(idx.get("indexSymbol","") or idx.get("index",""))
                    if "VIX" in name.upper():
                        ltp=float(idx.get("last",0) or idx.get("indexValue",0) or 0)
                        prev=float(idx.get("previousClose",0) or idx.get("yearHigh",0) or 0)
                        if ltp>0:
                            vix["ltp"]=round(ltp,2)
                            vix["prev"]=round(prev,2)
                            vix["chg"]=round(float(idx.get("change",0)),2)
                            vix["source"]="NSE India"
                            vix["status"]="<13 Safe" if ltp<13 else "13-16 Caution" if ltp<16 else ">16 HIGH RISK"
                            print(f"VIX from NSE: {ltp}")
                            return vix
        except Exception as e:
            print(f"VIX NSE error: {e}")

        # Method 3: Try alternate Yahoo endpoint
        try:
            r=requests.get(
                "https://query2.finance.yahoo.com/v10/finance/quoteSummary/%5EINDIAVIX?modules=price",
                headers=hdrs_common,timeout=8)
            if r.ok:
                d=r.json()
                price=d.get("quoteSummary",{}).get("result",[{}])[0].get("price",{})
                ltp=float(price.get("regularMarketPrice",{}).get("raw",0))
                prev=float(price.get("regularMarketPreviousClose",{}).get("raw",0))
                if ltp>0:
                    vix["ltp"]=round(ltp,2)
                    vix["prev"]=round(prev,2)
                    vix["chg"]=round(ltp-prev,2) if prev>0 else 0
                    vix["source"]="Yahoo v10"
                    vix["status"]="<13 Safe" if ltp<13 else "13-16 Caution" if ltp<16 else ">16 HIGH RISK"
                    print(f"VIX from Yahoo v10: {ltp}")
                    return vix
        except Exception as e:
            print(f"VIX Yahoo v10 error: {e}")

    except Exception as e:
        print(f"VIX fetch error: {e}")

    print("VIX: All sources failed — showing N/A")
    return vix

# ── API HELPERS ───────────────────────────────────────────────────────
def hdrs(tok=None):
    h={"Content-Type":"application/json","Accept":"application/json",
       "X-UserType":"USER","X-SourceID":"WEB",
       "X-ClientLocalIP":"192.168.1.1","X-ClientPublicIP":"106.193.147.98",
       "X-MACAddress":"fe80::216e:6507:4b90:3719","X-PrivateKey":API_KEY}
    if tok: h["Authorization"]=f"Bearer {tok}"
    return h

def login():
    totp=pyotp.TOTP(TOTP_SECRET).now()
    print(f"TOTP: {totp}")
    r=requests.post(f"{BASE}/rest/auth/angelbroking/user/v1/loginByPassword",
        json={"clientcode":CLIENT_ID,"password":PIN,"totp":totp},
        headers=hdrs(),timeout=15)
    d=r.json()
    if d.get("status") and d.get("data",{}).get("jwtToken"):
        print("✅ Login OK")
        return d["data"]["jwtToken"]
    raise Exception(f"Login failed: {d.get('message')}")

def get_quotes(tok, tokens):
    all_data=[]
    for i in range(0,len(tokens),50):
        batch=tokens[i:i+50]
        try:
            r=requests.post(f"{BASE}/rest/secure/angelbroking/market/v1/quote/",
                json={"mode":"FULL","exchangeTokens":{"NSE":batch}},
                headers=hdrs(tok),timeout=15)
            d=r.json()
            if d.get("status") and d.get("data"):
                all_data.extend(d["data"].get("fetched",[]))
        except Exception as e:
            print(f"Quote error: {e}")
        time.sleep(0.3)
    return all_data

def get_candles(tok, sym_tok, interval, fr, to):
    try:
        r=requests.post(f"{BASE}/rest/secure/angelbroking/historical/v1/getCandleData",
            json={"exchange":"NSE","symboltoken":sym_tok,"interval":interval,
                  "fromdate":fr,"todate":to},
            headers=hdrs(tok),timeout=15)
        return r.json()
    except:
        return {}

# ── INDICATORS ────────────────────────────────────────────────────────
def ema(prices,n):
    if not prices: return 0
    if len(prices)<n: return prices[-1]
    k=2/(n+1); e=sum(prices[:n])/n
    for p in prices[n:]: e=p*k+e*(1-k)
    return round(e,2)

def rsi(prices,n=14):
    if len(prices)<n+1: return 50
    g,l=[],[]
    for i in range(1,len(prices)):
        d=prices[i]-prices[i-1]
        g.append(max(d,0)); l.append(max(-d,0))
    ag=sum(g[:n])/n; al=sum(l[:n])/n
    for i in range(n,len(g)):
        ag=(ag*(n-1)+g[i])/n; al=(al*(n-1)+l[i])/n
    return 100 if al==0 else round(100-100/(1+ag/al))

def macd_bull(prices):
    if len(prices)<26: return True
    k12,k26=2/13,2/27; e12=e26=prices[0]
    for p in prices:
        e12=p*k12+e12*(1-k12); e26=p*k26+e26*(1-k26)
    return e12>e26

def st_bull(candles):
    if not candles: return True
    c=candles[-1]
    return float(c[4])>(float(c[2])+float(c[3]))/2

def vol_ratio(vols):
    if len(vols)<2: return 1.0
    avg=sum(vols[:-1])/len(vols[:-1])
    return round(vols[-1]/avg,1) if avg>0 else 1.0

def calc_vwap(candles):
    if not candles: return 0
    cum_tpv=0; cum_vol=0
    for c in candles[-20:]:
        tp=(float(c[2])+float(c[3])+float(c[4]))/3
        vol=float(c[5])
        cum_tpv+=tp*vol; cum_vol+=vol
    return round(cum_tpv/cum_vol,2) if cum_vol>0 else 0

# ── EXCEL JOURNAL ─────────────────────────────────────────────────────
def update_excel(signals,ist):
    fname="trading_journal.xlsx"
    try:
        wb=openpyxl.load_workbook(fname)
        ws=wb.active
    except:
        wb=openpyxl.Workbook(); ws=wb.active
        ws.title="SUBHA CAPITAL"
        headers=["Date","Stock","Sector","Side","Score","Pre-Mkt Price",
                 "Open","High","Low","VWAP","GTT Trigger","Limit Price",
                 "Stop Loss","Target 1","Target 2","Qty","Max Risk ₹",
                 "Bias","Phase","Result","P&L ₹","Notes"]
        hf=PatternFill(start_color="0D1221",end_color="0D1221",fill_type="solid")
        hfont=Font(bold=True,color="F0A500",size=10)
        for col,h in enumerate(headers,1):
            cell=ws.cell(row=1,column=col,value=h)
            cell.fill=hf; cell.font=hfont
            cell.alignment=Alignment(horizontal='center',vertical='center')
        ws.row_dimensions[1].height=28
        widths=[12,12,10,8,8,14,10,10,10,10,12,12,12,10,10,6,10,10,10,10,10,15]
        for i,w in enumerate(widths,1):
            ws.column_dimensions[get_column_letter(i)].width=w

    date_str=ist.strftime("%d-%b-%Y")
    for s in signals:
        isL=s["side"]=="LONG"
        risk=abs(s["entry"]-s["sl"])
        row=[date_str,s["sym"],s["sec"],s["side"],f"{s['score']}/6",
             s["ltp"],s.get("open",""),s.get("high",""),s.get("low",""),s.get("vwap",""),
             s.get("gtt_trigger",s["entry"]),s.get("gtt_limit",s["entry"]),
             s["sl"],s["t1"],s["t2"],s["qty"],round(risk*s["qty"],2),
             s.get("bias",""),s.get("phase",""),"","",""]
        rn=ws.max_row+1
        fc="002D1C" if isL else "2D0010"
        rf=PatternFill(start_color=fc,end_color=fc,fill_type="solid")
        for col,val in enumerate(row,1):
            cell=ws.cell(row=rn,column=col,value=val)
            cell.fill=rf; cell.alignment=Alignment(horizontal='center')
            if col==4: cell.font=Font(color="00D4A0" if isL else "FF4560",bold=True)
            elif col in [11,12]: cell.font=Font(color="F0A500",bold=True)
            elif col==13: cell.font=Font(color="FF4560",bold=True)
            elif col in [14,15]: cell.font=Font(color="00D4A0",bold=True)
    wb.save(fname)
    print(f"✅ Excel: {ws.max_row-1} signals")

# ── MAIN ──────────────────────────────────────────────────────────────
def main():
    ist=datetime.now(IST)
    phase=get_phase(ist)
    phase_label=get_phase_label(phase)
    print(f"═══ SUBHA CAPITAL v7 — {ist.strftime('%d %b %Y %H:%M IST')} ═══")
    print(f"Phase: {phase} — {phase_label}")

    mkt_open=(ist.hour>9 or (ist.hour==9 and ist.minute>=15)) and \
             (ist.hour<15 or (ist.hour==15 and ist.minute<30))

    # ── GIFT NIFTY (8:45 AM phase) ───────────────────────────────────
    print("Fetching Gift Nifty...")
    gift = fetch_gift_nifty()
    print(f"Gift Nifty: {gift['ltp']} {gift['chg']}% — {gift['bias']}")

    print("Fetching India VIX...")
    vix_data = fetch_vix()
    print(f"VIX: {vix_data['ltp']} | Status: {vix_data['status']} | Source: {vix_data['source']}")

    # ── ANGEL ONE LOGIN ───────────────────────────────────────────────
    jwt=login()

    # ── FETCH ALL QUOTES ──────────────────────────────────────────────
    stk_toks=[v["token"] for v in STOCKS.values()]
    idx_toks=list(INDICES.values())
    fetched=get_quotes(jwt,stk_toks+idx_toks)

    Q={}
    for q in fetched:
        t=q.get("symbolToken","")
        for sym,info in STOCKS.items():
            if info["token"]==t: Q[sym]=q; break
        for idx,tok in INDICES.items():
            if tok==t: Q[idx]=q; break
    print(f"Quotes: {len(Q)}")

    # ── PRICE FUNCTIONS ───────────────────────────────────────────────
    def F(s,field,default=0):
        return float(Q.get(s,{}).get(field,default) or default)

    def ltp(s):
        v=F(s,"ltp")
        return round(v,2) if v>0 else round(F(s,"close"),2)

    def prev_close(s):
        return round(F(s,"close"),2)

    def pct_chg(s):
        pc=F(s,"percentChange")
        if pc!=0: return round(pc,2)
        cur=ltp(s); prev=prev_close(s)
        if cur>0 and prev>0 and cur!=prev:
            return round((cur-prev)/prev*100,2)
        return 0.0

    def pts_chg(s):
        nc=F(s,"netChange")
        if nc!=0: return round(nc,2)
        return round(ltp(s)-prev_close(s),2)

    def day_high(s): return round(F(s,"high"),2) or ltp(s)
    def day_low(s):  return round(F(s,"low"),2) or ltp(s)
    def day_open(s): return round(F(s,"open"),2) or ltp(s)

    # ── INDEX DATA ────────────────────────────────────────────────────
    nifty_ltp  = ltp("NIFTY50")
    nifty_chg  = pct_chg("NIFTY50")
    nifty_pts  = pts_chg("NIFTY50")
    nifty_prev = prev_close("NIFTY50")
    fn_ltp     = ltp("FINNIFTY")
    fn_chg     = pct_chg("FINNIFTY")
    fn_pts     = pts_chg("FINNIFTY")
    fn_prev    = prev_close("FINNIFTY")

    print(f"Nifty: {nifty_ltp} | chg: {nifty_chg}% | pts: {nifty_pts}")
    print(f"FINNIFTY: {fn_ltp} | chg: {fn_chg}% | pts: {fn_pts}")

    # Bias
    bias="BULLISH" if nifty_chg>0.3 else "BEARISH" if nifty_chg<-0.3 else "NEUTRAL"

    # ── PRE-MARKET ANALYSIS (9:00 AM phase) ──────────────────────────
    # Top gainers and losers from Nifty 100
    pre_market = []
    for sym in list(STOCKS.keys())[:20]:  # Check top 20 stocks
        p=ltp(sym); c=pct_chg(sym)
        if p>0:
            pre_market.append({"sym":sym,"ltp":p,"chg":c,"sec":STOCKS[sym]["sec"]})
    pre_market.sort(key=lambda x:x["chg"],reverse=True)
    top_gainers=pre_market[:3]
    top_losers=pre_market[-3:][::-1]

    # ── FINNIFTY ATM + ORB ────────────────────────────────────────────
    fn_atm=round(fn_ltp/50)*50 if fn_ltp>0 else 0
    fn_orb={"open":0,"high":0,"low":0,"close":0,"signal":"WAIT",
            "ce_prem":0,"pe_prem":0}

    # Only fetch ORB after 9:45 AM
    if phase in ["ORB_945_PLUS","CLOSED"]:
        try:
            fn_date=ist.strftime("%Y-%m-%d")
            for interval,label in [("THIRTY_MINUTE","30min"),("FIFTEEN_MINUTE","15min")]:
                cd=get_candles(jwt,"26037",interval,
                              f"{fn_date} 09:00",f"{fn_date} 10:30")
                candle_list=cd.get("data",[]) if cd.get("status") else []
                print(f"FINNIFTY {label}: {len(candle_list)} candles")
                if candle_list:
                    if interval=="THIRTY_MINUTE":
                        c=candle_list[0]
                        fn_orb["open"]=round(float(c[1]),2)
                        fn_orb["high"]=round(float(c[2]),2)
                        fn_orb["low"]=round(float(c[3]),2)
                        fn_orb["close"]=round(float(c[4]),2)
                    else:
                        c2=candle_list[:2]
                        fn_orb["open"]=round(float(c2[0][1]),2)
                        fn_orb["high"]=round(max(float(x[2]) for x in c2),2)
                        fn_orb["low"]=round(min(float(x[3]) for x in c2),2)
                        fn_orb["close"]=round(float(c2[-1][4]),2)
                    if fn_ltp>0 and fn_orb["high"]>0:
                        if fn_ltp>fn_orb["high"]: fn_orb["signal"]="CE"
                        elif fn_ltp<fn_orb["low"]: fn_orb["signal"]="PE"
                        else: fn_orb["signal"]="MONITOR"
                    print(f"ORB: H={fn_orb['high']} L={fn_orb['low']} Sig={fn_orb['signal']}")
                    break
        except Exception as e:
            print(f"FINNIFTY ORB error: {e}")

    if fn_ltp>0:
        import math as _m
        days=max(1,3)
        fn_orb["ce_prem"]=round(fn_ltp*0.006*_m.sqrt(days/5),0)
        fn_orb["pe_prem"]=round(fn_orb["ce_prem"]*0.92,0)

    # ── SECTORS ───────────────────────────────────────────────────────
    sectors=[]
    for sec,stks in SECTOR_STOCKS.items():
        valid=[s for s in stks if ltp(s)>0]
        if not valid:
            valid=stks[:1]
        chgs=[pct_chg(s) for s in valid]
        # Weighted avg: use abs(chg) weighted
        sec_chg=round(sum(chgs)/len(chgs),2) if chgs else 0.0
        # Leader = stock with highest absolute % change
        ldr=max(valid, key=lambda s:abs(pct_chg(s))) if valid else stks[0]
        ldr_ltp=ltp(ldr); ldr_chg=pct_chg(ldr); ldr_pts=pts_chg(ldr)
        sectors.append({
            "name":sec,"chg":sec_chg,"leader":ldr,
            "leader_price":ldr_ltp,"leader_chg":ldr_chg,
            "leader_pts":ldr_pts,"idx_price":0,
            "stocks":[{"sym":s,"ltp":ltp(s),"chg":pct_chg(s)} for s in valid]
        })
        print(f"  Sector {sec}: {sec_chg}% | Leader: {ldr} {ldr_chg}%")
    sectors.sort(key=lambda x:x["chg"],reverse=True)
    print(f"Top: {sectors[0]['name']} {sectors[0]['chg']}% | Bot: {sectors[-1]['name']} {sectors[-1]['chg']}%")

    # ── CANDLES + ROCKERS (9:30 AM phase onwards) ─────────────────────
    longs=[]; shorts=[]
    candles={}

    if phase in ["ORB_930","ORB_945_PLUS","CLOSED","PREMARKET_900","PREMARKET_845"]:
        to_d=ist.strftime("%Y-%m-%d %H:%M")
        fr_d=(ist-timedelta(days=60)).strftime("%Y-%m-%d %H:%M")
        print(f"Fetching candles...")
        for i,(sym,info) in enumerate(STOCKS.items()):
            try:
                cd=get_candles(jwt,info["token"],"ONE_DAY",fr_d,to_d)
                if cd.get("status") and cd.get("data"):
                    candles[sym]=cd["data"]
                time.sleep(0.2)
                if (i+1)%25==0: print(f"  Candles: {i+1}/{len(STOCKS)}")
            except: pass
        print(f"Candles: {len(candles)}")

        # ROCKERS scan
        top4=[s["name"] for s in sectors[:4]]
        bot4=[s["name"] for s in sectors[-4:]]
        risk_amt=CAPITAL*RISK_PCT/100

        for sym,info in STOCKS.items():
            p=ltp(sym)
            if p<=0: continue
            c=candles.get(sym,[])
            closes=[float(x[4]) for x in c] if c else [p]
            vols=[float(x[5]) for x in c] if c else [1]
            h=day_high(sym); l=day_low(sym); op=day_open(sym)
            vwap_val=calc_vwap(c)
            e20=ema(closes,20); r=rsi(closes); vr=vol_ratio(vols)
            macd_b=macd_bull(closes); st_b=st_bull(c)

            if info["sec"] in top4:
                vwap_ok=p>vwap_val if vwap_val>0 else p>e20
                conds={"EMA20":p>e20,"SUPERT":st_b,"RSI":50<=r<=70,
                       "MACD":macd_b,"VOL":vr>=1.5,"VWAP":vwap_ok}
                sc=sum(conds.values())
                en=round(h*1.002,2); sl=round(en*(1-info["slp"]/100),2); rp=en-sl
                if rp<=0: continue
                longs.append({
                    "sym":sym,"sec":info["sec"],"ltp":p,"chg":pct_chg(sym),
                    "open":op,"high":h,"low":l,"vwap":vwap_val,
                    "score":sc,"rsi":r,"vol_ratio":vr,
                    "conds":{k:bool(v) for k,v in conds.items()},
                    "entry":en,"sl":sl,"t1":round(en+rp*1.5,2),"t2":round(en+rp*2.5,2),
                    "gtt_trigger":en,"gtt_limit":round(en*1.001,2),
                    "qty":max(1,math.floor(risk_amt/rp)),"side":"LONG",
                    "bias":bias,"phase":phase
                })

            if info["sec"] in bot4:
                vwap_ok=p<vwap_val if vwap_val>0 else p<e20
                conds={"EMA20":p<e20,"SUPERT":not st_b,"RSI":30<=r<=50,
                       "MACD":not macd_b,"VOL":vr>=1.5,"VWAP":vwap_ok}
                sc=sum(conds.values())
                en=round(l*0.998,2); sl=round(en*(1+info["slp"]/100),2); rp=sl-en
                if rp<=0: continue
                shorts.append({
                    "sym":sym,"sec":info["sec"],"ltp":p,"chg":pct_chg(sym),
                    "open":op,"high":h,"low":l,"vwap":vwap_val,
                    "score":sc,"rsi":r,"vol_ratio":vr,
                    "conds":{k:bool(v) for k,v in conds.items()},
                    "entry":en,"sl":sl,"t1":round(en-rp*1.5,2),"t2":round(en-rp*2.5,2),
                    "gtt_trigger":en,"gtt_limit":round(en*0.999,2),
                    "qty":max(1,math.floor(risk_amt/rp)),"side":"SHORT",
                    "bias":bias,"phase":phase
                })

        longs.sort(key=lambda x:x["score"],reverse=True)
        shorts.sort(key=lambda x:x["score"],reverse=True)
        print(f"Longs: {len(longs)} | Shorts: {len(shorts)}")

        # Excel journal
        all_signals=[{**s} for s in longs[:5]+shorts[:5]]
        if all_signals:
            try: update_excel(all_signals,ist)
            except Exception as e: print(f"Excel error: {e}")

    # ── SAVE DATA.JSON ────────────────────────────────────────────────
    data={
        "generated_at": ist.strftime("%d %b %Y %I:%M %p IST"),
        "phase": phase,
        "phase_label": phase_label,
        "market_status": "OPEN" if mkt_open else "PRE-MARKET" if ist.hour<9 or (ist.hour==9 and ist.minute<15) else "CLOSED",
        "is_live": True,
        "market": {
            "nifty50": {"ltp":nifty_ltp,"chg":nifty_chg,"pts":nifty_pts,"prev":nifty_prev},
            "vix":     {"ltp":vix_data["ltp"],"chg":vix_data["chg"],"prev":vix_data["prev"],
                        "status":vix_data["status"],"source":vix_data["source"]},
            "bias":    bias
        },
        "gift_nifty": gift,
        "pre_market": {
            "top_gainers": top_gainers,
            "top_losers":  top_losers,
            "analysis":    f"{'Gap Up' if gift['chg']>0.3 else 'Gap Down' if gift['chg']<-0.3 else 'Flat'} open expected. {bias} bias."
        },
        "finnifty": {
            "ltp":fn_ltp,"chg":fn_chg,"pts":fn_pts,"prev":fn_prev,"atm":fn_atm,
            "orb_open":fn_orb["open"],"orb_high":fn_orb["high"],
            "orb_low":fn_orb["low"],"orb_close":fn_orb["close"],
            "signal":fn_orb["signal"],
            "ce_premium":fn_orb["ce_prem"],"pe_premium":fn_orb["pe_prem"],
            "ce_sl":round(fn_orb["ce_prem"]*0.7,0),
            "ce_tgt":round(fn_orb["ce_prem"]*1.5,0),
            "pe_sl":round(fn_orb["pe_prem"]*0.7,0),
            "pe_tgt":round(fn_orb["pe_prem"]*1.5,0),
        },
        "sectors":  sectors,
        "long":     longs[:5],
        "short":    shorts[:5],
        "capital":  CAPITAL,
        "risk_pct": RISK_PCT,
        "risk_amt": CAPITAL*RISK_PCT/100,
        "universe": f"Nifty 100 — {len(STOCKS)} stocks",
    }

    with open("data.json","w") as f:
        json.dump(data,f,indent=2)
    print(f"✅ data.json saved! Phase: {phase}")
    print(f"═══ DONE ═══")

if __name__=="__main__":
    main()
