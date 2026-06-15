#!/usr/bin/env python3
"""
SUBHA CAPITAL — Market Data Fetcher
Runs on GitHub Actions at 8:45 AM IST
Fetches live data from Angel One
Saves to data.json for dashboard
"""

import os, json, math, pyotp, requests, pytz
from datetime import datetime, timedelta

# Config from GitHub Secrets
CLIENT_ID    = os.environ["ANGEL_CLIENT_ID"]
API_KEY      = os.environ["ANGEL_API_KEY"]
TOTP_SECRET  = os.environ["ANGEL_TOTP_SECRET"]
PIN          = os.environ["ANGEL_PIN"]
CAPITAL      = 100000
RISK_PCT     = 1.0
IST          = pytz.timezone("Asia/Kolkata")

BASE = "https://apiconnect.angelone.in"

STOCKS = {
    "HDFCBANK":    {"token":"1333",  "sec":"BANKING", "slp":0.9},
    "ICICIBANK":   {"token":"4963",  "sec":"BANKING", "slp":0.85},
    "INFY":        {"token":"1594",  "sec":"IT",      "slp":0.80},
    "TCS":         {"token":"11536", "sec":"IT",      "slp":0.75},
    "MARUTI":      {"token":"10999", "sec":"AUTO",    "slp":0.85},
    "MM":          {"token":"2031",  "sec":"AUTO",    "slp":0.80},
    "SUNPHARMA":   {"token":"3351",  "sec":"PHARMA",  "slp":0.75},
    "DRREDDY":     {"token":"881",   "sec":"PHARMA",  "slp":0.80},
    "TATASTEEL":   {"token":"3499",  "sec":"METAL",   "slp":1.0},
    "JSWSTEEL":    {"token":"11723", "sec":"METAL",   "slp":1.0},
    "HINDUNILVR":  {"token":"1394",  "sec":"FMCG",    "slp":0.70},
    "NESTLEIND":   {"token":"17963", "sec":"FMCG",    "slp":0.70},
    "RELIANCE":    {"token":"2885",  "sec":"ENERGY",  "slp":0.80},
    "ONGC":        {"token":"2475",  "sec":"ENERGY",  "slp":0.90},
    "DLF":         {"token":"14732", "sec":"REALTY",  "slp":1.0},
    "GODREJPROP":  {"token":"9742",  "sec":"REALTY",  "slp":1.0},
}

INDICES = {
    "NIFTY50":      "26000",
    "NIFTYBANK":    "26009",
    "NIFTYIT":      "26035",
    "NIFTYAUTO":    "26037",
    "NIFTYMETAL":   "26042",
    "NIFTYFMCG":    "26043",
    "NIFTYENERGY":  "26050",
    "NIFTYREALTY":  "26054",
    "INDIAVIX":     "26017",
}

SECTOR_MAP = {
    "BANKING":"NIFTYBANK","IT":"NIFTYIT","AUTO":"NIFTYAUTO",
    "METAL":"NIFTYMETAL","FMCG":"NIFTYFMCG","ENERGY":"NIFTYENERGY",
    "REALTY":"NIFTYREALTY","PHARMA":"NIFTY50",
}

def get_headers(token=None):
    h = {
        "Content-Type":"application/json","Accept":"application/json",
        "X-UserType":"USER","X-SourceID":"WEB",
        "X-ClientLocalIP":"192.168.1.1","X-ClientPublicIP":"106.193.147.98",
        "X-MACAddress":"fe80::216e:6507:4b90:3719","X-PrivateKey":API_KEY,
    }
    if token: h["Authorization"] = f"Bearer {token}"
    return h

def login():
    totp = pyotp.TOTP(TOTP_SECRET).now()
    print(f"TOTP: {totp}")
    r = requests.post(f"{BASE}/rest/auth/angelbroking/user/v1/loginByPassword",
        json={"clientcode":CLIENT_ID,"password":PIN,"totp":totp},
        headers=get_headers(), timeout=15)
    d = r.json()
    if d.get("status") and d.get("data",{}).get("jwtToken"):
        print("✅ Login success")
        return d["data"]["jwtToken"]
    raise Exception(f"Login failed: {d.get('message')}")

def get_quotes(token, all_tokens):
    r = requests.post(f"{BASE}/rest/secure/angelbroking/market/v1/quote/",
        json={"mode":"FULL","exchangeTokens":{"NSE":all_tokens}},
        headers=get_headers(token), timeout=15)
    return r.json()

def get_candles(token, sym_token, interval, from_d, to_d):
    r = requests.post(f"{BASE}/rest/secure/angelbroking/historical/v1/getCandleData",
        json={"exchange":"NSE","symboltoken":sym_token,"interval":interval,
              "fromdate":from_d,"todate":to_d},
        headers=get_headers(token), timeout=15)
    return r.json()

def calc_ema(prices, period):
    if len(prices) < period: return prices[-1] if prices else 0
    k = 2/(period+1)
    ema = sum(prices[:period])/period
    for p in prices[period:]: ema = p*k + ema*(1-k)
    return round(ema, 2)

def calc_rsi(prices, period=14):
    if len(prices) < period+1: return 50
    gains,losses = [],[]
    for i in range(1,len(prices)):
        d = prices[i]-prices[i-1]
        gains.append(max(d,0)); losses.append(max(-d,0))
    ag = sum(gains[:period])/period
    al = sum(losses[:period])/period
    for i in range(period,len(gains)):
        ag = (ag*(period-1)+gains[i])/period
        al = (al*(period-1)+losses[i])/period
    if al == 0: return 100
    return round(100 - 100/(1+ag/al))

def calc_macd_bull(prices):
    if len(prices) < 26: return True
    k12,k26 = 2/13,2/27
    e12 = e26 = prices[0]
    for p in prices: e12=p*k12+e12*(1-k12); e26=p*k26+e26*(1-k26)
    return e12 > e26

def calc_st_bull(candles):
    if not candles: return True
    last = candles[-1]
    return last[4] > (last[2]+last[3])/2

def calc_vol_ratio(volumes):
    if len(volumes) < 2: return 1.0
    avg = sum(volumes[:-1])/len(volumes[:-1])
    return round(volumes[-1]/avg,1) if avg > 0 else 1.0

def rockers_long(ltp, pdh, closes, volumes, candles):
    ema20 = calc_ema(closes,20)
    rsi = calc_rsi(closes)
    conds = {
        "EMA20": ltp > ema20,
        "SUPERT": calc_st_bull(candles),
        "RSI": 50 <= rsi <= 70,
        "MACD": calc_macd_bull(closes),
        "VOL": calc_vol_ratio(volumes) >= 1.5,
        "PDH": ltp > pdh,
    }
    vr = calc_vol_ratio(volumes)
    return sum(conds.values()), rsi, vr, conds

def rockers_short(ltp, pdl, closes, volumes, candles):
    ema20 = calc_ema(closes,20)
    rsi = calc_rsi(closes)
    conds = {
        "EMA20": ltp < ema20,
        "SUPERT": not calc_st_bull(candles),
        "RSI": 30 <= rsi <= 50,
        "MACD": not calc_macd_bull(closes),
        "VOL": calc_vol_ratio(volumes) >= 1.5,
        "PDL": ltp < pdl,
    }
    vr = calc_vol_ratio(volumes)
    return sum(conds.values()), rsi, vr, conds

def main():
    ist_now = datetime.now(IST)
    print(f"Running at {ist_now.strftime('%Y-%m-%d %H:%M IST')}")

    # Login
    jwt = login()

    # Fetch all quotes
    all_tokens = [v["token"] for v in STOCKS.values()] + list(INDICES.values())
    qdata = get_quotes(jwt, all_tokens)
    quotes = {}
    if qdata.get("status") and qdata.get("data"):
        for q in qdata["data"].get("fetched",[]):
            t = q.get("symbolToken","")
            for sym,info in {**STOCKS,**{k:{"token":v} for k,v in INDICES.items()}}.items():
                if info["token"] == t:
                    quotes[sym] = q; break
    print(f"Quotes: {len(quotes)}")

    def ltp(s): return float(quotes.get(s,{}).get("ltp",0) or 0)
    def chg(s):
        q=quotes.get(s,{})
        l,c=float(q.get("ltp",0) or 0),float(q.get("close",0) or 0)
        return round((l-c)/c*100,2) if c>0 else 0
    def pdh(s): return float(quotes.get(s,{}).get("high",ltp(s)) or ltp(s))
    def pdl(s): return float(quotes.get(s,{}).get("low",ltp(s)) or ltp(s))

    # Fetch candles
    to_d = ist_now.strftime("%Y-%m-%d %H:%M")
    fr_d = (ist_now-timedelta(days=60)).strftime("%Y-%m-%d %H:%M")
    candles = {}
    import time
    for sym,info in STOCKS.items():
        try:
            cd = get_candles(jwt,info["token"],"ONE_DAY",fr_d,to_d)
            if cd.get("status") and cd.get("data"):
                candles[sym] = cd["data"]
            time.sleep(0.3)
        except: pass
    print(f"Candles: {len(candles)}")

    # Sectors
    sectors = []
    for sec,idx in SECTOR_MAP.items():
        sectors.append({"name":sec,"chg":chg(idx),
            "leader":next((s for s,i in STOCKS.items() if i["sec"]==sec),"")})
    sectors.sort(key=lambda x:x["chg"],reverse=True)

    # ROCKERS
    top_secs = [s["name"] for s in sectors[:4]]
    bot_secs  = [s["name"] for s in sectors[-4:]]
    risk_amt = CAPITAL * RISK_PCT / 100

    long_list, short_list = [], []

    for sym,info in STOCKS.items():
        l = ltp(sym)
        if l <= 0: continue
        c = candles.get(sym,[])
        closes  = [x[4] for x in c] if c else [l]
        volumes = [x[5] for x in c] if c else [1]

        if info["sec"] in top_secs:
            sc,rsi,vr,conds = rockers_long(l,pdh(sym),closes,volumes,c)
            en = round(pdh(sym)*1.002,2)
            sl = round(en*(1-info["slp"]/100),2)
            rp = en-sl
            long_list.append({
                "sym":sym,"sec":info["sec"],"ltp":l,"chg":chg(sym),
                "score":sc,"rsi":rsi,"vol_ratio":vr,
                "conds":{k:bool(v) for k,v in conds.items()},
                "entry":en,"sl":sl,
                "t1":round(en+rp*1.5,2),"t2":round(en+rp*2.5,2),
                "qty":max(1,math.floor(risk_amt/rp)) if rp>0 else 1,
                "side":"LONG"
            })

        if info["sec"] in bot_secs:
            sc,rsi,vr,conds = rockers_short(l,pdl(sym),closes,volumes,c)
            en = round(pdl(sym)*0.998,2)
            sl = round(en*(1+info["slp"]/100),2)
            rp = sl-en
            short_list.append({
                "sym":sym,"sec":info["sec"],"ltp":l,"chg":chg(sym),
                "score":sc,"rsi":rsi,"vol_ratio":vr,
                "conds":{k:bool(v) for k,v in conds.items()},
                "entry":en,"sl":sl,
                "t1":round(en-rp*1.5,2),"t2":round(en-rp*2.5,2),
                "qty":max(1,math.floor(risk_amt/rp)) if rp>0 else 1,
                "side":"SHORT"
            })

    long_list.sort(key=lambda x:x["score"],reverse=True)
    short_list.sort(key=lambda x:x["score"],reverse=True)

    # Market bias
    nf_chg = chg("NIFTY50")
    vix = ltp("INDIAVIX")
    bias = "BULLISH" if nf_chg>0.3 and vix<15 else "BEARISH" if nf_chg<-0.3 or vix>16 else "NEUTRAL"

    # Save data.json
    data = {
        "generated_at": ist_now.strftime("%d %b %Y %I:%M %p IST"),
        "is_live": True,
        "market": {
            "nifty50": {"ltp":ltp("NIFTY50"),"chg":chg("NIFTY50")},
            "vix":     {"ltp":ltp("INDIAVIX"),"chg":chg("INDIAVIX")},
            "bias":    bias,
        },
        "sectors": sectors,
        "long":  long_list[:3],
        "short": short_list[:3],
        "capital": CAPITAL,
        "risk_pct": RISK_PCT,
        "risk_amt": risk_amt,
    }

    with open("data.json","w") as f:
        json.dump(data,f,indent=2)
    print("✅ data.json saved")
    print(f"Bias: {bias} | Long: {len(long_list)} | Short: {len(short_list)}")

if __name__ == "__main__":
    main()
