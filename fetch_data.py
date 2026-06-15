#!/usr/bin/env python3
"""
SUBHA CAPITAL — Market Data Fetcher v3
Fixed: VIX + Sector % display
"""

import os, json, math, pyotp, requests, pytz, time
from datetime import datetime, timedelta

CLIENT_ID    = os.environ["ANGEL_CLIENT_ID"]
API_KEY      = os.environ["ANGEL_API_KEY"]
TOTP_SECRET  = os.environ["ANGEL_TOTP_SECRET"]
PIN          = os.environ["ANGEL_PIN"]
CAPITAL      = 100000
RISK_PCT     = 1.0
IST          = pytz.timezone("Asia/Kolkata")
BASE         = "https://apiconnect.angelone.in"

STOCKS = {
    "HDFCBANK":   {"token":"1333",  "sec":"BANKING","slp":0.9},
    "ICICIBANK":  {"token":"4963",  "sec":"BANKING","slp":0.85},
    "INFY":       {"token":"1594",  "sec":"IT",     "slp":0.80},
    "TCS":        {"token":"11536", "sec":"IT",     "slp":0.75},
    "MARUTI":     {"token":"10999", "sec":"AUTO",   "slp":0.85},
    "MM":         {"token":"2031",  "sec":"AUTO",   "slp":0.80},
    "SUNPHARMA":  {"token":"3351",  "sec":"PHARMA", "slp":0.75},
    "DRREDDY":    {"token":"881",   "sec":"PHARMA", "slp":0.80},
    "TATASTEEL":  {"token":"3499",  "sec":"METAL",  "slp":1.0},
    "JSWSTEEL":   {"token":"11723", "sec":"METAL",  "slp":1.0},
    "HINDUNILVR": {"token":"1394",  "sec":"FMCG",   "slp":0.70},
    "NESTLEIND":  {"token":"17963", "sec":"FMCG",   "slp":0.70},
    "RELIANCE":   {"token":"2885",  "sec":"ENERGY", "slp":0.80},
    "ONGC":       {"token":"2475",  "sec":"ENERGY", "slp":0.90},
    "DLF":        {"token":"14732", "sec":"REALTY", "slp":1.0},
    "GODREJPROP": {"token":"9742",  "sec":"REALTY", "slp":1.0},
}

# NSE Indices — correct tokens
INDICES = {
    "NIFTY50":    {"token":"26000", "exch":"NSE"},
    "NIFTYBANK":  {"token":"26009", "exch":"NSE"},
    "NIFTYIT":    {"token":"26035", "exch":"NSE"},
    "NIFTYAUTO":  {"token":"26037", "exch":"NSE"},
    "NIFTYMETAL": {"token":"26042", "exch":"NSE"},
    "NIFTYFMCG":  {"token":"26043", "exch":"NSE"},
    "NIFTYENERGY":{"token":"26050", "exch":"NSE"},
    "NIFTYREALTY":{"token":"26054", "exch":"NSE"},
    "NIFTYPHARMA":{"token":"26038", "exch":"NSE"},  # Fixed pharma index
    "INDIAVIX":   {"token":"26017", "exch":"NSE"},  # VIX
}

# Correct sector to index mapping
SECTOR_MAP = {
    "BANKING": "NIFTYBANK",
    "IT":      "NIFTYIT",
    "AUTO":    "NIFTYAUTO",
    "METAL":   "NIFTYMETAL",
    "FMCG":    "NIFTYFMCG",
    "ENERGY":  "NIFTYENERGY",
    "REALTY":  "NIFTYREALTY",
    "PHARMA":  "NIFTYPHARMA",  # Fixed — dedicated pharma index
}

def hdrs(tok=None):
    h = {
        "Content-Type":"application/json","Accept":"application/json",
        "X-UserType":"USER","X-SourceID":"WEB",
        "X-ClientLocalIP":"192.168.1.1","X-ClientPublicIP":"106.193.147.98",
        "X-MACAddress":"fe80::216e:6507:4b90:3719","X-PrivateKey":API_KEY,
    }
    if tok: h["Authorization"] = f"Bearer {tok}"
    return h

def login():
    totp = pyotp.TOTP(TOTP_SECRET).now()
    print(f"TOTP: {totp}")
    r = requests.post(f"{BASE}/rest/auth/angelbroking/user/v1/loginByPassword",
        json={"clientcode":CLIENT_ID,"password":PIN,"totp":totp},
        headers=hdrs(), timeout=15)
    d = r.json()
    if d.get("status") and d.get("data",{}).get("jwtToken"):
        print("✅ Login OK")
        return d["data"]["jwtToken"]
    raise Exception(f"Login failed: {d.get('message')}")

def get_quotes(tok, nse_tokens):
    r = requests.post(f"{BASE}/rest/secure/angelbroking/market/v1/quote/",
        json={"mode":"FULL","exchangeTokens":{"NSE": nse_tokens}},
        headers=hdrs(tok), timeout=15)
    return r.json()

def get_candles(tok, sym_tok, interval, fr, to):
    r = requests.post(f"{BASE}/rest/secure/angelbroking/historical/v1/getCandleData",
        json={"exchange":"NSE","symboltoken":sym_tok,
              "interval":interval,"fromdate":fr,"todate":to},
        headers=hdrs(tok), timeout=15)
    return r.json()

def calc_ema(p, n):
    if not p: return 0
    if len(p)<n: return p[-1]
    k=2/(n+1); e=sum(p[:n])/n
    for x in p[n:]: e=x*k+e*(1-k)
    return round(e,2)

def calc_rsi(p, n=14):
    if len(p)<n+1: return 50
    g,l=[],[]
    for i in range(1,len(p)):
        d=p[i]-p[i-1]; g.append(max(d,0)); l.append(max(-d,0))
    ag=sum(g[:n])/n; al=sum(l[:n])/n
    for i in range(n,len(g)):
        ag=(ag*(n-1)+g[i])/n; al=(al*(n-1)+l[i])/n
    return 100 if al==0 else round(100-100/(1+ag/al))

def calc_macd(p):
    if len(p)<26: return True
    k12,k26=2/13,2/27; e12=e26=p[0]
    for x in p: e12=x*k12+e12*(1-k12); e26=x*k26+e26*(1-k26)
    return e12>e26

def calc_st(c):
    if not c: return True
    return c[-1][4]>(c[-1][2]+c[-1][3])/2

def calc_vr(v):
    if len(v)<2: return 1.0
    avg=sum(v[:-1])/len(v[:-1])
    return round(v[-1]/avg,1) if avg>0 else 1.0

def main():
    ist=datetime.now(IST)
    print(f"Time: {ist.strftime('%Y-%m-%d %H:%M IST')}")
    mkt_open=ist.hour<15 or (ist.hour==15 and ist.minute<30)
    print(f"Market: {'OPEN' if mkt_open else 'CLOSED'}")

    jwt=login()

    # All NSE tokens — stocks + indices together
    stock_tokens=[v["token"] for v in STOCKS.values()]
    index_tokens=[v["token"] for v in INDICES.values()]
    all_tokens=stock_tokens+index_tokens

    qd=get_quotes(jwt, all_tokens)
    Q={}
    if qd.get("status") and qd.get("data"):
        for q in qd["data"].get("fetched",[]):
            t=q.get("symbolToken","")
            # Match stocks
            for sym,info in STOCKS.items():
                if info["token"]==t: Q[sym]=q; break
            # Match indices
            for idx,info in INDICES.items():
                if info["token"]==t: Q[idx]=q; break
    print(f"Quotes fetched: {len(Q)}")
    print(f"Indices found: {[k for k in INDICES if k in Q]}")

    def price(s):
        q=Q.get(s,{})
        ltp=float(q.get("ltp",0) or 0)
        close=float(q.get("close",0) or 0)
        if ltp>0: return ltp
        if close>0: return close
        return float(q.get("open",0) or 0)

    def day_chg(s):
        q=Q.get(s,{})
        cur=price(s)
        if cur<=0: return 0.0
        prev=float(q.get("previousClose",0) or q.get("prevClose",0) or 0)
        if prev<=0: prev=float(q.get("open",0) or 0)
        if prev<=0: return 0.0
        return round((cur-prev)/prev*100,2)

    def hi(s):
        q=Q.get(s,{})
        return float(q.get("high",0) or price(s))

    def lo(s):
        q=Q.get(s,{})
        return float(q.get("low",0) or price(s))

    # VIX — special handling
    vix_price=price("INDIAVIX")
    vix_chg=day_chg("INDIAVIX")
    print(f"VIX: {vix_price} ({vix_chg}%)")

    # Nifty 50
    nifty_price=price("NIFTY50")
    nifty_chg=day_chg("NIFTY50")
    print(f"Nifty: {nifty_price} ({nifty_chg}%)")

    # Sectors — calculate from actual stock performance if index shows 0
    sectors=[]
    for sec,idx in SECTOR_MAP.items():
        idx_chg=day_chg(idx)
        idx_price=price(idx)
        ldr=next((s for s,i in STOCKS.items() if i["sec"]==sec),"")
        ldr_price=price(ldr)
        ldr_chg=day_chg(ldr)

        # If index shows 0, calculate from stock average
        if idx_chg==0.0 or idx_price==0:
            sec_stocks=[s for s,i in STOCKS.items() if i["sec"]==sec]
            chgs=[day_chg(s) for s in sec_stocks if price(s)>0]
            idx_chg=round(sum(chgs)/len(chgs),2) if chgs else 0.0
            print(f"Sector {sec}: using stock avg {idx_chg}%")
        else:
            print(f"Sector {sec}: index {idx_chg}%")

        # Use leader stock price if index price is 0
        display_price=idx_price if idx_price>0 else ldr_price
        sectors.append({
            "name":sec,"chg":idx_chg,"leader":ldr,
            "idx_price":display_price,
            "leader_price":ldr_price,"leader_chg":ldr_chg
        })

    sectors.sort(key=lambda x:x["chg"],reverse=True)
    print(f"Top sector: {sectors[0]['name']} {sectors[0]['chg']}%")
    print(f"Bot sector: {sectors[-1]['name']} {sectors[-1]['chg']}%")

    # Candles
    to_d=ist.strftime("%Y-%m-%d %H:%M")
    fr_d=(ist-timedelta(days=60)).strftime("%Y-%m-%d %H:%M")
    candles={}
    for sym,info in STOCKS.items():
        try:
            cd=get_candles(jwt,info["token"],"ONE_DAY",fr_d,to_d)
            if cd.get("status") and cd.get("data"):
                candles[sym]=cd["data"]
            time.sleep(0.3)
        except: pass
    print(f"Candles: {len(candles)}")

    # ROCKERS
    top4=[s["name"] for s in sectors[:4]]
    bot4=[s["name"] for s in sectors[-4:]]
    risk=CAPITAL*RISK_PCT/100
    longs,shorts=[],[]

    for sym,info in STOCKS.items():
        p=price(sym)
        if p<=0: continue
        c=candles.get(sym,[])
        closes=[x[4] for x in c] if c else [p]
        vols=[x[5] for x in c] if c else [1]
        h=hi(sym); l=lo(sym)

        if info["sec"] in top4:
            ema20=calc_ema(closes,20); rsi=calc_rsi(closes); vr=calc_vr(vols)
            conds={
                "EMA20":p>ema20,"SUPERT":calc_st(c),
                "RSI":50<=rsi<=70,"MACD":calc_macd(closes),
                "VOL":vr>=1.5,"PDH":p>h,
            }
            sc=sum(conds.values())
            en=round(h*1.002,2); sl=round(en*(1-info["slp"]/100),2); rp=en-sl
            longs.append({
                "sym":sym,"sec":info["sec"],"ltp":p,"chg":day_chg(sym),
                "score":sc,"rsi":rsi,"vol_ratio":vr,
                "conds":{k:bool(v) for k,v in conds.items()},
                "entry":en,"sl":sl,
                "t1":round(en+rp*1.5,2),"t2":round(en+rp*2.5,2),
                "qty":max(1,math.floor(risk/rp)) if rp>0 else 1,"side":"LONG"
            })

        if info["sec"] in bot4:
            ema20=calc_ema(closes,20); rsi=calc_rsi(closes); vr=calc_vr(vols)
            conds={
                "EMA20":p<ema20,"SUPERT":not calc_st(c),
                "RSI":30<=rsi<=50,"MACD":not calc_macd(closes),
                "VOL":vr>=1.5,"PDL":p<l,
            }
            sc=sum(conds.values())
            en=round(l*0.998,2); sl=round(en*(1+info["slp"]/100),2); rp=sl-en
            shorts.append({
                "sym":sym,"sec":info["sec"],"ltp":p,"chg":day_chg(sym),
                "score":sc,"rsi":rsi,"vol_ratio":vr,
                "conds":{k:bool(v) for k,v in conds.items()},
                "entry":en,"sl":sl,
                "t1":round(en-rp*1.5,2),"t2":round(en-rp*2.5,2),
                "qty":max(1,math.floor(risk/rp)) if rp>0 else 1,"side":"SHORT"
            })

    longs.sort(key=lambda x:x["score"],reverse=True)
    shorts.sort(key=lambda x:x["score"],reverse=True)

    bias="BULLISH" if nifty_chg>0.3 and (vix_price==0 or vix_price<15) else \
         "BEARISH" if nifty_chg<-0.3 or vix_price>16 else "NEUTRAL"

    data={
        "generated_at":ist.strftime("%d %b %Y %I:%M %p IST"),
        "market_status":"OPEN" if mkt_open else "CLOSED",
        "is_live":True,
        "market":{
            "nifty50":{"ltp":nifty_price,"chg":nifty_chg},
            "vix":{"ltp":vix_price,"chg":vix_chg},
            "bias":bias,
        },
        "sectors":sectors,
        "long":longs[:3],
        "short":shorts[:3],
        "capital":CAPITAL,"risk_pct":RISK_PCT,"risk_amt":risk,
    }

    with open("data.json","w") as f:
        json.dump(data,f,indent=2)
    print(f"✅ Saved! Bias:{bias} Longs:{len(longs)} Shorts:{len(shorts)}")

if __name__=="__main__":
    main()
