#!/usr/bin/env python3
"""
SUBHA CAPITAL — Market Data Fetcher v6
Nifty 100 Universe + Excel Backtest Journal
"""

import os,json,math,pyotp,requests,pytz,time,openpyxl
from openpyxl.styles import PatternFill,Font,Alignment,Border,Side
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

# Nifty 100 Universe
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

INDICES = {"NIFTY50":"26000","INDIAVIX":"26017","FINNIFTY":"26037"}

SECTOR_STOCKS = {
    "BANKING":  ["HDFCBANK","ICICIBANK","SBIN","AXISBANK","KOTAKBANK","INDUSINDBK"],
    "FINANCE":  ["BAJFINANCE","BAJAJFINSV","CHOLAFIN","PFC","RECLTD"],
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
        print("✅ Login OK"); return d["data"]["jwtToken"]
    raise Exception(f"Login failed: {d.get('message')}")

def get_quotes(tok,tokens):
    # Batch in groups of 50
    all_data=[]
    for i in range(0,len(tokens),50):
        batch=tokens[i:i+50]
        r=requests.post(f"{BASE}/rest/secure/angelbroking/market/v1/quote/",
            json={"mode":"FULL","exchangeTokens":{"NSE":batch}},
            headers=hdrs(tok),timeout=15)
        d=r.json()
        if d.get("status") and d.get("data"):
            all_data.extend(d["data"].get("fetched",[]))
        time.sleep(0.3)
    return all_data

def get_candles(tok,sym_tok,interval,fr,to):
    r=requests.post(f"{BASE}/rest/secure/angelbroking/historical/v1/getCandleData",
        json={"exchange":"NSE","symboltoken":sym_tok,"interval":interval,
              "fromdate":fr,"todate":to},
        headers=hdrs(tok),timeout=15)
    return r.json()

def ema(p,n):
    if not p: return 0
    if len(p)<n: return p[-1]
    k=2/(n+1); e=sum(p[:n])/n
    for x in p[n:]: e=x*k+e*(1-k)
    return round(e,2)

def rsi_calc(p,n=14):
    if len(p)<n+1: return 50
    g,l=[],[]
    for i in range(1,len(p)):
        d=p[i]-p[i-1]; g.append(max(d,0)); l.append(max(-d,0))
    ag=sum(g[:n])/n; al=sum(l[:n])/n
    for i in range(n,len(g)):
        ag=(ag*(n-1)+g[i])/n; al=(al*(n-1)+l[i])/n
    return 100 if al==0 else round(100-100/(1+ag/al))

def macd_bull(p):
    if len(p)<26: return True
    k12,k26=2/13,2/27; e12=e26=p[0]
    for x in p: e12=x*k12+e12*(1-k12); e26=x*k26+e26*(1-k26)
    return e12>e26

def supert(c):
    if not c: return True
    return c[-1][4]>(c[-1][2]+c[-1][3])/2

def vol_ratio(v):
    if len(v)<2: return 1.0
    avg=sum(v[:-1])/len(v[:-1])
    return round(v[-1]/avg,1) if avg>0 else 1.0

def update_excel(signals,ist):
    """Update trading journal Excel file"""
    fname="trading_journal.xlsx"
    
    # Load or create workbook
    try:
        wb=openpyxl.load_workbook(fname)
        ws=wb.active
    except:
        wb=openpyxl.Workbook()
        ws=wb.active
        ws.title="SUBHA CAPITAL Journal"
        
        # Header styling
        headers=["Date","Stock","Sector","Side","ROCKERS Score",
                 "Pre-Market Price","Day Open","Day High","Day Low",
                 "GTT Trigger","GTT Limit","Stop Loss","Target 1","Target 2",
                 "Qty","Max Risk ₹","Bias","Market"]
        
        hdr_fill=PatternFill(start_color="0D1221",end_color="0D1221",fill_type="solid")
        hdr_font=Font(bold=True,color="F0A500",size=10)
        
        for col,h in enumerate(headers,1):
            cell=ws.cell(row=1,column=col,value=h)
            cell.fill=hdr_fill
            cell.font=hdr_font
            cell.alignment=Alignment(horizontal='center',vertical='center')
        
        ws.row_dimensions[1].height=25
        
        # Column widths
        widths=[12,12,10,8,14,16,10,10,10,12,10,12,10,10,6,10,10,10]
        for i,w in enumerate(widths,1):
            ws.column_dimensions[get_column_letter(i)].width=w

    # Add today's signals
    date_str=ist.strftime("%d-%b-%Y")
    
    for s in signals:
        isL=s["side"]=="LONG"
        row=[
            date_str,
            s["sym"],
            s["sec"],
            s["side"],
            f"{s['score']}/6",
            s["ltp"],
            "",  # Day Open — fill after market
            "",  # Day High
            "",  # Day Low
            s.get("gtt_trigger",s["entry"]),
            s.get("gtt_limit",s["entry"]),
            s["sl"],
            s["t1"],
            s["t2"],
            s["qty"],
            round(abs(s["entry"]-s["sl"])*s["qty"],2),
            s.get("bias",""),
            "NSE F&O",
        ]
        
        # Row styling
        row_num=ws.max_row+1
        fill_color="002D1C" if isL else "2D0010"
        row_fill=PatternFill(start_color=fill_color,end_color=fill_color,fill_type="solid")
        
        for col,val in enumerate(row,1):
            cell=ws.cell(row=row_num,column=col,value=val)
            cell.fill=row_fill
            cell.alignment=Alignment(horizontal='center')
            # Color code important columns
            if col==4:  # Side
                cell.font=Font(color="00D4A0" if isL else "FF4560",bold=True)
            elif col in [10,11]:  # GTT levels
                cell.font=Font(color="F0A500",bold=True)
            elif col==12:  # SL
                cell.font=Font(color="FF4560",bold=True)
            elif col in [13,14]:  # Targets
                cell.font=Font(color="00D4A0",bold=True)
    
    wb.save(fname)
    print(f"✅ Excel updated: {fname} ({ws.max_row-1} signals total)")
    return fname

def main():
    ist=datetime.now(IST)
    print(f"Time: {ist.strftime('%Y-%m-%d %H:%M IST')}")
    mkt_open=ist.hour<15 or (ist.hour==15 and ist.minute<30)
    print(f"Stocks: {len(STOCKS)} | Market: {'OPEN' if mkt_open else 'CLOSED'}")

    jwt=login()

    # Batch fetch all quotes
    stk_toks=[v["token"] for v in STOCKS.values()]
    idx_toks=list(INDICES.values())
    all_fetched=get_quotes(jwt,stk_toks+idx_toks)
    
    Q={}
    for q in all_fetched:
        t=q.get("symbolToken","")
        for sym,info in STOCKS.items():
            if info["token"]==t: Q[sym]=q; break
        for idx,tok in INDICES.items():
            if tok==t: Q[idx]=q; break
    print(f"Quotes: {len(Q)}/{len(STOCKS)}")

    def px(s):
        q=Q.get(s,{})
        # Try ltp first, then close, then open
        for field in ['ltp','close','open']:
            v=float(q.get(field,0) or 0)
            if v>0: return round(v,2)
        return 0.0

    def prev_close(s):
        q=Q.get(s,{})
        # Angel One: "close" field = previous day closing price
        v=float(q.get("close",0) or 0)
        return round(v,2) if v>0 else 0.0

    def chg(s):
        q=Q.get(s,{})
        # Use percentChange directly from Angel One API
        pc=float(q.get("percentChange",0) or 0)
        if pc!=0: return round(pc,2)
        cur=px(s); prev=prev_close(s)
        if cur>0 and prev>0: return round((cur-prev)/prev*100,2)
        return 0.0

    def chg_pts(s):
        q=Q.get(s,{})
        # Use netChange directly from Angel One API
        nc=float(q.get("netChange",0) or 0)
        if nc!=0: return round(nc,2)
        return round(px(s)-prev_close(s),2)

    def hi(s):
        q=Q.get(s,{})
        v=float(q.get("high",0) or 0)
        return v if v>0 else px(s)

    def lo(s):
        q=Q.get(s,{})
        v=float(q.get("low",0) or 0)
        return v if v>0 else px(s)

    def open_px(s):
        q=Q.get(s,{})
        v=float(q.get("open",0) or 0)
        return v if v>0 else px(s)

    # Debug — print ALL fields for NIFTY50 to find correct field names
    nifty_q = Q.get("NIFTY50",{})
    vix_q = Q.get("INDIAVIX",{})
    fn_q = Q.get("FINNIFTY",{})
    print(f"NIFTY50 ALL FIELDS: {json.dumps(nifty_q)[:500]}")
    print(f"VIX ALL FIELDS: {json.dumps(vix_q)[:300]}")
    print(f"FINNIFTY ALL FIELDS: {json.dumps(fn_q)[:300]}")

    nifty_px=px("NIFTY50"); nifty_chg=chg("NIFTY50"); nifty_pts=chg_pts("NIFTY50")
    nifty_prev=prev_close("NIFTY50")
    vix_px=px("INDIAVIX"); vix_chg=chg("INDIAVIX"); vix_prev=prev_close("INDIAVIX")
    fn_px=px("FINNIFTY"); fn_chg=chg("FINNIFTY"); fn_pts=chg_pts("FINNIFTY")
    fn_prev=prev_close("FINNIFTY")
    print(f"Nifty: {nifty_px} prev={nifty_prev} chg={nifty_chg}% pts={nifty_pts}")
    print(f"VIX: {vix_px} prev={vix_prev} chg={vix_chg}%")
    print(f"FINNIFTY: {fn_px} prev={fn_prev} chg={fn_chg}%")

    # FINNIFTY ATM Strike (nearest 50)
    fn_atm = round(fn_px/50)*50 if fn_px>0 else 0

    # FINNIFTY 30-min ORB candle — try multiple intervals
    fn_orb = {"orb_open":0,"orb_high":0,"orb_low":0,"orb_close":0,"signal":"WAIT"}
    try:
        fn_date = ist.strftime("%Y-%m-%d")
        # Try THIRTY_MINUTE first, then fallback to ONE_MINUTE aggregated
        for interval in ["THIRTY_MINUTE","FIFTEEN_MINUTE"]:
            fn_candles = get_candles(jwt,"26037",interval,
                                     f"{fn_date} 09:00",f"{fn_date} 10:30")
            print(f"FINNIFTY candles ({interval}): status={fn_candles.get('status')} count={len(fn_candles.get('data',[]))}")
            if fn_candles.get("status") and fn_candles.get("data"):
                c = fn_candles["data"]
                if len(c)>=1:
                    if interval=="THIRTY_MINUTE":
                        # Single 30-min candle
                        first=c[0]
                    else:
                        # Two 15-min candles = 9:15 + 9:30 = combine for 30-min ORB
                        opens=[float(x[1]) for x in c[:2]]
                        highs=[float(x[2]) for x in c[:2]]
                        lows=[float(x[3]) for x in c[:2]]
                        closes=[float(x[4]) for x in c[:2]]
                        first=[c[0][0],opens[0],max(highs),min(lows),closes[-1],0]
                    fn_orb["orb_open"]  = round(float(first[1]),2)
                    fn_orb["orb_high"]  = round(float(first[2]),2)
                    fn_orb["orb_low"]   = round(float(first[3]),2)
                    fn_orb["orb_close"] = round(float(first[4]),2)
                    # Signal
                    if fn_px>0 and fn_orb["orb_high"]>0:
                        if fn_px > fn_orb["orb_high"]:
                            fn_orb["signal"] = "CE"
                        elif fn_px < fn_orb["orb_low"]:
                            fn_orb["signal"] = "PE"
                        else:
                            fn_orb["signal"] = "MONITOR"
                    print(f"ORB: O={fn_orb['orb_open']} H={fn_orb['orb_high']} L={fn_orb['orb_low']} C={fn_orb['orb_close']} Signal={fn_orb['signal']}")
                    break
    except Exception as e:
        print(f"FINNIFTY candle error: {e}")

    # FINNIFTY Options premium calculation
    fn_ce_premium=0; fn_pe_premium=0
    try:
        if fn_px>0:
            vix_use=vix_px if vix_px>0 else 14.0
            import math as _math
            days_to_expiry=max(1,2)
            time_factor=_math.sqrt(days_to_expiry/252)
            fn_ce_premium=round(fn_px*(vix_use/100)*time_factor,0)
            fn_pe_premium=round(fn_ce_premium*0.92,0)
            print(f"Options: CE=₹{fn_ce_premium} PE=₹{fn_pe_premium}")
    except: pass

    finnifty_data={
        "ltp":fn_px,"chg":fn_chg,"pts":fn_pts,"prev":fn_prev,
        "atm":fn_atm,
        "orb_open":fn_orb["orb_open"],"orb_high":fn_orb["orb_high"],
        "orb_low":fn_orb["orb_low"],"orb_close":fn_orb["orb_close"],
        "signal":fn_orb["signal"],
        "ce_premium":fn_ce_premium,"pe_premium":fn_pe_premium,
        "ce_sl":round(fn_ce_premium*0.7,0) if fn_ce_premium>0 else 0,
        "ce_tgt":round(fn_ce_premium*1.5,0) if fn_ce_premium>0 else 0,
        "pe_sl":round(fn_pe_premium*0.7,0) if fn_pe_premium>0 else 0,
        "pe_tgt":round(fn_pe_premium*1.5,0) if fn_pe_premium>0 else 0,
    }

    # Sectors — with leader stock price + % change
    sectors=[]
    for sec,stks in SECTOR_STOCKS.items():
        valid=[s for s in stks if px(s)>0]
        chgs=[chg(s) for s in valid]
        sec_chg=round(sum(chgs)/len(chgs),2) if chgs else 0.0
        ldr=stks[0]
        sectors.append({
            "name":sec,"chg":sec_chg,"leader":ldr,
            "leader_price":px(ldr),"leader_chg":chg(ldr),
            "leader_pts":chg_pts(ldr),"idx_price":0,
            "stocks":[{"sym":s,"ltp":px(s),"chg":chg(s),"pts":chg_pts(s)} for s in valid]
        })
    sectors.sort(key=lambda x:x["chg"],reverse=True)
    print(f"Top: {sectors[0]['name']} {sectors[0]['chg']}% | Bot: {sectors[-1]['name']} {sectors[-1]['chg']}%")

    # Candles — batched with rate limiting
    to_d=ist.strftime("%Y-%m-%d %H:%M")
    fr_d=(ist-timedelta(days=60)).strftime("%Y-%m-%d %H:%M")
    candles={}
    print(f"Fetching candles for {len(STOCKS)} stocks...")
    for i,(sym,info) in enumerate(STOCKS.items()):
        try:
            cd=get_candles(jwt,info["token"],"ONE_DAY",fr_d,to_d)
            if cd.get("status") and cd.get("data"):
                candles[sym]=cd["data"]
            time.sleep(0.25)
            if i%20==0: print(f"  Candles: {i}/{len(STOCKS)}")
        except: pass
    print(f"Candles done: {len(candles)}")

    # ROCKERS — scan ALL stocks
    top4=[s["name"] for s in sectors[:4]]
    bot4=[s["name"] for s in sectors[-4:]]
    risk=CAPITAL*RISK_PCT/100
    bias="BULLISH" if nifty_chg>0.3 and (vix_px==0 or vix_px<15) else \
         "BEARISH" if nifty_chg<-0.3 or (vix_px>0 and vix_px>16) else "NEUTRAL"
    longs,shorts=[],[]

    for sym,info in STOCKS.items():
        p=px(sym)
        if p<=0: continue
        c=candles.get(sym,[])
        closes=[x[4] for x in c] if c else [p]
        vols=[x[5] for x in c] if c else [1]
        h=hi(sym); l=lo(sym); op=open_px(sym)

        if info["sec"] in top4:
            e20=ema(closes,20); r=rsi_calc(closes); vr=vol_ratio(vols)
            conds={"EMA20":p>e20,"SUPERT":supert(c),"RSI":50<=r<=70,
                   "MACD":macd_bull(closes),"VOL":vr>=1.5,"PDH":p>h}
            sc=sum(conds.values())
            en=round(h*1.002,2); sl=round(en*(1-info["slp"]/100),2); rp=en-sl
            if rp<=0: continue
            longs.append({
                "sym":sym,"sec":info["sec"],"ltp":p,"chg":chg(sym),
                "open":op,"high":h,"low":l,
                "score":sc,"rsi":r,"vol_ratio":vr,
                "conds":{k:bool(v) for k,v in conds.items()},
                "entry":en,"sl":sl,"t1":round(en+rp*1.5,2),"t2":round(en+rp*2.5,2),
                "gtt_trigger":en,"gtt_limit":round(en*1.001,2),
                "qty":max(1,math.floor(risk/rp)),"side":"LONG","bias":bias
            })

        if info["sec"] in bot4:
            e20=ema(closes,20); r=rsi_calc(closes); vr=vol_ratio(vols)
            conds={"EMA20":p<e20,"SUPERT":not supert(c),"RSI":30<=r<=50,
                   "MACD":not macd_bull(closes),"VOL":vr>=1.5,"PDL":p<l}
            sc=sum(conds.values())
            en=round(l*0.998,2); sl=round(en*(1+info["slp"]/100),2); rp=sl-en
            if rp<=0: continue
            shorts.append({
                "sym":sym,"sec":info["sec"],"ltp":p,"chg":chg(sym),
                "open":op,"high":h,"low":l,
                "score":sc,"rsi":r,"vol_ratio":vr,
                "conds":{k:bool(v) for k,v in conds.items()},
                "entry":en,"sl":sl,"t1":round(en-rp*1.5,2),"t2":round(en-rp*2.5,2),
                "gtt_trigger":en,"gtt_limit":round(en*0.999,2),
                "qty":max(1,math.floor(risk/rp)),"side":"SHORT","bias":bias
            })

    longs.sort(key=lambda x:x["score"],reverse=True)
    shorts.sort(key=lambda x:x["score"],reverse=True)
    print(f"Longs: {len(longs)} | Shorts: {len(shorts)}")

    # Update Excel journal
    all_signals=[{**s,"bias":bias} for s in longs[:5]+shorts[:5]]
    if all_signals:
        try:
            update_excel(all_signals,ist)
        except Exception as e:
            print(f"Excel error: {e}")

    # Save data.json
    data={
        "generated_at":ist.strftime("%d %b %Y %I:%M %p IST"),
        "market_status":"OPEN" if mkt_open else "CLOSED",
        "is_live":True,
        "market":{
            "nifty50":{"ltp":nifty_px,"chg":nifty_chg,"pts":nifty_pts,"prev":nifty_prev},
            "vix":{"ltp":vix_px,"chg":vix_chg,"prev":vix_prev},
            "bias":bias
        },
        "finnifty":finnifty_data,
        "sectors":sectors,
        "long":longs[:5],"short":shorts[:5],
        "capital":CAPITAL,"risk_pct":RISK_PCT,"risk_amt":risk,
        "universe":f"Nifty 100 — {len(STOCKS)} stocks",
    }
    with open("data.json","w") as f: json.dump(data,f,indent=2)
    print(f"✅ All done! Bias:{bias}")

if __name__=="__main__":
    main()
