#!/usr/bin/env python3
"""
SUBHA CAPITAL — Market Data Fetcher V10
Data-backed intraday system for NSE

KEY IMPROVEMENTS:
- Universe: Top 20 F&O stocks only (most liquid)
- First candle move filter (skip if >1% already moved)
- ATR filter (skip if ATR < ₹15)
- Best 1 Long + 1 Short only
- ORB-based entry/SL/target
- 5x intraday margin qty
- Gap filter (skip if gap >2%)
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
MARGIN      = 5       # 5x intraday leverage
IST         = pytz.timezone("Asia/Kolkata")
BASE        = "https://apiconnect.angelone.in"

# ── TOP 20 F&O STOCKS — Most liquid, tightest spread ─────────────────
STOCKS = {
    "RELIANCE":   {"token":"2885",  "sec":"ENERGY",  "atr_min":20},
    "HDFCBANK":   {"token":"1333",  "sec":"BANKING", "atr_min":15},
    "ICICIBANK":  {"token":"4963",  "sec":"BANKING", "atr_min":15},
    "INFY":       {"token":"1594",  "sec":"IT",      "atr_min":20},
    "TCS":        {"token":"11536", "sec":"IT",      "atr_min":30},
    "AXISBANK":   {"token":"5900",  "sec":"BANKING", "atr_min":15},
    "SBIN":       {"token":"3045",  "sec":"BANKING", "atr_min":10},
    "BAJFINANCE": {"token":"317",   "sec":"FINANCE", "atr_min":30},
    "KOTAKBANK":  {"token":"1922",  "sec":"BANKING", "atr_min":20},
    "LT":         {"token":"11483", "sec":"INFRA",   "atr_min":40},
    "TATASTEEL":  {"token":"3499",  "sec":"METAL",   "atr_min":10},
    "JSWSTEEL":   {"token":"11723", "sec":"METAL",   "atr_min":15},
    "MARUTI":     {"token":"10999", "sec":"AUTO",    "atr_min":80},
    "MM":         {"token":"2031",  "sec":"AUTO",    "atr_min":30},
    "BHARTIARTL": {"token":"10604", "sec":"TELECOM", "atr_min":10},
    "DLF":        {"token":"14732", "sec":"REALTY",  "atr_min":10},
    "ADANIPORTS": {"token":"15083", "sec":"INFRA",   "atr_min":15},
    "SUNPHARMA":  {"token":"3351",  "sec":"PHARMA",  "atr_min":15},
    "TITAN":      {"token":"3506",  "sec":"CONSUMER","atr_min":30},
    "WIPRO":      {"token":"3787",  "sec":"IT",      "atr_min":10},
}

INDICES = {"NIFTY50":"26000","FINNIFTY":"26037"}

# ── PHASE DETECTION ───────────────────────────────────────────────────
def get_phase(ist):
    h,m=ist.hour,ist.minute
    if h==8 and m>=45: return "PREMARKET_845"
    if h==9 and m<15:  return "PREMARKET_900"
    if h==9 and m<30:  return "MARKET_OPEN_915"
    if h==9 and m<45:  return "ORB_930"
    if 9<=h<15:        return "ORB_945_PLUS"
    return "CLOSED"

def get_phase_label(p):
    return {
        "PREMARKET_845":  "8:45 AM — Pre-Market Brief",
        "PREMARKET_900":  "9:00 AM — Pre-Open Analysis",
        "MARKET_OPEN_915":"9:15 AM — Market Open Watch",
        "ORB_930":        "9:30 AM — ORB Confirmation",
        "ORB_945_PLUS":   "9:45 AM — FINNIFTY ORB Signal",
        "CLOSED":         "Market Closed — EOD Data",
    }.get(p,"—")

# ── GIFT NIFTY ────────────────────────────────────────────────────────
def fetch_gift_nifty():
    gift={"ltp":0,"chg":0,"bias":"—","source":"—"}
    hdrs={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
          "Accept":"application/json"}
    for sym in ["^NSEI","^CNXNIFTY","NIFTYBEES.NS"]:
        try:
            r=requests.get(
                f"https://query1.finance.yahoo.com/v8/finance/chart/{sym}?interval=1d&range=1d",
                headers=hdrs,timeout=10)
            if r.ok:
                result=r.json().get("chart",{}).get("result",[])
                if result:
                    meta=result[0].get("meta",{})
                    p=float(meta.get("regularMarketPrice",0))
                    prev=float(meta.get("previousClose",0) or meta.get("chartPreviousClose",0))
                    if p>0 and prev>0:
                        chg=round((p-prev)/prev*100,2)
                        gift={"ltp":round(p,2),"chg":chg,
                              "bias":"GAP UP ▲ BULLISH" if chg>0.5 else
                                     "SLIGHT GAP UP" if chg>0.2 else
                                     "GAP DOWN ▼ BEARISH" if chg<-0.5 else
                                     "SLIGHT GAP DOWN" if chg<-0.2 else
                                     "FLAT — WAIT",
                              "source":sym}
                        print(f"Gift Nifty: {p} {chg}% — {gift['bias']}")
                        return gift
        except: continue
    print("Gift Nifty: All sources failed")
    return gift

# ── VIX ───────────────────────────────────────────────────────────────
def fetch_vix():
    vix={"ltp":0,"chg":0,"prev":0,"status":"—","source":"—"}
    hdrs={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
          "Accept":"application/json"}
    # Method 1: Yahoo Finance
    try:
        s=requests.Session()
        s.headers.update(hdrs)
        s.get("https://finance.yahoo.com",timeout=6)
        r=s.get("https://query1.finance.yahoo.com/v8/finance/chart/%5EINDIAVIX?interval=1d&range=2d",
                timeout=8)
        if r.ok:
            result=r.json().get("chart",{}).get("result",[])
            if result:
                meta=result[0].get("meta",{})
                p=float(meta.get("regularMarketPrice",0))
                prev=float(meta.get("previousClose",0))
                if p>0:
                    vix={"ltp":round(p,2),"prev":round(prev,2),
                         "chg":round(p-prev,2),
                         "status":"<13 SAFE" if p<13 else "13-16 CAUTION" if p<16 else ">16 HIGH RISK",
                         "source":"Yahoo"}
                    print(f"VIX: {p} — {vix['status']}")
                    return vix
    except: pass
    # Method 2: Yahoo v2
    try:
        r=requests.get(
            "https://query2.finance.yahoo.com/v8/finance/chart/%5EINDIAVIX?interval=1d&range=2d",
            headers=hdrs,timeout=8)
        if r.ok:
            result=r.json().get("chart",{}).get("result",[])
            if result:
                meta=result[0].get("meta",{})
                p=float(meta.get("regularMarketPrice",0))
                if p>0:
                    vix["ltp"]=round(p,2)
                    vix["status"]="<13 SAFE" if p<13 else "13-16 CAUTION" if p<16 else ">16 HIGH RISK"
                    vix["source"]="Yahoo v2"
                    print(f"VIX v2: {p}")
                    return vix
    except: pass
    print("VIX: All sources failed")
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

def get_quotes(tok,tokens):
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

def get_candles(tok,sym_tok,interval,fr,to):
    try:
        r=requests.post(f"{BASE}/rest/secure/angelbroking/historical/v1/getCandleData",
            json={"exchange":"NSE","symboltoken":sym_tok,
                  "interval":interval,"fromdate":fr,"todate":to},
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

def atr_calc(candles,n=14):
    """Calculate 14-day ATR"""
    if len(candles)<2: return 0
    trs=[]
    for i in range(1,len(candles)):
        h=float(candles[i][2]); l=float(candles[i][3]); pc=float(candles[i-1][4])
        tr=max(h-l,abs(h-pc),abs(l-pc))
        trs.append(tr)
    if not trs: return 0
    atr=sum(trs[:n])/min(n,len(trs))
    for tr in trs[n:]:
        atr=(atr*(n-1)+tr)/n
    return round(atr,2)

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
    cum_tpv=cum_vol=0
    for c in candles[-20:]:
        tp=(float(c[2])+float(c[3])+float(c[4]))/3
        vol=float(c[5])
        cum_tpv+=tp*vol; cum_vol+=vol
    return round(cum_tpv/cum_vol,2) if cum_vol>0 else 0

# ── EXCEL JOURNAL ─────────────────────────────────────────────────────
def update_excel(signals,ist):
    fname="trading_journal.xlsx"
    try:
        wb=openpyxl.load_workbook(fname); ws=wb.active
    except:
        wb=openpyxl.Workbook(); ws=wb.active; ws.title="SUBHA CAPITAL"
        headers=["Date","Stock","Sector","Side","Score","ATR","First Candle%",
                 "Pre-Mkt Price","ORB High","ORB Low","ORB Range",
                 "Entry","SL","T1","T2","Qty","Max Risk ₹","Potential ₹",
                 "Bias","VIX","Result","P&L ₹","Notes"]
        hf=PatternFill(start_color="0D1221",end_color="0D1221",fill_type="solid")
        hfont=Font(bold=True,color="F0A500",size=10)
        for col,h in enumerate(headers,1):
            cell=ws.cell(row=1,column=col,value=h)
            cell.fill=hf; cell.font=hfont
            cell.alignment=Alignment(horizontal='center',vertical='center')
        ws.row_dimensions[1].height=28
        widths=[12,12,10,8,7,8,12,14,10,10,10,10,10,10,10,6,10,12,10,8,10,10,20]
        for i,w in enumerate(widths,1):
            ws.column_dimensions[get_column_letter(i)].width=w

    date_str=ist.strftime("%d-%b-%Y")
    for s in signals:
        isL=s["side"]=="LONG"
        row=[date_str,s["sym"],s["sec"],s["side"],f"{s['score']}/6",
             s.get("atr",0),f"{s.get('first_candle_pct',0):.2f}%",
             s["ltp"],s.get("orb_high",0),s.get("orb_low",0),s.get("orb_range",0),
             s["entry"],s["sl"],s["t1"],s["t2"],s["qty"],
             s.get("max_risk",0),s.get("potential_profit",0),
             s.get("bias",""),s.get("vix",0),"","",""]
        rn=ws.max_row+1
        fc="002D1C" if isL else "2D0010"
        rf=PatternFill(start_color=fc,end_color=fc,fill_type="solid")
        for col,val in enumerate(row,1):
            cell=ws.cell(row=rn,column=col,value=val)
            cell.fill=rf; cell.alignment=Alignment(horizontal='center')
            if col==4: cell.font=Font(color="00D4A0" if isL else "FF4560",bold=True)
            elif col in [12]: cell.font=Font(color="F0A500",bold=True)
            elif col==13: cell.font=Font(color="FF4560",bold=True)
            elif col in [14,15]: cell.font=Font(color="00D4A0",bold=True)
    wb.save(fname)
    print(f"✅ Excel: {ws.max_row-1} signals")

# ── MAIN ──────────────────────────────────────────────────────────────
def main():
    ist=datetime.now(IST)
    phase=get_phase(ist)
    phase_label=get_phase_label(phase)
    print(f"{'='*60}")
    print(f"SUBHA CAPITAL V10 — {ist.strftime('%d %b %Y %H:%M IST')}")
    print(f"Phase: {phase} — {phase_label}")
    print(f"Universe: Top 20 F&O stocks")
    print(f"{'='*60}")

    mkt_open=(ist.hour>9 or (ist.hour==9 and ist.minute>=15)) and \
             (ist.hour<15 or (ist.hour==15 and ist.minute<30))

    # ── GIFT NIFTY + VIX ─────────────────────────────────────────────
    print("\nFetching Gift Nifty...")
    gift=fetch_gift_nifty()
    print("\nFetching India VIX...")
    vix_data=fetch_vix()

    # ── LOGIN ─────────────────────────────────────────────────────────
    jwt=login()

    # ── QUOTES ───────────────────────────────────────────────────────
    all_toks=[v["token"] for v in STOCKS.values()]+list(INDICES.values())
    fetched=get_quotes(jwt,all_toks)
    Q={}
    for q in fetched:
        t=q.get("symbolToken","")
        for sym,info in STOCKS.items():
            if info["token"]==t: Q[sym]=q; break
        for idx,tok in INDICES.items():
            if tok==t: Q[idx]=q; break
    print(f"Quotes: {len(Q)}/{len(STOCKS)+len(INDICES)}")

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

    def day_high(s,c=[]):
        v=round(F(s,"high"),2)
        if v>0: return v
        return round(float(c[-1][2]),2) if c else ltp(s)

    def day_low(s,c=[]):
        v=round(F(s,"low"),2)
        if v>0: return v
        return round(float(c[-1][3]),2) if c else ltp(s)

    def day_open(s,c=[]):
        v=round(F(s,"open"),2)
        if v>0: return v
        return round(float(c[-1][1]),2) if c else ltp(s)

    # ── INDEX DATA ────────────────────────────────────────────────────
    nifty_ltp   = ltp("NIFTY50")
    nifty_chg   = pct_chg("NIFTY50")
    nifty_pts   = pts_chg("NIFTY50")
    nifty_prev  = prev_close("NIFTY50")
    fn_ltp      = ltp("FINNIFTY")
    fn_chg      = pct_chg("FINNIFTY")
    fn_pts      = pts_chg("FINNIFTY")
    fn_prev     = prev_close("FINNIFTY")
    print(f"Nifty: {nifty_ltp} | {nifty_chg}% | {nifty_pts}pts")
    print(f"FINNIFTY: {fn_ltp} | {fn_chg}% | {fn_pts}pts")

    bias="BULLISH" if nifty_chg>0.3 else "BEARISH" if nifty_chg<-0.3 else "NEUTRAL"

    # ── FINNIFTY ORB ──────────────────────────────────────────────────
    fn_atm=round(fn_ltp/50)*50 if fn_ltp>0 else 0
    fn_orb={"open":0,"high":0,"low":0,"close":0,"signal":"WAIT","ce_prem":0,"pe_prem":0}

    if phase in ["ORB_945_PLUS","CLOSED","ORB_930","MARKET_OPEN_915"]:
        try:
            fn_date=ist.strftime("%Y-%m-%d")
            for interval,label in [("THIRTY_MINUTE","30min"),("FIFTEEN_MINUTE","15min"),("ONE_MINUTE","1min")]:
                try:
                    cd=get_candles(jwt,"26037",interval,f"{fn_date} 09:00",f"{fn_date} 15:30")
                    cl=cd.get("data",[]) if cd.get("status") else []
                    print(f"FINNIFTY {label}: {len(cl)} candles")
                    if cl:
                        if interval=="THIRTY_MINUTE":
                            c=cl[0]
                            fn_orb["open"]=round(float(c[1]),2); fn_orb["high"]=round(float(c[2]),2)
                            fn_orb["low"]=round(float(c[3]),2);  fn_orb["close"]=round(float(c[4]),2)
                        elif interval=="FIFTEEN_MINUTE":
                            c2=cl[:2]
                            fn_orb["open"]=round(float(c2[0][1]),2)
                            fn_orb["high"]=round(max(float(x[2]) for x in c2),2)
                            fn_orb["low"]=round(min(float(x[3]) for x in c2),2)
                            fn_orb["close"]=round(float(c2[-1][4]),2)
                        elif interval=="ONE_MINUTE":
                            c30=cl[:30]
                            fn_orb["open"]=round(float(c30[0][1]),2)
                            fn_orb["high"]=round(max(float(x[2]) for x in c30),2)
                            fn_orb["low"]=round(min(float(x[3]) for x in c30),2)
                            fn_orb["close"]=round(float(c30[-1][4]),2)
                        if fn_orb["high"]>0:
                            if fn_ltp>fn_orb["high"]: fn_orb["signal"]="CE"
                            elif fn_ltp<fn_orb["low"]: fn_orb["signal"]="PE"
                            else: fn_orb["signal"]="MONITOR"
                            print(f"ORB OK: H={fn_orb['high']} L={fn_orb['low']} Sig={fn_orb['signal']}")
                            break
                except Exception as e:
                    print(f"FINNIFTY {label} error: {e}")
        except Exception as e:
            print(f"FINNIFTY ORB error: {e}")

    if fn_ltp>0:
        fn_orb["ce_prem"]=round(fn_ltp*0.005,0)
        fn_orb["pe_prem"]=round(fn_ltp*0.0046,0)

    # ── CANDLES ───────────────────────────────────────────────────────
    to_d=ist.strftime("%Y-%m-%d %H:%M")
    fr_d=(ist-timedelta(days=30)).strftime("%Y-%m-%d %H:%M")
    candles={}
    print(f"\nFetching candles for {len(STOCKS)} stocks...")
    for i,(sym,info) in enumerate(STOCKS.items()):
        try:
            cd=get_candles(jwt,info["token"],"ONE_DAY",fr_d,to_d)
            if cd.get("status") and cd.get("data"):
                candles[sym]=cd["data"]
            time.sleep(0.2)
        except: pass
        if (i+1)%10==0: print(f"  Candles: {i+1}/{len(STOCKS)}")
    print(f"Candles: {len(candles)}/{len(STOCKS)}")

    # ── ROCKERS SCAN — TOP 20 WITH SMART FILTERS ─────────────────────
    risk_amt=CAPITAL*RISK_PCT/100
    candidates=[]
    skipped=[]

    for sym,info in STOCKS.items():
        p=ltp(sym)
        if p<=0: continue

        c=candles.get(sym,[])
        closes=[float(x[4]) for x in c] if c else [p]
        vols=[float(x[5]) for x in c] if c else [1]

        h=day_high(sym,c); l=day_low(sym,c); op=day_open(sym,c)
        if h<=0: h=p
        if l<=0: l=p
        if op<=0: op=p

        # ATR calculation
        atr=atr_calc(c) if len(c)>14 else 0

        # VWAP
        vwap_val=calc_vwap(c)

        # First candle move % (how much stock already moved from open)
        first_candle_pct=abs(h-op)/op*100 if op>0 else 0

        # Gap % from prev close
        prev=prev_close(sym)
        gap_pct=abs(op-prev)/prev*100 if prev>0 else 0

        # Indicators
        e20=ema(closes,20); r=rsi(closes); vr=vol_ratio(vols)
        macd_b=macd_bull(closes); st_b=st_bull(c)
        vwap_long=p>vwap_val if vwap_val>0 else p>e20
        vwap_short=p<vwap_val if vwap_val>0 else p<e20

        # ── SMART FILTERS ─────────────────────────────────────────────
        skip_reason=None

        # Filter 1: Already moved too much in first candle
        if first_candle_pct>1.0:
            skip_reason=f"First candle moved {first_candle_pct:.1f}% — momentum used up"

        # Filter 2: ATR too low (stock won't move enough)
        elif atr>0 and atr<info["atr_min"]:
            skip_reason=f"ATR ₹{atr:.0f} below minimum ₹{info['atr_min']} — not enough move"

        # Filter 3: Gap too large (move already priced in)
        elif gap_pct>2.0:
            skip_reason=f"Gap {gap_pct:.1f}% at open — move already priced in"

        # Filter 4: VIX too high — skip equity
        elif vix_data["ltp"]>16:
            skip_reason=f"VIX {vix_data['ltp']} > 16 — too risky for equity"

        # Filter 5: Volume too low
        elif vr<1.0:
            skip_reason=f"Volume {vr}x below average — no participation"

        if skip_reason:
            skipped.append({
                "sym":sym,"sec":info["sec"],"ltp":p,
                "chg":pct_chg(sym),"reason":skip_reason,
                "first_candle_pct":round(first_candle_pct,2),
                "atr":atr,"gap_pct":round(gap_pct,2)
            })
            print(f"  SKIP {sym}: {skip_reason}")
            continue

        # ── 6-CONDITION ROCKERS SCORE ──────────────────────────────────
        # LONG conditions
        long_conds={
            "EMA20": p>e20,
            "SUPERT": st_b,
            "RSI": 50<=r<=70,
            "MACD": macd_b,
            "VOL": vr>=1.5,
            "VWAP": vwap_long,
        }
        long_score=sum(long_conds.values())

        # SHORT conditions
        short_conds={
            "EMA20": p<e20,
            "SUPERT": not st_b,
            "RSI": 30<=r<=50,
            "MACD": not macd_b,
            "VOL": vr>=1.5,
            "VWAP": vwap_short,
        }
        short_score=sum(short_conds.values())

        # ORB-based levels
        orb_range=round(h-l,2)
        orb_mid=round((h+l)/2,2)

        if long_score>=4:
            en=round(h+0.05,2)
            sl=round(orb_mid-0.05,2)
            rp=round(en-sl,2)
            if rp>0:
                t1=round(en+orb_range*0.5,2)
                t2=round(en+orb_range*1.0,2)
                qty=min(
                    int((CAPITAL*MARGIN)/en),
                    max(1,int(risk_amt/rp)*3)
                )
                candidates.append({
                    "sym":sym,"sec":info["sec"],"side":"LONG",
                    "ltp":p,"chg":pct_chg(sym),"prev":prev,
                    "open":op,"high":h,"low":l,"vwap":vwap_val,
                    "orb_high":h,"orb_low":l,"orb_mid":orb_mid,"orb_range":orb_range,
                    "atr":atr,"first_candle_pct":round(first_candle_pct,2),
                    "gap_pct":round(gap_pct,2),
                    "score":long_score,"rsi":r,"vol_ratio":vr,
                    "conds":{k:bool(v) for k,v in long_conds.items()},
                    "entry":en,"sl":sl,"t1":t1,"t2":t2,
                    "gtt_trigger":en,"gtt_limit":round(en+0.05,2),
                    "sl_trigger":sl,"sl_limit":round(sl-0.05,2),
                    "qty":qty,"risk_per_share":rp,
                    "max_risk":round(rp*qty,0),
                    "potential_profit":round((t1-en)*qty,0),
                    "bias":bias,"vix":vix_data["ltp"],
                    "phase":phase
                })

        if short_score>=4:
            en=round(l-0.05,2)
            sl=round(orb_mid+0.05,2)
            rp=round(sl-en,2)
            if rp>0:
                t1=round(en-orb_range*0.5,2)
                t2=round(en-orb_range*1.0,2)
                qty=min(
                    int((CAPITAL*MARGIN)/en),
                    max(1,int(risk_amt/rp)*3)
                )
                candidates.append({
                    "sym":sym,"sec":info["sec"],"side":"SHORT",
                    "ltp":p,"chg":pct_chg(sym),"prev":prev,
                    "open":op,"high":h,"low":l,"vwap":vwap_val,
                    "orb_high":h,"orb_low":l,"orb_mid":orb_mid,"orb_range":orb_range,
                    "atr":atr,"first_candle_pct":round(first_candle_pct,2),
                    "gap_pct":round(gap_pct,2),
                    "score":short_score,"rsi":r,"vol_ratio":vr,
                    "conds":{k:bool(v) for k,v in short_conds.items()},
                    "entry":en,"sl":sl,"t1":t1,"t2":t2,
                    "gtt_trigger":en,"gtt_limit":round(en-0.05,2),
                    "sl_trigger":sl,"sl_limit":round(sl+0.05,2),
                    "qty":qty,"risk_per_share":rp,
                    "max_risk":round(rp*qty,0),
                    "potential_profit":round((en-t1)*qty,0),
                    "bias":bias,"vix":vix_data["ltp"],
                    "phase":phase
                })

    # Sort by score descending
    candidates.sort(key=lambda x:x["score"],reverse=True)

    # Best 1 Long + 1 Short
    longs  = [c for c in candidates if c["side"]=="LONG"]
    shorts = [c for c in candidates if c["side"]=="SHORT"]

    print(f"\nCandidates: {len(longs)} Long | {len(shorts)} Short")
    print(f"Skipped: {len(skipped)} stocks")
    print(f"Top Long: {longs[0]['sym'] if longs else 'None'}")
    print(f"Top Short: {shorts[0]['sym'] if shorts else 'None'}")
    print(f"Bias: {bias}")

    # ── EXCEL ─────────────────────────────────────────────────────────
    trade_signals=[{**s} for s in longs[:1]+shorts[:1]]
    if trade_signals:
        try: update_excel(trade_signals,ist)
        except Exception as e: print(f"Excel error: {e}")

    # ── SAVE DATA.JSON ────────────────────────────────────────────────
    data={
        "generated_at": ist.strftime("%d %b %Y %I:%M %p IST"),
        "phase": phase,
        "phase_label": phase_label,
        "market_status": "OPEN" if mkt_open else "CLOSED",
        "market":{
            "nifty50":{"ltp":nifty_ltp,"chg":nifty_chg,"pts":nifty_pts,"prev":nifty_prev},
            "vix":{"ltp":vix_data["ltp"],"chg":vix_data["chg"],"prev":vix_data["prev"],
                   "status":vix_data["status"],"source":vix_data["source"]},
            "bias":bias
        },
        "gift_nifty": gift,
        "pre_market":{
            "analysis": f"Gift Nifty {gift['bias']}. {bias} bias. VIX {vix_data['ltp'] or 'N/A'}.",
            "top_gainers": sorted([{"sym":s,"ltp":ltp(s),"chg":pct_chg(s),"sec":info["sec"]}
                                   for s,info in STOCKS.items() if ltp(s)>0],
                                  key=lambda x:x["chg"],reverse=True)[:3],
            "top_losers":  sorted([{"sym":s,"ltp":ltp(s),"chg":pct_chg(s),"sec":info["sec"]}
                                   for s,info in STOCKS.items() if ltp(s)>0],
                                  key=lambda x:x["chg"])[:3],
        },
        "finnifty":{
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
        "sectors": [
            {"name":sec,"chg":round(sum(pct_chg(s) for s in stks if ltp(s)>0)/
                                   max(1,len([s for s in stks if ltp(s)>0])),2),
             "leader":max(stks,key=lambda s:abs(pct_chg(s)) if ltp(s)>0 else 0),
             "leader_price":ltp(max(stks,key=lambda s:abs(pct_chg(s)) if ltp(s)>0 else 0)),
             "leader_chg":pct_chg(max(stks,key=lambda s:abs(pct_chg(s)) if ltp(s)>0 else 0)),
             "stocks":[{"sym":s,"ltp":ltp(s),"chg":pct_chg(s)} for s in stks if ltp(s)>0]}
            for sec,stks in {
                "BANKING":["HDFCBANK","ICICIBANK","AXISBANK","SBIN","KOTAKBANK"],
                "IT":["INFY","TCS","WIPRO"],
                "FINANCE":["BAJFINANCE"],
                "AUTO":["MARUTI","MM"],
                "PHARMA":["SUNPHARMA"],
                "METAL":["TATASTEEL","JSWSTEEL"],
                "ENERGY":["RELIANCE"],
                "REALTY":["DLF"],
                "INFRA":["LT","ADANIPORTS"],
                "CONSUMER":["TITAN"],
                "TELECOM":["BHARTIARTL"],
            }.items()
        ],
        "rockers": candidates[:10],        # Top 10 for screener display
        "long":    longs[:5],              # Top 5 longs for screener
        "short":   shorts[:5],             # Top 5 shorts for screener
        "trade_long":  longs[:1],          # BEST 1 LONG for GTT
        "trade_short": shorts[:1],         # BEST 1 SHORT for GTT
        "skipped": skipped[:5],            # Skipped with reasons
        "capital":CAPITAL,"risk_pct":RISK_PCT,"risk_amt":risk_amt,
        "universe": f"Top 20 F&O stocks",
        "vix_value": vix_data["ltp"],
    }

    with open("data.json","w") as f:
        json.dump(data,f,indent=2)
    print(f"\n✅ data.json saved!")
    print(f"Trade: {longs[0]['sym'] if longs else 'No long'} LONG + {shorts[0]['sym'] if shorts else 'No short'} SHORT")
    print(f"{'='*60}")

if __name__=="__main__":
    main()
