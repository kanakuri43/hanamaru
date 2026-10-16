import json,sys,datetime as dt,urllib.request
def fetch(sym):
    if "." not in sym and sym[:1].isdigit(): sym+=".T"  # 日本株コードは .T 補完
    req=urllib.request.Request(f"https://query1.finance.yahoo.com/v8/finance/chart/{sym}?range=5y&interval=1d",headers={"User-Agent":"Mozilla/5.0"})
    r=json.load(urllib.request.urlopen(req,timeout=30))['chart']['result'][0]
    tz=r['meta'].get('gmtoffset',0)
    q=r['indicators']['quote'][0]
    rows=[(dt.datetime.fromtimestamp(t+tz,dt.timezone.utc).date(),o,h,l,c,v) for t,o,h,l,c,v in zip(r['timestamp'],q['open'],q['high'],q['low'],q['close'],q['volume']) if c]
    return sym,r['meta'],rows
def sma(a,n):return sum(a[-n:])/n
def ema(a,n):
    k=2/(n+1);e=a[0];out=[]
    for x in a:e=x*k+e*(1-k);out.append(e)
    return out
def rsi(a,n=14):
    g=l=0
    for i in range(1,n+1):d=a[i]-a[i-1];g+=max(d,0);l+=max(-d,0)
    g/=n;l/=n
    for i in range(n+1,len(a)):d=a[i]-a[i-1];g=(g*(n-1)+max(d,0))/n;l=(l*(n-1)+max(-d,0))/n
    return 100-100/(1+g/l) if l else 100
def macd(a):
    m=[x-y for x,y in zip(ema(a,12),ema(a,26))];s=ema(m,9);return m[-1],s[-1],m[-2]-s[-2]
def resample(D,H,L,C,key):
    out={}
    for d,h,l,c in zip(D,H,L,C):
        k=key(d)
        if k not in out:out[k]=[h,l,c]
        else:o=out[k];o[0]=max(o[0],h);o[1]=min(o[1],l);o[2]=c
    return list(out.values())
def report(name,h,l,c,mas):
    print(f"== {name} (n={len(c)}) close={c[-1]:.2f}")
    for n in mas:
        if len(c)>=n:print(f"  SMA{n}={sma(c,n):.2f} ({(c[-1]/sma(c,n)-1)*100:+.1f}%)"+(f"  5本前={sum(c[-n-5:-5])/n:.2f}" if len(c)>=n+5 else ""))
    if len(c)>30:
        print(f"  RSI14={rsi(c):.1f}");m,s,ph=macd(c);print(f"  MACD={m:.2f} sig={s:.2f} hist={m-s:.2f} 前hist={ph:.2f}")
    n=20;mu=sma(c,n);sd=(sum((x-mu)**2 for x in c[-n:])/n)**.5;print(f"  BB20 +2σ={mu+2*sd:.2f} -2σ={mu-2*sd:.2f}")
    hh=max(h[-14:]);ll=min(l[-14:]);print(f"  Stoch%K14={(c[-1]-ll)/(hh-ll)*100:.0f}")
    for k in (20,52,60):
        if len(h)>=k:print(f"  {k}本高値={max(h[-k:]):.2f} 安値={min(l[-k:]):.2f}")
if __name__=="__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sym,meta,rows=fetch(sys.argv[1])
    D=[x[0] for x in rows];H=[x[2] for x in rows];L=[x[3] for x in rows];C=[x[4] for x in rows];V=[x[5] or 0 for x in rows]
    print("symbol:",sym,meta.get('longName',''),"last:",D[-1],"days",len(C))
    print("recent:",[(str(d),round(c,2)) for d,c in zip(D[-10:],C[-10:])])
    print(f"vol5/vol20={sum(V[-5:])/5/max(sum(V[-20:])/20,1):.2f}")
    tr=[max(H[i]-L[i],abs(H[i]-C[i-1]),abs(L[i]-C[i-1])) for i in range(1,len(C))];print(f"ATR14={sum(tr[-14:])/14:.2f}")
    report("日足",H,L,C,(5,25,75,200))
    w=resample(D,H,L,C,lambda d:d.isocalendar()[:2]);report("週足",*zip(*w),(13,26,52))
    m=resample(D,H,L,C,lambda d:(d.year,d.month));report("月足",*zip(*m),(12,24,60))
    print("5年高値",max(H),D[H.index(max(H))])
    print("週足終値 直近30:",[round(x[2],2) for x in w[-30:]])
    print("月足終値 直近18:",[round(x[2],2) for x in m[-18:]])
    print("日足スイング安値(60日):",[(str(D[i]),round(L[i],2)) for i in range(len(L)-60,len(L)) if L[i]==min(L[max(0,i-5):i+6])])
