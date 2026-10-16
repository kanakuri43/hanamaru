# プライム市場の全銘柄を ta.py と同じ指標で機械判定し、◎条件を満たす時間軸を出力する
import io,sys,os,urllib.request,openpyxl
from concurrent.futures import ThreadPoolExecutor
sys.path.insert(0,os.path.join(os.path.dirname(os.path.abspath(__file__)),"..","tech-analysis"))
sys.dont_write_bytecode=True  # コピー時に __pycache__ が混ざらないように
from ta import fetch,sma,ema,rsi,resample
JPX="https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xlsx"

def prime_list():
    req=urllib.request.Request(JPX,headers={"User-Agent":"Mozilla/5.0"})
    ws=openpyxl.load_workbook(io.BytesIO(urllib.request.urlopen(req,timeout=60).read()),read_only=True).active
    return [(str(r[1]),r[2]) for r in list(ws.iter_rows(values_only=True))[1:] if str(r[3]).startswith("プライム")]

def judge(h,l,c,s,m,v=None):
    """◎判定。基本4条件(上昇トレンド/MACD GC 3本以内/RSI<70・Stoch<80/短期線まで2ATR以内)を満たさなければNone、
    満たすが精査条件(急騰/上値余地/出来高)で落ちたら理由の文字列、全部通れば指標dict"""
    if len(c)<max(m+5,61):return None
    sm,sm5,mm,mm5=sma(c,s),sum(c[-s-5:-5])/s,sma(c,m),sum(c[-m-5:-5])/m
    if not(sm>mm and sm>sm5 and mm>mm5 and c[-1]>mm):return None
    mac=[x-y for x,y in zip(ema(c,12),ema(c,26))];hist=[x-y for x,y in zip(mac,ema(mac,9))]
    if not(hist[-1]>0 and min(hist[-4:-1])<=0):return None
    r=rsi(c);hh,ll=max(h[-14:]),min(l[-14:]);st=(c[-1]-ll)/(hh-ll)*100 if hh>ll else 50
    if r>=70 or st>=80:return None
    atr=sum(max(h[i]-l[i],abs(h[i]-c[i-1]),abs(l[i]-c[i-1])) for i in range(len(c)-14,len(c)))/14
    d=(c[-1]-sm)/atr  # ponytail: サポート=SMA(s)のみ。スイング安値も見るならここに追加
    if d>2:return None
    stop=sm-atr;res=max(h[-61:-1])  # 損切り=短期線-1ATR、上値抵抗=直近60本高値(当日除く)
    up5=(c[-1]-c[-6])/atr
    if up5>3:return f"急騰(5本で+{up5:.1f}ATR)"
    rr=(res-c[-1])/(c[-1]-stop)
    if c[-1]<res and rr<1.5:return f"上値抵抗{res:g}が近い(RR{rr:.1f})"
    vr=sum(v[-5:])/5/max(sum(v[-20:])/20,1) if v else None
    if vr is not None and vr<0.7:return f"出来高細り(5/20日比{vr:.2f})"
    return dict(rsi=r,stoch=st,atr_dist=d,stop=stop,target=None if c[-1]>=res else res,rr=None if c[-1]>=res else rr)

def scan(item):
    code,name=item
    try:
        sym,meta,rows=fetch(code)
    except Exception:
        try:sym,meta,rows=fetch(code)  # Yahoo の一時エラー対策で1回だけ再試行
        except Exception as e:return code,name,None,str(e)
    D=[x[0] for x in rows];H=[x[2] for x in rows];L=[x[3] for x in rows];C=[x[4] for x in rows];V=[x[5] or 0 for x in rows]
    w=list(zip(*resample(D,H,L,C,lambda d:d.isocalendar()[:2])));mo=list(zip(*resample(D,H,L,C,lambda d:(d.year,d.month))))
    js=(("日足",judge(H,L,C,25,75,V)),("週足",judge(*w,13,26)),("月足",judge(*mo,12,24)))  # 出来高は日足のみ判定
    return sym,name,(D[-1],C[-1],{tf:j for tf,j in js if isinstance(j,dict)},{tf:j for tf,j in js if isinstance(j,str)}),None

def fmt(j):
    t=f"目標{j['target']:g}(RR{j['rr']:.1f})" if j['target'] else "高値更新中"
    return f"RSI{j['rsi']:.0f},St{j['stoch']:.0f},短期線{j['atr_dist']:+.1f}ATR,損切り{j['stop']:.0f},{t}"

if __name__=="__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    lst=prime_list();lim=int(sys.argv[1]) if len(sys.argv)>1 else None
    with ThreadPoolExecutor(8) as ex:res=list(ex.map(scan,lst[:lim]))
    err=[(c,e) for c,_,_,e in res if e]
    hits=[(s,n,*v[:3]) for s,n,v,e in res if v and v[2]]
    drops=[(s,n,v[3]) for s,n,v,e in res if v and v[3] and not v[2]]
    print(f"scanned={len(res)} errors={len(err)} hits={len(hits)} dropped={len(drops)} last={max((v[0] for _,_,v,_ in res if v),default='-')}")
    print("symbol	銘柄名	終値	月足	週足	日足	詳細")
    for s,n,d,c,h in sorted(hits,key=lambda x:(-len(x[4]),-max(j['rr'] or 9 for j in x[4].values()))):
        print(f"{s}	{n}	{c:g}	"+"".join(("◎" if tf in h else "")+"	" for tf in ("月足","週足","日足"))+" ".join(f"{tf}:{fmt(j)}" for tf,j in h.items()))
    print("\n# 基本条件は満たしたが精査条件で除外")
    for s,n,r in drops:print(f"DROP	{s}	{n}	"+" ".join(f"{tf}:{x}" for tf,x in r.items()))
    for c,e in err[:5]:print("ERR",c,e)
