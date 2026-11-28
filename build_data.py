import pandas as pd,json,gzip,base64,re,re as _re,collections
import sys
X=sys.argv[1]  # path to the new Dashboard .xlsx
rd=lambda s:pd.read_excel(X,s,dtype={'CUST_CODE':str,'HOUSE_NO_STD':str,'HOUSE_NO_ACT':str,'PROJ':str,'CONTRACT_TYPE_CODE':str}).fillna({}) 
c,p,f,pr=[rd(s) for s in ['เดือนนี้','เดือนก่อน','ฟ้องปีก่อน','โครงการ']]
pr=pr.drop_duplicates('PROJ_CODE').set_index('PROJ_CODE')
sued=set(f.CUST_CODE)
TP={'20':'เช่าซื้อ','50':'เช่า','70':'เช่าจัดประโยชน์','80':'เช่าซื้อ-คบส.'}
M=['M4_M6_PRICE','M7_M12_PRICE','M13_M24_PRICE','MORE24_PRICE']
def S(v):
    if v is None or (isinstance(v,float) and v!=v): return ''
    return str(v).strip()
def num(v):
    try:
        v=float(v);return 0.0 if v!=v else v
    except: return 0.0
def rng(n): return 'ไม่ค้าง' if n<=0 else 'ค้าง 1-2 งวด' if n<=2 else 'ค้าง 3 งวด' if n==3 else 'เกิน 3 งวด'
def ind(r): return 0.0 if S(r['CUST_CODE']) in sued else round(sum(num(r[m]) for m in M),2)
assert set(c.CONTRACT_TYPE_CODE)<=set(TP) and set(p.CONTRACT_TYPE_CODE)<=set(TP)
pd_={S(r['CUST_CODE']):r for _,r in p.iterrows()}
cd_={S(r['CUST_CODE']):r for _,r in c.iterrows()}
def due(x):
    """DUE_PRICE of that month's row (rounded), '' when the row is missing or the value is blank"""
    if x is None: return ''
    v=num(x['DUE_PRICE']) if S(x['DUE_PRICE'])!='' else None
    return '' if v is None else round(v)
def row(r,q,ch):
    """r = this-month row or None, q = previous-month row or None; ch: B both, N new (this month only), X closed (prev only)"""
    base=r if r is not None else q
    j=pr.loc[base['PROJ']]
    per=int(num(r['PERIOD'])) if r is not None else 0
    pp_=int(num(q['PERIOD'])) if q is not None else 0
    return [TP[base['CONTRACT_TYPE_CODE']],S(j['กอง']),S(j['สำนักงาน']),S(base['PROJ_NAME']),S(base['HOUSE_NO_STD']).lstrip('0'),S(base['HOUSE_NO_ACT']),
      S(r['DM_STATUS_NAME']) if r is not None else '',S(q['DM_STATUS_NAME']) if q is not None else '',
      rng(per) if r is not None else '',rng(pp_) if q is not None else '',per,pp_,
      round(num(r['ALL_PRICE'])) if r is not None else 0,round(num(q['ALL_PRICE'])) if q is not None else 0,
      round(ind(r)) if r is not None else 0,round(ind(q)) if q is not None else 0,'ยังไม่ติดตาม','ยังไม่ติดตาม','','',ch,S(base['CUST_CODE']),due(r),due(q)]
DR=[row(r,pd_.get(S(r['CUST_CODE'])),'B' if S(r['CUST_CODE']) in pd_ else 'N') for _,r in c.iterrows()]
DR+=[row(None,q,'X') for k,q in pd_.items() if k not in cd_]
json.dump(DR,open('DR_new.json','w'),ensure_ascii=False)
cnt=collections.Counter(r[20] for r in DR)
print('rows',len(DR),dict(cnt),'curr view',cnt['B']+cnt['N'],'prev view',cnt['B']+cnt['X'])
print('ALL curr',sum(r[12] for r in DR),'sheet',round(c.ALL_PRICE.sum()),'| ALL prev',sum(r[13] for r in DR),'sheet',round(p.ALL_PRICE.sum()))
print('ind curr',sum(r[14] for r in DR),'prev',sum(r[15] for r in DR),'prev sheet by rule',round(sum(ind(r) for _,r in p.iterrows())))
print('X rows debt',sum(r[13] for r in DR if r[20]=='X'),'N rows debt',sum(r[12] for r in DR if r[20]=='N'))
