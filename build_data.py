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
DR=[];miss=0
for _,r in c.iterrows():
    k=S(r['CUST_CODE']);q=pd_.get(k)
    j=pr.loc[r['PROJ']];per=int(num(r['PERIOD']));pp_=int(num(q['PERIOD'])) if q is not None else 0
    DR.append([TP[r['CONTRACT_TYPE_CODE']],S(j['กอง']),S(j['สำนักงาน']),S(r['PROJ_NAME']),S(r['HOUSE_NO_STD']).lstrip('0'),S(r['HOUSE_NO_ACT']),S(r['DM_STATUS_NAME']),S(q['DM_STATUS_NAME']) if q is not None else '',
      rng(per),(rng(pp_) if q is not None else ''),per,pp_,round(num(r['ALL_PRICE'])),round(num(q['ALL_PRICE'])) if q is not None else 0,round(ind(r)),round(ind(q)) if q is not None else 0,'ยังไม่ติดตาม','ยังไม่ติดตาม','',''])
json.dump(DR,open('DR_new.json','w'),ensure_ascii=False)
print('rows',len(DR),'curr rows',len(c),'ALL curr',round(sum(r[12] for r in DR),2),c.ALL_PRICE.sum(),'| prev-in-curr ALL',round(sum(r[13] for r in DR),2),'prev sheet',p.ALL_PRICE.sum())
print('ind curr',round(sum(r[14] for r in DR),2),'ind prev(of curr rows)',round(sum(r[15] for r in DR),2),'prev sheet full',round(sum(ind(r) for _,r in p.iterrows()),2))
print('ind curr no-rule',round(c[M].fillna(0).sum().sum(),2),'sued rows w/ indicator zeroed',sum(1 for r in c.itertuples() if r.CUST_CODE in sued))
print(collections.Counter(r[0] for r in DR),collections.Counter(r[1] for r in DR))
