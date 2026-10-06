import pandas as pd,warnings,sys,json,openpyxl
warnings.filterwarnings('ignore')
src=sys.argv[1]
df=pd.read_excel(src,sheet_name='FGI 8.5-01.02 PLANEJADO',header=7,usecols='B:S').dropna(subset=['OS'])
a=df[(df['PROJETO ATIVO PLANEJAMENTO']=='SIM')&(df['ORIGEM']!='CANCELADO')].copy()
a['m']=pd.to_datetime(a['Mês Planejado'],errors='coerce')
# tempos matrix from their per-stage tables
wb=openpyxl.load_workbook(src,data_only=True,read_only=True); t=wb['TEMPOS']
cols={'MEC':('X','Z'),'ELE':('AS','AU'),'PLAT':('AL','AN'),'CDP':('AE','AG'),'NORM/INSP/EMB':('AZ','BB')}
from openpyxl.utils import column_index_from_string as ci
rows=list(t.iter_rows(min_row=3,max_row=17,values_only=True))
tempo={}
for st,(k,h) in cols.items():
    for r in rows:
        tempo.setdefault(r[ci(k)-1],{})[st]=r[ci(h)-1]
keys=[]
for r in rows:  # ordered keys from MEC table
    keys.append((r[ci('V')-1],r[ci('W')-1],r[ci('X')-1]))
tempo={k:v for k,v in tempo.items()}
print({k:v for k,v in tempo.items()})
stages=['MEC','ELE','PLAT','CDP','NORM/INSP/EMB']
out=[]
for os_,g in a.groupby('OS',sort=False):
    f=g.iloc[0]; key=f['CONCATENAR2']; rec=dict(OS=int(os_),proj=str(f['CÓDIGO DO PROJETO']).strip(),se=str(f['SUBESTAÇÃO']).strip(),tag=str(f['TAG']).strip(),tipo=f['TIPO PNL'],comp=f['COMPLEXIDADE'],testes=f['TESTES'],giga=f['GIGA'],orig=f['ORIGEM'])
    for s in stages:
        gs=g[g['ETAPA']==s]
        if len(gs)==0: rec[s]=None; continue
        pad=tempo.get(key,{}).get(s) or 0
        h=gs['Horas / Mês'].sum()
        pct = round(h/pad,4) if pad else None
        if pad==0: print('NO PAD',os_,s,key,h)
        mm=gs['m'].min()
        rec[s]=(None if pd.isna(mm) else mm.strftime('%Y-%m-%d'), pct)
    out.append(rec)
# cronograma extras
c=pd.read_excel(src,sheet_name='Cronograma SET-OUT',header=1,usecols='A:V').dropna(subset=['OS'])
extra={}
for _,r in c.iterrows():
    def dt(v):
        try: return pd.to_datetime(v).strftime('%Y-%m-%d')
        except: return None
    extra[int(r['OS'])]=dict(alvo=dt(r['Data Alvo']) if str(r['Meta']).strip().upper()=='EMBALADO' else None,doc='S' if str(r['Data recebimento Documentação']).strip().upper()=='DISPONIVEL' else 'N',
        mat='S' if str(r['Data recebimento Material Critico']).strip().upper()=='RECEBIDO' else 'N')
for r in out: r.update(extra.get(r['OS'],{}))
json.dump(dict(rows=out,keys=keys,tempo=tempo),open('pmp_data.json','w'),ensure_ascii=False,default=str)
print(len(out),'paineis;',sum(1 for r in out if 'alvo' in r),'com cronograma')
# precedence violations
order=stages; v={}
import itertools
for r in out:
    ms=[(s,r[s][0]) for s in order if r[s] and r[s][0]]
    for (s1,m1),(s2,m2) in itertools.combinations(ms,2):
        if m1>m2: v[(s1,s2)]=v.get((s1,s2),0)+1
print('violations',v)
