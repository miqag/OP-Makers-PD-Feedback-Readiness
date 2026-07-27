import openpyxl, json, re

SRC='/sessions/cool-trusting-shannon/mnt/uploads/Fremont Leader Session Reflections (Responses).xlsx'
wb=openpyxl.load_workbook(SRC); ws=wb.active
rows=list(ws.iter_rows(values_only=True))
hdr=rows[0]; data=rows[1:]

SCALE=['Never','Sometimes','Most of the time','Always']
SCORE={'Never':1,'Sometimes':2,'Most of the time':3,'Always':4}

# indicator columns 3..13
inds=[]
for c in range(3,14):
    h=hdr[c]
    m=re.match(r'Baseline Assessment: (\w+) \[(.*)\]\s*$', h.strip())
    inds.append({'col':c,'domain':m.group(1),'text':m.group(2).strip()})

LEVELMAP={'District':'Central Office','Elementary School':'Elementary School',
          'Middle School':'Middle School','High School':'High School'}

resp=[]
for r in data:
    lvl_raw=(r[2] or '').strip()
    lvl=LEVELMAP.get(lvl_raw,'Not specified')
    ratings={}
    for i,ind in enumerate(inds):
        v=(r[ind['col']] or '').strip()
        ratings[i]= v if v in SCALE else None
    resp.append({'level':lvl,'ratings':ratings,
                 'q_well':(r[14] or '').strip(),
                 'q_next':(r[15] or '').strip(),
                 'q_support':(r[16] or '').strip()})

print('N =',len(resp))
from collections import Counter
print(Counter(x['level'] for x in resp))

def stats(subset):
    out=[]
    for i,ind in enumerate(inds):
        vals=[x['ratings'][i] for x in subset if x['ratings'][i]]
        c={s:vals.count(s) for s in SCALE}
        n=len(vals)
        avg=round(sum(SCORE[v] for v in vals)/n,2) if n else None
        out.append({'domain':ind['domain'],'text':ind['text'],'counts':c,'n':n,'avg':avg})
    return out

allstats=stats(resp)
for s in allstats:
    print(s['domain'],'|',s['avg'],s['n'],s['counts'],'|',s['text'][:60])

# domain averages
for d in ['Belonging','Consistency','Coherence']:
    sub=[s for s in allstats if s['domain']==d]
    tot=sum(sum(SCORE[k]*v for k,v in s['counts'].items()) for s in sub)
    n=sum(s['n'] for s in sub)
    print(d, round(tot/n,2), 'responses:',n)

levels=['Central Office','Elementary School','Middle School','High School','Not specified']
bylevel={L:stats([x for x in resp if x['level']==L]) for L in levels}

json.dump({'inds':inds,'resp':resp,'all':allstats,'bylevel':bylevel,
           'levels':levels,'scale':SCALE},open('data_raw.json','w'),indent=1)
