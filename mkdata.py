import openpyxl, json, re
SRC='/sessions/cool-trusting-shannon/mnt/uploads/Fremont Leader Session Reflections (Responses).xlsx'
ws=openpyxl.load_workbook(SRC).active
rows=list(ws.iter_rows(values_only=True)); hdr=rows[0]; data=rows[1:]
SCALE=['Never','Sometimes','Most of the time','Always']; SCORE={s:i+1 for i,s in enumerate(SCALE)}
inds=[]
for c in range(3,14):
    m=re.match(r'Baseline Assessment: (\w+) \[(.*)\]\s*$', hdr[c].strip())
    inds.append({'domain':m.group(1),'text':m.group(2).strip().replace('–','—')})
LM={'District':'Central Office','Elementary School':'Elementary School','Middle School':'Middle School','High School':'High School'}
RED=[('Fremont','[the district]'),('FMS','[our school]'),('JCAC','[our building]'),('Eileen','[our partner]')]
def red(t):
    t=(t or '').strip()
    for a,b in RED: t=t.replace(a,b)
    return re.sub(r'\s+',' ',t).strip()
resp=[]
for r in data:
    resp.append({'id':len(resp)+1,'level':LM.get((r[2] or '').strip(),'Not specified'),
        'r':[(r[c] or '').strip() if (r[c] or '').strip() in SCALE else None for c in range(3,14)],
        'well':red(r[14]),'next':red(r[15]),'support':red(r[16])})
def stats(sub):
    out=[]
    for i in range(11):
        v=[x['r'][i] for x in sub if x['r'][i]]
        out.append({'counts':[v.count(s) for s in SCALE],'n':len(v),
                    'avg':round(sum(SCORE[y] for y in v)/len(v),2) if v else None})
    return out
levels=['Central Office','Elementary School','Middle School','High School','Not specified']
payload={'scale':SCALE,'inds':inds,'levels':levels,'resp':resp,
  'stats':{'All':stats(resp),**{L:stats([x for x in resp if x['level']==L]) for L in levels}},
  'counts':{'All':len(resp),**{L:sum(1 for x in resp if x['level']==L) for L in levels}}}
json.dump(payload,open('payload.json','w'))
print(json.dumps(payload['counts']))
print('missing cells:',sum(1 for x in resp for y in x['r'] if y is None))
