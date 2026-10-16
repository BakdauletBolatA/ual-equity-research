import json,sys
f=json.load(open(sys.argv[1]))['facts']
pat=[p.lower() for p in sys.argv[2:]]
for ns in f:
  for tag,v in f[ns].items():
    if any(p in tag.lower() for p in pat):
      for u,arr in v['units'].items():
        ann={}; 
        for a in arr:
          if a.get('form','').startswith('10-K') and a.get('fp')=='FY':
            if 'start' in a:
              from datetime import date
              s=date.fromisoformat(a['start']);e=date.fromisoformat(a['end'])
              if (e-s).days<350: continue
            ann[a['end']]=a['val']
        q={}
        for a in arr:
          if a['end']>='2025-06-01' and a.get('form','').startswith('10-Q'):
            q[(a.get('start',''),a['end'])]=a['val']
        ks=sorted(ann)
        if ks and ks[-1]>='2024':
          print(tag,u,{k[:4]:ann[k] for k in ks if k>='2018'})
          if q: print('   Q:',{f'{k[0]}>{k[1]}':val for k,val in sorted(q.items())})
