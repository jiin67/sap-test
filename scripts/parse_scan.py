import re,sys
src,dst=sys.argv[1],sys.argv[2]
L=open(src,encoding='utf-16').read().splitlines()
rows={}
for s in L:
    m=re.match(r'\s+a(\d+):\s*(.*)',s)
    if m: rows[int(m.group(1))]=m.group(2)
cur=None; out=[]
for k in sorted(rows):
    body=rows[k]
    mi=re.match(r'\[1\]\d+\s+\[4\](\S+)',body)
    if mi: cur=mi.group(1); continue
    m3=re.match(r'\[4\](\d{6})\s*(?:\[11\](.*))?$',body)
    if m3: out.append((cur,int(m3.group(1)),(m3.group(2) or '').rstrip()))
print('absrows',min(rows),max(rows),len(rows))
inc={}
for c,n,t in out: inc.setdefault(c,[]).append(n)
for c,v in inc.items():
    gaps=sum(1 for a,b in zip(v,v[1:]) if b!=a+1)
    print(c,len(v),min(v),max(v),'gaps',gaps)
with open(dst,'w',encoding='utf-8') as f:
    for c,n,t in out:
        if t: f.write(f'{c}\t{n:5d}\t{t}\n')
