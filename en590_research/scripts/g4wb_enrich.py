import re, html, json, sys, os, hashlib, subprocess
from concurrent.futures import ThreadPoolExecutor
OUT=sys.argv[1]
def get(url):
    f=OUT+'/cache/'+hashlib.md5(url.encode()).hexdigest()+'.html'
    if not os.path.exists(f) or os.path.getsize(f)<5000:
        subprocess.run(['curl','-sSL','-A','Mozilla/5.0 (X11; Linux x86_64) Chrome/124','-m','40','-o',f,url])
    return open(f,encoding='utf8',errors='ignore').read() if os.path.exists(f) else ''
def lines(s):
    s=re.sub(r'<script.*?</script>|<style.*?</style>','',s,flags=re.S|re.I); s=re.sub(r'<[^>]+>','\n',s); s=html.unescape(s)
    return [re.sub(r'\s+',' ',l).strip() for l in s.splitlines() if l.strip()]
rows=json.load(open(OUT+'/g4wb_list.json'))
rows=[r for r in rows if r['date'][-2:] in ('25','26')]
def kv(L,key):
    for i,l in enumerate(L):
        if l==key and i+1<len(L): return L[i+1]
    return ''
def enrich(r):
    L=lines(get(r['url']))
    try: ks=L.index('Key Specifications')
    except ValueError: ks=0
    K=L[ks:ks+60]
    for k in ['Price','MOQ','Capacity','Origin','Incoterms','Standard','Payment Terms','Delivery Time']:
        r['p_'+k]=kv(K,k)
    try:
        d=L.index('Product Description'); 
        end=next((i for i in range(d,len(L)) if L[i] in ('Recent User Reviews','Similar Products','Product Specifications','Frequently Asked Questions')),d+80)
        r['desc']=' | '.join(L[d+1:end])[:3000]
    except ValueError: r['desc']=''
    ms=next((l for l in L if l.startswith('Member since')),''); r['member_since']=ms.replace('Member since','').strip()
    M=lines(get(r['member_url'])); T='\n'.join(M)
    def f(p):
        m=re.search(p,T); return m.group(1).strip().rstrip('.') if m else ''
    r['established']=f(r'established in (\d{4})')
    r['phone']=f(r'listed phone number is ([^\n]+)')
    r['contact']=f(r'listed contact person for .*? is ([^\n]+?)\.?\n')
    r['designation']=f(r' is listed as ([^\n]+?) at ')
    r['address']=f(r'is listed at ([^\n]+)')
    r['types']=f(r'is listed as an? ([^\n]+?) on go4WorldBusiness')
    r['supplies']=f(r'They supply ([^\n]+)')
    w=re.findall(r'https?://(?:www\.)?[a-z0-9\-\.]+\.[a-z]{2,}',' '.join(M)); 
    r['website']=next((x for x in w if not re.search(r'go4worldbusiness|cloudfront|google|facebook|twitter|linkedin|youtube|instagram|w3\.org|schema\.org|gstatic|jquery|bootstrap|fontawesome|wyzen',x)),'')
    r['verified_flag']='Verified' in ' '.join(M[:400])
    return r
with ThreadPoolExecutor(4) as ex: out=list(ex.map(enrich,rows))
json.dump(out,open(OUT+'/g4wb_enriched.json','w'),indent=1)
print(len(out))
