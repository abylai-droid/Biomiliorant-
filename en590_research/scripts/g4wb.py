import re, html, json, sys, time, subprocess, os, hashlib
OUT=sys.argv[1]; os.makedirs(OUT+'/cache',exist_ok=True)
BASE='https://www.go4worldbusiness.com'
def get(url):
    f=OUT+'/cache/'+hashlib.md5(url.encode()).hexdigest()+'.html'
    if not os.path.exists(f):
        subprocess.run(['curl','-sSL','-A','Mozilla/5.0 (X11; Linux x86_64) Chrome/124','-m','40','-o',f,url])
        time.sleep(0.7)
    return open(f,encoding='utf8',errors='ignore').read() if os.path.exists(f) else ''
def txt(s):
    s=re.sub(r'<script.*?</script>|<style.*?</style>','',s,flags=re.S)
    s=re.sub(r'<br\s*/?>|</p>|</li>|</div>','\n',s); s=re.sub(r'<[^>]+>',' ',s)
    s=html.unescape(s); return '\n'.join(re.sub(r'[ \t\xa0]+',' ',l).strip() for l in s.splitlines() if l.strip())
queries=['diesel-en590-10ppm','en590-10ppm','diesel-en590-10-ppm','en590-10-ppm','ulsd-10ppm','ultra-low-sulphur-diesel-10ppm','en590','diesel-en590','gasoil-10ppm','ulsd-en590']
rows={}
for q in queries:
    url=f'{BASE}/find?searchText={q}&BuyersOrSuppliers=suppliers'
    for page in range(3):
        s=get(url)
        blocks=s.split('entity-rows-container entity-rows-container-auto-height')[1:]
        for b in blocks:
            m=re.search(r'href="?(/pref_product/view/\d+/[^" >]+)',b)
            if not m: continue
            pu=BASE+m.group(1)
            if pu in rows: continue
            g=lambda p:(re.search(p,b,re.S).group(1).strip() if re.search(p,b,re.S) else '')
            rows[pu]=dict(url=pu,query=q,
              date=g(r'<small>\s*([A-Za-z]{3}-\d\d-\d\d)\s*</small>'),
              member=txt(g(r'star-text\w*">(.*?)</span>')),
              verified='Verified' in b.split('item-row')[0],
              title=txt(g(r'title="([^"]*)"')),
              price=txt(g(r'product-price ellipsis">(.*?)</div>')).replace('Ask for price','').strip(),
              moq=txt(g(r'product-moq">(.*?)</div>')).replace('MOQ:','').strip(),
              company=txt(g(r'product-member-company-name"[^>]*>(.*?)</span>')),
              member_url=BASE+g(r'href="(/member/view/[^"]+)"'),
              location=txt(g(r'Supplier From(.*?)</span>')).replace('\n',' '))
        n=re.search(r'href="(/find\?searchText=[^"]*searchAfter=[^"]+)"',s)
        if not n: break
        url=BASE+html.unescape(n.group(1))
    print(q,len(rows),file=sys.stderr)
json.dump(list(rows.values()),open(OUT+'/g4wb_list.json','w'),indent=1)
