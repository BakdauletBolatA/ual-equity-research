import sys,re,html
s=open(sys.argv[1],encoding='utf-8',errors='ignore').read()
s=re.sub(r'(?is)<(script|style).*?</\1>','',s)
s=re.sub(r'(?i)</(tr|p|div|br|h\d|li)>','\n',s); s=re.sub(r'(?i)<br[^>]*>','\n',s)
s=re.sub(r'(?i)</t[dh]>',' | ',s)
s=re.sub(r'<[^>]+>','',s); s=html.unescape(s).replace('\xa0',' ')
s=re.sub(r'[ \t]+',' ',s); s=re.sub(r'(\| )+\|','|',s); s=re.sub(r'\n\s*\n+','\n',s)
print(s)
