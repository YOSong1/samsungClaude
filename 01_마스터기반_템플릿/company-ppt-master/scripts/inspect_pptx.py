"""Read-only OOXML inventory. Usage: python inspect_pptx.py SOURCE.pptx OUTPUT.json"""
import collections, hashlib, json, posixpath, sys, zipfile
from pathlib import Path
from xml.etree import ElementTree as E

NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def inspect(source):
 source=Path(source)
 with zipfile.ZipFile(source) as z:
  def root(p): return E.fromstring(z.read(p))
  def rels(p): return [dict(n.attrib) for n in root(p)] if p in z.namelist() else []
  def obj(n):
   nv=n.find('.//p:cNvPr',NS); ph=n.find('.//p:ph',NS)
   return {'id':nv.get('id') if nv is not None else None,'name':nv.get('name') if nv is not None else None,'type':n.tag.rsplit('}',1)[-1],'placeholder':ph.attrib if ph is not None else None,'text':' '.join(a.text or '' for a in n.findall('.//a:t',NS))}
  p=root('ppt/presentation.xml'); r={a['Id']:a for a in rels('ppt/_rels/presentation.xml.rels')}
  layouts=[];masters=[];slides=[]
  for n in z.namelist():
   if n.startswith(('ppt/slideLayouts/','ppt/slideMasters/')) and n.endswith('.xml'):
    rt=root(n); cs=rt.find('p:cSld',NS)
    d={'path':n,'name':cs.get('name') if cs is not None else None,'placeholders':[obj(o) for o in rt.findall('.//p:sp',NS) if o.find('.//p:ph',NS) is not None]}
    (layouts if '/slideLayouts/' in n else masters).append(d)
  for i,n in enumerate(p.find('p:sldIdLst',NS),1):
   target=r[n.get('{'+NS['r']+'}id')]['Target']; part=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('ppt',target))
   rt=root(part); counts=collections.Counter(a.tag.rsplit('}',1)[-1] for a in rt.iter())
   sr=rels(posixpath.dirname(part)+'/_rels/'+posixpath.basename(part)+'.rels')
   slides.append({'number':i,'part':part,'layout':[a['Target'] for a in sr if a['Type'].endswith('/slideLayout')],'counts':{k:counts[k] for k in ['sp','grpSp','pic','tbl','chart','ph','cxnSp']},'objects':[obj(o) for o in rt.iter() if o.tag.rsplit('}',1)[-1] in ['sp','pic','graphicFrame','grpSp']],'relationships':sr})
  sz=p.find('p:sldSz',NS)
  return {'filename':source.name,'sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'slide_size_emu':sz.attrib,'slide_count':len(slides),'master_count':len(masters),'layout_count':len(layouts),'masters':masters,'layouts':layouts,'slides':slides,'limits':'Read-only structural inspection; does not validate visual layout, rendering, media playback or semantic correctness.'}

if __name__=='__main__':
 if len(sys.argv)!=3: raise SystemExit('Usage: python inspect_pptx.py SOURCE.pptx OUTPUT.json')
 src=Path(sys.argv[1]).resolve(); dst=Path(sys.argv[2]).resolve()
 if src==dst: raise SystemExit('Output must differ from source')
 data=inspect(src);dst.parent.mkdir(parents=True,exist_ok=True);dst.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')
 print(json.dumps({k:data[k] for k in ['slide_count','master_count','layout_count']},ensure_ascii=True))
