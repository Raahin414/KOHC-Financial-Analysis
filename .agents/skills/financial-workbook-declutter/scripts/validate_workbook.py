"""Read-only XLSX baseline checks. Standard library only; never rewrites a workbook.

snapshot BOOK OUT.json
compare BEFORE.json AFTER.json [--allow-text reviewed_text_changes.json] [--out report.json]
Allowlist: {"Sheet!A1": "exact revised text", "Sheet!B2": null}
Number formats and layout may change; numerical/formula integrity may not silently change.
"""
import argparse,hashlib,json,math,pathlib,posixpath,zipfile,xml.etree.ElementTree as E
N={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def snapshot(path):
 path=pathlib.Path(path); result={'file':str(path.resolve()),'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),'sheets':[],'cells':{},'charts':{},'drawings':{},'hyperlinks':{},'annotations':{},'names':[]}
 with zipfile.ZipFile(path)as z:
  strings=[]
  if 'xl/sharedStrings.xml'in z.namelist():strings=[''.join(e.itertext())for e in E.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si',N)]
  rel={e.attrib['Id']:e.attrib['Target']for e in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
  wb=E.fromstring(z.read('xl/workbook.xml'))
  for e in wb.findall('s:definedNames/s:definedName',N):result['names'].append([e.attrib,e.text])
  for s in wb.findall('s:sheets/s:sheet',N):
   name=s.attrib['name'];target=rel[s.attrib['{'+N['r']+'}id']];part=target.lstrip('/')if target.startswith('/')else'xl/'+target
   root=E.fromstring(z.read(part)); hidden=[dict(e.attrib)for e in root.findall('s:cols/s:col',N)if e.attrib.get('hidden')=='1']
   hidden=[{k:v for k,v in h.items()if k in ['min','max','hidden']}for h in hidden]
   hidden_rows=[e.attrib['r']for e in root.findall('s:sheetData/s:row',N)if e.attrib.get('hidden')=='1']
   result['sheets'].append({'name':name,'state':s.attrib.get('state','visible'),'hidden_columns':hidden,'hidden_rows':hidden_rows})
   relpart=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels');sheet_rels={}
   if relpart in z.namelist():sheet_rels={e.attrib['Id']:e.attrib.get('Target')for e in E.fromstring(z.read(relpart))}
   for e in root.findall('s:hyperlinks/s:hyperlink',N):
    h={k:v for k,v in e.attrib.items()if k!='{'+N['r']+'}id'};rid=e.attrib.get('{'+N['r']+'}id')
    if rid:h['target']=sheet_rels.get(rid)
    result['hyperlinks'][name+'!'+h['ref']]=h
   for c in root.findall('s:sheetData/s:row/s:c',N):
    key=name+'!'+c.attrib['r'];f=c.find('s:f',N);v=c.find('s:v',N);typ=c.attrib.get('t','n');val=v.text if v is not None else None
    if typ=='s':val=strings[int(val)]if val is not None else None
    elif typ=='inlineStr':val=''.join(c.find('s:is',N).itertext())
    elif typ=='n'and val is not None:val=float(val)
    if f is not None:result['cells'][key]={'kind':'formula','formula':f.text,'attributes':dict(f.attrib),'cache':val}
    elif val is not None:result['cells'][key]={'kind':'number'if typ in ['n','b']else'text','value':val}
  for part in z.namelist():
   if part.startswith('xl/charts/')and part.endswith('.xml'):result['charts'][part]=hashlib.sha256(z.read(part)).hexdigest()
   if part.startswith('xl/drawings/')and not part.endswith('/'):result['drawings'][part]=hashlib.sha256(z.read(part)).hexdigest()
   if ('comments'in part.lower()or part.startswith('xl/externalLinks/'))and not part.endswith('/'):result['annotations'][part]=hashlib.sha256(z.read(part)).hexdigest()
 return result
def equal(a,b):
 if isinstance(a,(int,float))and isinstance(b,(int,float)):return math.isclose(a,b,rel_tol=1e-11,abs_tol=1e-10)
 return a==b
def compare(a,b,allowed=None):
 allowed=allowed or {};issues=[];counts={'formulas':0,'numeric_inputs':0,'cached_results':0,'reviewed_text_changes':0}
 for k in ['sheets','charts','drawings','hyperlinks','annotations','names']:
  if a[k]!=b[k]:issues.append({'area':k,'reason':'protected structure or evidence changed'})
 for key in sorted(set(a['cells'])|set(b['cells'])):
  x=a['cells'].get(key);y=b['cells'].get(key)
  if key in allowed:
   if (x and x['kind']in ['formula','number'])or(y and y['kind']in ['formula','number']):issues.append({'cell':key,'reason':'text allowlist cannot authorize financial changes'});continue
   if (y or{}).get('value')!=allowed[key]:issues.append({'cell':key,'reason':'revised text differs from reviewed allowlist'})
   counts['reviewed_text_changes']+=1;continue
  if x is None or y is None or x['kind']!=y['kind']:issues.append({'cell':key,'reason':'cell added, removed or type changed'});continue
  if x['kind']=='formula':
   counts['formulas']+=1
   if x['formula']!=y['formula']or x['attributes']!=y['attributes']:issues.append({'cell':key,'reason':'formula changed'})
   if not equal(x['cache'],y['cache']):issues.append({'cell':key,'reason':'cached formula result changed','before':x['cache'],'after':y['cache']})
   counts['cached_results']+=1
  else:
   if x['kind']=='number':counts['numeric_inputs']+=1
   if not equal(x['value'],y['value']):issues.append({'cell':key,'reason':'stored value changed','before':x['value'],'after':y['value']})
 return {'passed':not issues,'counts':counts,'charts':len(b['charts']),'issues':issues}
def main():
 p=argparse.ArgumentParser(description=__doc__);sub=p.add_subparsers(dest='cmd',required=True)
 s=sub.add_parser('snapshot');s.add_argument('book');s.add_argument('out')
 c=sub.add_parser('compare');c.add_argument('before');c.add_argument('after');c.add_argument('--allow-text');c.add_argument('--out')
 args=p.parse_args()
 if args.cmd=='snapshot':pathlib.Path(args.out).write_text(json.dumps(snapshot(args.book),indent=2),encoding='utf-8');print('Read-only baseline saved.');return
 r=compare(json.loads(pathlib.Path(args.before).read_text(encoding='utf-8')),json.loads(pathlib.Path(args.after).read_text(encoding='utf-8')),json.loads(pathlib.Path(args.allow_text).read_text(encoding='utf-8'))if args.allow_text else None)
 if args.out:pathlib.Path(args.out).write_text(json.dumps(r,indent=2),encoding='utf-8')
 print(json.dumps(r,indent=2));raise SystemExit(0 if r['passed']else 1)
if __name__=='__main__':main()
