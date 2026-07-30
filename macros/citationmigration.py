"""Safe citation migration planning for Word/Zotero/EndNote."""
from __future__ import annotations
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
import json,re,unicodedata,zipfile,datetime
import xml.etree.ElementTree as ET

NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
DOI=re.compile(r'(?:https?://(?:dx\.)?doi\.org/|doi:\s*)?(10\.\d{4,9}/[-._;()/:a-z0-9]+)',re.I)
YEAR=re.compile(r'\b(?:18|19|20)\d{2}\b')
URL=re.compile(r'(?<!@)(?:https?://|www\.)[^\s<>\]\[(){}"\']+',re.I)
PINK=5

@dataclass
class Ref:
    title:str=''; author:str=''; year:str=''; doi:str=''; url:str=''; accessed:str=''; itemtype:str=''; raw:str=''
@dataclass
class PlanItem:
    source:Ref; decision:str; replacement:str; reason:str; score:float=0.0
@dataclass
class Plan:
    citationtype:str; originaltext:str; fieldstart:int=0; fieldend:int=0; resultstart:int=0; resultend:int=0; items:list[PlanItem]=field(default_factory=list)
    @property
    def replacement(self):
        known=[x.replacement for x in self.items if x.decision=='EndNote']
        unknown=[x.replacement for x in self.items if x.decision!='EndNote']
        parts=[]
        if known: parts.append('{'+ '; '.join(known) + '}')
        if unknown: parts.extend(unknown)
        return ' '.join(parts)

def norm(x):
    x=unicodedata.normalize('NFKD',x or '')
    x=''.join(c for c in x if not unicodedata.combining(c)).casefold()
    return re.sub(r'\s+',' ',re.sub(r'[^a-z0-9\s]',' ',x)).strip()
def doi(x):
    m=DOI.search(x or ''); return m.group(1).rstrip('.,;)').lower() if m else ''
def firstauthor(x):
    x=(x or '').strip()
    if not x:return ''
    if ',' in x:return norm(x.split(',',1)[0])
    return norm(x.split()[0])
def parse_library(path):
    chunks=[re.sub(r'\s+',' ',z).strip() for z in re.split(r'\n\s*\n+',Path(path).read_text(encoding='utf-8-sig')) if z.strip()]
    out=[]
    for raw in chunks:
        y=YEAR.findall(raw); year=y[-1] if y else ''
        before=raw[:raw.find(year)] if year else raw
        author=firstauthor(before)
        title=raw[raw.find(year)+4:].lstrip(' .,:-') if year else raw
        title=re.split(r'\.\s+(?=[A-Z])',title,maxsplit=1)[0].strip(' .')
        out.append(Ref(title,author,year,doi(raw),raw=raw))
    return out
def score(source,lib):
    if source.doi and source.doi==lib.doi:return 1.0,'Exact DOI match'
    st,lt=norm(source.title),norm(lib.title)
    exact=bool(st and st==lt and source.author==lib.author and source.year==lib.year)
    if exact:return 1.0,'Exact normalized title, author, and year'
    ts=0.0
    if st and lt:
        ts=.6*SequenceMatcher(None,st,lt).ratio()+.4*(len(set(st.split())&set(lt.split()))/max(1,len(set(st.split())|set(lt.split()))))
    return .7*ts+.15*float(source.author==lib.author and bool(source.author))+.15*float(source.year==lib.year and bool(source.year)),f'title {ts:.0%}'
def best(source,library):
    candidates=[]
    for item in library:
        if source.doi and item.doi and source.doi!=item.doi:continue
        if source.author and item.author and source.author!=item.author and not source.doi:continue
        s,r=score(source,item);candidates.append((s,r,item))
    candidates.sort(reverse=True,key=lambda x:x[0])
    return candidates[0] if candidates else (0.0,'No library candidate',None)
def zotero_items(code):
    if 'CSL_CITATION' not in code:return []
    payload=json.loads(code.split('CSL_CITATION',1)[1].strip());out=[]
    for item in payload.get('citationItems',[]):
        d=item.get('itemData',{}) or {}; authors=d.get('author',[]) or [];a=norm(authors[0].get('family','')) if authors else ''
        dates=d.get('issued',{}).get('date-parts',[]);y=str(dates[0][0]) if dates and dates[0] else ''
        out.append(Ref(d.get('title',''),a,y,doi(d.get('DOI','')),d.get('URL',''),str(d.get('accessed',{}).get('raw','') or ''),d.get('type',''),d.get('title','')))
    return out
def temp(ref):
    title=re.sub(r'[{},]',' ',ref.title);title=re.sub(r'\s+',' ',title).strip()
    return f'{ref.author}, {ref.year}, {title}'
def unresolved(ref):
    parts=[ref.author or 'Unknown author',ref.year or 'Unknown year',ref.title or 'Unknown title']
    return '[REFERENCE REQUIRED: ' + ', '.join(parts) + ']'
def inline(ref):
    url=re.sub(r'^(?:https?://)?(?:www\.)?','',ref.url or '[website address]',flags=re.I)
    title=ref.title or '[Webpage title]';access=ref.accessed or '[DD Month YYYY]'
    return f'({url}, {title}, Accessed {access})'
def plan_sources(sources,library):
    plan=[]
    for source in sources:
        if source.itemtype.casefold() in {'webpage','blogpost','forumpost','post','web'}:
            text=inline(source);missing='[' in text
            plan.append(PlanItem(source,'Inline' if not missing else 'Request',text,'Website source; complete placeholders before final use.'))
            continue
        s,reason,match=best(source,library)
        if match and s>=.85:plan.append(PlanItem(source,'EndNote',temp(match),reason,s))
        else:plan.append(PlanItem(source,'Request',unresolved(source),reason,s))
    return plan
