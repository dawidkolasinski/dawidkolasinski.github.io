#!/usr/bin/env python3
"""Extract full legal blocks and round-trip real DOCX hyperlinks for translation."""
import argparse
import copy
from hashlib import sha256
from html import escape
import json
from pathlib import Path
import re
import subprocess
import unicodedata
from urllib.parse import urljoin
from lxml import etree, html
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Inches, Pt, RGBColor

ROOT=Path(__file__).resolve().parents[2]
DEFAULT_INDEX=ROOT/'legal/source-index.json'
ORIGIN='https://dawidkolasinski.github.io/'
SOURCES=[('sunchucks_privacy','sunchucks','privacy','privacy/index.html'),('sunchucks_terms','sunchucks','terms','terms/index.html'),('moonchucks_privacy','moonchucks','privacy','moonchucks/privacy/index.html'),('moonchucks_terms','moonchucks','terms','moonchucks/terms/index.html')]
UI={'privacy':'Privacy','terms':'Terms','support':'Support','legalNavigation':'Legal pages','language':'Language','onThisPage':'On this page','documentSections':'Document sections','skipToContent':'Skip to content','mainNavigation':'Main navigation','footerNavigation':'Footer navigation','backToTop':'Back to top'}
# Google Documents transliterated only this transport prefix in the verified Hindi export.
# Unit numbers stay ASCII, and both the source pin and full unit coverage remain mandatory.
MARKER=re.compile(r'^\[\[(LGL|एलजीएल)([0-9]{4})\]\]$')
SOURCE_MARKER=re.compile(r'^\[\[LGLSOURCE([0-9a-f]{64})\]\]$')
LOCALES='en pl de es fr it ca cs da el fi hr hu nb bg hi ru sk sl sv tr uk bn gu id ja kk kn ko lt ml mr ms nl or pa ro ta te th vi en-AU en-GB en-IN es-419 es-US fr-CA pt-BR pt-PT zh-HK zh-Hans zh-Hant'.split()
ALIASES={'en-AU':'en','en-GB':'en','en-IN':'en','es-419':'es','es-US':'es','fr-CA':'fr','pt-PT':'pt-BR','zh-HK':'zh-Hant'}
ALLOWED_INLINE={'a','strong','em','br','li'}
RETIRED_ACCESS=re.compile(r'reward(?:ed)?|30\s+minut|reklam[^<.]{0,50}nagrod|czasow[^<.]{0,40}Premium',re.IGNORECASE)


def write_json(path,data):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True);path.write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')


def source_hash(source):
    payload={key:value for key,value in source.items() if key not in ('sourceHash','provenance','translationUnits')}
    return sha256(json.dumps(payload,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf-8')).hexdigest()


def check_retired_access(text,label):
    if RETIRED_ACCESS.search(text):raise ValueError('Retired rewarded Premium clause: '+label)


def load_source(path):
    source=json.loads(Path(path).read_text())
    if source_hash(source)!=source['sourceHash']:raise ValueError('Source manifest hash mismatch')
    if source['translationUnits']!=units_for(source):raise ValueError('Translation unit binding differs from source blocks')
    expected={path for _,app,kind,path in SOURCES}
    expected|={('pl/' if app=='sunchucks' else 'moonchucks/pl/')+kind+'/index.html' for _,app,kind,_ in SOURCES}
    if set(source['sourceFiles'])!=expected:raise ValueError('Canonical EN/PL source file coverage mismatch')
    for relative,digest in source['sourceFiles'].items():
        if sha256((ROOT/relative).read_bytes()).hexdigest()!=digest:raise ValueError('Source HTML hash mismatch: '+relative)
        check_retired_access((ROOT/relative).read_text(),relative)
    if source['documents']!=[extract(*args) for args in SOURCES]:raise ValueError('Source document blocks differ from canonical HTML')
    expected_hosts={doc['id']:hosting_variants(doc) for doc in source['documents'] if doc['kind']=='privacy'}
    if source['privacyHostingVariants']!=expected_hosts:raise ValueError('Hosting alternatives differ from the source generator')
    return source


def inner(node):
    return (escape(node.text or '')+''.join(etree.tostring(child,encoding='unicode',method='html',with_tail=True) for child in node)).strip()


def normalize_links(node,path):
    for a in node.iter('a'):
        a.set('href',urljoin(ORIGIN+path,a.get('href','')))


def extract(doc_id,app,kind,path):
    tree=html.fromstring((ROOT/path).read_bytes());main=tree.xpath('//main')[0]
    normalize_links(main,path)
    blocks=[]
    for node in main:
        if node.tag in ('nav','h1') or (node.tag=='p' and set(node.get('class','').split()) & {'meta','eyebrow'}):continue
        if node.tag not in ('p','h2','h3','ul','ol'):raise ValueError(f'Unexpected legal element {node.tag} in {path}')
        blocks.append({'id':f'block-{len(blocks)+1:03d}','tag':node.tag,'html':inner(node)})
    title=main.xpath('./h1')[0].text_content().strip();updated=main.xpath('./p[contains(concat(" ",normalize-space(@class)," ")," meta ")]')[0].text_content().strip()
    footer=tree.xpath('//footer/p')
    return {'id':doc_id,'app':app,'kind':kind,'sourceLanguage':'en','path':path,'title':title,'updatedLabel':updated,'footerHTML':inner(footer[0]) if footer else '', 'blocks':blocks}


def hosting_variants(doc):
    final=doc['blocks'][-1]
    assert final['tag']=='p' and 'GitHub' in final['html']
    name='Sunchucks' if doc['app']=='sunchucks' else 'Moonchucks'
    host={
    'github':final['html'],
    'cloudflare':'This website uses Cloudflare Pages. Cloudflare processes technical request data, including IP addresses, to deliver and protect the site. See <a href="https://www.cloudflare.com/privacypolicy/">Cloudflare privacy information</a>.',
    'firebase':'This website uses Firebase Hosting. Google processes IP addresses and request data to deliver the site and protect it from abuse. See <a href="https://firebase.google.com/support/privacy">Firebase privacy information</a>.'}
    if doc['app']=='sunchucks':
        for key in ('cloudflare','firebase'):host[key]='The current policy is posted here. '+host[key]
    else:
        for key in ('cloudflare','firebase'):host[key]+=f" The current {name} policy is published on this page and may be updated when the app's data practices change."
    return host


def units_for(source):
    units=[]
    def add(target,content,tag='p'):
        units.append({'id':f'LGL{len(units)+1:04d}','target':target,'tag':tag,'html':content})
    for doc in source['documents']:
        base={'document':doc['id']}
        for field in ('title','updatedLabel','footerHTML'):
            add({**base,'field':field},escape(doc[field]) if field!='footerHTML' else doc[field], 'h1' if field=='title' else 'p')
        for block in doc['blocks']:
            target={**base,'field':'blocks','block':block['id']}
            if block['tag'] in ('ul','ol'):
                fragment=html.fragment_fromstring('<div>'+block['html']+'</div>')
                for i,item in enumerate(fragment):
                    assert item.tag=='li'
                    add({**target,'item':i},inner(item),'li')
            else:add(target,block['html'],block['tag'])
    for key,text in source['uiLabels'].items():add({'field':'uiLabels','key':key},escape(text))
    for doc_id,variants in source['privacyHostingVariants'].items():
        for host,text in variants.items():add({'field':'privacyHostingVariants','document':doc_id,'host':host},text)
    return units


def build(index_path):
    for _,app,kind,path in SOURCES:
        check_retired_access((ROOT/path).read_text(),path)
        polish=('pl/' if app=='sunchucks' else 'moonchucks/pl/')+kind+'/index.html'
        check_retired_access((ROOT/polish).read_text(),polish)
    docs=[extract(*args) for args in SOURCES]
    source={'schemaVersion':1,'sourceLanguage':'en','documents':docs,'uiLabels':UI,'privacyHostingVariants':{doc['id']:hosting_variants(doc) for doc in docs if doc['kind']=='privacy'}}
    files=[]
    for _,app,kind,path in SOURCES:
        files.append(path)
        files.append(('pl/' if app=='sunchucks' else 'moonchucks/pl/')+kind+'/index.html')
    source['sourceFiles']={path:sha256((ROOT/path).read_bytes()).hexdigest() for path in files}
    source['sourceHash']=sha256(json.dumps(source,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode()).hexdigest()
    status=subprocess.check_output(['git','status','--porcelain=v1','--',*files],cwd=ROOT,text=True)
    source['provenance']={'kind':'local-working-tree','gitHead':subprocess.check_output(['git','rev-parse','HEAD'],cwd=ROOT,text=True).strip(),'dirtySourceFiles':status.splitlines(),'published':False}
    source['translationUnits']=units_for(source)
    write_json(index_path,source)
    # English is the exact source; regional English variants can reuse this locale.
    output=empty_locale(source,'en')
    output['documents']={doc['id']:{'title':doc['title'],'updatedLabel':doc['updatedLabel'],'footerHTML':doc['footerHTML'],'blocks':{b['id']:b['html'] for b in doc['blocks']}} for doc in docs}
    output['uiLabels']=copy.deepcopy(UI);output['privacyHostingVariants']=copy.deepcopy(source['privacyHostingVariants'])
    output['translation']={'method':'source','reviewStatus':'source','published':False}
    write_json(ROOT/'legal/locales/en.json',output)
    print(f'Source {source["sourceHash"]}: {len(docs)} documents, {sum(len(d["blocks"]) for d in docs)} top-level blocks, {len(source["translationUnits"])} full translation units')
    return source


def add_html(paragraph,markup):
    fragment=html.fragment_fromstring('<div>'+markup+'</div>')
    def text_run(text,parent,bold=False,italic=False):
        if not text:return
        r=OxmlElement('w:r');props=OxmlElement('w:rPr')
        if bold:props.append(OxmlElement('w:b'))
        if italic:props.append(OxmlElement('w:i'))
        if len(props):r.append(props)
        t=OxmlElement('w:t');t.set(qn('xml:space'),'preserve');t.text=text;r.append(t);parent.append(r)
    def append(node,parent,bold=False,italic=False):
        text_run(node.text,parent,bold,italic)
        for child in node:
            if child.tag=='a':
                link=OxmlElement('w:hyperlink');link.set(qn('r:id'),paragraph.part.relate_to(child.get('href'),RT.HYPERLINK,is_external=True));parent.append(link)
                append(child,link,bold,italic)
            elif child.tag=='br':
                r=OxmlElement('w:r');r.append(OxmlElement('w:br'));parent.append(r)
            elif child.tag in ('strong','b','em','i'):append(child,parent,bold or child.tag in ('strong','b'),italic or child.tag in ('em','i'))
            else:raise ValueError(f'Unexpected inline element {child.tag}')
            text_run(child.tail,parent,bold,italic)
    append(fragment,paragraph._p)


def export_docx(source,path):
    doc=Document();section=doc.sections[0];section.top_margin=section.bottom_margin=Inches(.7);section.left_margin=section.right_margin=Inches(.8)
    normal=doc.styles['Normal'];normal.font.name='Arial';normal.font.size=Pt(10);normal.paragraph_format.space_after=Pt(6)
    for name in ('Title','Heading 1','Heading 2','Heading 3'):
        doc.styles[name].font.name='Arial';doc.styles[name].font.color.rgb=RGBColor(0,0,0)
    doc.add_paragraph('App privacy policies and terms of use',style='Title')
    doc.add_paragraph('Complete Sunchucks and Moonchucks documents for translation. Keep each identifier and the linked phrases with their full paragraph. Hosting alternatives are separate complete paragraphs.')
    pin=doc.add_paragraph('[[LGLSOURCE'+source['sourceHash']+']]');pin.runs[0].font.size=Pt(7)
    for unit in source['translationUnits']:
        p=doc.add_paragraph('[['+unit['id']+']]');p.paragraph_format.space_before=Pt(7);p.paragraph_format.space_after=Pt(1);p.paragraph_format.keep_with_next=True;p.runs[0].font.size=Pt(7)
        style={'h1':'Heading 1','h2':'Heading 2','h3':'Heading 3'}.get(unit['tag'],'Normal')
        p=doc.add_paragraph(style=style);add_html(p,unit['html'])
    doc.core_properties.title='App privacy policies and terms of use';doc.core_properties.author='Dawid Kolasiński';doc.core_properties.subject='sourceHash '+source['sourceHash']
    for border in doc.styles.element.xpath('.//w:pBdr'):
        border.getparent().remove(border)
    Path(path).parent.mkdir(parents=True,exist_ok=True);doc.save(path)
    print(f'DOCX {path}: {len(source["translationUnits"])} units with real hyperlinks')


def empty_locale(source,locale):
    return {'schemaVersion':1,'sourceHash':source['sourceHash'],'locale':locale,'documents':{d['id']:{'title':'','updatedLabel':'','footerHTML':'','blocks':{b['id']:'' for b in d['blocks']}} for d in source['documents']},'uiLabels':{},'privacyHostingVariants':{key:{} for key in source['privacyHostingVariants']}}


def paragraph_html(p,part):
    out=[]
    def walk(node,bold=False,italic=False):
        tag=etree.QName(node).localname
        if tag in ('pPr','rPr','bookmarkStart','bookmarkEnd','proofErr'):return
        if tag=='hyperlink':
            relation=node.get(qn('r:id'))
            if not relation:raise ValueError('Hyperlink is missing a relationship')
            href=part.rels[relation].target_ref
            out.append('<a href="'+escape(href,quote=True)+'">')
            for child in node:walk(child)
            out.append('</a>');return
        if tag=='r':
            rpr=node.find(qn('w:rPr'));is_bold=rpr is not None and rpr.find(qn('w:b')) is not None;is_italic=rpr is not None and rpr.find(qn('w:i')) is not None
            if is_bold:out.append('<strong>')
            if is_italic:out.append('<em>')
            for child in node:walk(child)
            if is_italic:out.append('</em>')
            if is_bold:out.append('</strong>')
            return
        if tag=='t':out.append(escape(node.text or ''));return
        if tag in ('br','cr'):out.append('<br>');return
        if tag=='tab':out.append(' ');return
        for child in node:walk(child)
    walk(p)
    return ''.join(out).strip()


def plain(markup):
    return html.fragment_fromstring('<div>'+markup+'</div>').text_content().strip()


def hrefs(markup):
    return sorted(html.fragment_fromstring('<div>'+markup+'</div>').xpath('.//a/@href'))


def assemble_locale(source,locale,unit_html,translation=None,allow_source=False):
    """Assemble complete translated units without inventing document-input provenance."""
    expected={u['id'] for u in source['translationUnits']}
    if set(unit_html)!=expected:raise ValueError(f'Unit coverage mismatch missing={sorted(expected-set(unit_html))} extra={sorted(set(unit_html)-expected)}')
    result=empty_locale(source,locale);list_items={}
    for unit in source['translationUnits']:
        value=unit_html[unit['id']]
        if not plain(value):raise ValueError('Empty unit '+unit['id'])
        if hrefs(value)!=hrefs(unit['html']):raise ValueError('Hyperlinks changed in '+unit['id'])
        if not allow_source and len(plain(unit['html']))>80 and plain(value)==plain(unit['html']):raise ValueError('Untranslated full paragraph '+unit['id'])
        target=unit['target'];field=target['field']
        if field=='uiLabels':result[field][target['key']]=plain(value)
        elif field=='privacyHostingVariants':result[field][target['document']][target['host']]=value
        else:
            document=result['documents'][target['document']]
            if field=='blocks':
                if 'item' in target:list_items.setdefault((target['document'],target['block']),{})[target['item']]=value
                else:document[field][target['block']]=value
            else:document[field]=plain(value) if field in ('title','updatedLabel') else value
    for (doc_id,block_id),items in list_items.items():
        result['documents'][doc_id]['blocks'][block_id]=''.join('<li>'+items[i]+'</li>' for i in range(len(items)))
    result['translation']=copy.deepcopy(translation or {'method':'assistant-authored','reviewStatus':'targeted-editorial-review','independentNativeReview':False,'published':False})
    validate_locale(source,result,locale)
    return result


def import_docx(source,path,locale,output,allow_source=False):
    doc=Document(path);found={};current=None;values=[];pins=[]
    subject=doc.core_properties.subject or ''
    if subject.startswith('sourceHash '):pins.append(subject.removeprefix('sourceHash ').strip())
    def finish():
        if current:
            if current in found:raise ValueError('Duplicate unit '+current)
            found[current]='<br>'.join(values).strip()
    for p in doc.element.body.iter(qn('w:p')):
        value=paragraph_html(p,doc.part);text=plain(value);pin=SOURCE_MARKER.fullmatch(text)
        if pin:pins.append(pin.group(1));continue
        match=MARKER.fullmatch(text)
        if match:finish();current='LGL'+match.group(2);values=[]
        elif current and plain(value):values.append(value)
    finish()
    if not pins or any(pin!=source['sourceHash'] for pin in pins):raise ValueError('DOCX sourceHash is missing or differs from the frozen source')
    metadata={'method':'google-documents' if not allow_source else 'roundtrip-test','reviewStatus':'machine-translated-needs-review' if not allow_source else 'source','published':False,'inputFile':Path(path).name,'inputSHA256':sha256(Path(path).read_bytes()).hexdigest()}
    result=assemble_locale(source,locale,found,metadata,allow_source)
    write_json(output,result);print(f'Imported {locale}: {len(found)} units, exact sourceHash and hyperlink coverage')
    return result


def canonical_pl(source,output=None):
    result=empty_locale(source,'pl')
    for doc_id,app,kind,_ in SOURCES:
        path=('pl/' if app=='sunchucks' else 'moonchucks/pl/')+kind+'/index.html'
        doc=extract(doc_id,app,kind,path)
        expected=next(item for item in source['documents'] if item['id']==doc_id)
        if [b['tag'] for b in doc['blocks']]!=[b['tag'] for b in expected['blocks']]:raise ValueError('Polish source structure differs: '+doc_id)
        result['documents'][doc_id]={'title':doc['title'],'updatedLabel':doc['updatedLabel'],'footerHTML':doc['footerHTML'],'blocks':{b['id']:b['html'] for b in doc['blocks']}}
        if kind=='privacy':
            variants={'github':doc['blocks'][-1]['html']}
            common={
                'cloudflare':'Ta strona korzysta z Cloudflare Pages. Cloudflare przetwarza techniczne dane żądań, w tym adresy IP, na potrzeby dostarczania i ochrony strony. Szczegóły: <a href="https://www.cloudflare.com/privacypolicy/">informacje o prywatności Cloudflare</a>.',
                'firebase':'Ta strona korzysta z Firebase Hosting. Google przetwarza adresy IP i dane żądań na potrzeby dostarczania strony i ochrony przed nadużyciami. Szczegóły: <a href="https://firebase.google.com/support/privacy">informacje o prywatności Firebase</a>.'}
            for host,text in common.items():
                if app=='sunchucks':text='Aktualna polityka jest opublikowana na tej stronie. '+text
                else:text+=' Aktualna polityka Moonchucks jest publikowana na tej stronie i może być aktualizowana wraz ze zmianą zasad przetwarzania danych przez aplikację.'
                variants[host]=text
            result['privacyHostingVariants'][doc_id]=variants
    result['uiLabels']={'privacy':'Prywatność','terms':'Warunki','support':'Pomoc','legalNavigation':'Strony prawne','language':'Język','onThisPage':'Na tej stronie','documentSections':'Sekcje dokumentu','skipToContent':'Przejdź do treści','mainNavigation':'Nawigacja główna','footerNavigation':'Nawigacja w stopce','backToTop':'Powrót na górę'}
    result['translation']={'method':'existing-canonical-polish','reviewStatus':'source','published':False,'hostingAlternatives':'local-translation'}
    if output is not None:write_json(output,result)
    return result


def number_tokens(markup):
    text=plain(markup)
    return sorted(''.join(str(unicodedata.decimal(char)) for char in token) for token in re.findall(r'\d+',text))


def normalized_hrefs(markup):
    return sorted(href.replace('/pl/','/') for href in hrefs(markup))


def validate_html(value,source_value,label,check_numbers=True):
    if not isinstance(value,str) or not plain(value):raise ValueError('Empty or invalid text: '+label)
    check_retired_access(plain(value),label)
    fragment=html.fragment_fromstring('<div>'+value+'</div>')
    for node in fragment.iterdescendants():
        if node.tag not in ALLOWED_INLINE:raise ValueError('Unsupported HTML in '+label+': '+node.tag)
        expected={'href'} if node.tag=='a' else set()
        if set(node.attrib)!=expected:raise ValueError('Unsupported HTML attributes in '+label)
    if normalized_hrefs(value)!=normalized_hrefs(source_value):raise ValueError('Hyperlink mismatch: '+label)
    if check_numbers and number_tokens(value)!=number_tokens(source_value):raise ValueError('Numeric clause mismatch: '+label)
    if len(plain(source_value))>120 and len(plain(value))<len(plain(source_value))*.18:raise ValueError('Truncated paragraph: '+label)


def validate_locale(source,data,locale):
    if locale not in LOCALES:raise ValueError('Unsupported locale: '+locale)
    if data.get('schemaVersion')!=1 or data.get('sourceHash')!=source['sourceHash'] or data.get('locale')!=locale:raise ValueError('Locale identity/hash mismatch: '+locale)
    if set(data['documents'])!={doc['id'] for doc in source['documents']}:raise ValueError('Document coverage mismatch: '+locale)
    for doc in source['documents']:
        actual=data['documents'][doc['id']];label=locale+'/'+doc['id']
        if set(actual)!={'title','updatedLabel','footerHTML','blocks'}:raise ValueError('Document field coverage mismatch: '+label)
        for field in ('title','updatedLabel','footerHTML'):validate_html(actual[field],escape(doc[field]) if field!='footerHTML' else doc[field],label+'/'+field,check_numbers=field!='updatedLabel')
        if set(actual['blocks'])!={b['id'] for b in doc['blocks']}:raise ValueError('Block coverage mismatch: '+label)
        for block in doc['blocks']:
            value=actual['blocks'][block['id']]
            validate_html(value,block['html'],label+'/'+block['id'])
            if block['tag'] in ('ul','ol'):
                fragment=html.fragment_fromstring('<div>'+value+'</div>');original=html.fragment_fromstring('<div>'+block['html']+'</div>')
                if [n.tag for n in fragment]!=[n.tag for n in original]:raise ValueError('List structure mismatch: '+label+'/'+block['id'])
        if locale not in ('en','en-AU','en-GB','en-IN'):
            long_blocks=[b for b in doc['blocks'] if len(plain(b['html']))>120]
            if any(plain(actual['blocks'][b['id']])==plain(b['html']) for b in long_blocks):raise ValueError('Untranslated full paragraph: '+label)
    if set(data['uiLabels'])!=set(source['uiLabels']):raise ValueError('UI label coverage mismatch: '+locale)
    for key,value in data['uiLabels'].items():validate_html(escape(value),escape(source['uiLabels'][key]),locale+'/ui/'+key)
    if set(data['privacyHostingVariants'])!=set(source['privacyHostingVariants']):raise ValueError('Hosting document coverage mismatch: '+locale)
    for doc_id,variants in source['privacyHostingVariants'].items():
        actual=data['privacyHostingVariants'][doc_id]
        if set(actual)!=set(variants):raise ValueError('Hosting provider coverage mismatch: '+locale+'/'+doc_id)
        for host,value in actual.items():validate_html(value,variants[host],locale+'/'+doc_id+'/'+host)


def finalize(source,directory,reviewed):
    directory=Path(directory);reviewed=Path(reviewed)
    # Stage every file before replacing any catalog; missing review never becomes an English fallback.
    catalogs={'en':json.loads((ROOT/'legal/locales/en.json').read_text()),'pl':canonical_pl(source)}
    for locale in LOCALES:
        if locale in catalogs:continue
        path=reviewed/(locale+'.json')
        if locale in ALIASES and not path.is_file():continue
        if not path.is_file():raise ValueError('Missing reviewed locale: '+locale)
        catalogs[locale]=json.loads(path.read_text());validate_locale(source,catalogs[locale],locale)
        if catalogs[locale].get('translation',{}).get('reviewStatus') not in ('targeted-editorial-review','reviewed-local'):
            raise ValueError('Locale has not completed local editorial review: '+locale)
    for locale,base in ALIASES.items():
        if locale in catalogs:continue
        catalogs[locale]=copy.deepcopy(catalogs[base]);catalogs[locale]['locale']=locale
        catalogs[locale]['translation']['derivedFrom']=base
    for locale,data in catalogs.items():validate_locale(source,data,locale)
    for locale in LOCALES:write_json(directory/(locale+'.json'),catalogs[locale])
    write_json(ROOT/'legal/locale-sources.json',{'schemaVersion':1,'sourceHash':source['sourceHash'],'localeSources':{locale:catalogs[locale].get('translation',{}).get('derivedFrom',locale) for locale in LOCALES}})
    print(f'Finalized {len(catalogs)} explicit locales from canonical EN/PL and reviewed translations')


def main():
    p=argparse.ArgumentParser();p.add_argument('--index',type=Path,default=DEFAULT_INDEX);commands=p.add_subparsers(dest='command',required=True)
    commands.add_parser('build')
    polish=commands.add_parser('polish');polish.add_argument('--output',type=Path,default=ROOT/'legal/locales/pl.json')
    check=commands.add_parser('check');check.add_argument('--directory',type=Path,default=ROOT/'legal/locales');check.add_argument('--allow-incomplete',action='store_true')
    final=commands.add_parser('finalize');final.add_argument('--reviewed',type=Path,required=True);final.add_argument('--directory',type=Path,default=ROOT/'legal/locales')
    export=commands.add_parser('export');export.add_argument('--output',type=Path,required=True)
    imp=commands.add_parser('import');imp.add_argument('docx',type=Path);imp.add_argument('--locale',required=True);imp.add_argument('--output',type=Path,required=True);imp.add_argument('--allow-source',action='store_true')
    args=p.parse_args()
    if args.command=='build':build(args.index);return
    source=load_source(args.index)
    if args.command=='export':export_docx(source,args.output)
    elif args.command=='import':import_docx(source,args.docx,args.locale,args.output,args.allow_source)
    elif args.command=='polish':canonical_pl(source,args.output)
    elif args.command=='finalize':finalize(source,args.directory,args.reviewed)
    else:
        paths=sorted(args.directory.glob('*.json'))
        if not args.allow_incomplete and {p.stem for p in paths}!=set(LOCALES):raise ValueError('Complete locale coverage requires all 52 explicit catalogs')
        for path in paths:validate_locale(source,json.loads(path.read_text()),path.stem)
        print(f'PASS {len(paths)} catalogs: source hash, block/field coverage, hyperlinks, lists, numerical clauses and all hosting alternatives')

if __name__=='__main__':main()
