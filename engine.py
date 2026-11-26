"""Read-only multi-master engine. Sources never modified. Python 3.10+."""
from __future__ import annotations
import collections, hashlib, json, math, os, posixpath, re, sqlite3, tempfile, zipfile
from datetime import datetime, timezone
from xml.etree import ElementTree as ET

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
M='{'+NS+'}'
STANDARD={'sku':'D','name':'E','hpp':'N','tok_price':'CC','tok_fee':'DA','tok_rate':'DB','tok_margin':'DH', 'shp_price':'EM','shp_fee':'FL','shp_rate':'FM','shp_margin':'FS'}
REQUIRED=['sku','name','hpp','tok_price','tok_fee','tok_rate','tok_margin','shp_price','shp_fee','shp_rate','shp_margin']
MARKETS={'TOKOPEDIA':('tok_price','tok_fee','tok_rate','tok_margin'),'SHOPEE':('shp_price','shp_fee','shp_rate','shp_margin')}

DDL='''
PRAGMA foreign_keys=ON;
CREATE TABLE IF NOT EXISTS masters(master_id TEXT PRIMARY KEY, source_name TEXT NOT NULL, last_hash TEXT, last_run INTEGER, updated_at TEXT);
CREATE TABLE IF NOT EXISTS runs(run_id INTEGER PRIMARY KEY AUTOINCREMENT,master_id TEXT NOT NULL,source_name TEXT,file_hash TEXT,sync_time TEXT,sheet_count INTEGER,source_rows INTEGER,records INTEGER,issues INTEGER);
CREATE TABLE IF NOT EXISTS items(run_id INTEGER NOT NULL,master_id TEXT NOT NULL,sheet TEXT NOT NULL,source_row INTEGER NOT NULL,marketplace TEXT NOT NULL,sku TEXT,name TEXT,hpp REAL,price REAL,old_fee REAL,old_pct REAL,old_margin REAL,validation TEXT,PRIMARY KEY(run_id,sheet,source_row,marketplace));
CREATE INDEX IF NOT EXISTS idx_items_master ON items(master_id,run_id);
CREATE INDEX IF NOT EXISTS idx_items_sku ON items(marketplace,sku);
CREATE TABLE IF NOT EXISTS sheet_maps(master_id TEXT,sheet TEXT,map_json TEXT,status TEXT,reason TEXT,updated_at TEXT,PRIMARY KEY(master_id,sheet));
CREATE TABLE IF NOT EXISTS changes(run_id INTEGER,master_id TEXT,sheet TEXT,marketplace TEXT,sku TEXT,field TEXT,old_value TEXT,new_value TEXT,kind TEXT);
CREATE TABLE IF NOT EXISTS custom_maps(master_id TEXT,sheet TEXT,map_json TEXT,PRIMARY KEY(master_id,sheet));
'''

def db_connect(path):
    con=sqlite3.connect(path,timeout=30);con.row_factory=sqlite3.Row;con.executescript(DDL);return con

def excel_col(ref):
    m=re.match(r'^([A-Z]+)',ref or '')
    return m.group(1) if m else ''

def num(v):
    try:
        if v is None or isinstance(v,bool) or str(v).strip()=='':return None
        n=float(v)
        return n if math.isfinite(n) else None
    except (ValueError,TypeError):return None

def cell_value(c,shared):
    v=c.find(M+'v')
    if v is not None and v.text is not None:
        if c.get('t')=='s':
            try:return shared[int(v.text)]
            except (ValueError,IndexError):return None
        return v.text
    if c.get('t')=='inlineStr':return ''.join(e.text or '' for e in c.iter(M+'t'))
    return None

def read_workbook(path):
    """Yield worksheet name, color flag and row iterator; zip remains open during iteration."""
    with zipfile.ZipFile(path) as z:
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            with z.open('xl/sharedStrings.xml') as fh:
                for _,e in ET.iterparse(fh,events=('end',)):
                    if e.tag==M+'si':
                        shared.append(''.join(t.text or '' for t in e.iter(M+'t')));e.clear()
        wb=ET.fromstring(z.read('xl/workbook.xml'))
        rel=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap={e.get('Id'):e.get('Target') for e in rel}
        for sh in wb.find(M+'sheets'):
            sheet=sh.get('name');target=relmap.get(sh.get('{'+REL+'}id'),'')
            fn=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl',target))
            if fn not in z.namelist():continue
            with z.open(fn) as fh:prefix=fh.read(40000)
            match=re.search(rb'<(?:\w+:)?sheetPr\b.*?</(?:\w+:)?sheetPr>',prefix,re.S)
            # Fallback to root sheetPr node when prefix namespaces have uncommon attributes.
            green=False
            if match:
                c=re.search(rb'<(?:\w+:)?tabColor\b[^>]*\brgb="([0-9a-fA-F]+)"',match.group(0))
                green=bool(c and c.group(1).decode().upper().endswith('92D050'))
            def rowiter():
                with z.open(fn) as fh:
                    for _,el in ET.iterparse(fh,events=('end',)):
                        if el.tag==M+'row':
                            rownum=int(el.get('r','0'))
                            vals={}
                            for cell in el.findall(M+'c'):
                                co=excel_col(cell.get('r'))
                                value=cell_value(cell,shared)
                                if co and value is not None:vals[co]=value
                            el.clear()
                            yield rownum,vals
            yield sheet,green,rowiter()

def column_index(col):
    n=0
    for c in col:n=n*26+ord(c)-64
    return n

def column_name(n):
    x=''
    while n>0:
        n,j=divmod(n-1,26);x=chr(65+j)+x
    return x

def try_detect_mapping(sample):
    # Offset templates are recognized ONLY when commission reconciliation confirms them.
    for delta in range(-25,26):
        if delta==0:continue
        mapping={k:column_name(column_index(v)+delta) for k,v in STANDARD.items() if column_index(v)+delta>0}
        if len(mapping)!=len(STANDARD):continue
        ok,reason=assess_mapping(sample,mapping)
        if ok:return mapping,f'Geser {delta:+d} kolom; {reason}'
    return None,None

def assess_mapping(rows,mapping):
    """Validate candidate on sample product rows; reject blind cross-template offset matches."""
    inspected=0;checks=0;matched=0;sku_count=0;numeric_count=0
    for rn,r in rows:
        if rn<5:continue
        sku=str(r.get(mapping['sku'],'') or '').strip()
        if not sku or sku.upper() in ('SKU','NAMA SKU'):continue
        inspected+=1;sku_count+=bool(re.search('[A-Za-z0-9]',sku))
        if num(r.get(mapping['hpp'])) is not None:numeric_count+=1
        for _,keys in MARKETS.items():
            p,f,rate,_=keys
            price=num(r.get(mapping[p]));fee=num(r.get(mapping[f]));pct=num(r.get(mapping[rate]))
            if price is None or fee is None or pct is None or price<=0:continue
            checks+=1
            if abs(price*pct-fee)<=max(3,abs(fee)*0.002):matched+=1
        if inspected>=65:break
    if inspected<1:return False,'Tidak ada SKU pada baris data'
    if checks<2:return False,'Bukti hitungan komisi tidak cukup (harga × tarif = nominal)'
    ratio=matched/checks
    if ratio<0.80:return False,f'Perhitungan komisi tidak cocok ({matched}/{checks} contoh); mapping ditahan'
    if sku_count/inspected<0.8:return False,'Kolom SKU tidak meyakinkan'
    return True,f'Terverifikasi {matched}/{checks} perhitungan ({ratio:.0%}) pada sampel'

def validate(rec):
    errs=[]
    if not rec['sku']:errs.append('SKU_KOSONG')
    if rec['price'] is None or rec['price']<=0:errs.append('HARGA_TIDAK_VALID')
    if rec['old_pct'] is None or not 0<=rec['old_pct']<=100:errs.append('KOMISI_TIDAK_VALID')
    if rec['old_fee'] is None or rec['old_fee']<0:errs.append('BIAYA_AFF_TIDAK_VALID')
    if rec['old_margin'] is None:errs.append('MARGIN_TIDAK_VALID')
    if rec['hpp'] is None or rec['hpp']<0:errs.append('HPP_TIDAK_VALID')
    if not errs and abs(rec['price']*rec['old_pct']/100-rec['old_fee'])>max(3,rec['old_fee']*0.002):errs.append('BIAYA_KOMISI_TIDAK_COCOK')
    return 'OK' if not errs else '; '.join(errs)

# ---------------------------------------------------------------------------
# V5.3 · Algoritma "Margin Berjarak Aman"
# Margin master = margin bersih (sudah dikurangi komisi affiliate awal).
#   Margin sebelum affiliate (M) = margin master + biaya komisi awal.
# Aturan pokok: setelah komisi baru (C), SISA MARGIN >= k x KOMISI, dengan k > 1,
#   sehingga komisi tidak pernah melebihi sisa margin dan tidak mepet dengannya.
#   M - C >= k*C  ->  C <= M / (1+k)   (batas komisi per produk)
# Mode menentukan jarak aman k:
#   Agresif     k=1.25 (sisa margin >= 1,25x komisi; komisi maks ~44% margin)
#   Seimbang    k=1.50 (sisa margin >= 1,5x komisi;  komisi maks 40% margin)
#   Konservatif k=2.00 (sisa margin >= 2x komisi;    komisi maks ~33% margin)
# Setiap produk disesuaikan ke batas komisinya (dibulatkan ke bawah, kelipatan 1%):
#   NAIK bila komisi awal di bawah batas, TURUNKAN bila di atas batas,
#   TETAP bila sudah tepat, NONAKTIF bila komisi 1% pun tidak memenuhi jarak aman.
# ---------------------------------------------------------------------------
MODES={'AGRESIF':1.25,'SEIMBANG':1.50,'KONSERVATIF':2.00}

def _rp(v):
    return 'Rp '+f'{v:,.0f}'.replace(',','.')

def analyze(rec, mode='AGRESIF', min_profit=0, max_pct=30, target_pct=30, step=1.0, max_increase=None):
    if step != 1.0: raise ValueError('Kelipatan wajib 1 poin persentase')
    if mode not in MODES: raise ValueError('Mode tidak dikenal')
    k=MODES[mode]
    out={'status':'VALIDASI DATA','proposed_pct':None,'proposed_fee':None,'margin_before':None,'margin_new':None,'safe_cap':None,
         'hard_cap':None,'affiliate_share':None,'reason':'','balance':None,'auto_eligible':False}
    if rec['validation']!='OK':out['reason']=rec['validation'];return out
    if max_increase is not None and (not math.isfinite(float(max_increase)) or max_increase < 0):raise ValueError('Batas kenaikan tidak valid')
    price=float(rec['price']);old=float(rec['old_pct']);fee=float(rec['old_fee']);after=float(rec['old_margin'])
    before=after+fee;out['margin_before']=before
    if before<=0:
        out.update(status='REVIEW NONAKTIF',reason='Margin bersih + komisi awal tidak positif; tidak ada ruang untuk komisi affiliate.')
        return out
    min_profit=max(float(min_profit),0.0);limit=min(float(max_pct),float(target_pct))
    if before<min_profit:
        out.update(status='REVIEW NONAKTIF',reason=f'Margin sebelum affiliate {_rp(before)} sudah di bawah margin minimum {_rp(min_profit)}; tidak ada ruang untuk komisi.')
        return out
    margin_cap=before/(1+k)*100/price
    cap=max(0.0,min(limit,margin_cap,(before-min_profit)*100/price))
    out.update(safe_cap=cap,hard_cap=cap)
    cap_i=math.floor(cap+1e-9);eps=1e-8
    if cap_i<1:
        out.update(status='REVIEW NONAKTIF',reason=f'Margin sebelum affiliate {_rp(before)} terlalu tipis: komisi 1% pun membuat sisa margin kurang dari {k:g}x komisi.')
        return out
    if old>cap+eps:
        new=cap_i
        why='melewati batas komisi yang diatur' if limit<min(margin_cap,(before-min_profit)*100/price) else f'membuat sisa margin kurang dari {k:g}x komisi'
        reason=f'Komisi awal {old:g}% {why}; disesuaikan turun ke {new}%.'
    elif cap_i>old+eps:
        new=cap_i
        if max_increase is not None:new=max(min(new,math.floor(old+float(max_increase)+1e-9)),math.ceil(old-eps))
        reason=f'Margin masih longgar; komisi disesuaikan naik ke {new}%.'
        if new<cap_i:reason+=f' Dibatasi kenaikan per evaluasi (batas margin {cap_i}%).'
    else:
        new=old
        reason='Komisi awal sudah pas dengan margin produk; tidak perlu diubah.'
    cost=price*new/100;remain=before-cost
    if remain<-0.01 or remain+0.01<k*cost or remain+0.01<min_profit:
        out.update(status='VALIDASI DATA',reason='Pemeriksaan akhir perlindungan margin gagal');return out
    status='NAIK' if new>old+eps else ('TURUNKAN' if new<old-eps else 'TETAP')
    if status!='TETAP' or old>0:
        reason+=f' Sisa margin {_rp(remain)} = {remain/cost:.2f}x komisi {_rp(cost)}.' if cost>0 else ''
    out.update(status=status,proposed_pct=round(new,4),proposed_fee=cost,margin_new=remain,balance=remain-cost,
               affiliate_share=cost/before,reason=reason,auto_eligible=status in ('NAIK','TURUNKAN'))
    return out

def sync_many(paths, db_path, overrides=None):
    """Independent source snapshot for each master; transaction per file. Invalid tabs skipped."""
    overrides=overrides or {};results=[]
    con=db_connect(db_path)
    for r in con.execute("SELECT master_id,sheet,map_json FROM custom_maps"):
        overrides.setdefault((r["master_id"],r["sheet"]),json.loads(r["map_json"]))
    try:
        for path in paths:
            filename=os.path.basename(path);master_id=filename.casefold()
            filehash=hashlib.sha256(open(path,'rb').read()).hexdigest()
            old=con.execute('SELECT last_hash,last_run FROM masters WHERE master_id=?',(master_id,)).fetchone()
            has_override=any(k[0]==master_id for k in overrides)
            if old and old['last_hash']==filehash and not has_override:
                results.append({'file':filename,'status':'UNCHANGED','run_id':old['last_run']});continue
            now=datetime.now(timezone.utc).isoformat()
            parsed=[];maprows=[];skipped=[];countsheet=0;source=0;issues=0
            for sheet,green,iterator in read_workbook(path):
                if not green:continue
                countsheet+=1
                it=iter(iterator);sample=[]
                for j,(rn,r) in enumerate(it):
                    sample.append((rn,r))
                    if j>=85:break
                mapping=overrides.get((master_id,sheet)) or STANDARD
                ok,reason=assess_mapping(sample,mapping)
                if not ok and (master_id,sheet) not in overrides:
                    detected,detected_reason=try_detect_mapping(sample)
                    if detected:
                        mapping=detected;ok=True;reason=detected_reason
                if not ok:
                    skipped.append({'sheet':sheet,'reason':reason});maprows.append((sheet,mapping,'SKIPPED',reason));continue
                maprows.append((sheet,mapping,'VALIDATED',reason))
                for rn,r in __import__('itertools').chain(sample,it):
                    if rn<5:continue
                    sku=str(r.get(mapping['sku'],'') or '').strip()
                    if not sku or sku.upper() in ('SKU','NAMA SKU'):continue
                    source+=1
                    for market,keys in MARKETS.items():
                        p,f,rate,margin=keys
                        rateval=num(r.get(mapping[rate]))
                        rec={'master_id':master_id,'sheet':sheet,'source_row':rn,'marketplace':market,'sku':sku,'name':str(r.get(mapping['name'],'') or '').strip(),'hpp':num(r.get(mapping['hpp'])),'price':num(r.get(mapping[p])),'old_fee':num(r.get(mapping[f])),'old_pct':rateval*100 if rateval is not None else None,'old_margin':num(r.get(mapping[margin]))}
                        rec['validation']=validate(rec);parsed.append(rec)
                        if rec['validation']!='OK':issues+=1
            if not countsheet:
                results.append({'file':filename,'status':'NO_GREEN_SHEETS','sheets':0});continue
            # Duplicate SKU inside same master and marketplace, including cross-sheet, blocks export.
            counts=collections.Counter((r['marketplace'],r['sku']) for r in parsed)
            for r in parsed:
                if counts[(r['marketplace'],r['sku'])]>1:
                    if r['validation']=='OK':issues+=1
                    r['validation']='SKU_DUPLIKAT_DI_MASTER'
            try:
                con.execute('BEGIN')
                cur=con.execute('INSERT INTO runs(master_id,source_name,file_hash,sync_time,sheet_count,source_rows,records,issues) VALUES (?,?,?,?,?,?,?,?)',(master_id,filename,filehash,now,countsheet,source,len(parsed),issues))
                rid=cur.lastrowid
                con.executemany('INSERT INTO items VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',[(rid,r['master_id'],r['sheet'],r['source_row'],r['marketplace'],r['sku'],r['name'],r['hpp'],r['price'],r['old_fee'],r['old_pct'],r['old_margin'],r['validation']) for r in parsed])
                con.execute('INSERT INTO masters(master_id,source_name,last_hash,last_run,updated_at) VALUES(?,?,?,?,?) ON CONFLICT(master_id) DO UPDATE SET source_name=excluded.source_name,last_hash=excluded.last_hash,last_run=excluded.last_run,updated_at=excluded.updated_at',(master_id,filename,filehash,rid,now))
                con.executemany('INSERT INTO sheet_maps VALUES(?,?,?,?,?,?) ON CONFLICT(master_id,sheet) DO UPDATE SET map_json=excluded.map_json,status=excluded.status,reason=excluded.reason,updated_at=excluded.updated_at',[(master_id,sheet,json.dumps(mapping),status,reason,now) for sheet,mapping,status,reason in maprows])
                if old:
                    # Compare only nonduplicated keys, avoiding ungrounded row merges.
                    prev=con.execute('SELECT * FROM items WHERE run_id=?',(old['last_run'],)).fetchall()
                    previous={(r['sheet'],r['marketplace'],r['sku']):r for r in prev if r['validation']!='SKU_DUPLIKAT_DI_MASTER'}
                    current={(r['sheet'],r['marketplace'],r['sku']):r for r in parsed if r['validation']!='SKU_DUPLIKAT_DI_MASTER'}
                    for k in set(previous)|set(current):
                        a=previous.get(k);b=current.get(k)
                        if a is None or b is None:
                            con.execute('INSERT INTO changes VALUES(?,?,?,?,?,?,?,?,?)',(rid,master_id,*k,'ROW',str(a is not None),str(b is not None),'ADDED' if b else 'REMOVED'))
                        else:
                            for fld in ('price','hpp','old_pct','old_fee','old_margin'):
                                v1=a[fld];v2=b[fld]
                                if v1!=v2:con.execute('INSERT INTO changes VALUES(?,?,?,?,?,?,?,?,?)',(rid,master_id,*k,fld,str(v1),str(v2),'MODIFIED'))
                con.commit()
            except Exception:
                con.rollback();raise
            results.append({'file':filename,'status':'IMPORTED','run_id':rid,'green_sheets':countsheet,'validated_sheets':sum(x[2]=='VALIDATED' for x in maprows),'source_rows':source,'records':len(parsed),'issues':issues,'skipped':skipped})
    finally:con.close()
    return results
