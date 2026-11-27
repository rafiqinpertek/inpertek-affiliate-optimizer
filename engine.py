"""Read-only multi-master engine. Sources never modified. Python 3.10+."""
from __future__ import annotations
import collections, hashlib, json, math, os, posixpath, re, sqlite3, tempfile, zipfile
from lxml import etree as LET
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
CREATE INDEX IF NOT EXISTS idx_changes_master ON changes(master_id,run_id);
CREATE TABLE IF NOT EXISTS custom_maps(master_id TEXT,sheet TEXT,map_json TEXT,PRIMARY KEY(master_id,sheet));
'''

# Skema setara untuk Postgres (Neon). DOUBLE PRECISION dipakai agar angka Rupiah tidak kehilangan presisi.
PG_DDL=[
 'CREATE TABLE IF NOT EXISTS masters(master_id TEXT PRIMARY KEY, source_name TEXT NOT NULL, last_hash TEXT, last_run BIGINT, updated_at TEXT)',
 'CREATE TABLE IF NOT EXISTS runs(run_id BIGSERIAL PRIMARY KEY,master_id TEXT NOT NULL,source_name TEXT,file_hash TEXT,sync_time TEXT,sheet_count INTEGER,source_rows INTEGER,records INTEGER,issues INTEGER)',
 'CREATE TABLE IF NOT EXISTS items(run_id BIGINT NOT NULL,master_id TEXT NOT NULL,sheet TEXT NOT NULL,source_row INTEGER NOT NULL,marketplace TEXT NOT NULL,sku TEXT,name TEXT,hpp DOUBLE PRECISION,price DOUBLE PRECISION,old_fee DOUBLE PRECISION,old_pct DOUBLE PRECISION,old_margin DOUBLE PRECISION,validation TEXT,PRIMARY KEY(run_id,sheet,source_row,marketplace))',
 'CREATE INDEX IF NOT EXISTS idx_items_master ON items(master_id,run_id)',
 'CREATE INDEX IF NOT EXISTS idx_items_sku ON items(marketplace,sku)',
 'CREATE TABLE IF NOT EXISTS sheet_maps(master_id TEXT,sheet TEXT,map_json TEXT,status TEXT,reason TEXT,updated_at TEXT,PRIMARY KEY(master_id,sheet))',
 'CREATE TABLE IF NOT EXISTS changes(run_id BIGINT,master_id TEXT,sheet TEXT,marketplace TEXT,sku TEXT,field TEXT,old_value TEXT,new_value TEXT,kind TEXT)',
 'CREATE INDEX IF NOT EXISTS idx_changes_master ON changes(master_id,run_id)',
 'CREATE TABLE IF NOT EXISTS custom_maps(master_id TEXT,sheet TEXT,map_json TEXT,PRIMARY KEY(master_id,sheet))',
]

ITEM_COLS=('run_id','master_id','sheet','source_row','marketplace','sku','name','hpp','price','old_fee','old_pct','old_margin','validation')
CHANGE_COLS=('run_id','master_id','sheet','marketplace','sku','field','old_value','new_value','kind')

def is_postgres(target):
    return isinstance(target,str) and target.startswith(('postgres://','postgresql://'))

class Database:
    """Satu antarmuka untuk SQLite (lokal/sementara) dan Postgres (Neon, permanen).
    Query ditulis dengan placeholder '?' lalu diterjemahkan otomatis untuk Postgres."""
    def __init__(self,target):
        self.kind='postgres' if is_postgres(target) else 'sqlite'
        if self.kind=='postgres':
            import psycopg
            from psycopg.rows import dict_row
            self.con=psycopg.connect(target,row_factory=dict_row,connect_timeout=20,prepare_threshold=None)
            with self.con.cursor() as cur:
                for stmt in PG_DDL:cur.execute(stmt)
            self.con.commit()
        else:
            self.con=sqlite3.connect(target,timeout=30);self.con.row_factory=sqlite3.Row;self.con.executescript(DDL)
    def _sql(self,sql):
        return sql.replace('?','%s') if self.kind=='postgres' else sql
    def execute(self,sql,params=()):
        if self.kind=='postgres':
            cur=self.con.cursor();cur.execute(self._sql(sql),params);return cur
        return self.con.execute(sql,params)
    def executemany(self,sql,rows):
        rows=list(rows)
        if not rows:return
        if self.kind=='postgres':
            with self.con.cursor() as cur:cur.executemany(self._sql(sql),rows)
        else:self.con.executemany(sql,rows)
    def bulk_insert(self,table,cols,rows):
        """Sisip banyak baris sekaligus. Postgres memakai COPY (jauh lebih cepat lewat internet)."""
        rows=list(rows)
        if not rows:return
        if self.kind=='postgres':
            with self.con.cursor() as cur:
                with cur.copy(f"COPY {table} ({','.join(cols)}) FROM STDIN") as cp:
                    for r in rows:cp.write_row(r)
        else:
            self.con.executemany(f"INSERT INTO {table}({','.join(cols)}) VALUES({','.join('?' for _ in cols)})",rows)
    def insert_id(self,sql,params,id_col):
        if self.kind=='postgres':
            return self.execute(sql+f' RETURNING {id_col}',params).fetchone()[id_col]
        return self.con.execute(sql,params).lastrowid
    def begin(self):
        if self.kind=='sqlite':self.con.execute('BEGIN')
        # Postgres: transaksi dimulai otomatis dan diakhiri commit()/rollback().
    def commit(self):self.con.commit()
    def rollback(self):self.con.rollback()
    def close(self):
        try:
            if self.kind=='postgres':self.con.rollback()
        finally:self.con.close()
    def query_df(self,sql,params=()):
        import pandas as pd
        cur=self.execute(sql,params)
        cols=[d[0] for d in cur.description] if cur.description else []
        rows=cur.fetchall()
        data=[tuple(r[c] for c in cols) if self.kind=='postgres' else tuple(r) for r in rows]
        return pd.DataFrame.from_records(data,columns=cols)

def db_connect(target):
    return Database(target)

def storage_label(target):
    return 'Permanen · Neon Postgres' if is_postgres(target) else 'Sementara · SQLite di server'

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

_DIGITS='0123456789'

def _col_of(ref):
    """Huruf kolom dari referensi sel (mis. 'CC12' -> 'CC'); hasil sama persis dengan excel_col, tapi lebih cepat."""
    if not ref:return ''
    col=ref.rstrip(_DIGITS)
    if col and col.isalpha() and col.isupper() and col.isascii():return col
    return excel_col(ref)

_ROW_END=re.compile(rb'</(?:\w+:)?row>')
_V_SIMPLE=re.compile(rb'^\s*<(?:\w+:)?v>([^<&]*)</(?:\w+:)?v>\s*$')

def _cell_regex(cols):
    alt=b'|'.join(re.escape(c.encode()) for c in sorted(cols,key=len,reverse=True))
    return re.compile(rb'<(?:\w+:)?c\b([^>]*?)\br=(["\'])('+alt+rb')(\d+)\2([^>]*?)(?:/>|>(.*?)</(?:\w+:)?c>)',re.S)

_T_ATTR=re.compile(rb'\bt=(["\'])(.*?)\1')

class SheetRows:
    """Iterator baris sebuah sheet. Atribut `want` (set huruf kolom) dapat diisi setelah mapping
    diketahui; sejak itu hanya sel pada kolom yang dipakai yang diambil, langsung dari XML
    (jalur cepat). Sel berbentuk tidak biasa dibaca ulang dengan parser XML lengkap agar hasil identik."""
    def __init__(self,z,fn,shared):
        self.z=z;self.fn=fn;self.shared=shared;self.want=None;self._gen=None;self.last_row=0
    def __iter__(self):
        if self._gen is None:self._gen=self._rows()
        return self._gen
    def __next__(self):
        return next(iter(self))
    def _rows(self):
        shared=self.shared
        with self.z.open(self.fn) as fh:
            for _,el in LET.iterparse(fh,events=('end',),tag=M+'row',huge_tree=True):
                rownum=int(el.get('r','0'))
                vals={}
                for cell in el.iterchildren(M+'c'):
                    co=_col_of(cell.get('r'))
                    if co:
                        value=cell_value(cell,shared)
                        if value is not None:vals[co]=value
                el.clear(keep_tail=True)
                while el.getprevious() is not None:del el.getparent()[0]
                self.last_row=rownum
                yield rownum,vals
                if self.want is not None:break
        if self.want is not None:
            yield from self._fast_rows(self.last_row)
    def _slow_value(self,raw):
        root=LET.fromstring(b'<root xmlns="'+NS.encode()+b'">'+raw+b'</root>')
        cell=root[0] if len(root) else None
        return cell_value(cell,self.shared) if cell is not None else None
    def _fast_rows(self,after_row):
        rx=_cell_regex(self.want);shared=self.shared;import html as _h
        cur=None;vals={}
        with self.z.open(self.fn) as fh:
            buf=b''
            while True:
                chunk=fh.read(8*1024*1024)
                if chunk:
                    buf+=chunk
                    ends=list(_ROW_END.finditer(buf))
                    if not ends:continue
                    cut=ends[-1].end();part=buf[:cut];buf=buf[cut:]
                else:
                    part=buf;buf=b''
                for m in rx.finditer(part):
                    rownum=int(m.group(4))
                    if rownum<=after_row:continue
                    if rownum!=cur:
                        if cur is not None:yield cur,vals
                        cur=rownum;vals={}
                    attrs=m.group(1)+m.group(5);inner=m.group(6)
                    t=_T_ATTR.search(attrs);t=t.group(2) if t else None
                    if inner is None:value=None
                    else:
                        sv=_V_SIMPLE.match(inner)
                        if sv and sv.group(1):
                            txt=sv.group(1).decode('utf-8')
                            if t==b's':
                                try:value=shared[int(txt)]
                                except (ValueError,IndexError):value=None
                            else:value=txt
                        else:
                            value=self._slow_value(m.group(0))
                    if value is not None:vals[m.group(3).decode()]=value
                if not chunk:break
        if cur is not None:yield cur,vals

def read_workbook(path):
    """Yield worksheet name, color flag and row iterator; zip remains open during iteration."""
    with zipfile.ZipFile(path) as z:
        shared=[]
        names=set(z.namelist())
        if 'xl/sharedStrings.xml' in names:
            with z.open('xl/sharedStrings.xml') as fh:
                for _,e in LET.iterparse(fh,events=('end',),tag=M+'si',huge_tree=True):
                    shared.append(''.join(t.text or '' for t in e.iter(M+'t')));e.clear(keep_tail=True)
                    while e.getprevious() is not None:del e.getparent()[0]
        wb=ET.fromstring(z.read('xl/workbook.xml'))
        rel=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap={e.get('Id'):e.get('Target') for e in rel}
        for sh in wb.find(M+'sheets'):
            sheet=sh.get('name');target=relmap.get(sh.get('{'+REL+'}id'),'')
            fn=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl',target))
            if fn not in names:continue
            with z.open(fn) as fh:prefix=fh.read(40000)
            match=re.search(rb'<(?:\w+:)?sheetPr\b.*?</(?:\w+:)?sheetPr>',prefix,re.S)
            # Fallback to root sheetPr node when prefix namespaces have uncommon attributes.
            green=False
            if match:
                c=re.search(rb'<(?:\w+:)?tabColor\b[^>]*\brgb="([0-9a-fA-F]+)"',match.group(0))
                green=bool(c and c.group(1).decode().upper().endswith('92D050'))
            yield sheet,green,SheetRows(z,fn,shared)

def file_sha256(path):
    h=hashlib.sha256()
    with open(path,'rb') as f:
        for chunk in iter(lambda:f.read(1024*1024),b''):h.update(chunk)
    return h.hexdigest()

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

def diff_changes(rid, master_id, prev_rows, cur_rows):
    """Bandingkan dua snapshot master yang sama. SKU duplikat diabaikan agar tidak menggabungkan baris yang salah."""
    previous={(r['sheet'],r['marketplace'],r['sku']):r for r in prev_rows if r['validation']!='SKU_DUPLIKAT_DI_MASTER'}
    current={(r['sheet'],r['marketplace'],r['sku']):r for r in cur_rows if r['validation']!='SKU_DUPLIKAT_DI_MASTER'}
    change_rows=[]
    for k in set(previous)|set(current):
        a=previous.get(k);b=current.get(k)
        if a is None or b is None:
            change_rows.append((rid,master_id,*k,'ROW',str(a is not None),str(b is not None),'ADDED' if b else 'REMOVED'))
        else:
            for fld in ('price','hpp','old_pct','old_fee','old_margin'):
                v1=a[fld];v2=b[fld]
                if v1!=v2:change_rows.append((rid,master_id,*k,fld,str(v1),str(v2),'MODIFIED'))
    return change_rows

def sync_many(paths, db_path, overrides=None):
    """Independent source snapshot for each master; transaction per file. Invalid tabs skipped."""
    overrides=overrides or {};results=[]
    con=db_connect(db_path)
    for r in con.execute("SELECT master_id,sheet,map_json FROM custom_maps"):
        overrides.setdefault((r["master_id"],r["sheet"]),json.loads(r["map_json"]))
    try:
        for path in paths:
            filename=os.path.basename(path);master_id=filename.casefold()
            filehash=file_sha256(path)
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
                if isinstance(iterator,SheetRows):iterator.want=set(mapping.values())
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
                con.begin()
                rid=con.insert_id('INSERT INTO runs(master_id,source_name,file_hash,sync_time,sheet_count,source_rows,records,issues) VALUES (?,?,?,?,?,?,?,?)',(master_id,filename,filehash,now,countsheet,source,len(parsed),issues),'run_id')
                con.bulk_insert('items',ITEM_COLS,[(rid,r['master_id'],r['sheet'],r['source_row'],r['marketplace'],r['sku'],r['name'],r['hpp'],r['price'],r['old_fee'],r['old_pct'],r['old_margin'],r['validation']) for r in parsed])
                con.execute('INSERT INTO masters(master_id,source_name,last_hash,last_run,updated_at) VALUES(?,?,?,?,?) ON CONFLICT(master_id) DO UPDATE SET source_name=excluded.source_name,last_hash=excluded.last_hash,last_run=excluded.last_run,updated_at=excluded.updated_at',(master_id,filename,filehash,rid,now))
                con.executemany('INSERT INTO sheet_maps VALUES(?,?,?,?,?,?) ON CONFLICT(master_id,sheet) DO UPDATE SET map_json=excluded.map_json,status=excluded.status,reason=excluded.reason,updated_at=excluded.updated_at',[(master_id,sheet,json.dumps(mapping),status,reason,now) for sheet,mapping,status,reason in maprows])
                if old:
                    # Compare only nonduplicated keys, avoiding ungrounded row merges.
                    prev=con.execute('SELECT * FROM items WHERE run_id=?',(old['last_run'],)).fetchall()
                    change_rows=diff_changes(rid,master_id,prev,parsed)
                    con.bulk_insert('changes',CHANGE_COLS,change_rows)
                con.commit()
            except Exception:
                con.rollback();raise
            results.append({'file':filename,'status':'IMPORTED','run_id':rid,'green_sheets':countsheet,'validated_sheets':sum(x[2]=='VALIDATED' for x in maprows),'source_rows':source,'records':len(parsed),'issues':issues,'skipped':skipped})
    finally:con.close()
    return results


# ---------------------------------------------------------------------------
# Penghapusan data tersimpan. Setiap fungsi berjalan dalam satu transaksi:
# berhasil semua atau tidak sama sekali.
# ---------------------------------------------------------------------------
_MASTER_TABLES=('items','changes','runs','sheet_maps','custom_maps','masters')

def delete_run(db_path, run_id):
    """Hapus satu upload (snapshot). Riwayat perubahan di sekitarnya dihitung ulang:
    upload sesudahnya dibandingkan ulang dengan upload sebelumnya yang masih ada.
    Jika upload ini satu-satunya, master ikut terhapus."""
    con=db_connect(db_path)
    try:
        run=con.execute('SELECT run_id,master_id FROM runs WHERE run_id=?',(int(run_id),)).fetchone()
        if not run:return {'status':'NOT_FOUND'}
        master_id=run['master_id'];rid=int(run['run_id'])
        prev=con.execute('SELECT run_id,file_hash FROM runs WHERE master_id=? AND run_id<? ORDER BY run_id DESC LIMIT 1',(master_id,rid)).fetchone()
        nxt=con.execute('SELECT run_id FROM runs WHERE master_id=? AND run_id>? ORDER BY run_id LIMIT 1',(master_id,rid)).fetchone()
        con.begin()
        try:
            con.execute('DELETE FROM items WHERE run_id=?',(rid,))
            con.execute('DELETE FROM changes WHERE run_id=?',(rid,))
            con.execute('DELETE FROM runs WHERE run_id=?',(rid,))
            if nxt:
                nid=int(nxt['run_id'])
                con.execute('DELETE FROM changes WHERE run_id=?',(nid,))
                if prev:
                    a=con.execute('SELECT * FROM items WHERE run_id=?',(int(prev['run_id']),)).fetchall()
                    b=con.execute('SELECT * FROM items WHERE run_id=?',(nid,)).fetchall()
                    con.bulk_insert('changes',CHANGE_COLS,diff_changes(nid,master_id,a,b))
            remaining=con.execute('SELECT run_id,file_hash FROM runs WHERE master_id=? ORDER BY run_id DESC LIMIT 1',(master_id,)).fetchone()
            if remaining:
                con.execute('UPDATE masters SET last_run=?,last_hash=? WHERE master_id=?',(int(remaining['run_id']),remaining['file_hash'],master_id))
                status='RUN_DELETED'
            else:
                for t in ('sheet_maps','custom_maps','masters'):con.execute(f'DELETE FROM {t} WHERE master_id=?',(master_id,))
                status='MASTER_DELETED'
            con.commit()
        except Exception:
            con.rollback();raise
        return {'status':status,'master_id':master_id}
    finally:con.close()

def delete_master(db_path, master_id):
    """Hapus satu master beserta semua upload dan riwayat perubahannya."""
    con=db_connect(db_path)
    try:
        con.begin()
        try:
            for t in _MASTER_TABLES:con.execute(f'DELETE FROM {t} WHERE master_id=?',(master_id,))
            con.commit()
        except Exception:
            con.rollback();raise
        return {'status':'MASTER_DELETED','master_id':master_id}
    finally:con.close()

def delete_all(db_path):
    """Hapus semua data tersimpan (semua master, upload, dan riwayat)."""
    con=db_connect(db_path)
    try:
        con.begin()
        try:
            for t in _MASTER_TABLES:con.execute(f'DELETE FROM {t}')
            con.commit()
        except Exception:
            con.rollback();raise
        return {'status':'ALL_DELETED'}
    finally:con.close()
