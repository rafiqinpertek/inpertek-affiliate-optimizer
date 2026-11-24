"""Read-only snapshot analyzer for INSERT master. Python 3.10+."""
import zipfile, xml.etree.ElementTree as ET, sqlite3, json, hashlib, re, time, os
from datetime import datetime, timezone
from decimal import Decimal, ROUND_DOWN
M='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
FIELDS={'sku':'D','name':'E','hpp':'N','tok_price':'CC','tok_fee':'DA','tok_rate':'DB','tok_margin':'DH','tok_margin_pct':'DI','shp_price':'EM','shp_fee':'FL','shp_rate':'FM','shp_margin':'FS','shp_margin_pct':'FT'}
NEEDED=set(FIELDS.values()); GREEN='92D050'

SCHEMA='''
CREATE TABLE IF NOT EXISTS sync_runs(run_id INTEGER PRIMARY KEY AUTOINCREMENT, file_hash TEXT NOT NULL, source_name TEXT, sync_time TEXT NOT NULL, sheets_green INTEGER, source_rows INTEGER, listing_rows INTEGER, issues INTEGER, notes TEXT);
CREATE TABLE IF NOT EXISTS records(run_id INTEGER, source_sheet TEXT, source_row INTEGER, sku TEXT, name TEXT, marketplace TEXT, price REAL, affiliate_nominal REAL, commission_pct REAL, margin_after REAL, margin_pct REAL, hpp REAL, row_hash TEXT, validation TEXT, PRIMARY KEY(run_id,source_sheet,source_row,marketplace));
CREATE INDEX IF NOT EXISTS idx_records_key ON records(sku,marketplace);
CREATE TABLE IF NOT EXISTS changes(run_id INTEGER, marketplace TEXT, source_sheet TEXT, sku TEXT, field TEXT, previous_value TEXT, new_value TEXT, event_type TEXT);
CREATE TABLE IF NOT EXISTS issues(run_id INTEGER, source_sheet TEXT, source_row INTEGER, marketplace TEXT, sku TEXT, issue TEXT);
CREATE TABLE IF NOT EXISTS mappings(source_sheet TEXT PRIMARY KEY, mapping_json TEXT, validated INTEGER, updated_at TEXT);
'''

def init(db):
    con=sqlite3.connect(db);con.executescript(SCHEMA);return con

def to_number(x):
    if x is None or isinstance(x,bool):return None
    try:return float(x)
    except (ValueError,TypeError):return None

def read_cells(row, shared):
    vals={}
    for cell in row.iter(M+'c'):
        coord=cell.get('r',''); column=re.match(r'[A-Z]+',coord)
        if not column or column.group() not in NEEDED:continue
        v=cell.find(M+'v'); typ=cell.get('t')
        if v is not None and v.text is not None:
            val=shared[int(v.text)] if typ=='s' else v.text
        elif typ=='inlineStr':
            val=''.join(t.text or '' for t in cell.iter(M+'t'))
        else:val=None
        vals[column.group()]=val
    return vals

def workbook_rows(path):
    with zipfile.ZipFile(path) as z:
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            with z.open('xl/sharedStrings.xml') as handle:
                for e,el in ET.iterparse(handle,events=('end',)):
                    if el.tag==M+'si':
                        shared.append(''.join(t.text or '' for t in el.iter(M+'t')));el.clear()
        wb=ET.fromstring(z.read('xl/workbook.xml'))
        rel=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap={el.attrib['Id']:el.attrib['Target'] for el in rel}
        for sh in wb.find(M+'sheets'):
            name=sh.get('name');target=relmap[sh.attrib[R+'id']]
            fn=target.lstrip('/') if target.startswith('/') else 'xl/'+target
            # read first 30KB, sheetPr is first property in normal Excel XML
            with z.open(fn) as fh: prefix=fh.read(32768)
            match=re.search(rb'<sheetPr\b.*?</sheetPr>',prefix,re.S)
            color=re.search(rb'<tabColor\b[^>]*rgb="([0-9A-Fa-f]+)"',match.group(0)) if match else None
            if not color or not color.group(1).decode().upper().endswith(GREEN):continue
            def reader():
                with z.open(fn) as fh:
                    for _,el in ET.iterparse(fh,events=('end',)):
                        if el.tag==M+'row':
                            number=int(el.get('r','0'))
                            if number>=5: yield number,read_cells(el,shared)
                            el.clear()
            yield name,reader()

def validate(rec):
    errs=[]
    if not rec['sku']:errs.append('SKU_KOSONG')
    if rec['price'] is None or rec['price']<=0:errs.append('HARGA_TIDAK_VALID')
    if rec['commission_pct'] is None or not (0<=rec['commission_pct']<=100):errs.append('KOMISI_TIDAK_VALID')
    if rec['margin_after'] is None:errs.append('MARGIN_TIDAK_VALID')
    if rec['affiliate_nominal'] is None:errs.append('AFF_RP_TIDAK_VALID')
    if not errs:
        expected=rec['price']*rec['commission_pct']/100
        if abs(expected-rec['affiliate_nominal'])>max(2,abs(expected)*0.001):errs.append('KOMISI_NOMINAL_TIDAK_COCOK')
    return ', '.join(errs) if errs else 'OK'

def make_record(sheet,number,c,market):
    prefix='tok' if market=='TOKOPEDIA' else 'shp'
    rate=to_number(c.get(prefix+'_rate'))
    if rate is not None:rate*=100
    rec={'source_sheet':sheet,'source_row':number,'sku':str(c.get('sku') or '').strip(),
      'name':str(c.get('name') or '').strip(),'marketplace':market,'price':to_number(c.get(prefix+'_price')),
      'affiliate_nominal':to_number(c.get(prefix+'_fee')),'commission_pct':rate,
      'margin_after':to_number(c.get(prefix+'_margin')),'margin_pct':to_number(c.get(prefix+'_margin_pct')),
      'hpp':to_number(c.get('hpp'))}
    rec['validation']=validate(rec)
    rec['row_hash']=hashlib.sha256(json.dumps({k:rec[k] for k in ['sku','name','marketplace','price','affiliate_nominal','commission_pct','margin_after','margin_pct','hpp']},sort_keys=True).encode()).hexdigest()
    return rec

def sync(path,db):
    h=hashlib.sha256()
    with open(path,'rb') as fh:
        for b in iter(lambda:fh.read(1024*1024),b''):h.update(b)
    con=init(db)
    last=con.execute('SELECT run_id,file_hash FROM sync_runs ORDER BY run_id DESC LIMIT 1').fetchone()
    if last and last[1]==h.hexdigest():
        con.close(); return {'status':'UNCHANGED','run_id':last[0]}
    now=datetime.now(timezone.utc).isoformat()
    green=0;source_rows=0;counts=0;issues=0
    cur=con.cursor()
    try:
        cur.execute('BEGIN')
        cur.execute('INSERT INTO sync_runs (file_hash,source_name,sync_time,sheets_green,source_rows,listing_rows,issues,notes) VALUES (?,?,?,?,?,?,?,?)',(h.hexdigest(),os.path.basename(path),now,0,0,0,0,''))
        run_id=cur.lastrowid
        for sheet, rows in workbook_rows(path):
            green+=1
            cur.execute('INSERT OR REPLACE INTO mappings VALUES (?,?,?,?)',(sheet,json.dumps(FIELDS),1,now))
            for nr,cs in rows:
                sku=cs.get('D')
                if sku is None or not str(sku).strip():continue
                if str(sku).strip().upper() in ('SKU','NAMA SKU'):continue
                source_rows+=1
                c={key:cs.get(col) for key,col in FIELDS.items()}
                for market in ['TOKOPEDIA','SHOPEE']:
                    rec=make_record(sheet,nr,c,market);counts+=1
                    cur.execute('''INSERT INTO records VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',(run_id,*[rec[k] for k in ['source_sheet','source_row','sku','name','marketplace','price','affiliate_nominal','commission_pct','margin_after','margin_pct','hpp','row_hash','validation']]))
                    if rec['validation']!='OK':
                        issues+=1
                        cur.execute('INSERT INTO issues VALUES (?,?,?,?,?,?)',(run_id,sheet,nr,market,rec['sku'],rec['validation']))
        if not green:raise ValueError('NO_GREEN_SHEETS: no eligible green tabs found')
        # Within-sheet repeated SKU is ambiguous: do not recommend commission.
        duplicates=con.execute('SELECT source_sheet,sku,marketplace,COUNT(*) FROM records WHERE run_id=? GROUP BY source_sheet,sku,marketplace HAVING COUNT(*)>1',(run_id,)).fetchall()
        for sh,sku,market,n in duplicates:
            cur.execute("UPDATE records SET validation='SKU_DUPLIKAT_DALAM_SHEET' WHERE run_id=? AND source_sheet=? AND sku=? AND marketplace=?",(run_id,sh,sku,market))
            cur.execute('INSERT INTO issues VALUES (?,?,?,?,?,?)',(run_id,sh,0,market,sku,f'SKU_DUPLIKAT_DALAM_SHEET: {n} baris'))
            issues+=n
        # key includes sheet because products may appear in multiple brands/sheets
        if last:
            dupkeys={(sh,sku,market) for sh,sku,market,_ in duplicates}
            prev={(r[0],r[1],r[2]):r[3:] for r in con.execute('SELECT source_sheet,sku,marketplace,price,commission_pct,margin_after,affiliate_nominal,hpp FROM records WHERE run_id=?',(last[0],))}
            current={(r[0],r[1],r[2]):r[3:] for r in con.execute('SELECT source_sheet,sku,marketplace,price,commission_pct,margin_after,affiliate_nominal,hpp FROM records WHERE run_id=?',(run_id,))}
            for key in set(prev)|set(current):
                if key in dupkeys:
                    cur.execute('INSERT INTO changes VALUES (?,?,?,?,?,?,?,?)',(run_id,key[2],key[0],key[1],'SKU','DUPLIKAT','DUPLIKAT','VALIDATION_REQUIRED'))
                    continue
                if key not in prev or key not in current:
                    cur.execute('INSERT INTO changes VALUES (?,?,?,?,?,?,?,?)',(run_id,key[2],key[0],key[1],'ROW',str(key in prev),str(key in current),'ADDED' if key in current else 'REMOVED'))
                else:
                    for i,field in enumerate(('price','commission_pct','margin_after','affiliate_nominal','hpp')):
                        old,new=prev[key][i],current[key][i]
                        if old!=new:
                            cur.execute('INSERT INTO changes VALUES (?,?,?,?,?,?,?,?)',(run_id,key[2],key[0],key[1],field,str(old),str(new),'MODIFIED'))
        cur.execute('UPDATE sync_runs SET sheets_green=?,source_rows=?,listing_rows=?,issues=? WHERE run_id=?',(green,source_rows,counts,issues,run_id))
        con.commit()
        return {'status':'IMPORTED','run_id':run_id,'green_sheets':green,'source_rows':source_rows,'listing_rows':counts,'validation_issues':issues,'changes':con.execute('SELECT COUNT(*) FROM changes WHERE run_id=?',(run_id,)).fetchone()[0]}
    except Exception:
        con.rollback();raise
    finally:con.close()

def recommend(row,target_share=0.5,ceil_rate=20,step=0.5,max_increase=2):
    # financial illustration: safeguard 50% of margin before affiliate; not a business policy.
    if row['validation']!='OK':return ('VALIDASI',None,None,None)
    before=row['margin_after']+row['affiliate_nominal']
    if before<0:return ('REVIEW NONAKTIF',0,before,0)
    target=max(0,before*target_share)
    max_safe=max(0,(before-target)/row['price']*100)
    allowed=min(max_safe,ceil_rate)
    if allowed<step:return ('REVIEW NONAKTIF',0,before,allowed)
    rate=min(allowed,row['commission_pct']+max_increase)
    rate=float((Decimal(str(rate))/Decimal(str(step))).to_integral_value(rounding=ROUND_DOWN)*Decimal(str(step)))
    old=row['commission_pct']
    if rate>old+0.00001:action='POTENSI NAIK'
    elif rate<old-0.00001:action='TURUNKAN'
    else:action='TETAP'
    return (action,rate,before,allowed)


def recommend_balanced(row, protected_share=0.5, min_profit=0.0,
                       target_rate=20.0, ceiling_rate=20.0, step=0.5,
                       max_increase=None, opportunity_mode=False,
                       opportunity_min_ratio=0.75, use_headroom=0.5):
    """Dynamic, read-only commission recommendation. Rates are in percentage POINTS.

    Normal proposal guarantees nonnegative remaining margin, margin >= affiliate cost,
    and configured retained profit. For increases, use a configurable fraction of
    safe headroom instead of a fixed 2-point raise. Exceptions require manual review.
    """
    result = dict(status='VALIDASI DATA', proposed_rate=None,
                  margin_before=None, safe_cap=None, margin_after_new=None,
                  commission_new=None, balance_ratio=None, notes='',
                  allowed_auto=False, original_rate=None,
                  margin_old=None, old_commission=None, opportunity=False)
    if str(row.get('validation')) != 'OK':
        result['notes']='Data sumber belum valid: '+str(row.get('validation'))
        return result
    try:
        import math
        price=float(row['price']); old=float(row['commission_pct'])
        aff=float(row['affiliate_nominal']); margin=float(row['margin_after'])
        numeric=[price,old,aff,margin,protected_share,min_profit,target_rate,ceiling_rate,step,opportunity_min_ratio,use_headroom]
        if max_increase is not None: numeric.append(max_increase)
        if not all(math.isfinite(float(v)) for v in numeric) or price<=0 or old<0 or old>100 or aff<0 or step<=0 or not (0<=protected_share<=1) or not (0<=use_headroom<=1) or min_profit<0 or target_rate<0 or ceiling_rate<0:
            result['notes']='Input atau parameter tidak valid';return result
        before=margin+aff
        floor=max(float(min_profit),max(0,before)*float(protected_share),0.0)
        result.update(margin_before=before,margin_old=margin,old_commission=aff,original_rate=old)
        def round_down(v): return max(0.0,math.floor((v+1e-10)/step)*step)
        def apply(rate):
            c=price*rate/100; rem=before-c
            result.update(proposed_rate=rate,commission_new=c,margin_after_new=rem,
                          balance_ratio=(rem/c if c>0 else None))
        # balance: before - commission >= commission
        balance_cap=before/(2*price)*100
        # retain at least floor, including never-negative margin
        retained_cap=(before-floor)/price*100
        safe_cap=max(0.0,min(float(ceiling_rate),float(target_rate),balance_cap,retained_cap))
        result['safe_cap']=safe_cap
        max_standard=round_down(safe_cap)
        if before<=0 or max_standard<step:
            apply(0.0)
            result.update(status='REVIEW NONAKTIF',allowed_auto=False,
                notes='Ruang komisi standar tidak mencukupi. Komisi 0% hanya rekomendasi review; master dan marketplace tidak berubah.')
            return result
        if old>max_standard+1e-8:
            proposed=max_standard
            action='TURUNKAN'
            reason='Komisi awal melewati batas yang melindungi margin. Tarif diturunkan ke batas aman.'
        else:
            headroom=max(0.0,max_standard-old)
            # automatic but moderate: propose a fraction of the available headroom
            proposed=round_down(old+headroom*float(use_headroom))
            if max_increase is not None: proposed=min(proposed,round_down(old+float(max_increase)))
            # rounding must not lower a valid existing rate just because the step differs
            if proposed<old: proposed=old
            action='POTENSI NAIK' if proposed>old+1e-8 else 'TETAP'
            reason=('Kenaikan otomatis memakai sebagian ruang komisi aman; tidak memakai angka kenaikan tetap.'
                    if action=='POTENSI NAIK' else 'Komisi saat ini masih sesuai batas dan belum ada ruang kenaikan pada kelipatan terpilih.')
        apply(proposed)
        if result['margin_after_new'] < -1e-7 or result['margin_after_new'] + 1e-7 < result['commission_new'] or result['margin_after_new'] + 1e-7 < floor:
            result.update(status='VALIDASI PERHITUNGAN',allowed_auto=False,notes='Pemeriksaan margin gagal; rekomendasi ditahan.')
            return result
        result.update(status=action,allowed_auto=True,
                      notes=reason+f' Margin minimal terjaga Rp{floor:,.0f}; sisa margin >= komisi.')
        # Optional exception: allow less than commission, but NEVER negative, and explicitly review.
        if opportunity_mode and before>0 and action!='TURUNKAN':
            exceptional_cap=min(float(ceiling_rate),float(target_rate),retained_cap)
            if max_increase is not None: exceptional_cap=min(exceptional_cap,old+float(max_increase))
            exceptional=round_down(exceptional_cap)
            if exceptional>proposed+1e-8:
                c=price*exceptional/100; rem=before-c
                if rem>=0 and rem+1e-7>=floor and rem<c-1e-7 and rem+1e-7>=float(opportunity_min_ratio)*c:
                    apply(exceptional)
                    result.update(status='PELUANG KHUSUS - REVIEW',allowed_auto=False,opportunity=True,
                        notes='PENGECUALIAN: margin tetap nonnegatif dan memenuhi floor, tetapi margin tersisa < biaya komisi. Hanya simulasi, wajib persetujuan dan analisis performa sebelum upload.')
        return result
    except (ValueError,TypeError,KeyError) as exc:
        result['notes']='Gagal menghitung: '+str(exc)
        return result
