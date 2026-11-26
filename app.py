"""INPERTEK multi-master read-only commission optimizer. No Supabase required."""
import html, os, tempfile
from io import BytesIO
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from engine import db_connect, sync_many, analyze

st.set_page_config(page_title='INPERTEK | Affiliate Optimizer V5.3',page_icon='📈',layout='wide')

# Application visual shell: premium, restrained dark-light hybrid.
st.markdown("""
<style>
/* V5.3 SPEKTRUM - TAMPILAN SAJA. Warna dasar diatur di .streamlit/config.toml */
@import url('https://fonts.googleapis.com/css2?family=Unbounded:wght@500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap');
:root{
  --abyss:#10132A; --deck:#171B3A; --deck-2:#1D2247; --rim:#2B3166; --rim-soft:#232853;
  --iris:#8B7BFF; --ion:#5CE1FF; --mint:#7CF2C4; --flare:#FFB547; --rose:#FF6B8B; --magenta:#E36BFF;
  --frost:#ECEBFF; --haze:#A3A6D4; --dim:#7478A8;
  --spectrum:linear-gradient(90deg,#5CE1FF 0%,#8B7BFF 52%,#E36BFF 100%);
}
html, body, [data-testid="stApp"] { font-family:'Manrope',sans-serif; }
[data-testid="stApp"] {
  background:
    radial-gradient(1100px 520px at 92% -8%, rgba(139,123,255,.20), transparent 62%),
    radial-gradient(800px 480px at -8% 105%, rgba(92,225,255,.08), transparent 60%),
    var(--abyss);
}
[data-testid="stHeader"] { background:transparent; }
[data-testid="stMainBlockContainer"] { max-width:1480px; padding-top:2.4rem; padding-bottom:4rem; }
[data-testid="stMain"] h2, [data-testid="stMain"] h3 { font-family:'Unbounded',sans-serif; font-weight:600; letter-spacing:-.01em; font-size:1.15rem!important; color:var(--frost); }
[data-testid="stCaptionContainer"] p { color:var(--haze)!important; }
::selection { background:rgba(139,123,255,.45); }

/* Sidebar: panel kendali */
[data-testid="stSidebar"] { border-right:1px solid var(--rim-soft); }
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p { color:var(--frost); font-weight:600; }
[data-testid="stSidebar"] hr { border-color:var(--rim-soft); margin:1.2rem 0; }
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] { background:var(--deck); border:1px dashed var(--rim); border-radius:14px; }
.brand { display:flex; align-items:center; gap:12px; margin:.1rem 0 .4rem; }
.brand-mark { width:30px; height:30px; border-radius:9px; background:var(--spectrum); box-shadow:0 0 22px rgba(139,123,255,.5); flex:none; }
.brand-name { font-family:'Unbounded',sans-serif; font-weight:700; font-size:15px; color:#fff; line-height:1.1; }
.brand-sub { font-size:12px; color:var(--haze); margin-top:2px; }
.step { display:flex; align-items:center; gap:10px; margin:.4rem 0 .7rem; }
.step-n { width:26px; height:26px; border-radius:50%; border:1.5px solid var(--iris); color:var(--iris); font:600 11px 'Unbounded',sans-serif; display:grid; place-items:center; flex:none; }
.step-t { font-weight:800; font-size:15px; color:var(--frost); }

/* Hero */
.hero { display:grid; grid-template-columns:minmax(0,1.5fr) minmax(0,1fr); gap:28px 48px; align-items:end; margin:0 0 30px; }
.hero h1 { font-family:'Unbounded',sans-serif!important; font-weight:600!important; font-size:clamp(26px,3vw,42px)!important; line-height:1.08!important; letter-spacing:-.03em; color:#fff!important; margin:0!important; padding:0!important; }
.hero p { color:var(--haze); font-size:14.5px; line-height:1.65; max-width:62ch; margin:0 0 4px; }
.hero-line { grid-column:1 / -1; height:2px; border-radius:2px; background:var(--spectrum); transform-origin:left; animation:draw 1.1s cubic-bezier(.2,.7,.1,1) both; box-shadow:0 0 18px rgba(139,123,255,.6); }
@keyframes draw { from { transform:scaleX(0); opacity:.2; } to { transform:scaleX(1); opacity:1; } }

/* Panel spektrum */
.deck { background:linear-gradient(180deg,var(--deck-2),var(--deck)); border:1px solid var(--rim); border-radius:20px; padding:26px 28px 22px; margin:8px 0 30px; box-shadow:0 24px 60px -30px rgba(5,6,20,.9), inset 0 1px 0 rgba(255,255,255,.04); }
.deck-head { display:flex; justify-content:space-between; align-items:baseline; gap:16px; flex-wrap:wrap; margin-bottom:22px; }
.deck-title { font-family:'Unbounded',sans-serif; font-weight:600; font-size:16px; color:var(--frost); }
.deck-note { font-size:12.5px; color:var(--dim); }
.stats { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:0; margin-bottom:24px; }
.stat { padding:2px 22px; border-left:1px solid var(--rim-soft); }
.stat:first-child { padding-left:0; border-left:0; }
.stat b { display:block; font-family:'Unbounded',sans-serif; font-weight:600; font-size:clamp(24px,2.5vw,36px); color:#fff; letter-spacing:-.03em; line-height:1.1; font-variant-numeric:tabular-nums; }
.stat span { font-size:13px; color:var(--haze); }
.stat.hot b { color:var(--ion); text-shadow:0 0 24px rgba(92,225,255,.35); }
.stat.warn b { color:var(--flare); }
.band { display:flex; height:14px; border-radius:99px; overflow:hidden; background:var(--rim-soft); gap:2px; }
.band i { display:block; height:100%; min-width:3px; }
.legend { display:flex; flex-wrap:wrap; gap:8px 22px; margin:14px 0 0; padding:0; list-style:none; }
.legend li { display:flex; align-items:center; gap:8px; font-size:13px; color:var(--haze); }
.legend li i { width:10px; height:10px; border-radius:3px; display:inline-block; }
.legend li b { color:var(--frost); font-variant-numeric:tabular-nums; }

/* Tab sebagai segmented control (selector role-based: aman untuk Streamlit lama & baru) */
[data-testid="stTabs"] [role="tablist"] { gap:4px; background:var(--deck); border:1px solid var(--rim-soft); border-radius:14px; padding:5px; overflow-x:auto; box-shadow:none; border-bottom:1px solid var(--rim-soft); }
[data-testid="stTabs"] [role="tab"] { border-radius:10px; min-height:40px; padding:0 16px; white-space:nowrap; color:var(--haze); display:flex; align-items:center; border:0; background:transparent; }
[data-testid="stTabs"] [role="tab"] p { font-weight:700; font-size:13px; color:inherit; }
[data-testid="stTabs"] [role="tab"]:hover { color:var(--frost); background:rgba(139,123,255,.08); }
[data-testid="stTabs"] [role="tab"][aria-selected="true"] { background:linear-gradient(180deg,#3A3480,#2C2A66); color:#fff; box-shadow:inset 0 0 0 1px rgba(139,123,255,.55), 0 0 18px -4px rgba(139,123,255,.6); }
[data-testid="stTabs"] .react-aria-SelectionIndicator, [data-testid="stTabs"] [role="tablist"]::after { display:none; }
[data-testid="stTabs"] [data-baseweb="tab-highlight"], [data-testid="stTabs"] [data-baseweb="tab-border"] { display:none; }
[data-testid="stTabs"] [role="tabpanel"] { padding-top:1.2rem; }

/* Komponen */
[data-testid="stMetric"] { background:var(--deck); border:1px solid var(--rim-soft); border-radius:16px; padding:16px 20px; }
[data-testid="stMetricValue"] { font-family:'Unbounded',sans-serif; font-weight:600; color:#fff; }
[data-testid="stDataFrame"] { border:1px solid var(--rim-soft); border-radius:14px; overflow:hidden; }
[data-testid="stExpander"] { border:1px solid var(--rim-soft); border-radius:14px; background:var(--deck); }
[data-testid="stAlert"] { border-radius:12px; }
.stButton button[kind="primary"] { background:var(--spectrum); border:0; color:#0C0F23; font-weight:800; box-shadow:0 8px 26px -10px rgba(139,123,255,.9); }
.stButton button[kind="primary"]:hover { filter:brightness(1.08); color:#0C0F23; }
.stButton button[kind="primary"]:disabled { background:var(--rim-soft); color:var(--dim); box-shadow:none; }
[data-testid="stDownloadButton"] button { border:1px solid var(--rim); background:var(--deck-2); color:var(--frost); font-weight:700; border-radius:10px; }
[data-testid="stDownloadButton"] button:hover { border-color:var(--ion); color:#fff; box-shadow:0 0 0 3px rgba(92,225,255,.12); }
button:focus-visible, [role="tab"]:focus-visible { outline:2px solid var(--ion)!important; outline-offset:2px; }

@media (max-width:900px) {
  .hero { grid-template-columns:1fr; gap:16px; }
  .stats { grid-template-columns:repeat(2,minmax(0,1fr)); row-gap:18px; }
  .stat:nth-child(3) { padding-left:0; border-left:0; }
  .deck { padding:20px; }
}
@media (prefers-reduced-motion:reduce) { .hero-line { animation:none; } }
</style>
""",unsafe_allow_html=True)

DB=os.path.join(os.path.dirname(__file__),'optimizer_v5.sqlite')
MONEY=['HARGA JUAL (RP)','HPP (RP)','BIAYA AFF AWAL (RP)','MARGIN MASTER (RP)','MARGIN SEBELUM AFF (RP)','BIAYA KOMISI USULAN (RP)','SISA MARGIN USULAN (RP)','SELISIH MARGIN - KOMISI (RP)']
GUIDE={
'MASTER':'File master yang menjadi asal produk.','SHEET':'Worksheet hijau tempat data berasal.','BARIS':'Nomor baris pada worksheet asli.',
'MARKETPLACE':'TOKOPEDIA mencakup TikTok Shop sesuai keputusan internal.','SKU':'Kode produk di master.','NAMA PRODUK':'Nama asli pada master.',
'HARGA JUAL (RP)':'Harga jual pada blok marketplace.', 'HPP (RP)':'Harga pokok dari master.',
'KOMISI AWAL (%)':'Tarif affiliate saat snapshot master dibaca.', 'BIAYA AFF AWAL (RP)':'Biaya affiliate yang sudah tercermin pada margin master.',
'MARGIN MASTER (RP)':'Margin setelah biaya affiliate lama menurut master.', 'MARGIN SEBELUM AFF (RP)':'Margin master + biaya affiliate lama. Biaya affiliate hanya ditambahkan kembali sekali.',
'BATAS AMAN (%)':'Batas maksimum komisi simulasi dari aturan perlindungan margin.', 'KOMISI USULAN (%)':'Persentase baru yang direkomendasikan.',
'BIAYA KOMISI USULAN (RP)':'Harga jual dikali persentase baru (asumsi basis = harga jual).',
'SISA MARGIN USULAN (RP)':'Margin sebelum affiliate dikurangi biaya komisi baru.', 'SELISIH MARGIN - KOMISI (RP)':'Sisa margin dikurangi biaya komisi baru.',
'PERUBAHAN KOMISI (POIN %)':'Selisih tarif usulan dengan tarif awal, dalam poin persentase.','RASIO MARGIN / KOMISI':'Rasio sisa margin terhadap biaya affiliate usulan; untuk standar minimal 1.','STATUS':'Naik/turun/tetap/nonaktif/validasi/peluang khusus sesuai perhitungan.', 'VALIDASI':'OK berarti data lolos pemeriksaan dasar, bukan verifikasi batas resmi marketplace.',
'ALASAN':'Pertimbangan sistem untuk tindakan terhadap SKU.'}

def rupiah(v):
    try:
        if pd.isna(v):return '-'
        n=float(v);return ('-Rp' if n<0 else 'Rp')+f'{abs(n):,.0f}'.replace(',','.')
    except (ValueError,TypeError):return '-'

# --- Tampilan saja: komponen visual tanpa logika bisnis ---
STATUS_STYLE=[('NAIK','Naik','#5CE1FF'),('TETAP','Tetap','#8B7BFF'),('TURUNKAN','Turunkan','#FFB547'),
              ('PELUANG KHUSUS - REVIEW','Peluang khusus','#E36BFF'),('REVIEW NONAKTIF','Review nonaktif','#FF6B8B'),('VALIDASI DATA','Validasi data','#5B5F8F')]

def step_header(n,title):
    st.markdown(f'<div class="step"><span class="step-n">{n}</span><span class="step-t">{html.escape(title)}</span></div>',unsafe_allow_html=True)

def render_deck(master_count,df):
    total=len(df);counts=df['STATUS'].value_counts()
    known=[(k,label,color,int(counts.get(k,0))) for k,label,color in STATUS_STYLE]
    other=total-sum(c for *_,c in known)
    if other>0:known.append(('LAINNYA','Lainnya','#3A3F70',other))
    seg=''.join(f'<i style="flex:{c};background:{color}" title="{html.escape(label)}: {c}"></i>' for _,label,color,c in known if c>0)
    leg=''.join(f'<li><i style="background:{color}"></i>{html.escape(label)} <b>{c:,}</b></li>'.replace(',','.') for _,label,color,c in known if c>0)
    naik=int(counts.get('NAIK',0));issues=int((df['VALIDASI']!='OK').sum())
    def fmt(v):return f'{v:,}'.replace(',','.')
    st.markdown(f'''<section class="deck" aria-label="Ringkasan rekomendasi">
<div class="deck-head"><div class="deck-title">Spektrum rekomendasi</div><div class="deck-note">Sebaran status untuk {fmt(total)} baris marketplace sesuai filter</div></div>
<div class="stats">
<div class="stat"><b>{fmt(master_count)}</b><span>master dipilih</span></div>
<div class="stat"><b>{fmt(total)}</b><span>baris marketplace</span></div>
<div class="stat hot"><b>{fmt(naik)}</b><span>kandidat naik</span></div>
<div class="stat{' warn' if issues else ''}"><b>{fmt(issues)}</b><span>perlu validasi</span></div>
</div>
<div class="band" role="img" aria-label="Distribusi status rekomendasi">{seg}</div>
<ul class="legend">{leg}</ul>
</section>''',unsafe_allow_html=True)

def prepare_export(df):
    wb=Workbook();info=wb.active;info.title='PANDUAN'
    for row in [
        ['INPERTEK AFFILIATE OPTIMIZER V5.2 - REVIEW INTERNAL'],
        ['SUMBER','Snapshot master diunggah; master asli tidak diubah'],
        ['PERINGATAN','BUKAN TEMPLATE UPLOAD RESMI MARKETPLACE'],
        ['ALUR','Validasi listing & ketentuan marketplace, review, lalu susun file impor resmi'],
        ['METODE','Asumsi dasar komisi = harga jual. Margin sebelum affiliate = margin master + biaya affiliate awal'],
        ['BATAS','Sisa margin >= biaya komisi baru, >= margin minimum, dan tidak negatif'],
        ['KEAMANAN','Data master dan database jangan diunggah ke GitHub. Unduh hasil saat analisis selesai'],
        ['DUPLIKAT','SKU berulang antar master / sheet ditahan dari daftar kandidat'],
    ]:info.append(row)
    info.column_dimensions['A'].width=23;info.column_dimensions['B'].width=115
    for title,sub in [('REKOMENDASI',df),('KANDIDAT_REVIEW',df[(df['STATUS'].isin(['NAIK','TURUNKAN'])) & (df['VALIDASI']=='OK') & (~df['DUPLIKAT LINTAS MASTER']) & (df['KOMISI USULAN (%)']>0)])]:
        ws=wb.create_sheet(title)
        fields=[c for c in sub.columns if c!='DUPLIKAT LINTAS MASTER']
        ws.append(fields)
        for vals in sub[fields].itertuples(index=False,name=None):ws.append([None if pd.isna(x) else x for x in vals])
        ws.freeze_panes='F2';ws.auto_filter.ref=ws.dimensions
        for j,col in enumerate(fields,1):
            ws.column_dimensions[get_column_letter(j)].width=55 if col in ('NAMA PRODUK','ALASAN') else 24
            if col in MONEY:
                for r in range(2,ws.max_row+1):ws.cell(r,j).number_format='"Rp" #,##0;[Red]("Rp" #,##0);"Rp" 0'
            if col.endswith('(%)'):
                for r in range(2,ws.max_row+1):ws.cell(r,j).number_format='0.00"%"'
    for ws in wb:
        ws.sheet_view.showGridLines=False
        ws.row_dimensions[1].height=27
        for c in ws[1]:c.fill=PatternFill('solid',fgColor='17365D');c.font=Font(bold=True,color='FFFFFF',name='Calibri',size=10)
    file=BytesIO();wb.save(file);return file.getvalue()

st.markdown('''<header class="hero">
<h1>Affiliate Commission Optimizer</h1>
<p>Hitung ulang komisi affiliate per SKU di Shopee dan Tokopedia tanpa menggerus margin. Hanya sheet hijau dengan mapping terverifikasi yang diproses, master asli tidak pernah diubah, dan file ekspor adalah bahan review, bukan template impor marketplace.</p>
<div class="hero-line"></div>
</header>''',unsafe_allow_html=True)
try:db_connect(DB).close()
except Exception:st.error('Database SQLite tidak dapat dibuka.');st.stop()

with st.sidebar:
    st.markdown('<div class="brand"><div class="brand-mark"></div><div><div class="brand-name">INPERTEK</div><div class="brand-sub">Affiliate Optimizer v5.3</div></div></div>',unsafe_allow_html=True)
    st.divider()
    step_header(1,'Upload master')
    uploads=st.file_uploader('Pilih satu atau beberapa master (.xlsx)',type=['xlsx'],accept_multiple_files=True,help='Prioritas sheet tab hijau. Tidak boleh upload data sensitif ke GitHub.')
    if st.button('Analisis / sinkronkan master',disabled=not uploads,type='primary',use_container_width=True):
        with tempfile.TemporaryDirectory() as d:
            paths=[]
            for u in uploads:
                safe=os.path.basename(u.name)
                p=os.path.join(d,safe)
                with open(p,'wb') as f:
                    for chunk in iter(lambda:u.read(1024*1024),b''):f.write(chunk)
                u.seek(0);paths.append(p)
            try:
                with st.spinner('Memeriksa worksheet, mapping, nominal affiliate, dan snapshot...'):
                    results=sync_many(paths,DB)
                for r in results:
                    if r['status']=='IMPORTED':
                        st.success(f"{r['file']}: {r['validated_sheets']}/{r['green_sheets']} sheet terverifikasi; {r['records']} baris marketplace; {r['issues']} catatan validasi")
                        for s in r['skipped'][:15]:st.warning(f"{r['file']} · {s['sheet']}: {s['reason']}")
                    elif r['status']=='UNCHANGED':st.info(f"{r['file']}: file sama dengan snapshot terakhir; tidak diimpor ulang.")
                    else:st.warning(f"{r['file']}: {r['status']}")
            except Exception:
                st.error('Impor gagal. Periksa file XLSX dan pengaturan workbook. Data sebelumnya tetap tersedia.')
    st.divider()
    step_header(2,'Kebijakan komisi')
    st.caption('Default: agresif, tetapi sisa margin tidak boleh minus dan idealnya minimal sebesar komisi baru.')
    mode=st.selectbox('Agresivitas',['AGRESIF','SEIMBANG','KONSERVATIF'],index=0)
    headroom={'AGRESIF':.85,'SEIMBANG':.60,'KONSERVATIF':.35}[mode]
    retain=st.slider('Minimum margin sebelum affiliate yang dipertahankan (%)',0,95,30,5)/100
    min_profit=st.number_input('Margin minimum (Rp)',min_value=0.0,value=0.0,step=1000.0)
    cap=st.number_input('Batas komisi simulasi (%)',min_value=.5,max_value=100.0,value=30.0,step=.5)
    target=st.number_input('Target maksimum komisi (%)',min_value=.5,max_value=100.0,value=30.0,step=.5)
    st.info('Kelipatan komisi usulan ditetapkan 1 poin persentase: 3% → 4% → 5%. Tidak ada opsi 0,5%.')
    step=1.0
    limit_raise=st.toggle('Batasi kenaikan per evaluasi (opsional)',value=False)
    max_raise=st.number_input('Batas kenaikan (poin persentase)',min_value=1,max_value=100,value=5,step=1,disabled=not limit_raise) if limit_raise else None
    opportunity=st.toggle('Tampilkan skenario peluang khusus (review manual)',value=False,help='Margin boleh lebih kecil daripada biaya komisi, tetapi tidak negatif, dan dikecualikan dari kandidat ekspor otomatis.')
    opportunity_ratio=st.slider('Minimal rasio margin / komisi untuk peluang khusus',0.50,0.99,0.75,0.01,disabled=not opportunity)
    st.caption('Batas dan kelipatan adalah ASUMSI simulasi. Periksa aturan resmi dan dasar pengenaan komisi tiap marketplace sebelum upload.')

con=db_connect(DB)
masters=pd.read_sql_query('SELECT master_id,source_name,last_run,updated_at FROM masters ORDER BY source_name',con)
if masters.empty:
    st.warning('Belum ada snapshot. Upload master Excel melalui sidebar lalu klik Analisis / sinkronkan master.')
    con.close();st.stop()
selected=st.multiselect('Pilih master untuk analisis',masters.source_name.tolist(),default=masters.source_name.tolist())
chosen=masters[masters.source_name.isin(selected)]
if chosen.empty:st.warning('Pilih minimal satu master.');con.close();st.stop()
ids=chosen.last_run.astype(int).tolist()
q=','.join('?' for _ in ids)
base=pd.read_sql_query(f'SELECT i.*, m.source_name AS master_name FROM items i JOIN masters m ON m.master_id=i.master_id WHERE i.run_id IN ({q})',con,params=ids)
mapping=pd.read_sql_query('SELECT sm.master_id, m.source_name, sm.sheet, sm.status, sm.reason, sm.map_json FROM sheet_maps sm JOIN masters m ON m.master_id=sm.master_id',con)
run_stats=pd.read_sql_query('SELECT r.*,m.source_name AS latest_name FROM runs r JOIN masters m ON m.last_run=r.run_id',con)
change=pd.read_sql_query(f'SELECT * FROM changes WHERE run_id IN ({q}) ORDER BY run_id DESC',con,params=ids)
con.close()
with st.sidebar:
    step_header(3,'Filter data')
    market=st.selectbox('Marketplace',['SEMUA','SHOPEE','TOKOPEDIA'])
    sheet_filter=st.selectbox('Worksheet sumber',['SEMUA']+sorted(base.sheet.dropna().unique().tolist()))
    keyword=st.text_input('Cari SKU / nama produk')
    validity=st.selectbox('Validasi',['SEMUA','HANYA OK','BERMASALAH'])
if market!='SEMUA':base=base[base.marketplace==market]
if sheet_filter!='SEMUA':base=base[base.sheet==sheet_filter]
if keyword:base=base[base.sku.fillna('').str.contains(keyword,case=False,regex=False)|base.name.fillna('').str.contains(keyword,case=False,regex=False)]
if validity=='HANYA OK':base=base[base.validation=='OK']
if validity=='BERMASALAH':base=base[base.validation!='OK']
if base.empty:st.info('Tidak ada produk sesuai filter.');st.stop()

# Analyze in the view, independent of master. Cross-master duplicates are never eligible for export.
rows=[]
for r in base.to_dict('records'):
    result=analyze(r,retained_share=retain,min_profit=min_profit,max_pct=cap,target_pct=target,headroom=headroom,step=step,max_increase=max_raise,opportunity_mode=opportunity,opportunity_ratio=opportunity_ratio)
    rows.append({'MASTER':r['master_name'],'SHEET':r['sheet'],'BARIS':r['source_row'],'MARKETPLACE':r['marketplace'],'SKU':r['sku'],'NAMA PRODUK':r['name'],
                 'HARGA JUAL (RP)':r['price'],'HPP (RP)':r['hpp'],'KOMISI AWAL (%)':r['old_pct'],'BIAYA AFF AWAL (RP)':r['old_fee'],'MARGIN MASTER (RP)':r['old_margin'],
                 'MARGIN SEBELUM AFF (RP)':result['margin_before'],'BATAS AMAN (%)':result['safe_cap'],'KOMISI USULAN (%)':result['proposed_pct'],
                 'PERUBAHAN KOMISI (POIN %)':(result['proposed_pct']-r['old_pct'] if result['proposed_pct'] is not None and r['old_pct'] is not None else None),'BIAYA KOMISI USULAN (RP)':result['proposed_fee'],'SISA MARGIN USULAN (RP)':result['margin_new'],'SELISIH MARGIN - KOMISI (RP)':result['balance'],
                 'RASIO MARGIN / KOMISI':(result['margin_new']/result['proposed_fee'] if result['proposed_fee'] and result['margin_new'] is not None else None),'STATUS':result['status'],'VALIDASI':r['validation'],'ALASAN':result['reason']})
output=pd.DataFrame(rows)
dups=output.duplicated(subset=['MARKETPLACE','SKU'],keep=False)
output['DUPLIKAT LINTAS MASTER']=dups
if dups.any():
    output.loc[dups,'ALASAN']=output.loc[dups,'ALASAN']+' SKU ini muncul beberapa kali pada master/sheet terpilih; kandidat ditahan dari ekspor.'

render_deck(len(chosen),output)

tabs=st.tabs(['Analisis komisi','Kandidat ekspor','Validasi data','Perubahan master','Detail SKU & panduan','Mapping & histori'])
with tabs[0]:
    st.caption('Bandingkan tarif komisi awal, tarif usulan, biaya affiliate, serta sisa margin dalam satu tabel. Arahkan kursor ke judul kolom untuk penjelasan.')
    states=st.multiselect('Filter rekomendasi',sorted(output.STATUS.unique().tolist()))
    shown=output[output.STATUS.isin(states)] if states else output
    display=shown.drop(columns=['DUPLIKAT LINTAS MASTER']).copy()
    for c in MONEY:display[c]=display[c].map(rupiah)
    cfg={c:st.column_config.TextColumn(c,help=GUIDE.get(c,'Periksa sumber dan satuan pada kolom ini.'),width='large' if c in ('NAMA PRODUK','ALASAN') else 'medium') for c in display.columns}
    st.dataframe(display,column_config=cfg,hide_index=True,use_container_width=True,height=480)
    st.download_button('Download analisis lengkap (Excel)',data=prepare_export(output),file_name='INPERTEK_ANALISIS_V52.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.caption('Angka Rupiah ditampilkan dengan pemisah ribuan Indonesia; file Excel tetap menyimpan angka numerik.')
with tabs[1]:
    valid=output[(output.STATUS.isin(['NAIK','TURUNKAN'])) & (output.VALIDASI=='OK') & (~output['DUPLIKAT LINTAS MASTER']) & (output['KOMISI USULAN (%)']>0)].copy()
    st.metric('Kandidat perubahan yang layak DIREVIEW',len(valid))
    st.warning('Ini daftar REVIEW internal, belum format upload resmi Shopee/Tokopedia. Tidak ada keputusan otomatis yang dikirim ke marketplace.')
    view_candidate=valid[['MASTER','SHEET','BARIS','MARKETPLACE','SKU','NAMA PRODUK','KOMISI AWAL (%)','KOMISI USULAN (%)','BIAYA KOMISI USULAN (RP)','SISA MARGIN USULAN (RP)','STATUS','ALASAN']].copy()
    for col in ['BIAYA KOMISI USULAN (RP)','SISA MARGIN USULAN (RP)']:view_candidate[col]=view_candidate[col].map(rupiah)
    st.dataframe(view_candidate,use_container_width=True,hide_index=True)
    st.download_button('Download kandidat perubahan (Excel)',prepare_export(output),file_name='INPERTEK_KANDIDAT_REVIEW_V52.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.caption('Gunakan sheet KANDIDAT_REVIEW dalam Excel. SKU ambigu lintas master ditahan.')
with tabs[2]:
    st.subheader('Validasi data dan SKU ambigu')
    st.warning('Produk dengan data bermasalah ditahan dari kandidat. Format komisi dihitung hanya jika validasi OK dan tidak ada duplikat.')
    problems=output[(output.VALIDASI!='OK') | (output['DUPLIKAT LINTAS MASTER'])].copy()
    st.metric('Baris bermasalah / ambigu dalam pilihan ini',len(problems))
    if not problems.empty:
        st.dataframe(problems[['MASTER','SHEET','BARIS','MARKETPLACE','SKU','VALIDASI','DUPLIKAT LINTAS MASTER','ALASAN']],hide_index=True,use_container_width=True)
        st.download_button('Download validasi (CSV)',problems.to_csv(index=False).encode('utf-8-sig'),'INPERTEK_VALIDASI_V52.csv','text/csv')
    else:st.success('Tidak ada masalah pada data yang sedang dipilih.')
with tabs[3]:
    st.subheader('Perubahan antar-snapshot MASTER YANG SAMA')
    st.caption('Upload hanya satu master terbaru: perbandingan hanya pada master tersebut. Master lain tidak ikut ditimpa.')
    if change.empty:st.info('Belum ada perubahan yang tercatat pada master yang dipilih.')
    else:
        cshow=change.copy()
        cshow['MASTER']=cshow['master_id'].map(dict(zip(masters.master_id,masters.source_name)))
        st.dataframe(cshow[['MASTER','sheet','marketplace','sku','field','old_value','new_value','kind','run_id']],hide_index=True,use_container_width=True)
        st.download_button('Download riwayat perubahan (CSV)',change.to_csv(index=False).encode('utf-8-sig'),'INPERTEK_PERUBAHAN_V52.csv','text/csv')
with tabs[4]:
    st.subheader('Detail perhitungan untuk satu SKU')
    st.caption('Buka satu produk untuk melihat sumber data, alasan, dan alur hitung usulan komisi secara transparan.')
    if not output.empty:
        # Use a compact index selector and preserve source sheet and original row in its label.
        opts=output.index.tolist()
        selected_i=st.selectbox('Pilih SKU / marketplace / master',opts,format_func=lambda i:f"{output.loc[i,'MASTER']} | {output.loc[i,'SKU']} | {output.loc[i,'MARKETPLACE']} | {output.loc[i,'SHEET']} · baris {output.loc[i,'BARIS']}")
        r=output.loc[selected_i]
        a,b,c=st.columns(3)
        a.metric('Komisi awal',f"{r['KOMISI AWAL (%)']:.2f}%" if pd.notna(r['KOMISI AWAL (%)']) else '-')
        b.metric('Komisi usulan',f"{r['KOMISI USULAN (%)']:.0f}%" if pd.notna(r['KOMISI USULAN (%)']) else '-')
        c.metric('Status',r['STATUS'])
        st.write('**Sumber:**',r['MASTER'],'→',r['SHEET'],'→ Baris',r['BARIS'],'→',r['MARKETPLACE'])
        st.write('**Alasan rekomendasi:**',r['ALASAN'])
        if pd.notna(r['MARGIN SEBELUM AFF (RP)']):
            st.write('**1. Margin sebelum affiliate:**',rupiah(r['MARGIN MASTER (RP)']),' + ',rupiah(r['BIAYA AFF AWAL (RP)']),' = ',rupiah(r['MARGIN SEBELUM AFF (RP)']))
            if pd.notna(r['KOMISI USULAN (%)']):
                st.write('**2. Biaya komisi baru:**',rupiah(r['HARGA JUAL (RP)']),f" × {r['KOMISI USULAN (%)']:.0f}% = ",rupiah(r['BIAYA KOMISI USULAN (RP)']))
                st.write('**3. Margin baru:**',rupiah(r['MARGIN SEBELUM AFF (RP)']),' − ',rupiah(r['BIAYA KOMISI USULAN (RP)']),' = ',rupiah(r['SISA MARGIN USULAN (RP)']))
                if r['STATUS']=='PELUANG KHUSUS - REVIEW':st.warning('Peluang khusus: margin positif tetapi lebih kecil daripada biaya komisi. Tidak masuk kandidat otomatis.')
                elif r['SELISIH MARGIN - KOMISI (RP)']>=-0.01:st.success('Sisa margin minimal sebesar biaya komisi baru.')
    st.divider()
    st.subheader('Kamus kolom')
    st.dataframe(pd.DataFrame([{'KOLOM':k,'ARTI / CARA MEMBACA':v} for k,v in GUIDE.items()]),hide_index=True,use_container_width=True)
    st.markdown('**Status:** NAIK = kapasitas margin mendukung kenaikan; TURUNKAN = komisi awal terlalu besar; TETAP = belum perlu berubah; REVIEW NONAKTIF = kapasitas margin belum cukup; VALIDASI DATA = sumber belum aman dihitung; PELUANG KHUSUS - REVIEW = hanya simulasi manual.')
    st.latex(r'M_{sebelum}=M_{master}+C_{awal}')
    st.latex(r'C_{baru}=Harga\ Jual\times\frac{Persentase\ Komisi}{100}')
    st.latex(r'M_{baru}=M_{sebelum}-C_{baru}')
    st.latex(r'M_{baru}\geq C_{baru}\quad(standar)')
    st.caption('Tarif komisi rekomendasi selalu kelipatan 1%. Basis komisi diasumsikan harga jual dan wajib diperiksa terhadap ketentuan marketplace aktual.')
with tabs[5]:
    st.subheader('Mapping per master dan sheet')
    mapp=mapping[mapping.source_name.isin(selected)]
    st.dataframe(mapp[['source_name','sheet','status','reason','map_json']],hide_index=True,use_container_width=True)
    st.subheader('Snapshot terakhir per master')
    st.dataframe(masters[masters.source_name.isin(selected)],hide_index=True,use_container_width=True)
    st.subheader('Riwayat sinkronisasi yang masih tersimpan')
    st.dataframe(run_stats[run_stats.source_name.isin(selected)] if 'source_name' in run_stats.columns else run_stats,hide_index=True,use_container_width=True)
    st.caption('SQLite pada Streamlit Cloud bersifat sementara. Unduh Excel/CSV untuk menyimpan hasil penting.')
