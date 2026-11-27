"""INPERTEK multi-master read-only commission optimizer. No Supabase required."""
import html, os, tempfile
from datetime import datetime, timedelta, timezone
from io import BytesIO
import pandas as pd
import streamlit as st
import xlsxwriter
from functools import partial
from engine import db_connect, sync_many, analyze, MODES, is_postgres, storage_label, delete_run, delete_master, delete_all

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
.store { display:flex; align-items:center; gap:8px; font-size:12px; font-weight:600; padding:6px 10px; border-radius:99px; margin:.2rem 0 0; width:fit-content; }
.store i { width:7px; height:7px; border-radius:50%; display:inline-block; }
.store.ok { color:#7CF2C4; background:rgba(124,242,196,.10); border:1px solid rgba(124,242,196,.30); }
.store.ok i { background:#7CF2C4; box-shadow:0 0 8px #7CF2C4; }
.store.temp { color:#FFB547; background:rgba(255,181,71,.10); border:1px solid rgba(255,181,71,.30); }
.store.temp i { background:#FFB547; }
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
[class*="st-key-del_btn"] button { background:#FF6B8B!important; color:#0C0F23!important; box-shadow:0 8px 26px -12px rgba(255,107,139,.9)!important; }
[class*="st-key-del_btn"] button:disabled { background:var(--rim-soft)!important; color:var(--dim)!important; box-shadow:none!important; }
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

def resolve_db():
    """DATABASE_URL (Streamlit Secrets / env) -> Postgres permanen; jika kosong -> SQLite sementara."""
    url=None
    try:url=st.secrets.get('DATABASE_URL')
    except Exception:url=None
    url=url or os.environ.get('DATABASE_URL')
    return url.strip() if url else os.path.join(os.path.dirname(__file__),'optimizer_v5.sqlite')
DB=resolve_db();PERSISTENT=is_postgres(DB)
WIB=timezone(timedelta(hours=7))

def wib(ts):
    try:return datetime.fromisoformat(str(ts)).astimezone(WIB).strftime('%d %b %Y, %H:%M WIB')
    except Exception:return str(ts)

# Snapshot (run) tidak pernah berubah setelah disimpan, jadi aman di-cache berdasarkan run_id.
@st.cache_data(show_spinner=False,max_entries=16)
def load_items(db,ids):
    con=db_connect(db)
    try:
        q=','.join('?' for _ in ids)
        return con.query_df(f'SELECT i.*, m.source_name AS master_name FROM items i JOIN masters m ON m.master_id=i.master_id WHERE i.run_id IN ({q})',list(ids))
    finally:con.close()

@st.cache_data(show_spinner=False,max_entries=16)
def load_history(db,master_ids,latest_run):
    con=db_connect(db)
    try:
        q=','.join('?' for _ in master_ids)
        runs=con.query_df(f'SELECT r.*, m.source_name AS master_name FROM runs r JOIN masters m ON m.master_id=r.master_id WHERE r.master_id IN ({q}) ORDER BY r.run_id DESC',list(master_ids))
        changes=con.query_df(f'SELECT * FROM changes WHERE master_id IN ({q}) ORDER BY run_id DESC',list(master_ids))
        return runs,changes
    finally:con.close()

MONEY=['HARGA JUAL (RP)','HPP (RP)','BIAYA AFF AWAL (RP)','MARGIN MASTER (RP)','MARGIN SEBELUM AFF (RP)','BIAYA KOMISI USULAN (RP)','SISA MARGIN USULAN (RP)','SELISIH MARGIN - KOMISI (RP)']
GUIDE={
'MASTER':'File master yang menjadi asal produk.','SHEET':'Worksheet hijau tempat data berasal.','BARIS':'Nomor baris pada worksheet asli.',
'MARKETPLACE':'TOKOPEDIA mencakup TikTok Shop sesuai keputusan internal.','SKU':'Kode produk di master.','NAMA PRODUK':'Nama asli pada master.',
'HARGA JUAL (RP)':'Harga jual pada blok marketplace.', 'HPP (RP)':'Harga pokok dari master.',
'KOMISI AWAL (%)':'Tarif affiliate saat snapshot master dibaca.', 'BIAYA AFF AWAL (RP)':'Biaya affiliate yang sudah tercermin pada margin master.',
'MARGIN MASTER (RP)':'Margin setelah biaya affiliate lama menurut master.', 'MARGIN SEBELUM AFF (RP)':'Margin master + biaya affiliate lama. Biaya affiliate hanya ditambahkan kembali sekali.',
'BATAS KOMISI (%)':'Komisi tertinggi yang membuat sisa margin tetap ≥ jarak aman × komisi (1,25x / 1,5x / 2x sesuai mode). Usulan = batas ini dibulatkan ke bawah.', 'PORSI AFFILIATE DARI MARGIN (%)':'Bagian margin sebelum affiliate yang dibayarkan ke affiliate dengan komisi usulan.', 'KOMISI USULAN (%)':'Persentase baru yang direkomendasikan.',
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
    """Bangun file Excel review. Dipanggil hanya saat tombol download diklik."""
    file=BytesIO()
    wb=xlsxwriter.Workbook(file,{'in_memory':True,'nan_inf_to_errors':True,'strings_to_numbers':False,'strings_to_formulas':False,'strings_to_urls':False})
    head=wb.add_format({'bold':True,'font_color':'#FFFFFF','bg_color':'#17365D','font_name':'Calibri','font_size':10,'valign':'vcenter'})
    money=wb.add_format({'num_format':'"Rp" #,##0;[Red]("Rp" #,##0);"Rp" 0'})
    pct=wb.add_format({'num_format':'0.00"%"'})
    info=wb.add_worksheet('PANDUAN')
    guide=[
        ['INPERTEK AFFILIATE OPTIMIZER V5.3 - REVIEW INTERNAL'],
        ['SUMBER','Snapshot master diunggah; master asli tidak diubah'],
        ['PERINGATAN','BUKAN TEMPLATE UPLOAD RESMI MARKETPLACE'],
        ['ALUR','Validasi listing & ketentuan marketplace, review, lalu susun file impor resmi'],
        ['METODE','Asumsi dasar komisi = harga jual. Margin sebelum affiliate = margin master + biaya affiliate awal'],
        ['BATAS','Sisa margin >= jarak aman x biaya komisi baru (1,25x / 1,5x / 2x sesuai mode), >= margin minimum, dan tidak negatif'],
        ['KEAMANAN','Data master dan database jangan diunggah ke GitHub. Unduh hasil saat analisis selesai'],
        ['DUPLIKAT','SKU berulang antar master / sheet ditahan dari daftar kandidat'],
    ]
    for i,row in enumerate(guide):
        info.write_row(i,0,row,head if i==0 else None)
    info.set_column(0,0,23);info.set_column(1,1,115);info.set_row(0,27);info.hide_gridlines(2)
    cand=df[(df['STATUS'].isin(['NAIK','TURUNKAN'])) & (df['VALIDASI']=='OK') & (~df['DUPLIKAT LINTAS MASTER']) & (df['KOMISI USULAN (%)']>0)]
    for title,sub in [('REKOMENDASI',df),('KANDIDAT_REVIEW',cand)]:
        ws=wb.add_worksheet(title)
        fields=[c for c in sub.columns if c!='DUPLIKAT LINTAS MASTER']
        ws.write_row(0,0,fields,head);ws.set_row(0,27)
        for j,col in enumerate(fields):
            fmt=money if col in MONEY else (pct if col.endswith('(%)') else None)
            ws.set_column(j,j,55 if col in ('NAMA PRODUK','ALASAN') else 24,fmt)
        data=sub[fields].astype(object).where(sub[fields].notna(),None)
        for i,vals in enumerate(data.itertuples(index=False,name=None),1):
            ws.write_row(i,0,vals)
        ws.freeze_panes(1,5);ws.autofilter(0,0,max(len(sub),1),len(fields)-1);ws.hide_gridlines(2)
    wb.close();return file.getvalue()

def csv_bytes(df):
    return df.to_csv(index=False).encode('utf-8-sig')

PAGE_SIZE=500
def paged_table(df,key,money=(),**kwargs):
    """Tampilkan tabel per halaman agar halaman ringan; file download tetap berisi semua baris."""
    n=len(df)
    if n>PAGE_SIZE:
        pages=(n-1)//PAGE_SIZE+1
        c1,c2=st.columns([1,4],vertical_alignment='bottom')
        p=c1.number_input('Halaman',min_value=1,max_value=pages,value=1,step=1,key=f'pg_{key}_{n}')
        c2.caption(f'Menampilkan baris {(p-1)*PAGE_SIZE+1:,}–{min(p*PAGE_SIZE,n):,} dari {n:,} ({pages:,} halaman). File download berisi semua baris.'.replace(',','.'))
        df=df.iloc[(p-1)*PAGE_SIZE:p*PAGE_SIZE]
    if money:
        df=df.copy()
        for c in money:
            if c in df.columns:df[c]=df[c].map(rupiah)
    st.dataframe(df,**kwargs)


st.markdown('''<header class="hero">
<h1>Affiliate Commission Optimizer</h1>
<p>Hitung ulang komisi affiliate per SKU di Shopee dan Tokopedia tanpa menggerus margin. Hanya sheet hijau dengan mapping terverifikasi yang diproses, master asli tidak pernah diubah, dan file ekspor adalah bahan review, bukan template impor marketplace.</p>
<div class="hero-line"></div>
</header>''',unsafe_allow_html=True)
@st.cache_data(show_spinner=False,max_entries=8)
def analyze_all(db,ids,mode,min_profit,cap,target,step,max_raise):
    """Analisis per baris untuk seluruh snapshot terpilih. Di-cache: hanya dihitung ulang bila data atau kebijakan berubah."""
    base=load_items(db,ids)
    rows=[]
    for r in base.to_dict('records'):
        result=analyze(r,mode=mode,min_profit=min_profit,max_pct=cap,target_pct=target,step=step,max_increase=max_raise)
        rows.append({'MASTER':r['master_name'],'SHEET':r['sheet'],'BARIS':r['source_row'],'MARKETPLACE':r['marketplace'],'SKU':r['sku'],'NAMA PRODUK':r['name'],
                     'HARGA JUAL (RP)':r['price'],'HPP (RP)':r['hpp'],'KOMISI AWAL (%)':r['old_pct'],'BIAYA AFF AWAL (RP)':r['old_fee'],'MARGIN MASTER (RP)':r['old_margin'],
                     'MARGIN SEBELUM AFF (RP)':result['margin_before'],'BATAS KOMISI (%)':result['hard_cap'],'KOMISI USULAN (%)':result['proposed_pct'],
                     'PERUBAHAN KOMISI (POIN %)':(result['proposed_pct']-r['old_pct'] if result['proposed_pct'] is not None and r['old_pct'] is not None else None),'BIAYA KOMISI USULAN (RP)':result['proposed_fee'],'SISA MARGIN USULAN (RP)':result['margin_new'],'SELISIH MARGIN - KOMISI (RP)':result['balance'],'PORSI AFFILIATE DARI MARGIN (%)':(result['affiliate_share']*100 if result['affiliate_share'] is not None else None),
                     'RASIO MARGIN / KOMISI':(result['margin_new']/result['proposed_fee'] if result['proposed_fee'] and result['margin_new'] is not None else None),'STATUS':result['status'],'VALIDASI':r['validation'],'ALASAN':result['reason']})
    return pd.DataFrame(rows)

try:db_connect(DB).close()
except Exception:
    if PERSISTENT:st.error('Tidak bisa terhubung ke database permanen. Periksa DATABASE_URL di Settings → Secrets Streamlit, lalu reboot app. Data lama tetap aman di Neon.')
    else:st.error('Database sementara tidak dapat dibuka.')
    st.stop()
if st.session_state.get('flash'):st.success(st.session_state.pop('flash'))

with st.sidebar:
    st.markdown('<div class="brand"><div class="brand-mark"></div><div><div class="brand-name">INPERTEK</div><div class="brand-sub">Affiliate Optimizer v5.3</div></div></div>',unsafe_allow_html=True)
    if PERSISTENT:st.markdown('<div class="store ok"><i></i>Riwayat tersimpan permanen</div>',unsafe_allow_html=True)
    else:st.markdown('<div class="store temp"><i></i>Riwayat sementara: hilang saat app restart</div>',unsafe_allow_html=True)
    st.divider()
    step_header(1,'Upload master')
    uploads=st.file_uploader('Pilih satu atau beberapa master (.xlsx)',type=['xlsx'],accept_multiple_files=True,key=f"uploader_{st.session_state.get('upload_round',0)}",help='Prioritas sheet tab hijau. Tidak boleh upload data sensitif ke GitHub.')
    if st.button('Analisis / sinkronkan master',disabled=not uploads,type='primary',width='stretch'):
        with tempfile.TemporaryDirectory() as d:
            paths=[]
            for u in uploads:
                safe=os.path.basename(u.name)
                p=os.path.join(d,safe)
                with open(p,'wb') as f:
                    for chunk in iter(lambda:u.read(1024*1024),b''):f.write(chunk)
                u.seek(0);paths.append(p)
            try:
                results=[];bar=st.progress(0.0,text='Menyiapkan...')
                for k,p in enumerate(paths):
                    bar.progress(k/len(paths),text=f'Memproses {os.path.basename(p)} ({k+1}/{len(paths)})')
                    results+=sync_many([p],DB)
                bar.empty()
                load_history.clear();analyze_all.clear()
                st.session_state['upload_round']=st.session_state.get('upload_round',0)+1
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
    st.caption('Aturan pokok: setelah komisi baru, sisa margin selalu lebih besar dari komisi dengan jarak aman. Margin master dibaca sebagai margin bersih (sudah dikurangi komisi awal).')
    mode=st.selectbox('Agresivitas',['AGRESIF','SEIMBANG','KONSERVATIF'],index=0,help='Jarak aman sisa margin terhadap komisi. Agresif: sisa margin ≥ 1,25x komisi. Seimbang: ≥ 1,5x. Konservatif: ≥ 2x. Komisi tiap produk disesuaikan naik/turun ke batas ini.')
    ratio=MODES[mode]
    st.caption(f'Sisa margin minimal {ratio:g}x komisi, jadi komisi paling besar {1/(1+ratio):.0%} dari margin sebelum affiliate.')
    min_profit=st.number_input('Margin minimum (Rp)',min_value=0.0,value=0.0,step=1000.0)
    cap=st.number_input('Batas komisi simulasi (%)',min_value=.5,max_value=100.0,value=30.0,step=.5)
    target=st.number_input('Target maksimum komisi (%)',min_value=.5,max_value=100.0,value=30.0,step=.5)
    st.info('Kelipatan komisi usulan ditetapkan 1 poin persentase: 3% → 4% → 5%. Tidak ada opsi 0,5%.')
    step=1.0
    limit_raise=st.toggle('Batasi kenaikan per evaluasi (opsional)',value=False)
    max_raise=st.number_input('Batas kenaikan (poin persentase)',min_value=1,max_value=100,value=5,step=1,disabled=not limit_raise) if limit_raise else None
    st.caption('Batas dan kelipatan adalah ASUMSI simulasi. Periksa aturan resmi dan dasar pengenaan komisi tiap marketplace sebelum upload.')

con=db_connect(DB)
masters=con.query_df('SELECT master_id,source_name,last_run,updated_at FROM masters ORDER BY source_name')
if masters.empty:
    st.warning('Belum ada snapshot. Upload master Excel melalui sidebar lalu klik Analisis / sinkronkan master.')
    con.close();st.stop()
selected=st.multiselect('Pilih master untuk analisis',masters.source_name.tolist(),default=masters.source_name.tolist())
chosen=masters[masters.source_name.isin(selected)]
if chosen.empty:st.warning('Pilih minimal satu master.');con.close();st.stop()
ids=chosen.last_run.astype(int).tolist()
mapping=con.query_df('SELECT sm.master_id, m.source_name, sm.sheet, sm.status, sm.reason, sm.map_json FROM sheet_maps sm JOIN masters m ON m.master_id=sm.master_id')
con.close()
with st.spinner('Menghitung rekomendasi...'):
    full=analyze_all(DB,tuple(sorted(ids)),mode,float(min_profit),float(cap),float(target),step,max_raise)
    runs_all,change=load_history(DB,tuple(sorted(chosen.master_id)),int(max(ids)))
with st.sidebar:
    step_header(3,'Filter data')
    market=st.selectbox('Marketplace',['SEMUA','SHOPEE','TOKOPEDIA'])
    sheet_filter=st.selectbox('Worksheet sumber',['SEMUA']+sorted(full.SHEET.dropna().unique().tolist()))
    keyword=st.text_input('Cari SKU / nama produk')
    validity=st.selectbox('Validasi',['SEMUA','HANYA OK','BERMASALAH'])
output=full
if market!='SEMUA':output=output[output.MARKETPLACE==market]
if sheet_filter!='SEMUA':output=output[output.SHEET==sheet_filter]
if keyword:output=output[output.SKU.fillna('').astype(str).str.contains(keyword,case=False,regex=False)|output['NAMA PRODUK'].fillna('').astype(str).str.contains(keyword,case=False,regex=False)]
if validity=='HANYA OK':output=output[output.VALIDASI=='OK']
if validity=='BERMASALAH':output=output[output.VALIDASI!='OK']
if output.empty:st.info('Tidak ada produk sesuai filter.');st.stop()
output=output.reset_index(drop=True)

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
    display=shown.drop(columns=['DUPLIKAT LINTAS MASTER'])
    cfg={c:st.column_config.TextColumn(c,help=GUIDE.get(c,'Periksa sumber dan satuan pada kolom ini.'),width='large' if c in ('NAMA PRODUK','ALASAN') else 'medium') for c in display.columns}
    paged_table(display,'analisis',money=MONEY,column_config=cfg,hide_index=True,width='stretch',height=480)
    st.download_button('Download analisis lengkap (Excel)',data=partial(prepare_export,output.copy()),file_name='INPERTEK_ANALISIS_V52.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.caption('Angka Rupiah ditampilkan dengan pemisah ribuan Indonesia; file Excel tetap menyimpan angka numerik.')
with tabs[1]:
    valid=output[(output.STATUS.isin(['NAIK','TURUNKAN'])) & (output.VALIDASI=='OK') & (~output['DUPLIKAT LINTAS MASTER']) & (output['KOMISI USULAN (%)']>0)].copy()
    st.metric('Kandidat perubahan yang layak DIREVIEW',len(valid))
    st.warning('Ini daftar REVIEW internal, belum format upload resmi Shopee/Tokopedia. Tidak ada keputusan otomatis yang dikirim ke marketplace.')
    view_candidate=valid[['MASTER','SHEET','BARIS','MARKETPLACE','SKU','NAMA PRODUK','KOMISI AWAL (%)','KOMISI USULAN (%)','BIAYA KOMISI USULAN (RP)','SISA MARGIN USULAN (RP)','STATUS','ALASAN']]
    paged_table(view_candidate,'kandidat',money=['BIAYA KOMISI USULAN (RP)','SISA MARGIN USULAN (RP)'],width='stretch',hide_index=True)
    st.download_button('Download kandidat perubahan (Excel)',partial(prepare_export,output.copy()),file_name='INPERTEK_KANDIDAT_REVIEW_V52.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.caption('Gunakan sheet KANDIDAT_REVIEW dalam Excel. SKU ambigu lintas master ditahan.')
with tabs[2]:
    st.subheader('Validasi data dan SKU ambigu')
    st.warning('Produk dengan data bermasalah ditahan dari kandidat. Format komisi dihitung hanya jika validasi OK dan tidak ada duplikat.')
    problems=output[(output.VALIDASI!='OK') | (output['DUPLIKAT LINTAS MASTER'])].copy()
    st.metric('Baris bermasalah / ambigu dalam pilihan ini',len(problems))
    if not problems.empty:
        paged_table(problems[['MASTER','SHEET','BARIS','MARKETPLACE','SKU','VALIDASI','DUPLIKAT LINTAS MASTER','ALASAN']],'validasi',hide_index=True,width='stretch')
        st.download_button('Download validasi (CSV)',partial(csv_bytes,problems.copy()),'INPERTEK_VALIDASI_V52.csv','text/csv')
    else:st.success('Tidak ada masalah pada data yang sedang dipilih.')
with tabs[3]:
    st.subheader('Riwayat perubahan master')
    st.caption('Setiap upload master dibandingkan dengan upload sebelumnya dari master yang sama. Riwayat ini '+('tersimpan permanen.' if PERSISTENT else 'masih sementara: hilang saat app restart. Hubungkan Neon agar permanen.'))
    FIELD={'price':'Harga jual','hpp':'HPP','old_pct':'Komisi (%)','old_fee':'Biaya affiliate','old_margin':'Margin bersih','ROW':'Produk'}
    KIND={'MODIFIED':'Berubah','ADDED':'SKU baru','REMOVED':'SKU hilang'}
    def fmt_val(field,v):
        if field=='ROW':return '-'
        try:x=float(v)
        except (TypeError,ValueError):return '-' if v in (None,'None') else str(v)
        return f'{x:.2f}%'.replace('.',',') if field=='old_pct' else rupiah(x)
    uploads_with_change=runs_all[runs_all.run_id.isin(change.run_id.unique())] if not change.empty else runs_all.iloc[0:0]
    if change.empty:
        st.info('Belum ada perubahan. Riwayat muncul setelah master yang sama di-upload lagi dengan isi berbeda.')
    else:
        opts=['Semua upload']+[f"{r.master_name}  |  {wib(r.sync_time)}" for r in uploads_with_change.itertuples()]
        run_by_label={f"{r.master_name}  |  {wib(r.sync_time)}":r.run_id for r in uploads_with_change.itertuples()}
        pick=st.selectbox('Upload',opts,index=1 if len(opts)>1 else 0,help='Pilih satu upload untuk melihat apa saja yang berubah dibanding upload sebelumnya.')
        cv=change if pick=='Semua upload' else change[change.run_id==run_by_label[pick]]
        c1,c2,c3,c4=st.columns(4)
        c1.metric('Komisi berubah',int((cv.field=='old_pct').sum()))
        c2.metric('Harga berubah',int((cv.field=='price').sum()))
        c3.metric('SKU baru',int((cv.kind=='ADDED').sum()))
        c4.metric('SKU hilang',int((cv.kind=='REMOVED').sum()))
        kinds=st.multiselect('Jenis data',[FIELD[k] for k in FIELD if k in set(cv.field)],help='Kosongkan untuk menampilkan semua.')
        times=dict(zip(runs_all.run_id,runs_all.sync_time))
        names=dict(zip(masters.master_id,masters.source_name))
        view=pd.DataFrame({'MASTER':cv.master_id.map(names),'WAKTU UPLOAD':cv.run_id.map(times).map(wib),'SHEET':cv.sheet,'MARKETPLACE':cv.marketplace,'SKU':cv.sku,
                           'DATA':cv.field.map(FIELD).fillna(cv.field),'NILAI LAMA':[fmt_val(f,v) for f,v in zip(cv.field,cv.old_value)],
                           'NILAI BARU':[fmt_val(f,v) for f,v in zip(cv.field,cv.new_value)],'JENIS':cv.kind.map(KIND).fillna(cv.kind)})
        if kinds:view=view[view.DATA.isin(kinds)]
        paged_table(view,'perubahan',hide_index=True,width='stretch',height=420)
        st.download_button('Download riwayat perubahan (CSV)',partial(csv_bytes,view.copy()),'INPERTEK_PERUBAHAN_V53.csv','text/csv')
with tabs[4]:
    st.subheader('Detail perhitungan untuk satu SKU')
    st.caption('Buka satu produk untuk melihat sumber data, alasan, dan alur hitung usulan komisi secara transparan.')
    if not output.empty:
        # Use a compact index selector and preserve source sheet and original row in its label.
        qsku=st.text_input('Cari SKU atau nama produk',key='detail_q',placeholder='Ketik sebagian SKU atau nama')
        pool=output
        if qsku:
            pool=output[output.SKU.fillna('').astype(str).str.contains(qsku,case=False,regex=False)|output['NAMA PRODUK'].fillna('').astype(str).str.contains(qsku,case=False,regex=False)]
            if pool.empty:st.warning('Tidak ada SKU yang cocok; menampilkan daftar awal.');pool=output
        opts=pool.index.tolist()[:300]
        if len(pool)>300:st.caption(f'Menampilkan 300 dari {len(pool):,} SKU. Ketik di kolom pencarian untuk mempersempit.'.replace(',','.'))
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
                elif r['SELISIH MARGIN - KOMISI (RP)']>=-0.01 and pd.notna(r['RASIO MARGIN / KOMISI']):
                    st.success(f"Sisa margin {r['RASIO MARGIN / KOMISI']:.2f}x biaya komisi baru. Affiliate menerima {r['PORSI AFFILIATE DARI MARGIN (%)']:.0f}% dari margin sebelum affiliate.")
    st.divider()
    st.subheader('Kamus kolom')
    st.dataframe(pd.DataFrame([{'KOLOM':k,'ARTI / CARA MEMBACA':v} for k,v in GUIDE.items()]),hide_index=True,width='stretch')
    st.markdown('**Status:** NAIK = kapasitas margin mendukung kenaikan; TURUNKAN = komisi awal terlalu besar; TETAP = komisi awal sudah pas dengan margin; REVIEW NONAKTIF = kapasitas margin belum cukup; VALIDASI DATA = sumber belum aman dihitung; PELUANG KHUSUS - REVIEW = hanya simulasi manual.')
    st.latex(r'M_{sebelum}=M_{master}+C_{awal}')
    st.latex(r'C_{baru}=Harga\ Jual\times\frac{Persentase\ Komisi}{100}')
    st.latex(r'M_{baru}=M_{sebelum}-C_{baru}')
    st.latex(r'M_{baru}\geq k\cdot C_{baru}\iff C_{baru}\leq\frac{M_{sebelum}}{1+k}')
    st.latex(r'k=1{,}25\ (agresif),\quad 1{,}5\ (seimbang),\quad 2\ (konservatif)')
    st.caption('Setiap produk disesuaikan ke batas komisinya (kelipatan 1%, dibulatkan ke bawah): naik bila di bawah, turun bila di atas, tetap bila sudah pas, nonaktif bila 1% pun tidak muat.')
    st.caption('Tarif komisi rekomendasi selalu kelipatan 1%. Basis komisi diasumsikan harga jual dan wajib diperiksa terhadap ketentuan marketplace aktual.')
with tabs[5]:
    st.subheader('Mapping per master dan sheet')
    mapp=mapping[mapping.source_name.isin(selected)]
    st.dataframe(mapp[['source_name','sheet','status','reason','map_json']],hide_index=True,width='stretch')
    st.subheader('Snapshot terakhir per master')
    st.dataframe(masters[masters.source_name.isin(selected)],hide_index=True,width='stretch')
    st.subheader('Semua upload tersimpan')
    hist=runs_all.assign(WAKTU=runs_all.sync_time.map(wib))[['master_name','WAKTU','sheet_count','source_rows','records','issues','run_id']]
    hist.columns=['MASTER','WAKTU UPLOAD','SHEET','BARIS SUMBER','BARIS MARKETPLACE','CATATAN VALIDASI','ID SNAPSHOT']
    st.dataframe(hist,hide_index=True,width='stretch')
    st.caption(f'Penyimpanan: {storage_label(DB)}. '+('Riwayat tidak hilang saat app restart atau server berhenti.' if PERSISTENT else 'Riwayat hilang saat app restart. Isi DATABASE_URL di Secrets agar permanen.'))
    st.divider()
    with st.expander('Hapus data tersimpan'):
        st.caption('Penghapusan bersifat permanen dan tidak bisa dibatalkan. Download Excel/CSV dulu bila datanya masih diperlukan.')
        what=st.radio('Yang mau dihapus',['Satu upload','Satu master beserta riwayatnya','Semua data'],horizontal=True,key='del_what')
        dcon=db_connect(DB)
        try:all_runs=dcon.query_df('SELECT r.run_id,r.master_id,r.sync_time,r.records,m.source_name FROM runs r JOIN masters m ON m.master_id=r.master_id ORDER BY r.run_id DESC')
        finally:dcon.close()
        target=None
        if what=='Satu upload':
            labels={f"{r.source_name}  |  {wib(r.sync_time)}  |  {r.records} baris  |  ID {r.run_id}":int(r.run_id) for r in all_runs.itertuples()}
            pick=st.selectbox('Pilih upload',list(labels),key='del_run_pick')
            target=labels.get(pick);word='HAPUS'
            last_ids=set(masters.last_run.astype(int))
            if target in last_ids:
                note='Ini upload terbaru master tersebut. Setelah dihapus, analisis kembali memakai upload sebelumnya. Kalau tidak ada upload lain, masternya ikut terhapus.'
            else:
                note='Upload lama ini dihapus. Riwayat perubahan upload sesudahnya dihitung ulang terhadap upload sebelumnya.'
        elif what=='Satu master beserta riwayatnya':
            name_to_id=dict(zip(masters.source_name,masters.master_id))
            pick=st.selectbox('Pilih master',list(name_to_id),key='del_master_pick')
            target=name_to_id.get(pick);word=pick
            n=int((all_runs.master_id==target).sum())
            note=f'Semua {n} upload dan seluruh riwayat perubahan {pick} akan dihapus.'
        else:
            word='HAPUS SEMUA'
            note=f'Semua {len(masters)} master, {len(all_runs)} upload, dan seluruh riwayat perubahan akan dihapus.'
        st.warning(note)
        typed=st.text_input(f'Ketik  {word}  untuk konfirmasi',key=f'del_confirm_{what}',placeholder=word)
        if st.button('Hapus permanen',type='primary',disabled=typed.strip()!=word,key='del_btn'):
            try:
                if what=='Satu upload':
                    res=delete_run(DB,target)
                    msg='Upload dihapus.' if res['status']=='RUN_DELETED' else 'Upload dihapus. Itu upload satu-satunya, jadi masternya ikut terhapus.'
                elif what=='Satu master beserta riwayatnya':
                    delete_master(DB,target);msg=f'Master {word} beserta semua riwayatnya dihapus.'
                else:
                    delete_all(DB);msg='Semua data tersimpan dihapus.'
            except Exception as e:
                st.error(f'Penghapusan gagal, tidak ada data yang berubah. Detail: {e}')
            else:
                load_items.clear();load_history.clear();analyze_all.clear()
                for k in [k for k in st.session_state if str(k).startswith('del_')]:del st.session_state[k]
                st.session_state['flash']=msg;st.rerun()
