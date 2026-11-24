"""INPERTEK multi-master read-only commission optimizer. No Supabase required."""
import os, tempfile
from io import BytesIO
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from engine import db_connect, sync_many, analyze

st.set_page_config(page_title='INPERTEK | Affiliate Optimizer V5.2',page_icon='📈',layout='wide')

# Application visual shell: premium, restrained dark-light hybrid.
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap');
:root { --ink:#182a42; --muted:#60758b; --border:#dce5ef; --navy:#111e31; --teal:#0f827b; --blue:#2b6cb0; }
html, body, [class*="css"], [data-testid="stApp"] { font-family:'DM Sans',sans-serif; }
[data-testid="stAppViewContainer"] { background:linear-gradient(135deg,#f4f7fb 0%,#fafcff 60%,#f4f8fb 100%); color:var(--ink); }
[data-testid="stMainBlockContainer"] { padding-top:1.7rem; padding-bottom:4rem; max-width:1600px; }
[data-testid="stSidebar"] { background:linear-gradient(175deg,#15263c 0%,#101c2c 100%); }
[data-testid="stSidebar"] * { color:#e8eff7; }
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p { color:#d4e0ef!important; font-weight:600; }
[data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color:white!important; letter-spacing:.01em; }
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {color:#9fb3c8!important;}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {background:#1d324c;border:1px dashed #52718e;border-radius:12px;}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] *, [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] small {color:#e8eff7!important;}
[data-testid="stSidebar"] [data-testid="stSelectbox"] > div > div,
[data-testid="stSidebar"] [data-testid="stTextInput"] input,
[data-testid="stSidebar"] [data-testid="stNumberInput"] input { background:#20354e!important; color:#f4f7fc!important; border-color:#476078!important; }
[data-testid="stSidebar"] [data-testid="stSlider"] [data-baseweb="slider"] { filter:saturate(.75); }
[data-testid="stSidebar"] hr {border-color:#344960;}
[data-testid="stSidebar"] [data-testid="stButton"] button[kind="primary"] {background:#16a39c;border:0;color:white;font-weight:700;}
[data-testid="stSidebar"] [data-testid="stButton"] button[kind="primary"]:hover {background:#128b86;}
.dashboard-hero { position:relative; overflow:hidden; background:linear-gradient(112deg,#172941 0%,#203c59 68%,#244f61 100%); border:1px solid #294461; border-radius:18px; padding:28px 34px; margin:4px 0 23px; box-shadow:0 12px 30px rgba(22,42,68,.09); }
.dashboard-hero::after {content:'';position:absolute;right:-85px;top:-140px;height:340px;width:340px;border:1px solid rgba(170,225,231,.17);border-radius:50%;box-shadow:0 0 0 65px rgba(161,222,226,.035),0 0 0 130px rgba(161,222,226,.025);pointer-events:none;}
.hero-eyebrow { font-size:11px;font-weight:800;letter-spacing:.16em;text-transform:uppercase;color:#81d7cf; margin-bottom:7px; }
.dashboard-hero h1 { margin:0!important; padding:0!important; color:#f7fbff!important; font-size:clamp(22px,2.4vw,35px)!important; line-height:1.25; letter-spacing:-.035em; font-weight:800; }
.dashboard-hero p { color:#bdcedf!important; font-size:13px; margin:9px 0 0!important; }
.hero-tags {display:flex;gap:9px;flex-wrap:wrap;margin-top:16px;}
.hero-tag {display:inline-flex;align-items:center;border:1px solid #45627a;background:#ffffff09;border-radius:7px;color:#dceaf1;font-size:11px;padding:5px 10px;font-weight:650;}
.section-kicker {font-size:11px;font-weight:800;color:#637d91;letter-spacing:.12em;text-transform:uppercase;margin:12px 0 0;}
[data-testid="stMetric"] {background:white;border:1px solid #e2e9f1;border-radius:14px;padding:18px 19px;min-height:112px;box-shadow:0 5px 17px rgba(32,64,100,.035);}
[data-testid="stMetricLabel"] {font-size:12px;color:#60758b;font-weight:650;}
[data-testid="stMetricValue"] {font-size:clamp(20px,2vw,30px);color:#1a344f;font-weight:800;letter-spacing:-.03em;}
[data-testid="stTabs"] [data-baseweb="tab-list"] {gap:5px;background:#eaf0f6;border-radius:12px;padding:5px;overflow-x:auto;}
[data-testid="stTabs"] button[data-baseweb="tab"] {border-radius:8px;min-height:42px;font-size:12px;font-weight:700;color:#50657c;padding:0 12px;white-space:nowrap;}
[data-testid="stTabs"] button[aria-selected="true"] {background:#fff;color:#164b70;box-shadow:0 2px 8px rgba(20,45,75,.10);}
[data-testid="stTabs"] [data-baseweb="tab-highlight"], [data-testid="stTabs"] [data-baseweb="tab-border"] {display:none;}
[data-testid="stDataFrame"] {border-radius:13px;overflow:hidden;border:1px solid var(--border);background:white;}
[data-testid="stAlert"] {border-radius:11px;}
[data-testid="stDownloadButton"] button {border-radius:9px;border:1px solid #c5d6e5;background:#fff;color:#204d72;font-weight:700;}
[data-testid="stDownloadButton"] button:hover {border-color:#4392ac;color:#116f7b;}
[data-testid="stExpander"] {border:1px solid #e0e8f0;border-radius:11px;background:white;}
.stButton button {border-radius:9px;font-weight:700;}
hr {border-color:#e1e8ee;}
[data-testid="stCaptionContainer"] p {color:#728498;font-size:12px;}
@media (max-width:720px){.dashboard-hero{padding:23px 20px}.dashboard-hero h1{font-size:23px!important}[data-testid="stMetric"]{padding:13px 12px;min-height:98px}}
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

st.markdown('''<div class="dashboard-hero">
<div class="hero-eyebrow">INPERTEK / INTERNAL ANALYTICS</div>
<h1>Affiliate Commission Optimizer</h1>
<p>Multi-master analysis with margin-first decision intelligence</p>
<div class="hero-tags"><span class="hero-tag">V5.2 · ELEGANT HYBRID</span><span class="hero-tag">MASTER READ-ONLY</span><span class="hero-tag">KOMISI 1% STEP</span></div>
</div>''',unsafe_allow_html=True)
st.info('Satu dashboard untuk banyak master. Hanya sheet hijau dengan mapping terverifikasi yang diproses. Master asli tidak diubah; hasil ekspor adalah bahan review, bukan template impor resmi marketplace.')
try:db_connect(DB).close()
except Exception:st.error('Database SQLite tidak dapat dibuka.');st.stop()

with st.sidebar:
    st.markdown('### ◈ INPERTEK')
    st.caption('AFFILIATE OPTIMIZER / V5.2')
    st.divider()
    st.header('01 / Upload master')
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
    st.header('02 / Kebijakan komisi')
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
    st.header('03 / Filter data')
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

st.markdown('<div class="section-kicker">Ringkasan performa analisis</div>',unsafe_allow_html=True)
m1,m2,m3,m4=st.columns(4)
m1.metric('Master dipilih',len(chosen))
m2.metric('Baris marketplace',len(output))
m3.metric('Kandidat naik',int((output.STATUS=='NAIK').sum()))
m4.metric('Perlu validasi',int((output.VALIDASI!='OK').sum()))

st.markdown('<div class="section-kicker">Workspace analisis</div>',unsafe_allow_html=True)
tabs=st.tabs(['📈 Analisis Komisi','✅ Kandidat Ekspor','⚠️ Validasi Data','🔄 Perubahan Master','📗 Panduan & Detail SKU','🧩 Mapping & Histori'])
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
