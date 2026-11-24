"""INPERTEK multi-master read-only commission optimizer. No Supabase required."""
import os, tempfile
from io import BytesIO
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from engine import db_connect, sync_many, analyze

st.set_page_config(page_title='INPERTEK | Affiliate Optimizer V5',page_icon='📈',layout='wide')
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
'STATUS':'Naik/turun/tetap/nonaktif/validasi sesuai perhitungan.', 'VALIDASI':'OK berarti data lolos pemeriksaan dasar, bukan verifikasi batas resmi marketplace.',
'ALASAN':'Pertimbangan sistem untuk tindakan terhadap SKU.'}

def rupiah(v):
    try:
        if pd.isna(v):return '-'
        n=float(v);return ('-Rp' if n<0 else 'Rp')+f'{abs(n):,.0f}'.replace(',','.')
    except (ValueError,TypeError):return '-'

def prepare_export(df):
    wb=Workbook();info=wb.active;info.title='PANDUAN'
    for row in [
        ['INPERTEK AFFILIATE OPTIMIZER V5 - REVIEW INTERNAL'],
        ['SUMBER','Snapshot master diunggah; master asli tidak diubah'],
        ['PERINGATAN','BUKAN TEMPLATE UPLOAD RESMI MARKETPLACE'],
        ['ALUR','Validasi listing & ketentuan marketplace, review, lalu susun file impor resmi'],
        ['METODE','Asumsi dasar komisi = harga jual. Margin sebelum affiliate = margin master + biaya affiliate awal'],
        ['BATAS','Sisa margin >= biaya komisi baru, >= margin minimum, dan tidak negatif'],
        ['KEAMANAN','Data master dan database jangan diunggah ke GitHub. Unduh hasil saat analisis selesai'],
        ['DUPLIKAT','SKU berulang antar master / sheet ditahan dari daftar kandidat'],
    ]:info.append(row)
    info.column_dimensions['A'].width=23;info.column_dimensions['B'].width=115
    for title,sub in [('REKOMENDASI',df),('KANDIDAT_REVIEW',df[(df['STATUS'].isin(['NAIK','TURUNKAN'])) & (df['VALIDASI']=='OK') & (~df['DUPLIKAT LINTAS MASTER'])])]:
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

st.title('INPERTEK | Affiliate Commission Optimizer V5')
st.caption('MULTI-MASTER · SHOPEE + TOKOPEDIA/TIKTOK SHOP · SQLite sementara · MASTER READ-ONLY')
st.info('Satu dashboard untuk banyak master. Hanya sheet hijau dengan mapping yang LOLOS VERIFIKASI yang dianalisis. Tidak ada perubahan pada master atau komisi marketplace. File ekspor adalah bahan REVIEW, bukan template impor marketplace.')
try:db_connect(DB).close()
except Exception:st.error('Database SQLite tidak dapat dibuka.');st.stop()

with st.sidebar:
    st.header('1 · Upload master')
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
    st.header('2 · Kebijakan komisi')
    st.caption('Default: agresif, tetapi sisa margin tidak boleh minus dan idealnya minimal sebesar komisi baru.')
    mode=st.selectbox('Agresivitas',['AGRESIF','SEIMBANG','KONSERVATIF'],index=0)
    headroom={'AGRESIF':.85,'SEIMBANG':.60,'KONSERVATIF':.35}[mode]
    retain=st.slider('Minimum margin sebelum affiliate yang dipertahankan (%)',0,95,30,5)/100
    min_profit=st.number_input('Margin minimum (Rp)',min_value=0.0,value=0.0,step=1000.0)
    cap=st.number_input('Batas komisi simulasi (%)',min_value=.5,max_value=100.0,value=30.0,step=.5)
    target=st.number_input('Target maksimum komisi (%)',min_value=.5,max_value=100.0,value=30.0,step=.5)
    step=st.selectbox('Kelipatan komisi (%)',[.5,1.0],index=0)
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
    st.header('3 · Filter')
    market=st.selectbox('Marketplace',['SEMUA','SHOPEE','TOKOPEDIA'])
    keyword=st.text_input('Cari SKU / nama produk')
    validity=st.selectbox('Validasi',['SEMUA','HANYA OK','BERMASALAH'])
if market!='SEMUA':base=base[base.marketplace==market]
if keyword:base=base[base.sku.fillna('').str.contains(keyword,case=False,regex=False)|base.name.fillna('').str.contains(keyword,case=False,regex=False)]
if validity=='HANYA OK':base=base[base.validation=='OK']
if validity=='BERMASALAH':base=base[base.validation!='OK']
if base.empty:st.info('Tidak ada produk sesuai filter.');st.stop()

# Analyze in the view, independent of master. Cross-master duplicates are never eligible for export.
rows=[]
for r in base.to_dict('records'):
    result=analyze(r,retained_share=retain,min_profit=min_profit,max_pct=cap,target_pct=target,headroom=headroom,step=step)
    rows.append({'MASTER':r['master_name'],'SHEET':r['sheet'],'BARIS':r['source_row'],'MARKETPLACE':r['marketplace'],'SKU':r['sku'],'NAMA PRODUK':r['name'],
                 'HARGA JUAL (RP)':r['price'],'HPP (RP)':r['hpp'],'KOMISI AWAL (%)':r['old_pct'],'BIAYA AFF AWAL (RP)':r['old_fee'],'MARGIN MASTER (RP)':r['old_margin'],
                 'MARGIN SEBELUM AFF (RP)':result['margin_before'],'BATAS AMAN (%)':result['safe_cap'],'KOMISI USULAN (%)':result['proposed_pct'],
                 'BIAYA KOMISI USULAN (RP)':result['proposed_fee'],'SISA MARGIN USULAN (RP)':result['margin_new'],'SELISIH MARGIN - KOMISI (RP)':result['balance'],
                 'STATUS':result['status'],'VALIDASI':r['validation'],'ALASAN':result['reason']})
output=pd.DataFrame(rows)
dups=output.duplicated(subset=['MARKETPLACE','SKU'],keep=False)
output['DUPLIKAT LINTAS MASTER']=dups
if dups.any():
    output.loc[dups,'ALASAN']=output.loc[dups,'ALASAN']+' SKU ini muncul beberapa kali pada master/sheet terpilih; kandidat ditahan dari ekspor.'

m1,m2,m3,m4=st.columns(4)
m1.metric('Master dipilih',len(chosen))
m2.metric('Baris marketplace',len(output))
m3.metric('Kandidat naik',int((output.STATUS=='NAIK').sum()))
m4.metric('Perlu validasi',int((output.VALIDASI!='OK').sum()))

tabs=st.tabs(['📈 Analisis Komisi','✅ Kandidat Ekspor','🧩 Mapping & Validasi','🔄 Riwayat Perubahan','📗 Panduan'])
with tabs[0]:
    states=st.multiselect('Filter rekomendasi',sorted(output.STATUS.unique().tolist()))
    shown=output[output.STATUS.isin(states)] if states else output
    display=shown.drop(columns=['DUPLIKAT LINTAS MASTER']).copy()
    for c in MONEY:display[c]=display[c].map(rupiah)
    cfg={c:st.column_config.TextColumn(c,help=GUIDE.get(c,'Periksa sumber dan satuan pada kolom ini.'),width='large' if c in ('NAMA PRODUK','ALASAN') else 'medium') for c in display.columns}
    st.dataframe(display,column_config=cfg,hide_index=True,use_container_width=True,height=480)
    st.download_button('Download analisis lengkap (Excel)',data=prepare_export(output),file_name='INPERTEK_ANALISIS_V5.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.caption('Angka Rupiah ditampilkan dengan pemisah ribuan Indonesia; file Excel tetap menyimpan angka numerik.')
with tabs[1]:
    valid=output[(output.STATUS.isin(['NAIK','TURUNKAN'])) & (output.VALIDASI=='OK') & (~output['DUPLIKAT LINTAS MASTER'])].copy()
    st.metric('Kandidat perubahan yang layak DIREVIEW',len(valid))
    st.warning('Ini daftar REVIEW internal, belum format upload resmi Shopee/Tokopedia. Tidak ada keputusan otomatis yang dikirim ke marketplace.')
    st.dataframe(valid[['MASTER','MARKETPLACE','SKU','NAMA PRODUK','KOMISI AWAL (%)','KOMISI USULAN (%)','STATUS','ALASAN']],use_container_width=True,hide_index=True)
    st.download_button('Download kandidat perubahan (Excel)',prepare_export(output),file_name='INPERTEK_KANDIDAT_REVIEW_V5.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.caption('Gunakan sheet KANDIDAT_REVIEW dalam Excel. SKU ambigu lintas master ditahan.')
with tabs[2]:
    st.subheader('Mapping yang berhasil diverifikasi')
    mapp=mapping[mapping.source_name.isin(selected)]
    st.dataframe(mapp[['source_name','sheet','status','reason']],hide_index=True,use_container_width=True)
    st.caption('Versi ini memiliki template kolom utama terverifikasi. Sheet dengan format lain DILEWATI, tidak dipaksakan. Untuk format baru, mapping khusus perlu ditambahkan dan diuji.')
    st.subheader('Contoh data bermasalah')
    st.dataframe(output[output.VALIDASI!='OK'][['MASTER','SHEET','BARIS','MARKETPLACE','SKU','VALIDASI']].head(1000),use_container_width=True,hide_index=True)
with tabs[3]:
    st.subheader('Perubahan antar upload master yang sama')
    if change.empty:st.info('Belum ada perubahan karena setiap master baru memiliki satu snapshot atau file yang diunggah belum berubah.')
    else:st.dataframe(change,hide_index=True,use_container_width=True)
    st.caption('Database SQLite cloud bersifat sementara. Unduh hasil Excel sebelum aplikasi di-reboot.')
with tabs[4]:
    for k,v in GUIDE.items():st.markdown(f'**{k}** — {v}')
    st.markdown('**Aturan normal:** margin sebelum affiliate = margin master + biaya affiliate lama; biaya baru = harga jual × komisi baru; margin tersisa ≥ biaya baru, ≥ persentase margin dilindungi, ≥ minimum Rupiah dan tidak negatif.')
    st.markdown('**Catatan:** hanya struktur dan data yang konsisten yang boleh dipakai. Kenaikan agresif adalah simulasi kapasitas margin, bukan jaminan peningkatan penjualan affiliate.')
