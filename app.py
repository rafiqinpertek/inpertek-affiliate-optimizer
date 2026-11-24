"""Dashboard pilot INSERT: master read-only, explanations and balanced affiliate rate rules."""
import os
import tempfile
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st
from engine import recommend_balanced
from cloud_store import init_cloud, read_query, sync_cloud

st.set_page_config(page_title='INPERTEK | Affiliate Commission Optimizer',layout='wide',page_icon='📊')
def rupiah(v):
    if pd.isna(v):return '-'
    return ('-Rp' if float(v)<0 else 'Rp')+f'{abs(float(v)):,.0f}'.replace(',','.')

# Cloud-only: fail closed if the DATABASE_URL secret is missing.
# Never display database URLs, passwords, or raw connection exceptions.
try:
    DATABASE_URL = st.secrets.get('DATABASE_URL', os.environ.get('DATABASE_URL', ''))
except Exception:
    DATABASE_URL = os.environ.get('DATABASE_URL', '')
if not DATABASE_URL:
    st.error('Database PostgreSQL belum diatur. Masukkan DATABASE_URL di Settings → Secrets pada Streamlit Cloud.')
    st.stop()
try:
    init_cloud(DATABASE_URL)
except Exception:
    st.error('Database belum bisa dihubungi. Periksa DATABASE_URL, SSL, dan pengaturan jaringan Supabase. Kredensial tidak ditampilkan demi keamanan.')
    st.stop()

st.title('INPERTEK | Affiliate Commission Optimizer V4 · PostgreSQL')
st.caption('PILOT MASTER INSERT · TOKOPEDIA (termasuk TikTok Shop) + SHOPEE · Master dan komisi marketplace tidak diubah.')
st.info('**MASTER READ-ONLY:** file master tidak pernah ditulis/diubah oleh aplikasi. Hasil analisis dan file kandidat update dibuat terpisah. **Arti rekomendasi:** perhitungan adalah simulasi dari margin master SETELAH komisi affiliate. Aturan standar menjaga **sisa margin setelah komisi usulan minimal sama dengan biaya komisi usulan**. Pengecualian peluang khusus perlu persetujuan manual.')

with st.sidebar:
    st.header('1 · Sinkronisasi master')
    upload=st.file_uploader('Unggah 01. INSERT.xlsx terbaru',type=['xlsx'],help='Prioritaskan worksheet bertab hijau. File asli tidak diedit. Upload ulang file yang sudah berubah untuk deteksi perbedaan.')
    if st.button('Sinkronkan dan bandingkan',disabled=upload is None,use_container_width=True):
        with tempfile.NamedTemporaryFile(suffix='.xlsx',delete=False) as tmp:
            tmp.write(upload.getbuffer());path=tmp.name
        try:
            with st.spinner('Membaca worksheet hijau dan membandingkan snapshot...'):
                res=sync_cloud(path,DATABASE_URL)
            st.success('Sinkronisasi: '+str(res))
        except Exception as e:
            st.error('Sinkronisasi gagal. Periksa format master INSERT, tab hijau, dan koneksi database. Detail kredensial disembunyikan.')
        finally:
            os.remove(path)

try:
    runs=read_query(DATABASE_URL,'SELECT * FROM sync_runs ORDER BY run_id DESC')
except Exception:
    st.error('Tidak bisa membaca histori PostgreSQL. Periksa pengaturan Supabase.')
    st.stop()
if runs.empty:
    st.warning('Database PostgreSQL belum berisi snapshot. Unggah master INSERT untuk membuat snapshot pertama. Master hanya dibaca, tidak dimodifikasi.')
    st.stop()
latest=runs.iloc[0]
run_id=int(latest.run_id)
try:
    base=read_query(DATABASE_URL,'SELECT * FROM records WHERE run_id=?',(run_id,))
except Exception:
    st.error('Gagal mengambil data produk dari PostgreSQL.')
    st.stop()

with st.sidebar:
    st.header('2 · Filter produk')
    market=st.selectbox('Marketplace',['SEMUA','TOKOPEDIA','SHOPEE'],help='TOKOPEDIA mewakili Tokopedia dan TikTok Shop sesuai struktur master.')
    sheet=st.selectbox('Sheet sumber',['SEMUA']+sorted(base.source_sheet.dropna().unique().tolist()))
    query=st.text_input('Cari SKU / nama produk',help='Cari tanpa membuka sheet sumber.')
    st.header('3 · Aturan komisi (simulasi)')
    target_rate=st.number_input('Target komisi maksimum yang diinginkan (%)',0.0,100.0,10.0,0.5,help='Tarif tujuan, bukan komisi yang harus diberikan. Sistem tidak melewati batas keamanan finansial.')
    share=st.slider('Minimal porsi margin sebelum affiliate yang dilindungi (%)',0,100,50,5,help='Contoh 50%: minimal separuh margin sebelum affiliate harus tersisa setelah komisi baru.')/100
    min_profit=st.number_input('Minimal margin tersisa (Rp)',min_value=0.0,value=0.0,step=1000.0,help='Ambang minimum Rupiah; sistem memakai batas yang lebih ketat dibanding target persentase.')
    cap=st.number_input('Batas tarif komisi simulasi (%)',min_value=0.0,max_value=100.0,value=20.0,step=0.5,help='Bukan ketentuan resmi marketplace; sesuaikan dengan aturan program aktif.')
    use_headroom=st.slider('Porsi ruang komisi aman yang digunakan untuk kenaikan (%)',0,100,50,5,help='Kenaikan adaptif: jika ruang aman besar, kenaikan bisa lebih besar. Jika ruang kecil, kenaikan ikut mengecil.')/100
    optional_cap=st.toggle('Batasi kenaikan per evaluasi (opsional)',value=False)
    inc=st.number_input('Batas kenaikan (poin persentase)',0.0,30.0,2.0,0.5,disabled=not optional_cap) if optional_cap else None
    step=st.selectbox('Kelipatan usulan komisi (%)',[0.5,1.0])
    st.divider()
    opportunity=st.toggle('Tampilkan skenario peluang khusus',value=False,help='Jika aktif, sistem boleh memperlihatkan pengecualian margin tersisa < komisi, tetapi statusnya wajib review manual, bukan disetujui otomatis.')
    opp_ratio=st.slider('Minimal rasio margin/komisi untuk skenario khusus',0.5,0.99,0.75,0.05,disabled=not opportunity,help='Hanya untuk simulasi peluang khusus, misal 0,75 berarti sisa margin >= 75% biaya komisi.')

x,y,z,w=st.columns(4)
x.metric('Worksheet hijau',int(latest.sheets_green),help='Jumlah tab berwarna hijau yang terpilih pada impor terakhir.')
y.metric('Baris produk sumber',int(latest.source_rows),help='Baris SKU yang terbaca dari worksheet hijau, belum dideduplikasi lintas sheet.')
z.metric('Baris marketplace',int(latest.listing_rows),help='Setiap baris produk dapat menghasilkan 2 baris: TOKOPEDIA dan SHOPEE.')
w.metric('Catatan validasi',int(latest.issues),help='Catatan kesalahan data pada sumber, bukan jumlah SKU unik bermasalah.')
st.caption(f"Snapshot #{run_id} · sinkronisasi terakhir: {latest.sync_time} · sumber: {latest.source_name}")

filtered=base
if market!='SEMUA':filtered=filtered[filtered.marketplace==market]
if sheet!='SEMUA':filtered=filtered[filtered.source_sheet==sheet]
if query:
    filtered=filtered[filtered.sku.fillna('').str.contains(query,case=False,regex=False)|filtered.name.fillna('').str.contains(query,case=False,regex=False)]

view=filtered.copy()
if not view.empty:
    outcomes=[recommend_balanced(r,protected_share=share,min_profit=min_profit,target_rate=target_rate,
                ceiling_rate=cap,step=step,max_increase=inc,opportunity_mode=opportunity,
                use_headroom=use_headroom,
                opportunity_min_ratio=opp_ratio) for r in view.to_dict('records')]
    for key,col in [('status','KEPUTUSAN'),('proposed_rate','KOMISI_USULAN_%'),('margin_before','MARGIN_SEBELUM_AFF_RP'),
                    ('safe_cap','BATAS_KOMISI_AMAN_%'),('commission_new','BIAYA_KOMISI_USULAN_RP'),
                    ('margin_after_new','SISA_MARGIN_USULAN_RP'),('balance_ratio','RASIO_MARGIN_TERHADAP_KOMISI'),
                    ('notes','ALASAN_REKOMENDASI'),('allowed_auto','MEMENUHI_ATURAN_NORMAL')]:
        view[col]=[r[key] for r in outcomes]
    view['SELISIH_MARGIN_MINUS_KOMISI_RP']=view['SISA_MARGIN_USULAN_RP']-view['BIAYA_KOMISI_USULAN_RP']
    view['KOMISI_BERUBAH_POIN_%']=view['KOMISI_USULAN_%']-view['commission_pct']

COLUMN_HELP={
    'source_sheet':'Nama worksheet hijau tempat nilai diambil.',
    'source_row':'Nomor baris asli dalam master INSERT.',
    'sku':'Kode produk internal yang dibaca dari master.',
    'name':'Nama produk pada master.',
    'marketplace':'Blok hitung yang dipakai. TOKOPEDIA termasuk TikTok Shop.',
    'price':'Harga jual produk pada blok marketplace (Rupiah).',
    'commission_pct':'Persentase komisi affiliate awal pada master.',
    'affiliate_nominal':'Nilai biaya affiliate awal yang telah masuk perhitungan margin master (Rupiah).',
    'margin_after':'Margin master SETELAH komisi affiliate awal (Rupiah).',
    'MARGIN_SEBELUM_AFF_RP':'Margin master + biaya affiliate awal, untuk mengembalikan biaya komisi sekali.',
    'BATAS_KOMISI_AMAN_%':'Batas komisi yang menjaga margin tersisa >= biaya komisi DAN target margin minimum. Dihitung otomatis untuk masing-masing produk.',
    'KOMISI_USULAN_%':'Tarif yang dihitung dari ruang margin aman tiap produk, target, batas program, dan kelipatan komisi.',
    'BIAYA_KOMISI_USULAN_RP':'Harga jual/dasar komisi dikalikan persentase usulan. Rumus pilot mengasumsikan dasar komisi = harga jual.',
    'SISA_MARGIN_USULAN_RP':'Margin sebelum affiliate dikurangi nominal komisi usulan.',
    'SELISIH_MARGIN_MINUS_KOMISI_RP':'Sisa margin usulan minus biaya komisi usulan. Idealnya >= 0; pengecualian ditandai REVIEW.',
    'RASIO_MARGIN_TERHADAP_KOMISI':'Sisa margin dibagi nominal biaya komisi. >=1 berarti sisa margin tidak lebih kecil daripada komisi.',
    'KOMISI_BERUBAH_POIN_%':'Komisi usulan dikurangi komisi awal dalam poin persentase.',
    'KEPUTUSAN':'POTENSI NAIK / TETAP / TURUNKAN / REVIEW NONAKTIF / VALIDASI DATA / PELUANG KHUSUS - REVIEW.',
    'ALASAN_REKOMENDASI':'Penjelasan kenapa status dan angka usulan dipilih.',
    'validation':'Status validasi nilai sumber; OK berarti lolos pemeriksaan dasar, belum menjamin ketentuan platform.'
}
LABEL={
    'source_sheet':'SHEET SUMBER','source_row':'BARIS SUMBER','sku':'SKU','name':'NAMA PRODUK',
    'marketplace':'MARKETPLACE','price':'HARGA JUAL (RP)','commission_pct':'KOMISI AWAL (%)',
    'affiliate_nominal':'BIAYA KOMISI AWAL (RP)','margin_after':'MARGIN MASTER SESUDAH AFF (RP)',
    'MARGIN_SEBELUM_AFF_RP':'MARGIN SEBELUM AFF (RP)','BATAS_KOMISI_AMAN_%':'BATAS KOMISI AMAN (%)',
    'KOMISI_USULAN_%':'KOMISI USULAN (%)','BIAYA_KOMISI_USULAN_RP':'BIAYA KOMISI USULAN (RP)',
    'SISA_MARGIN_USULAN_RP':'SISA MARGIN SETELAH USULAN (RP)',
    'SELISIH_MARGIN_MINUS_KOMISI_RP':'SELISIH MARGIN - KOMISI (RP)',
    'RASIO_MARGIN_TERHADAP_KOMISI':'RASIO MARGIN / KOMISI',
    'KOMISI_BERUBAH_POIN_%':'PERUBAHAN KOMISI (POIN %)','KEPUTUSAN':'STATUS USULAN',
    'ALASAN_REKOMENDASI':'ALASAN / CATATAN', 'validation':'VALIDASI SUMBER'
}
cols=['source_sheet','source_row','sku','name','marketplace','price','commission_pct','affiliate_nominal','margin_after',
      'MARGIN_SEBELUM_AFF_RP','BATAS_KOMISI_AMAN_%','KOMISI_USULAN_%','BIAYA_KOMISI_USULAN_RP',
      'SISA_MARGIN_USULAN_RP','SELISIH_MARGIN_MINUS_KOMISI_RP','RASIO_MARGIN_TERHADAP_KOMISI',
      'KOMISI_BERUBAH_POIN_%','KEPUTUSAN','ALASAN_REKOMENDASI','validation']
configs={k:st.column_config.NumberColumn(LABEL[k],help=COLUMN_HELP[k],format=('Rp %.0f' if k in ['price','affiliate_nominal','margin_after','MARGIN_SEBELUM_AFF_RP','BIAYA_KOMISI_USULAN_RP','SISA_MARGIN_USULAN_RP','SELISIH_MARGIN_MINUS_KOMISI_RP'] else '%.2f')) if k in ['source_row','price','commission_pct','affiliate_nominal','margin_after','MARGIN_SEBELUM_AFF_RP','BATAS_KOMISI_AMAN_%','KOMISI_USULAN_%','BIAYA_KOMISI_USULAN_RP','SISA_MARGIN_USULAN_RP','SELISIH_MARGIN_MINUS_KOMISI_RP','RASIO_MARGIN_TERHADAP_KOMISI','KOMISI_BERUBAH_POIN_%'] else st.column_config.TextColumn(LABEL[k],help=COLUMN_HELP[k],width='large' if k in ('name','ALASAN_REKOMENDASI') else 'medium') for k in cols}

tabs=st.tabs(['📈 Analisis komisi','🔄 Perubahan master','⚠️ Validasi data','📘 Panduan kolom & aturan','🗂️ Mapping & histori'])
with tabs[0]:
    st.subheader('Analisis per SKU dan marketplace')
    st.caption('Tabel dapat digeser horizontal. Arahkan kursor ke judul kolom untuk melihat artinya. Pilih SKU di panel detail di bawah untuk melihat penjelasan lengkap. Seluruh master dibaca saja.')
    if view.empty:
        st.info('Tidak ada SKU yang sesuai filter.')
    else:
        statuses=st.multiselect('Filter status rekomendasi',sorted(view.KEPUTUSAN.unique().tolist()))
        shown=view[view.KEPUTUSAN.isin(statuses)] if statuses else view
        metrics=st.columns(4)
        metrics[0].metric('Baris tampil',len(shown))
        metrics[1].metric('Potensi naik',int((shown.KEPUTUSAN=='POTENSI NAIK').sum()))
        metrics[2].metric('Perlu turun',int((shown.KEPUTUSAN=='TURUNKAN').sum()))
        metrics[3].metric('Perlu review',int(shown.KEPUTUSAN.isin(['REVIEW NONAKTIF','VALIDASI DATA','PELUANG KHUSUS - REVIEW']).sum()))
        rp_columns=['price','affiliate_nominal','margin_after','MARGIN_SEBELUM_AFF_RP','BIAYA_KOMISI_USULAN_RP','SISA_MARGIN_USULAN_RP','SELISIH_MARGIN_MINUS_KOMISI_RP']
        display=shown[cols].copy()
        for c in rp_columns:
            display[c]=display[c].map(rupiah)
        # Money columns are rendered as Rp strings with Indonesian thousands separators.
        display_configs={k:(st.column_config.TextColumn(LABEL[k],help=COLUMN_HELP[k],width='medium') if k in rp_columns else configs[k]) for k in cols}
        st.dataframe(display,column_config=display_configs,use_container_width=True,hide_index=True,height=420)
        export=shown[cols].rename(columns=LABEL)
        st.download_button('Download analisis lengkap (CSV)',export.to_csv(index=False).encode('utf-8-sig'),file_name='INSERT_ANALISIS_LENGKAP_V3.csv',mime='text/csv')
        # Separate export: never put a recommendation directly into master.
        # Ineligible, exception and ambiguous SKU rows MUST NOT be offered for upload.
        candidates=shown[(shown['validation']=='OK') &
                         (shown['KEPUTUSAN'].isin(['POTENSI NAIK','TURUNKAN'])) &
                         (shown['KOMISI_USULAN_%'].notna()) &
                         (shown['KOMISI_USULAN_%']>0) &
                         (shown['MEMENUHI_ATURAN_NORMAL']) &
                         (shown['SISA_MARGIN_USULAN_RP']>=-0.00001) &
                         (shown['SELISIH_MARGIN_MINUS_KOMISI_RP']>=-0.00001)].copy()
        # Identical SKU across green sheets is not a reliable single platform listing.
        duplicated=candidates.duplicated(subset=['marketplace','sku'],keep=False)
        ambiguous=candidates.loc[duplicated].copy()
        candidates=candidates.loc[~duplicated].copy()
        c1,c2=st.columns(2)
        c1.metric('Kandidat perubahan komisi (review)',len(candidates))
        c2.metric('Kandidat SKU ambigu (ditahan)',len(ambiguous))
        st.caption('Daftar kandidat hanya perubahan NAIK/TURUN yang lolos aturan margin; komisi TETAP, nonaktif, validasi, peluang khusus dan SKU duplikat lintas-sheet tidak masuk daftar ini.')
        export_cols=['marketplace','sku','name','commission_pct','KOMISI_USULAN_%','KEPUTUSAN',
                     'price','margin_after','BIAYA_KOMISI_USULAN_RP','SISA_MARGIN_USULAN_RP',
                     'SELISIH_MARGIN_MINUS_KOMISI_RP','source_sheet','source_row','ALASAN_REKOMENDASI']
        candidate_output=candidates[export_cols].rename(columns=LABEL)
        def make_excel():
            wb=Workbook(); summary=wb.active;summary.title='PANDUAN'
            rules=[
                ['INPERTEK AFFILIATE INSERT - KANDIDAT REVIEW'],
                ['SUMBER','Snapshot file master dibaca saja, tidak diubah'],
                ['PENTING','BELUM merupakan format upload resmi Shopee / Tokopedia / TikTok Shop'],
                ['ALUR','Review kandidat -> cocokkan dengan listing marketplace -> sesuaikan template impor resmi -> upload secara manual'],
                ['ATURAN','Kandidat hanya perubahan NAIK / TURUN, data valid dan sisa margin >= komisi baru serta >= 0'],
                ['ASUMSI','Dasar biaya komisi pilot = harga jual; validasi dengan ketentuan marketplace'],
                ['MARGIN','Margin adalah simulasi dari master; angka marketplace aktual dapat berbeda'],
                ['DUPLIKAT','SKU yang muncul di lebih dari satu worksheet untuk marketplace sama ditahan untuk verifikasi'],
                ['KEAMANAN','Aplikasi tidak mengubah master dan tidak mengunggah ke marketplace'],
            ]
            for x in rules:summary.append(x)
            summary.column_dimensions['A'].width=24;summary.column_dimensions['B'].width=120
            data=wb.create_sheet('KANDIDAT_REVIEW')
            data.append(list(candidate_output.columns))
            for row in candidate_output.itertuples(index=False,name=None):
                data.append([None if pd.isna(v) else v for v in row])
            money_names={LABEL[x] for x in ['price','margin_after','BIAYA_KOMISI_USULAN_RP','SISA_MARGIN_USULAN_RP','SELISIH_MARGIN_MINUS_KOMISI_RP']}
            for idx,header in enumerate(candidate_output.columns,1):
                data.column_dimensions[get_column_letter(idx)].width=54 if header in ('NAMA PRODUK','ALASAN / CATATAN') else (28 if header in ('SKU','SHEET SUMBER') else 23)
                if header in money_names:
                    for rr in range(2,data.max_row+1):
                        data.cell(rr,idx).number_format='"Rp" #,##0;[Red]("Rp" #,##0);"Rp" 0'
                if header in ('KOMISI AWAL (%)','KOMISI USULAN (%)'):
                    for rr in range(2,data.max_row+1):data.cell(rr,idx).number_format='0.00"%"'
            data.freeze_panes='C2';data.auto_filter.ref=data.dimensions
            for sheet_ in (summary,data):
                sheet_.sheet_view.showGridLines=False
                sheet_.row_dimensions[1].height=27
                for cell in sheet_[1]:
                    cell.fill=PatternFill('solid',fgColor='17365D')
                    cell.font=Font(name='Calibri',size=10,bold=True,color='FFFFFF')
                    cell.alignment=Alignment(wrap_text=True,vertical='center')
            out=BytesIO();wb.save(out);return out.getvalue()
        st.download_button('Download KANDIDAT REVIEW (Excel)',data=make_excel(),
            file_name='INPERTEK_INSERT_KANDIDAT_REVIEW_KOMISI_V3.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            help='Excel terpisah, numerik asli untuk Rp, tanpa mengubah master.')
        st.caption('PERHATIAN: Excel ini belum format impor resmi marketplace. Cocokkan SKU, komisi aktif, program affiliate, batas komisi, dan template upload sebelum digunakan. Jangan langsung upload file analisis.')
        if not ambiguous.empty:
            with st.expander('Lihat SKU ganda lintas-sheet yang ditahan'):
                st.dataframe(ambiguous[['marketplace','sku','source_sheet','source_row']],use_container_width=True,hide_index=True)
        st.divider()
        st.subheader('Detail perhitungan dan alasan untuk 1 produk')
        if not shown.empty:
            opts=shown.index.tolist()
            chosen=st.selectbox('Pilih baris produk',opts,format_func=lambda i:f"{shown.loc[i,'sku']} · {shown.loc[i,'marketplace']} · {shown.loc[i,'source_sheet']} · baris {shown.loc[i,'source_row']}")
            r=shown.loc[chosen]
            aa,bb,cc=st.columns(3)
            aa.metric('Komisi awal',f"{r['commission_pct']:.2f}%" if pd.notna(r['commission_pct']) else '-')
            bb.metric('Komisi usulan',f"{r['KOMISI_USULAN_%']:.2f}%" if pd.notna(r['KOMISI_USULAN_%']) else '-')
            cc.metric('Status',str(r['KEPUTUSAN']))
            st.write('**Mengapa hasilnya begitu?**',r['ALASAN_REKOMENDASI'])
            if pd.notna(r['MARGIN_SEBELUM_AFF_RP']):
                st.write(f"**Rumus 1:** Margin sebelum affiliate = {rupiah(r['margin_after'])} + {rupiah(r['affiliate_nominal'])} = **{rupiah(r['MARGIN_SEBELUM_AFF_RP'])}**")
                if pd.notna(r['KOMISI_USULAN_%']):
                    st.write(f"**Rumus 2:** Komisi usulan {r['KOMISI_USULAN_%']:.2f}% × harga {rupiah(r['price'])} = **{rupiah(r['BIAYA_KOMISI_USULAN_RP'])}**")
                    st.write(f"**Rumus 3:** Margin setelah usulan = {rupiah(r['MARGIN_SEBELUM_AFF_RP'])} - {rupiah(r['BIAYA_KOMISI_USULAN_RP'])} = **{rupiah(r['SISA_MARGIN_USULAN_RP'])}**")
                    if r['SELISIH_MARGIN_MINUS_KOMISI_RP']<0:
                        st.warning('SISA MARGIN LEBIH KECIL DARI KOMISI. Ini hanya skenario PELUANG KHUSUS yang memerlukan review, bukan persetujuan otomatis.')
                    else:
                        st.success('Aturan keseimbangan terpenuhi: sisa margin >= biaya komisi usulan.')
            st.caption(f"Sumber: 01. INSERT.xlsx · {r['source_sheet']} · baris {r['source_row']} · {r['marketplace']}")
with tabs[1]:
    st.subheader('Perubahan data sejak snapshot sebelumnya')
    changes=read_query(DATABASE_URL,'SELECT * FROM changes WHERE run_id=?',(run_id,))
    if changes.empty:st.info('Belum ada perubahan. Snapshot pertama jadi baseline. Upload master yang sudah diperbarui untuk membandingkan.')
    else:
        translation={'marketplace':'MARKETPLACE','source_sheet':'SHEET SUMBER','sku':'SKU','field':'KOLOM YANG BERUBAH','previous_value':'NILAI LAMA','new_value':'NILAI TERBARU','event_type':'JENIS PERUBAHAN'}
        st.dataframe(changes.rename(columns=translation),use_container_width=True,hide_index=True)
        st.download_button('Download perubahan CSV',changes.to_csv(index=False).encode('utf-8-sig'),'INSERT_PERUBAHAN.csv')
with tabs[2]:
    st.subheader('Daftar data yang perlu diperiksa')
    bad=read_query(DATABASE_URL,'SELECT * FROM issues WHERE run_id=?',(run_id,))
    st.warning('SKU dengan data kosong/komisi tidak cocok/duplikat tidak dihitung sebagai rekomendasi aman. Kolom validasi menjelaskan penyebabnya.')
    st.dataframe(bad.rename(columns={'source_sheet':'SHEET','source_row':'BARIS','marketplace':'MARKETPLACE','sku':'SKU','issue':'ALASAN VALIDASI'}),use_container_width=True,hide_index=True)
with tabs[3]:
    st.subheader('Kamus setiap kolom pada tabel')
    st.dataframe(pd.DataFrame([{'KOLOM':LABEL[k],'ARTI / CARA MEMBACA':COLUMN_HELP[k]} for k in cols]),hide_index=True,use_container_width=True)
    st.subheader('Cara membaca rekomendasi')
    st.markdown('''
- **POTENSI NAIK:** besarnya kenaikan otomatis dihitung dari ruang komisi aman (porsi dapat diubah), bukan angka tetap. Bukan jaminan peningkatan penjualan.
- **TETAP:** komisi awal sudah cocok dengan batas dan target yang dipilih.
- **TURUNKAN:** komisi awal terlalu tinggi terhadap aturan perlindungan margin; tarif lebih rendah disimulasikan.
- **REVIEW NONAKTIF:** ruang komisi tidak mencukupi menurut kebijakan; bukan perintah menonaktifkan otomatis.
- **PELUANG KHUSUS - REVIEW:** sengaja memperlihatkan alternatif dengan sisa margin lebih kecil dari komisi. **Tidak aman secara aturan normal dan harus ada persetujuan manual.**
- **VALIDASI DATA:** input belum layak dihitung; jangan gunakan hasil komisi sebelum data sumber diperbaiki.
''')
    st.subheader('Rumus utama')
    st.latex(r'M_{sebelum}=M_{master\ sesudah}+C_{awal}')
    st.latex(r'C_{usulan}=P\times k_{usulan}')
    st.latex(r'M_{sisa}=M_{sebelum}-C_{usulan}')
    st.latex(r'M_{sisa}\geq C_{usulan}\quad\mathrm{(aturan\ standar)}')
    st.caption('P pada pilot = harga jual pada blok marketplace. Asumsi tersebut harus diperiksa apabila dasar affiliate sebenarnya berbeda. Mode pengecualian tidak pernah mengizinkan margin menjadi negatif.')
with tabs[4]:
    st.subheader('Mapping awal dari master INSERT')
    st.json({'SKU':'D','NAMA':'E','HPP':'N','TOKOPEDIA':{'HARGA':'CC','BIAYA_AFFILIATE_RP':'DA','KOMISI_PERSEN':'DB','MARGIN_RP':'DH','MARGIN_PERSEN':'DI'},'SHOPEE':{'HARGA':'EM','BIAYA_AFFILIATE_RP':'FL','KOMISI_PERSEN':'FM','MARGIN_RP':'FS','MARGIN_PERSEN':'FT'}})
    st.dataframe(runs,use_container_width=True,hide_index=True)
    st.caption('Tab hijau dipilih sebagai sumber prioritas. Mapping aktual per sheet harus tetap divalidasi; pilot ini belum sinkron langsung dengan Drive.')
