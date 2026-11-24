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
