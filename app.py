with tab3:
    st.markdown("### 🔍 Laporan Validasi Pengisian Kuota Mapel Per Rombel")
    if not st.session_state.jadwal_terplot:
        st.info("Jalankan optimasi jadwal terlebih dahulu untuk melihat hasil audit.")
    else:
        audit_records = []
        for kls in semua_kelas:
            tingkat = kls[0] # Mengambil angka depan kelas (7, 8, atau 9)
            matches_kls = [d for d in st.session_state.jadwal_terplot if d['kelas'] == kls]
            df_kls = pd.DataFrame(matches_kls)
            
            unique_mapel = df_kls['kode_mapel'].nunique() if not df_kls.empty else 0
            total_jp_terisi = len(df_kls) if not df_kls.empty else 0
            
            # Kelas 7 target mapel uniknya bisa berbeda karena Seni Budaya diganti Prakarya
            target_mapel = 11 
            
            # Cek status kelengkapan
            if unique_mapel >= target_mapel:
                status = "✅ LENGKAP"
            else:
                status = f"⚠️ KURANG ({target_mapel - unique_mapel} Mapel Belum Terdaftar)"
            
            audit_records.append({
                "Kelas": kls,
                "Jumlah Mapel Terplot": unique_mapel,
                "Total JP Terisi": total_jp_terisi,
                "Status Kelengkapan": status
            })
            
        df_audit = pd.DataFrame(audit_records)
        st.dataframe(df_audit, use_container_width=True)
        
        st.markdown("### 💡 Analisis Mengapa Ada Slot Kosong:")
        st.write("- **Kelas 7:** Sudah aman karena Seni Budaya diisi oleh Prakarya (Guru: Thiara M).")
        st.write("- **Kelas 8 & 9:** Jika masih ada yang berstatus *⚠️ KURANG*, pastikan kembali di Tab 1 apakah **Bahasa Jawa kelas 8 (baru terisi 8A-8B)** dan **Bahasa Inggris kelas 9E** sudah ditambahkan nama gurunya.")
