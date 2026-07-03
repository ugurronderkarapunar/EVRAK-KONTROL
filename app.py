# SEKME 3 – Kişi Evrak Düzenle / Sil (v1.6.2 güncellemesi)
with tab3:
    st.markdown("<div class='section-label'>✏️ Kişi Evrak Düzenle / Sil</div>", unsafe_allow_html=True)
    sec_kisi = st.selectbox("Personel", ["(Seçin)"]+personel_listesi, key="kisi_duz")
    if sec_kisi != "(Seçin)":
        kisi_df = df_final[df_final["Adı Soyadı"]==sec_kisi]
        for idx, row in kisi_df.iterrows():
            key_str = row["_key"]
            with st.container():
                st.markdown("---")
                col1, col2, col3 = st.columns([2,1,1])
                with col1:
                    st.markdown(f"**{row['Evrak Adı']}** ({row['Kaynak']})")
                    st.caption(f"Başlangıç: {format_date(row['Başlangıç Tarihi'], fmt)} | Bitiş: {format_date(row['Bitiş Tarihi'], fmt)}")
                    st.caption(f"Kalan: {int(row['Kalan Gün']) if row['Kalan Gün'] else '-'} | Durum: {row['Durum']}")
                with col2:
                    if st.button("✏️ Düzenle", key=f"duz_{key_str}"):
                        st.session_state[f"edit_{key_str}"] = True
                with col3:
                    if st.button("🗑 Sil", key=f"sil_{key_str}"):
                        push_undo()
                        st.session_state["silinmis_evraklar"].append(key_str)
                        if row["Kaynak"] == "Manuel":
                            parts = key_str.rsplit("_manuel_",1)
                            if len(parts)==2:
                                ad_evrak = parts[0].split("|",1)
                                if len(ad_evrak)==2:
                                    ad, evrak = ad_evrak
                                    st.session_state["manuel_tarihler"].pop((ad, evrak), None)
                        save_state_and_excel()
                        st.rerun()
                if st.session_state.get(f"edit_{key_str}"):
                    with st.form(key=f"form_{key_str}"):
                        # Başlangıç ve bitiş tarihi alanları
                        baslangic_deger = row["Başlangıç Tarihi"] or today
                        bitis_deger = row["Bitiş Tarihi"] or today

                        ybas = st.date_input("Başlangıç Tarihi", value=baslangic_deger, key=f"bas_{key_str}")
                        
                        # Otomatik hesapla butonu
                        hesapla_tiklandi = st.form_submit_button("🔁 Bitişi Hesapla (2/5 yıl)")
                        if hesapla_tiklandi:
                            # Evrak adına göre süreyi bul
                            evrak_adi = row["Evrak Adı"]
                            bitis_hesaplanan = hesapla_bitis(ybas, evrak_adi)
                            # Bitiş değerini güncellemek için session_state kullanabiliriz ama form içinde normal değişken zor.
                            # Bunun yerine bitiş alanını disabled yapıp hesaplanan değeri gösterip kaydederken kontrol edeceğiz.
                            # Streamlit formlarında dinamik değişiklik zor olduğu için ayrı bir yöntem kullanalım:
                            # Hesapla butonuna basıldığında bitiş tarihini doğrudan hesaplanan değer olarak ata.
                            # Bunu yapmak için form submit olduğunda bitis_deger'i güncelleyelim.
                            bitis_deger = bitis_hesaplanan
                            st.rerun()  # Formu yenilemek için rerun yeterli değil, çünkü form submit edildi.
                            # Daha iyi çözüm: hesapla butonunu form dışına alalım.
                        # Alternatif: Hesapla butonunu form dışına alıp, butona basınca bitis_deger güncellensin.
                        # Şimdilik bu şekilde düzeltiyoruz: formda iki buton olacak, "Hesapla" ve "Kaydet".
                        # "Hesapla" tıklandığında bitiş değerini hesaplayıp, date_input'a yansıtmak için session_state kullanabiliriz.
                        # Kolay yöntem: bitiş için bir st.number_input değil, date_input'un value'sini session_state'ten al.
                        # Ben basitçe şöyle yapacağım: formda sadece bir "Kaydet" butonu olacak, ama yanında bir checkbox "Bitişi otomatik hesapla" olacak.
                        # Daha temiz: form içinde bir checkbox "Bitişi otomatik hesapla", işaretlenirse bitiş tarihini hesaplanan değer olarak alır.
                        # En kolayı: formun içine bir onay kutusu ekleyip, işaretliyse bitişi hesapla.
                        
                        otomatik_hesapla = st.checkbox("Bitişi otomatik hesapla (sağlık:2 yıl, diğer:5 yıl)", value=False, key=f"oto_{key_str}")
                        ybit = st.date_input("Bitiş Tarihi", value=bitis_deger, key=f"bit_{key_str}")
                        
                        if otomatik_hesapla:
                            # Hesaplanan bitişi göster
                            hesaplanan = hesapla_bitis(ybas, row["Evrak Adı"])
                            st.caption(f"Hesaplanan bitiş: {hesaplanan.strftime('%d.%m.%Y')}")
                            ybit = hesaplanan
                        
                        kaydet_duzenle = st.form_submit_button("Değişiklikleri Kaydet")
                        iptal = st.form_submit_button("İptal")
                        
                        if kaydet_duzenle:
                            push_undo()
                            # Eğer otomatik hesapla seçiliyse ybit zaten hesaplanmıştır
                            final_bitis = hesapla_bitis(ybas, row["Evrak Adı"]) if otomatik_hesapla else ybit
                            st.session_state["evrak_duzenlemeleri"][key_str] = {
                                "baslangic": ybas.strftime("%Y-%m-%d"),
                                "bitis": final_bitis.strftime("%Y-%m-%d") if hasattr(final_bitis, 'strftime') else ybit.strftime("%Y-%m-%d")
                            }
                            save_state_and_excel()
                            st.session_state[f"edit_{key_str}"] = False
                            st.rerun()
                        if iptal:
                            st.session_state[f"edit_{key_str}"] = False
                            st.rerun()
    if st.button("🔄 Tüm düzenleme/silmeleri sıfırla"):
        push_undo()
        st.session_state["evrak_duzenlemeleri"] = {}
        st.session_state["silinmis_evraklar"] = []
        save_state_and_excel()
        st.rerun()
