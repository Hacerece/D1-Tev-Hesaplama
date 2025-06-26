import streamlit as st
import pandas as pd
from io import BytesIO
from pandas import DataFrame

def load_excel(uploaded_file):
    return pd.read_excel(uploaded_file, sheet_name=None)

def filter_imports(df):
    columns = [
        'TCGB Gümrük İdaresi', 'TCGB Tescil No', 'TCGB Tescil Tarihi',
        'Alıcı / Gönderici Unvan', 'Kalem No', 'Satır Kodu','Atr','E-Atr','Eur1','E-Eur1','Eur1med','E-Eur1med',
        'GTİP Kodu (12 li)', 'GTİP açıklaması', 'Madde Adı',
        'Tamamlayıcı Ölçü Birim', 'Miktar', 'Brüt Kg', 'Net Kg',
        'İstatistiki Birim Kodu', 'İstatistiki Miktar', 'İstatistiki Kıymet ($)',
        'Kalem Rejim Kodu', 'Menşe Ülke Adı', 'Sevk Ülkesi', 'Çıkış Ülkesi',
        'Varış Ülkesi', 'Ticaret Yapılan Ülke', 'Kap Ürün Bilgisi',
        'Özel Durum', 'Muafiyet Kodu', 'Fatura Bedeli', 'Döviz Türü',
        'Gümrük Vergisi Kalem Rto', 'Gümrük Vergisi USD'
    ]
    df = df[df['Gümrük Vergisi Kalem Rto'].astype(str).str.replace(',', '').str.replace('.', '').str.isdigit()]
    df = df[df['Gümrük Vergisi Kalem Rto'].astype(float) != 0]
    filtered = df[df['Özel Durum'] == 0]
    available_columns = [col for col in columns if col in df.columns]
    return filtered[available_columns]

def filter_exports(df, ab_countries):
    df_ab = df[df['Varış Ülkesi'].isin(ab_countries)]
    df_non_ab = df[~df['Varış Ülkesi'].isin(ab_countries)]
    return df_ab, df_non_ab

ab_countries = [
    "ALMANYA", "AVUSTURYA", "BELÇİKA", "BULGARİSTAN", "ÇEKYA", "DANİMARKA", "ESTONYA", "FİNLANDİYA",
    "FRANSA", "HİRVATİSTAN", "HOLLANDA", "İRLANDA", "İSPANYA", "İSVEC", "İTALYA", "LETONYA", "LİTVANYA",
    "LÜKSEMBURG", "MACARİSTAN", "MALTA", "POLONYA", "PORTEKİZ", "ROMANYA", "SLOVAKYA", "SLOVENYA", "YUNANİSTAN"
]

st.set_page_config(page_title="İthalat ve İhracat Raporları", page_icon="📊", layout="wide")

st.title("İthalat ve İhracat Verilerini Yükleyin")
st.markdown("**İthalat ve İhracat verilerinizi yükleyin ve istediğiniz raporları alın.**")

uploaded_file = st.file_uploader("Excel Dosyasını Yükleyin", type=["xlsx"])

if uploaded_file is not None:
    data = load_excel(uploaded_file)
    sheet_names = list(data.keys())
    st.write(f"Yüklenen dosyada şu sheet'ler var: {', '.join(sheet_names)}")

    options = ['Gerç.İth.List.', 'Gerç.İhr.List.', 'Sarfiyat']
    selected_option = st.selectbox("İşlem Yapmak İstediğiniz Veri Tipini Seçin", options)

    if len(sheet_names) >= 3:
        ithalat_df_all = data[sheet_names[0]]
        ihracat_df_all = data[sheet_names[1]]
        sarfiyat_df_all = data[sheet_names[2]]

        ithalat_df = filter_imports(ithalat_df_all)
        ithalat_pivot = ithalat_df.groupby(['Satır Kodu'])['İstatistiki Miktar'].sum().reset_index().rename(columns={'İstatistiki Miktar': 'Toplam İstatistiki Miktar'})

        ab_df, non_ab_df = filter_exports(ihracat_df_all, ab_countries)
        ab_pivot = ab_df.groupby('Satır Kodu')['İstatistiki Miktar'].sum().reset_index().rename(columns={'İstatistiki Miktar': 'Toplam Miktar'})
        non_ab_pivot = non_ab_df.groupby('Satır Kodu')['İstatistiki Miktar'].sum().reset_index().rename(columns={'İstatistiki Miktar': 'Toplam Miktar'})

        # SARFİYAT: satır kodu hem "Madde Adı" sütununda hem de kolon adlarında kontrol edilir
        sarfiyat_df = None
        if sarfiyat_df_all is not None:
            ithalat_kodlari = ithalat_pivot['Satır Kodu'].astype(str).unique()
            dunya_kodlari = non_ab_pivot['Satır Kodu'].astype(str).unique()

            satir_maskesi = sarfiyat_df_all['Madde Adı'].astype(str).isin(ithalat_kodlari)
            kolonlar = ['Madde Adı', 'Parametreler'] + [col for col in sarfiyat_df_all.columns if str(col) in dunya_kodlari]

            sarfiyat_df = sarfiyat_df_all[satir_maskesi][kolonlar]
            
        # SARFİYAT: Satır kodu hem Madde Adı'nda hem de kolon adlarında kontrol edilir
        sarfiyat_df = None
        if sarfiyat_df_all is not None:
            ithalat_kodlari = ithalat_pivot['Satır Kodu'].astype(str).unique()
            dunya_kodlari = non_ab_pivot['Satır Kodu'].astype(str).unique()

            satir_maskesi = sarfiyat_df_all['Madde Adı'].astype(str).isin(ithalat_kodlari)
            kolonlar = ['Madde Adı', 'Parametreler'] + [col for col in sarfiyat_df_all.columns if str(col) in dunya_kodlari]

            sarfiyat_df_filtered = sarfiyat_df_all[satir_maskesi][kolonlar]

            # İthalat pivotuyla birleştir ve yeni sütunu ekle
            sarfiyat_df = sarfiyat_df_filtered.merge(
                ithalat_pivot,
                how="left",
                left_on="Madde Adı",
                right_on="Satır Kodu"
            ).drop(columns=["Satır Kodu"])

            sarfiyat_df.rename(columns={"Toplam İstatistiki Miktar": "Gerçekleşen İthalat Miktarı"}, inplace=True)

        # 3. Dünya Ülkeleri Pivot verisini alın
        dunya_pivot_dict = dict(zip(non_ab_pivot['Satır Kodu'].astype(str), non_ab_pivot['Toplam Miktar']))

        # Yeni "Toplam Miktar" satırını oluştur
        toplam_miktar_row = {
            "Madde Adı": "",
            "Parametreler": "Toplam Miktar"
        }

        # Tablodaki sarfiyat kolonları üzerinde dön
        for col in sarfiyat_df.columns:
            if col in dunya_pivot_dict:
                toplam_miktar_row[col] = dunya_pivot_dict[col]
            elif col not in ["Madde Adı", "Parametreler"]:
                toplam_miktar_row[col] = 0  # Eşleşmeyen kolonlara 0 yaz

        # Yeni satırı en başa ekle
        sarfiyat_df = pd.concat([
            pd.DataFrame([toplam_miktar_row]),
            sarfiyat_df
        ], ignore_index=True)
        
        # Kullanıcıdan gelen sarfiyat tablosunun 3. satırı (index 2), ürün adlarını içeriyor
        urun_adi_satiri = sarfiyat_df_all.iloc[3]

        # Yeni satır: Kullanılan Ürün
        kullanilan_urun_row = {
            "Madde Adı": "",
            "Parametreler": "Kullanılan Ürün"
        }

        # Her ürün kolonu için ürün adını al
        for col in sarfiyat_df.columns:
            if col not in ["Madde Adı", "Parametreler"] and col in urun_adi_satiri.index:
                kullanilan_urun_row[col] = urun_adi_satiri[col]

        # "Toplam Miktar" satırı zaten en başta
        # Şimdi onun altına "Kullanılan Ürün" satırını ekle
        sarfiyat_df = pd.concat([
            sarfiyat_df.iloc[[0]],                          # Toplam Miktar
            pd.DataFrame([kullanilan_urun_row]),           # Kullanılan Ürün
            sarfiyat_df.iloc[1:]                           # Kalan sarfiyat
        ], ignore_index=True)
        

        # Blok bazlı: Fire ve Toplam Birim Kullanım satırlarını "Birim Kullanım"ın altına ekle
        yeni_sarfiyat_df = pd.DataFrame(columns=sarfiyat_df.columns)

        i = 0
        while i < len(sarfiyat_df):
            satir = sarfiyat_df.iloc[i]
            yeni_sarfiyat_df.loc[len(yeni_sarfiyat_df)] = satir

            if satir["Parametreler"] == "Birim Kullanım Miktarı (adet)":
                # Aynı madde adına sahip satırları orijinal df'de bul
                madde_adi = satir["Madde Adı"]
                orijinal_blok = sarfiyat_df_all[sarfiyat_df_all["Madde Adı"] == madde_adi]

                if not orijinal_blok.empty:
                    idx = orijinal_blok.index[0]
                    try:
                        try:
                            fire_row = sarfiyat_df_all.iloc[idx + 1].copy()
                            if "fire" in str(fire_row["Parametreler"]).lower():
                                if pd.isna(fire_row["Madde Adı"]) or fire_row["Madde Adı"] == "":
                                    fire_row["Madde Adı"] = madde_adi

                                # Yeni satırı tam uyumlu şekilde oluştur
                                fire_row_dict = {col: "" for col in sarfiyat_df.columns}
                                for col in sarfiyat_df.columns:
                                    if col in fire_row:
                                        fire_row_dict[col] = fire_row[col]

                                yeni_sarfiyat_df.loc[len(yeni_sarfiyat_df)] = fire_row_dict
                        except Exception as e:
                            st.warning(f"Fire satırı eklenirken hata oluştu: {e}")

                    except:
                        pass

                    try:
                        toplam_row = sarfiyat_df_all.iloc[idx + 2].copy()
                        if toplam_row["Parametreler"] == "Toplam Birim Kullanım":
                            yeni_sarfiyat_df.loc[len(yeni_sarfiyat_df)] = toplam_row
                    except:
                        pass
            i += 1

        sarfiyat_df = yeni_sarfiyat_df.reset_index(drop=True)

        # Hesaplanan Toplam Mamul Kullanımını alt satıra ekle
        yeni_df_with_mamul = pd.DataFrame(columns=sarfiyat_df.columns)
        i = 0
        while i < len(sarfiyat_df):
            satir = sarfiyat_df.iloc[i]

            if satir["Parametreler"] == "Birim Kullanım Miktarı (adet)":
                # Hesaplamayı şimdilik tut
                madde_adi = satir["Madde Adı"]
                mamul_row = {col: "" for col in sarfiyat_df.columns}
                mamul_row["Madde Adı"] = madde_adi
                mamul_row["Parametreler"] = "Toplam Mamul Kullanımı"

                for col in sarfiyat_df.columns:
                    if col not in ["Madde Adı", "Parametreler"] and pd.notna(satir[col]):
                        try:
                            adet = float(satir[col])
                            miktar = float(toplam_miktar_row.get(col, 0))
                            mamul_row[col] = adet * miktar
                        except:
                            mamul_row[col] = ""

                # 1. Birim Kullanım'ı ekle
                yeni_df_with_mamul.loc[len(yeni_df_with_mamul)] = satir

                # 2. Fire satırı varsa ekle
                if i + 1 < len(sarfiyat_df):
                    fire_row = sarfiyat_df.iloc[i + 1]
                    if str(fire_row["Parametreler"]).lower().startswith("fire"):
                        yeni_df_with_mamul.loc[len(yeni_df_with_mamul)] = fire_row
                        i += 1

                # 3. Toplam Birim Kullanım varsa ekle
                if i + 1 < len(sarfiyat_df):
                    toplam_row = sarfiyat_df.iloc[i + 1]
                    if str(toplam_row["Parametreler"]).strip().lower() == "toplam birim kullanım":
                        yeni_df_with_mamul.loc[len(yeni_df_with_mamul)] = toplam_row
                        i += 1

                # 4. En son: Toplam Mamul Kullanımı'nı ekle
                yeni_df_with_mamul.loc[len(yeni_df_with_mamul)] = mamul_row

            else:
                # Diğer tüm satırları aynen aktar
                yeni_df_with_mamul.loc[len(yeni_df_with_mamul)] = satir

            i += 1

        sarfiyat_df = yeni_df_with_mamul.reset_index(drop=True)

        # 1. İlk satır: "Toplam Miktar" satırı
        toplam_miktar_row = sarfiyat_df.iloc[0]

        # 1. "Toplam Mamul Kullanımı" satırlarını filtrele
        mamul_maskesi = sarfiyat_df["Parametreler"] == "Toplam Mamul Kullanımı"
        mamul_satirlari = sarfiyat_df[mamul_maskesi].copy()
        
        for idx, row in sarfiyat_df[mamul_maskesi].iterrows():
            madde_adi = row["Madde Adı"]
            
            # Aynı madde adına sahip satırdan gerçekleşen ithalatı bul
            ithalat_degeri = sarfiyat_df.loc[
                (sarfiyat_df["Madde Adı"] == madde_adi) &
                (sarfiyat_df["Parametreler"] == "Birim Kullanım Miktarı (adet)")
            ]["Gerçekleşen İthalat Miktarı"]

            if not ithalat_degeri.empty:
                sarfiyat_df.at[idx, "Gerçekleşen İthalat Miktarı"] = ithalat_degeri.values[0]

        # 2. Her satırdaki sayısal sütunları toplayarak yeni sütun ekle
        sarfiyat_df.loc[mamul_maskesi, "Toplam Mamul Kullanımı"] = mamul_satirlari.apply(
            lambda row: sum([float(row[col]) for col in sarfiyat_df.columns if col not in ["Madde Adı", "Parametreler"] and pd.notna(row[col]) and isinstance(row[col], (int, float))]),
            axis=1
        )
        # Fark ve TEV durumu hesapla (yalnızca mamul satırları için)
        sarfiyat_df.loc[mamul_maskesi, "Fark"] = (
            sarfiyat_df.loc[mamul_maskesi, "Toplam Mamul Kullanımı"] -
            sarfiyat_df.loc[mamul_maskesi, "Gerçekleşen İthalat Miktarı"]
        )

        sarfiyat_df.loc[mamul_maskesi, "TEV Durumu"] = sarfiyat_df.loc[mamul_maskesi, "Fark"].apply(
            lambda x: "TEV Var" if pd.notna(x) and x < 0 else "TEV Yok"
        )
        

        # Görselleştirme
        if selected_option == 'Gerç.İth.List.':
            st.subheader("Filtrelenmiş İthalat Verileri")
            st.dataframe(ithalat_df)
            st.subheader("İthalat Pivot Tablosu")
            st.dataframe(ithalat_pivot)

        elif selected_option == 'Gerç.İhr.List.':
            st.subheader("Vergili İhracat (AB Ülkeleri)")
            st.dataframe(ab_df)
            st.subheader("AB Ülkeleri Pivot Tablosu")
            st.dataframe(ab_pivot)
            st.subheader("3. Dünya Ülkelerine İhracat")
            st.dataframe(non_ab_df)
            st.subheader("3. Dünya Ülkeleri Pivot Tablosu")
            st.dataframe(non_ab_pivot)

        elif selected_option == 'Sarfiyat':
            if sarfiyat_df is not None:
                st.subheader("Filtrelenmiş Sarfiyat Verileri")
                st.dataframe(sarfiyat_df)
            else:
                st.warning("Sarfiyat verisi bulunamadı.")

        output_combined = BytesIO()
        with pd.ExcelWriter(output_combined, engine="xlsxwriter") as writer:
            ithalat_df.to_excel(writer, sheet_name="İthalat Verileri", index=False)
            ithalat_pivot.to_excel(writer, sheet_name="İthalat Pivot", index=False)
            ab_df.to_excel(writer, sheet_name="Vergili İhracat", index=False)
            ab_pivot.to_excel(writer, sheet_name="AB Ülkeleri Pivot", index=False)
            non_ab_df.to_excel(writer, sheet_name="3. Dünya Ülkeleri", index=False)
            non_ab_pivot.to_excel(writer, sheet_name="3. Dünya Ülkeleri Pivot", index=False)
            if sarfiyat_df is not None:
                sarfiyat_df.to_excel(writer, sheet_name="Sarfiyat", index=False)

        output_combined.seek(0)
        st.download_button("Tüm Verileri İndir", data=output_combined, file_name="tüm_veriler_raporu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

