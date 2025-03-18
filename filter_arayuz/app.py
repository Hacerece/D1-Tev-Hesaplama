import streamlit as st
import pandas as pd
from io import BytesIO

# Excel dosyasını yükleme fonksiyonu
def load_excel(uploaded_file):
    return pd.read_excel(uploaded_file, sheet_name=None)

# İthalat için filtreleme işlemi
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
    # Gümrük Vergisi Kalem Rto değeri sayısal olanları al
    df = df[df['Gümrük Vergisi Kalem Rto'].astype(str).str.replace(',', '').str.replace('.', '').str.isdigit()]
    df = df[df['Gümrük Vergisi Kalem Rto'].astype(float) != 0]  # Sayısala çevirerek 0 olmayanları al
    
    filtered = df[df['Özel Durum'] == 0]
    available_columns = [col for col in columns if col in df.columns]
    return filtered[available_columns]

# İhracat için filtreleme işlemi
def filter_exports(df, ab_countries):
    df_ab = df[df['Varış Ülkesi'].isin(ab_countries)]
    df_non_ab = df[~df['Varış Ülkesi'].isin(ab_countries)]
    return df_ab[['Kalem No', 'Satır Kodu','İstatistiki Miktar', 'GTİP Kodu (12 li)', 'GTİP açıklaması', 'Madde Adı', 'Tamamlayıcı Ölçü Birim', 'Miktar', 'Brüt Kg', 'Varış Ülkesi', 'Çıkış Ülkesi','Atr','E-Atr','Eur1','E-Eur1','Eur1med','E-Eur1med']], \
           df_non_ab[['Kalem No', 'Satır Kodu', 'İstatistiki Miktar', 'GTİP Kodu (12 li)', 'GTİP açıklaması', 'Madde Adı', 'Tamamlayıcı Ölçü Birim', 'Miktar', 'Brüt Kg', 'Varış Ülkesi', 'Çıkış Ülkesi','Atr','E-Atr','Eur1','E-Eur1','Eur1med','E-Eur1med']]

# AB ülkeleri listesi
ab_countries = [
    "ALMANYA", "AVUSTURYA", "BELÇİKA", "BULGARİSTAN", "ÇEKYA", "DANİMARKA", "ESTONYA", "FİNLANDİYA",
    "FRANSA", "HİRVATİSTAN", "HOLLANDA", "İRLANDA", "İSPANYA", "İSVEC", "İTALYA", "LETONYA", "LİTVANYA",
    "LÜKSEMBURG", "MACARİSTAN", "MALTA", "POLONYA", "PORTEKİZ", "ROMANYA", "SLOVAKYA", "SLOVENYA", "YUNANİSTAN"
]

# Streamlit Arayüzü
st.set_page_config(page_title="İthalat ve İhracat Raporları", page_icon="📊", layout="wide")

# Başlık
st.title("İthalat ve İhracat Verilerini Yükleyin")
st.markdown(""" **İthalat ve İhracat verilerinizi yükleyin ve istediğiniz raporları alın.** Excel dosyanızdaki verileri uygun şekilde filtreleyebilir ve raporları oluşturabilirsiniz. """)
# Dosya yükleme
uploaded_file = st.file_uploader("Excel Dosyasını Yükleyin", type=["xlsx"])

# Dosya yüklendiyse işlem yapalım
if uploaded_file is not None:
    data = load_excel(uploaded_file)

    # Sheet isimlerini kontrol etme
    sheet_names = data.keys()
    st.write(f"Yüklenen dosyada şu sheet'ler var: {', '.join(sheet_names)}")

    # Seçim butonları
    options = ['Gerç.İth.List.', 'Gerç.İhr.List.', 'Sarfiyat']
    selected_option = st.selectbox("İşlem Yapmak İstediğiniz Veri Tipini Seçin", options)

    # Sayfa sırasına göre veri seçme
    if len(sheet_names) >= 3:
        ithalat_df_all = data[list(data.keys())[0]]  # 1. Sayfa (İthalat)
        ihracat_df_all = data[list(data.keys())[1]]  # 2. Sayfa (İhracat)
        sarfiyat_df_all = data[list(data.keys())[2]] if len(data) > 2 else None  # 3. Sayfa (Sarfiyat)

        # İthalat verisini filtreleyelim
        ithalat_df = filter_imports(ithalat_df_all)
        
        # İthalat Pivot Hesaplaması
        ithalat_pivot = ithalat_df.groupby(['Satır Kodu']).agg({
            'İstatistiki Miktar': 'sum'
        }).reset_index()
        ithalat_pivot = ithalat_pivot.rename(columns={'İstatistiki Miktar': 'Toplam İstatistiki Miktar'})

        # İhracat verileri
        ab_df, non_ab_df = filter_exports(ihracat_df_all, ab_countries)

        # AB Ülkeleri için pivot hesaplaması
        ab_pivot = ab_df.groupby('Satır Kodu')['İstatistiki Miktar'].sum().reset_index()
        ab_pivot = ab_pivot.rename(columns={'İstatistiki Miktar': 'Toplam Miktar'})

        # 3. Dünya Ülkeleri için pivot hesaplaması
        non_ab_pivot = non_ab_df.groupby('Satır Kodu')['İstatistiki Miktar'].sum().reset_index()
        non_ab_pivot = non_ab_pivot.rename(columns={'İstatistiki Miktar': 'Toplam Miktar'})

        # Sarfiyat verileri
        if sarfiyat_df_all is not None:
            sarfiyat_df = sarfiyat_df_all

        # Ekranda veri görselleştirme
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
            if sarfiyat_df_all is not None:
                st.subheader("Sarfiyat Verileri")
                st.dataframe(sarfiyat_df_all)
            else:
                st.warning("Sarfiyat verisi bulunmamaktadır.")

        # Excel dosyasına yazma ve indirme
        output_combined = BytesIO()
        with pd.ExcelWriter(output_combined, engine="xlsxwriter") as writer:
            # İthalat
            ithalat_df.to_excel(writer, sheet_name="İthalat Verileri", index=False)
            ithalat_pivot.to_excel(writer, sheet_name="İthalat Pivot", index=False)

            # İhracat
            ab_df.to_excel(writer, sheet_name="Vergili İhracat", index=False)
            ab_pivot.to_excel(writer, sheet_name="AB Ülkeleri Pivot", index=False)
            non_ab_df.to_excel(writer, sheet_name="3. Dünya Ülkeleri", index=False)
            non_ab_pivot.to_excel(writer, sheet_name="3. Dünya Ülkeleri Pivot", index=False)

            # Sarfiyat
            if sarfiyat_df_all is not None:
                sarfiyat_df.to_excel(writer, sheet_name="Sarfiyat", index=False)

        output_combined.seek(0)
        st.download_button("Tüm Verileri İndir", data=output_combined, file_name="tüm_veriler_raporu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
