import streamlit as st
import pandas as pd

st.title("Excel Filtreleme Aracı")

# Kullanıcıdan dosya yüklemesini iste
uploaded_file = st.file_uploader("Bir Excel dosyası yükleyin", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=None)  # Tüm sayfaları oku

    # Sayfa seçimi
    sheet_name = st.selectbox("Bir sayfa seçin", df.keys())
    data = df[sheet_name]

    # "Gümrük Vergisi Kalem Rto" sütununda 0 olanları filtreleme
    if "Gümrük Vergisi Kalem Rto" in data.columns:
        data = data[data["Gümrük Vergisi Kalem Rto"] != 0]

    # Varsayılan olarak filtrelenecek sütunlar
    default_columns = [
        "TCGB Gümrük İdaresi", "TCGB Tescil No", "TCGB Tescil Tarihi", "Alıcı / Gönderici Unvan",
        "Kalem No", "Satır Kodu", "GTİP Kodu (12 li)", "GTİP açıklaması", "Madde Adı",
        "Tamamlayıcı Ölçü Birim", "Miktar", "Brüt Kg", "Net Kg", "İstatistiki Birim Kodu",
        "İstatistiki Miktar", "İstatistiki Kıymet ($)", "Kalem Rejim Kodu", "Menşe Ülke Adı",
        "Sevk Ülkesi", "Çıkış Ülkesi", "Varış Ülkesi", "Ticaret Yapılan Ülke", "Kap Ürün Bilgisi",
        "Özel Durum", "Muafiyet Kodu", "Fatura Bedeli", "Döviz Türü", "Gümrük Vergisi Kalem Rto", "Gümrük Vergisi USD"
    ]

    # Kullanıcıya ek sütun seçme imkanı ver
    available_columns = [col for col in data.columns if col not in default_columns]
    selected_columns = st.multiselect(
        "Eklemek istediğiniz ekstra sütunları seçin:", available_columns
    )

    # Filtreleme sütunu ve değeri
    filter_column = st.selectbox("Filtreleme sütunu seçin (Opsiyonel)", [""] + data.columns.tolist())

    if filter_column:
        filter_value = st.text_input(f"{filter_column} sütununda filtrelenecek değeri girin:")

        if filter_value:
            data = data[data[filter_column].astype(str).str.contains(filter_value, na=False, case=False)]

    # Seçili sütunları belirle (varsayılan + kullanıcı seçimi)
    final_columns = default_columns + selected_columns
    filtered_data = data[final_columns]

    st.write("Filtrelenmiş Veri", filtered_data.head())

    # Kullanıcının çıktı olarak indirmesi için Excel dosyası oluştur
    if not filtered_data.empty:
        output_file = "filtered_data.xlsx"
        filtered_data.to_excel(output_file, index=False)
        
        with open(output_file, "rb") as f:
            st.download_button("Excel olarak indir", f, file_name="filtrelenmis_veri.xlsx")
    else:
        st.warning("Filtrelenmiş veri yok.")  
