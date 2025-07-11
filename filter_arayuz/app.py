import streamlit as st
import pandas as pd
from io import BytesIO
from pandas import DataFrame
from openpyxl.utils import column_index_from_string

# Sabitler ve yardımcı fonksiyonlar
COLUMNS_IMPORT = {
    "A": "TCGB Gümrük İdaresi",
    "B": "TCGB Tescil No",
    "C": "TCGB Tescil Tarihi",
    "D": "Alıcı / Gönderici VKN",
    "E": "Alıcı / Gönderici Unvan",
    "F": "Beyan Sahibi",
    "G": "Beyan Sahibi Unvan",
    "H": "Kalem No",
    "I": "Satır Kodu",
    "J": "Devredilen Satır Kodu",
    "K": "GTIP No (8li)",
    "L": "GTİP Kodu (12 li)",
    "M": "GTİP açıklaması",
    "N": "Madde Adı",
    "O": "Tamamlayıcı Ölçü Birim",
    "P": "Miktar",
    "Q": "Brüt Kg",
    "R": "Net Kg",
    "S": "İstatistiki Birim Kodu",
    "T": "İstatistiki Miktar",
    "U": "İstatistiki Kıymet ($)",
    "V": "Kalem Rejim Kodu",
    "W": "Menşe Ülke Adı",
    "X": "Sevk Ülkesi",
    "Y": "Çıkış Ülkesi",
    "Z": "Varış Ülkesi",
    "AA": "Ticaret Yapılan Ülke",
    "AB": "Kap Ürün Bilgisi",
    "AC": "Özel Durum",
    "AD": "Açıklama44 Beyanname",
    "AE": "Kalem Açıklama44",
    "AF": "Muafiyet Kodu",
    "AG": "Fatura Bedeli",
    "AH": "Döviz Türü",
    "AI": "Fatura",
    "AJ": "Atr",
    "AK": "Atr Tarihi",
    "AL": "E-Atr",
    "AM": "E_Atr Tarihi",
    "AN": "Eur1",
    "AO": "Eur1 Tarihi",
    "AP": "E-Eur1",
    "AQ": "E-Eur1 Tarihi",
    "AR": "Eur1med",
    "AS": "Eur1Med Tarihi",
    "AT": "E-Eur1med",
    "AU": "E-Eur1med Tarihi",
    "AV": "Eur1 Fatura",
    "AW": "Eur1Fatura Tarihi",
    "AX": "Eur1 Sertifika",
    "AY": "Fatura Tarihi",
    "AZ": "Form AV",
    "BA": "Form A Tarihi",
    "BB": "Inf2",
    "BC": "Inf2 Tarihi",
    "BD": "Tedarikçi Beyan",
    "BE": "Tedarikçi Beyan Tarihi",
    "BF": "Tam Beyan Usül",
    "BG": "Navlun Tutarı",
    "BH": "Navlun Dvz",
    "BI": "Sigorta Tutarı",
    "BJ": "Sigorta Dvz",
    "BK": "GK8DI",
    "BL": "GK168",
    "BM": "GK8",
    "BN": "Karşı Unvan",
    "BO": "Telafi Edici Beyan",
    "BP": "Gümrük Vergisi Kalem Rto",
    "BQ": "Gümrük Vergisi USD",
    "BR": "İGV Kalem Rto",
    "BS": "İGV USD",
    "BT": "TEV Kalem Rto",
    "BU": "TEV Kalem USD",
    "BV": "KDV Rto",
    "BW": "KDV Usd",
    "BX": "OTV Rto",
    "BY": "OTV Usd",
    "BZ": "Döviz Kuru",
    "CA": "Döviz/Usd Kur",
    "CB": "Banka Adı",
}

COLUMNS_EXPORT = {
    "A": "TCGB Gümrük İdaresi",
    "B": "TCGB Tescil No",
    "C": "TCGB Tescil Tarihi",
    "D": "Kapanma Tarihi",
    "E": "Alıcı / Gönderici Vergi No",
    "F": "Alıcı / Gönderici Unvan",
    "G": "Beyan Sahibi",
    "H": "Beyan Sahibi Unvan",
    "I": "Kalem No",
    "J": "Satır Kodu",
    "K": "Devredilen Satır Kodu",
    "L": "GTIP No (8li)",
    "M": "GTİP Kodu (12 li)",
    "N": "GTİP açıklaması",
    "O": "Madde Adı",
    "P": "Tamamlayıcı Ölçü Birim",
    "Q": "Miktar",
    "R": "Brüt Kg",
    "S": "Net Kg",
    "T": "İstatistiki Birim Kodu",
    "U": "İstatistiki Miktar",
    "V": "İstatistiki Kıymet ($)",
    "W": "Kalem Rejim Kodu",
    "X": "Menşe Ülke Adı",
    "Y": "Sevk Ülkesi",
    "Z": "Çıkış Ülkesi",
    "AA": "Varış Ülkesi",
    "AB": "Ticaret Yapılan Ülke",
    "AC": "Kap Ürün Bilgisi",
    "AD": "Özel Durum",
    "AE": "Açıklama44 Beyanname",
    "AF": "Kalem Açıklama44",
    "AG": "Muafiyet Kodu",
    "AH": "Fatura Bedeli",
    "AI": "Döviz Türü",
    "AJ": "Fatura",
    "AK": "Atr",
    "AL": "Atr Tarihi",
    "AM": "E-Atr",
    "AN": "E-Atr Tarihi",
    "AO": "Eur1",
    "AP": "Eur1 Tarihi",
    "AQ": "E-Eur1",
    "AR": "E-Eur1 Tarihi",
    "AS": "Eur1med",
    "AT": "Eur1Med Tarihi",
    "AU": "E-Eur1med",
    "AV": "E-Eur1med Tarihi",
    "AW": "Eur1 Fatura",
    "AX": "Eur1Fatura Tarihi",
    "AY": "Eur1 Sertifika",
    "AZ": "Fatura Tarihi",
    "BA": "Form A",
    "BB": "Form A Tarihi",
    "BC": "Inf2",
    "BD": "Inf2 Tarihi",
    "BE": "Tedarikçi Beyan",
    "BF": "Tedarikçi Beyan Tarihi",
    "BG": "Tam Beyan Usül",
    "BH": "Navlun Tutarı",
    "BI": "Navlun Dvz",
    "BJ": "Sigorta Tutarı",
    "BK": "Sigorta Dvz",
    "BL": "GK8DI",
    "BM": "GK168",
    "BN": "GK8",
    "BO": "Karşı Unvan",
    "BP": "Telafi Edici Beyan",
    "BQ": "TEV Kalem Rto",
    "BR": "TEV Kalem USD",
    "BS": "Döviz Kuru",
    "BT": "Döviz/Usd Kur",
    "BU": "Banka Adı"
}
UCUNCU_DUNYA_ULKELERI = [
    "ÜRDÜN", "FAREO ADALARI", "GÜNEY KORE", "MALEZYA", "SİNGAPUR",
    "ABD VİRJİN ADALARI", "AFGANİSTAN", "A.B.D.", "AMERİKAN OKYANUSYASI", 
    "ANDORRA", "ANGOLA", "ANGUILLA", "ANTIGUA VE BERMUDA", "ARJANTİN", 
    "ARUBA", "AVUSTRALYA", "AVUSTRALYA OKYANUSU", "AVUSTURYA", 
    "AZERBEYCAN-NAHÇIVAN", "BAHAMA", "BAHREYN", "BANGLADEŞ", "BARBADOS", 
    "BELİZE", "BENİN", "BERMUDA", "BEYAZ RUSYA", "BHUTAN", 
    "BİRLEŞİK ARAP EMİRLİKLERİ", "BOLİVYA", "BOSTVANA", "BREZİLYA", "BRUNEİ",
    "BURKİNA FASO", "BURMA", "BURUNDİ", "CAPE VERDE", "CAYMAN ADALARI",
    "CEBELİ TARIK", "CEUTA VE MELİLLA", "CEZAYİR", "CİBUTİ", "COOK ADALARI",
    "ÇAD", "ÇEÇEN CUMHURİYETİ", "ÇİN HALK CUMHURİYETİ", "DAĞISTAN CUMHURİYETİ",
    "DOMİNİK CUMHURİYETİ", "DOMİNİKA", "DUBAİ", "EKVATOR", "EKVATOR GİNESİ",
    "EL SALVADOR", "ENDONEZYA", "ERMENİSTAN", "ETİYOPYA", "FALKLAND ADALARI",
    "FAROE ADALARI", "FİJİ", "FİLDİŞİ SAHİLİ", "FİLİPİNLER", "FRANSIZ GUYANASI",
    "GABON", "GAMBİYA", "GANA", "GİNE", "GİNE-BİSSAU", "GRENADA",
    "GRÖNLAND", "GUADELUP", "GUATEMALA", "GUYANA", "GÜNEY AFRİKA CUMHURİYETİ",
    "GÜNEY YEMEN", "HAİTİ", "HİNDİSTAN", "HOLLANDA ANTİLLERİ", "HONDURAS",
    "HONG KONG", "IRAK", "İNGİLİZ HİNT OKY.TOPRAKLARI", "İNGİLİZ VİRJİN ADALARI",
    "İRAN", "JAMAİKA", "JAPONYA", "KAMBOÇYA", "KAMERUN", "KANADA",
    "KANARYA ADALARI", "KATAR", "KAZAKİSTAN", "KENYA", "KIRGIZİSTAN",
    "KİRİBATİ", "KOLOMBİYA", "KOMORO ADALARI", "KONGO", "KOSTA RİKA",
    "KUVEYT", "KUZEY KIBRIS T.C.", "KUZEY KORE DEMOKRATİK HALK CUM.",
    "KUZEY YEMEN", "KÜBA", "LAOS", "LESOTHO", "LİBERYA", "LİBYA", "LÜBNAN",
    "MADAGASKAR", "MAKAO", "MALAVİ", "MALDİV ADALARI", "MALİ", "MARTİNİK",
    "MAYOTTE", "MEKSİKA", "MERKEZİ AFRİKA CUMHURİYETİ", "MOĞOLİSTAN", "MONACO",
    "MORİTANYA", "MOZAMBİK", "NAMİBYA", "NAURU", "NEPAL", "NİJER", "NİJERYA",
    "NİKARAGUA", "ÖZBEKİSTAN", "PAKİSTAN", "PANAMA", "PAPUA YENİ GİNE",
    "PARAGUAY", "PERU", "PİTCAİRN", "REUNİON", "RUANDA", "RUM KESİMİ",
    "RUSYA FEDERASYONU", "SAO TOME AND PRINCIPE", "SENEGAL",
    "SEYŞEL ADALARI VE BAĞLANTILARI", "SIERRA LEONE", "SOLOMON ADALARI",
    "SOMALİ", "SRİ LANKA", "ST. CHRİSTOPHER VE NEVİS", "ST. HELENA VE BAĞLANTILARI",
    "ST. LUCİA", "ST. PİERRE VE MİQUELON", "ST. VİNCENT", "SUDAN", "SURİNAM",
    "SUUDİ ARABİSTAN", "SVAZİLAND", "TACİKİSTAN", "TANZANYA", "TATARİSTAN",
    "TAYLAND", "TAYVAN", "TOGO", "TONGA", "TRİNİDAD VE TOBAGO",
    "TURKS VE CAİCOS ADASI", "TUVALU", "TÜRKİYE", "TÜRKMENİSTAN", "UGANDA",
    "UKRAYNA", "UMMAN", "URUGUAY", "VANUATU", "VATİKAN", "VENEZUELLA",
    "VİETNAM", "WALLİS VE FUTUNA ADALARI", "YAKUTİSTAN",
    "YENİ KALODENYA VE BAĞLANTILARI", "YENİ ZELANDA", "YENİ ZELANDA OKYANUSU",
    "YUGOSLAVYA", "ZAİRE", "ZAMBİA", "ZİMBABVE", "HONG-KONG", "ÇİN HALK CUMHUR.",
    "BİR.ARAP EMİRLİK.", "GÜNEY KORE CUM.", "VİETNAM SOSYALİST", "KOLOMBİA",
    "GÜNEY AFRİKA CUM.", "AZERBAYCAN-NAHÇ.", "KOSTARİKA"
]

AB_COUNTRIES = [
    "ALMANYA", "AVUSTURYA", "BELÇİKA", "BULGARİSTAN", "ÇEKYA", "DANİMARKA", "ESTONYA", "FİNLANDİYA",
    "FRANSA", "HİRVATİSTAN", "HOLLANDA", "İRLANDA", "İSPANYA", "İSVEC", "İTALYA", "LETONYA", "LİTVANYA",
    "LÜKSEMBURG", "MACARİSTAN", "MALTA", "POLONYA", "PORTEKİZ", "ROMANYA", "SLOVAKYA", "SLOVENYA", "YUNANİSTAN",
    "NORVEÇ", "İSVİÇRE", "İZLANDA", "LİHTENŞTAYN",  "NORVEÇ", "İSVİÇRE", "İZLANDA", "LİHTENŞTAYN", "ŞİLİ", "FİLİSTİN", "SIRBİSTAN", "KARADAĞ",
    "GÜRCİSTAN", "ARNAVUTLUK", "BOSNA HERSEK", "İSRAİL", "MAKEDONYA", "GÜNEY KORE", "MORİTYUS", "MOLDOVA", "FİLİSTİN (GAZZE)", "GAZZE"
]

OZEL_ULKELER = [
    "NORVEÇ", "İSVİÇRE", "İZLANDA", "LİHTENŞTAYN", "ŞİLİ", "FİLİSTİN", "SIRBİSTAN", "KARADAĞ",
    "GÜRCİSTAN", "ARNAVUTLUK", "BOSNA HERSEK", "İSRAİL", "MAKEDONYA", "GÜNEY KORE", "MORİTYUS", "MOLDOVA", "FİLİSTİN (GAZZE)", "GAZZE"
]

def apply_column_mapping(df, mapping):
    mapped_columns = {}
    for col_letter, col_name in mapping.items():
        col_index = column_index_from_string(col_letter) - 1
        if col_index < len(df.columns):
            mapped_columns[df.columns[col_index]] = col_name
    return df.rename(columns=mapped_columns)

def load_excel(uploaded_file):
    try:
        sheet_data = pd.read_excel(uploaded_file, sheet_name=None)
        for sheet_name, df in sheet_data.items():
            if "İth" in sheet_name or "Gerç.İth" in sheet_name:
                df = apply_column_mapping(df, COLUMNS_IMPORT)
            elif "İhr" in sheet_name or "Gerç.İhr" in sheet_name:
                df = apply_column_mapping(df, COLUMNS_EXPORT)
            sheet_data[sheet_name] = df
        return sheet_data
    except Exception as e:
        st.error(f"Excel yükleme hatası: {str(e)}")
        return None

def filter_imports(df):
    columns = [
        'TCGB Gümrük İdaresi', 'TCGB Tescil No', 'TCGB Tescil Tarihi',
        'Alıcı / Gönderici Unvan', 'Kalem No', 'Satır Kodu', 'Atr', 'E-Atr',
        'Eur1', 'E-Eur1', 'Eur1med', 'E-Eur1med',
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

def get_madde_blok(df: pd.DataFrame, start_idx: int, madde_adi: str):
    blok = []

    for j in range(start_idx, min(start_idx + 4, len(df))):
        row = df.iloc[j]
        param = str(row["Parametreler"]).lower().strip()

        if (
            "birim kullanım miktarı" in param or
            "fire" in param or
            "toplam birim kullanım" in param
        ):
            row_copy = row.copy()
            if pd.isna(row_copy["Madde Adı"]) or row_copy["Madde Adı"] == "":
                row_copy["Madde Adı"] = madde_adi
            blok.append(row_copy)

    return blok

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

        ab_df, non_ab_df = filter_exports(ihracat_df_all, AB_COUNTRIES)
        ab_pivot = ab_df.groupby('Satır Kodu')['İstatistiki Miktar'].sum().reset_index().rename(columns={'İstatistiki Miktar': 'Toplam Miktar'})
        non_ab_pivot = non_ab_df.groupby('Satır Kodu')['İstatistiki Miktar'].sum().reset_index().rename(columns={'İstatistiki Miktar': 'Toplam Miktar'})
        ozel_df = ihracat_df_all[ihracat_df_all['Varış Ülkesi'].isin(OZEL_ULKELER)]
        kontrol_df = ozel_df.copy() if not ozel_df.empty else pd.DataFrame()

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

            if "birim kullanım miktarı" in str(satir["Parametreler"]).lower():
                madde_adi = satir["Madde Adı"]

                match_index = sarfiyat_df_all[
                    (sarfiyat_df_all["Madde Adı"] == madde_adi) &
                    (sarfiyat_df_all["Parametreler"].str.lower().str.contains("birim kullanım miktarı"))
                ].index

                if not match_index.empty:
                    blok_satirlari = get_madde_blok(sarfiyat_df_all, match_index[0], madde_adi)
                    for row in blok_satirlari[1:]:  # ilk satır zaten eklendi
                        yeni_row = {col: "" for col in sarfiyat_df.columns}
                        for col in sarfiyat_df.columns:
                            if col in row:
                                yeni_row[col] = row[col]
                        yeni_sarfiyat_df.loc[len(yeni_sarfiyat_df)] = yeni_row

            i += 1

        sarfiyat_df = yeni_sarfiyat_df.reset_index(drop=True)

        # Hesaplanan Toplam Mamul Kullanımını alt satıra ekle
        yeni_df_with_mamul = pd.DataFrame(columns=sarfiyat_df.columns)
        i = 0
        while i < len(sarfiyat_df):
            satir = sarfiyat_df.iloc[i]

            if str(satir["Parametreler"]).lower().startswith("birim kullanım miktarı"):
                madde_adi = satir["Madde Adı"]
                mamul_row = {col: "" for col in sarfiyat_df.columns}
                mamul_row["Madde Adı"] = madde_adi
                mamul_row["Parametreler"] = "Toplam Mamul Kullanımı"

                for col in sarfiyat_df.columns:
                    if col not in ["Madde Adı", "Parametreler"] and pd.notna(satir[col]):
                        try:
                            birim = float(satir[col])
                            miktar = float(toplam_miktar_row.get(col, 0))
                            mamul_row[col] = birim * miktar
                        except:
                            mamul_row[col] = ""

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
                (sarfiyat_df["Parametreler"].str.lower().str.startswith("birim kullanım miktarı"))
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
            if not kontrol_df.empty:
                kontrol_df.to_excel(writer, sheet_name="Kontrol Listesi", index=False)
            if sarfiyat_df is not None:
                sarfiyat_df.to_excel(writer, sheet_name="Sarfiyat", index=False)

        output_combined.seek(0)
        st.download_button("Tüm Verileri İndir", data=output_combined, file_name="tüm_veriler_raporu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

