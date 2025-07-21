import streamlit as st
import pandas as pd
import re
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
    "AU": "E-Eur1med Tarihi",
    "AV": "Eur1 Fatura",
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

def secili_sarfiyat_sayfasi(s1: pd.DataFrame, s2: pd.DataFrame = None):
    # s2 parametresini isteğe bağlı yapıyoruz
    if s2 is not None:
        try:
            parametreler = s1["Parametreler"].dropna().astype(str).str.lower()
            if parametreler.str.contains(r"birim kullanım miktarı\s*\((kg|kilo)\)").any():
                print("4. sarfiyat sayfası (kg) kullanıldı.")
                return s2  # 4. sayfa
            else:
                print("3. sarfiyat sayfası kullanıldı.")
                return s1  # 3. sayfa
        except Exception as e:
            print(f"Sarfiyat sayfası seçim hatası: {e}. Varsayılan olarak 3. sayfa kullanıldı.")
            return s1  # varsayılan olarak s1 kullan
    else:
        print("Sadece 3. sarfiyat sayfası mevcut veya 4. sayfa kullanılmadı.")
        return s1 # Sadece 3. sayfa varsa onu kullan

def birim_turleri_satir4ten(sarfiyat_df: pd.DataFrame, use_second_unit_sheet: bool):
    """
    3. veya 4. sayfadaki sarfiyat sayfasının 3. (kod) ve 4. (birim) satırlarından ürün birimi belirler.
    Dönüş: {ürün_kodu: "kg" | "non-kg"}
    """
    birim_map = {}
    if not use_second_unit_sheet:
        st.info("İkinci birim sayfası kullanılmadığı için birim türleri 3. sayfadan belirleniyor.")
        # Eğer 2. birim sayfası kullanılmıyorsa, tüm birimleri 'non-kg' olarak varsayabiliriz
        # veya sarfiyat_df_all'dan ilgili satırları çekip birim türünü belirleyebiliriz.
        # Basitçe 'non-kg' varsayımı:
        for col in sarfiyat_df.columns:
            if col not in ["Madde Adı", "Parametreler"]:
                birim_map[str(col)] = "non-kg" # Satır kodları kolon başlıklarında
        return birim_map

    try:
        kod_satiri = sarfiyat_df.iloc[2]  # Satır kodları
        birim_satiri = sarfiyat_df.iloc[3]  # Birim bilgileri

        for col in sarfiyat_df.columns:
            if col in ["Madde Adı", "Parametreler"]:
                continue
            urun_kodu = str(kod_satiri[col]).strip()
            birim = str(birim_satiri[col]).strip().lower()

            if urun_kodu == "" or birim == "":
                continue
            elif "kg" in birim or "kilo" in birim:
                birim_map[urun_kodu] = "kg"
            else:
                birim_map[urun_kodu] = "non-kg"
    except Exception as e:
        st.warning(f"Birim türü alınamadı: {e}")
    return birim_map

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
        sarfiyat1 = data[sheet_names[2]] # 3. sayfa her zaman var

        sarfiyat2_exists = len(sheet_names) >= 4 # 4. sayfa var mı kontrol et
        sarfiyat2 = data[sheet_names[3]] if sarfiyat2_exists else None

        ithalat_df = filter_imports(ithalat_df_all)
        ithalat_pivot = ithalat_df.groupby(['Satır Kodu'])['İstatistiki Miktar'].sum().reset_index().rename(columns={'İstatistiki Miktar': 'Toplam İstatistiki Miktar'})

        ab_df, non_ab_df = filter_exports(ihracat_df_all, AB_COUNTRIES)
        ab_pivot = ab_df.groupby('Satır Kodu')['İstatistiki Miktar'].sum().reset_index().rename(columns={'İstatistiki Miktar': 'Toplam Miktar'})
        non_ab_pivot = non_ab_df.groupby('Satır Kodu')['İstatistiki Miktar'].sum().reset_index().rename(columns={'İstatistiki Miktar': 'Toplam Miktar'})
        ozel_df = ihracat_df_all[ihracat_df_all['Varış Ülkesi'].isin(OZEL_ULKELER)]
        kontrol_df = ozel_df.copy() if not ozel_df.empty else pd.DataFrame()

        # Sarfiyat sayfası seçimi
        sarfiyat_df_all = secili_sarfiyat_sayfasi(sarfiyat1, sarfiyat2)
        st.info("Kullanılan sarfiyat sayfası: " + (sheet_names[3] if sarfiyat2_exists and sarfiyat_df_all.equals(data[sheet_names[3]]) else sheet_names[2]))

        # SARFİYAT: satır kodu hem "Madde Adı" sütununda hem de kolon adlarında kontrol edilir
        sarfiyat_df = None
        if sarfiyat_df_all is not None:
            ithalat_kodlari = ithalat_pivot['Satır Kodu'].astype(str).unique()
            dunya_kodlari = non_ab_pivot['Satır Kodu'].astype(str).unique()

            satir_maskesi = sarfiyat_df_all['Madde Adı'].astype(str).isin(ithalat_kodlari)
            kolonlar = ['Madde Adı', 'Parametreler'] + [col for col in sarfiyat_df_all.columns if str(col) in dunya_kodlari]
            
            # Tüm kolonların sarfiyat_df_all içinde olduğundan emin olun
            kolonlar = [col for col in kolonlar if col in sarfiyat_df_all.columns]
            
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

        birim_turleri = birim_turleri_satir4ten(sarfiyat_df_all, sarfiyat2_exists)

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
        # sarfiyat_df_all, seçilen sarfiyat sayfasıdır (sarfiyat1 veya sarfiyat2)
        urun_adi_satiri_3 = sarfiyat1.iloc[3] # 3. sayfanın 4. satırı (ürün adları)

        # Yeni satır: Kullanılan Ürün (3. Sayfadan)
        kullanilan_urun_row_3 = {
            "Madde Adı": "",
            "Parametreler": "Kullanılan Ürün"
        }

        for col in sarfiyat_df.columns:
            if col not in ["Madde Adı", "Parametreler"] and col in urun_adi_satiri_3.index:
                kullanilan_urun_row_3[col] = urun_adi_satiri_3[col]

        # "Toplam Miktar" satırı zaten en başta
        sarfiyat_df = pd.concat([
            sarfiyat_df.iloc[[0]],                          # Toplam Miktar
            pd.DataFrame([kullanilan_urun_row_3]),         # Kullanılan Ürün (3. Sayfa)
            sarfiyat_df.iloc[1:]                           # Kalan sarfiyat
        ], ignore_index=True)

        # Eğer 4. sayfa varsa ve kullanılıyorsa, "2.Br Birim Kullanım" satırını ekle
        if sarfiyat2_exists and sarfiyat_df_all.equals(sarfiyat2):
            urun_adi_satiri_4 = sarfiyat2.iloc[1] # 4. sayfanın 2. satırı (örneğin C3:D3:E3:F3 gibi satır varsa index=1)

            kullanilan_urun_row_4 = {col: "" for col in sarfiyat_df.columns}
            kullanilan_urun_row_4["Madde Adı"] = "2.Br Birim Kullanım"
            kullanilan_urun_row_4["Parametreler"] = ""

            for col in sarfiyat_df.columns:
                if col not in ["Madde Adı", "Parametreler"] and col in urun_adi_satiri_4.index:
                    kullanilan_urun_row_4[col] = urun_adi_satiri_4[col]
            
            # "Toplam Miktar" ve "Kullanılan Ürün (3. Sayfa)" satırından sonra ekle
            sarfiyat_df = pd.concat([
                sarfiyat_df.iloc[[0,1]],                  # Toplam Miktar, Kullanılan Ürün (3. Sayfa)
                pd.DataFrame([kullanilan_urun_row_4]),    # Kullanılan Ürün (4. Sayfa)
                sarfiyat_df.iloc[2:]                      # Diğer satırlar
            ], ignore_index=True)
        
        # Fire ve Toplam Birim Kullanım satırlarını ekleme
        yeni_sarfiyat_df = pd.DataFrame(columns=sarfiyat_df.columns)

        i = 0
        while i < len(sarfiyat_df):
            satir = sarfiyat_df.iloc[i]
            yeni_sarfiyat_df.loc[len(yeni_sarfiyat_df)] = satir

            if "birim kullanım miktarı" in str(satir["Parametreler"]).lower():
                madde_adi = satir["Madde Adı"]
                parametre = str(satir["Parametreler"]).lower() if pd.notnull(satir["Parametreler"]) else ""

                # Parantez içindeki birimi al (regex ile)
                match = re.search(r'\((.*?)\)', parametre)
                parantez_ici = match.group(1) if match else ""

                # Kaynak sayfa seçimi
                kaynak_sarfiyat = sarfiyat_df_all # Sarfiyat_df_all zaten seçilen (3. veya 4.) sayfayı temsil ediyor

                match_index = kaynak_sarfiyat[
                    (kaynak_sarfiyat["Madde Adı"] == madde_adi) &
                    (kaynak_sarfiyat["Parametreler"].str.lower().str.contains("birim kullanım miktarı"))
                ].index

                if not match_index.empty:
                    blok_satirlari = get_madde_blok(kaynak_sarfiyat, match_index[0], madde_adi)
                    for row_blok in blok_satirlari[1:]:  # ilk satır zaten eklendi
                        yeni_row = {col: "" for col in sarfiyat_df.columns}
                        for col in sarfiyat_df.columns:
                            if col in row_blok:
                                yeni_row[col] = row_blok[col]
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

                parametre = str(satir["Parametreler"]).lower() if pd.notnull(satir["Parametreler"]) else ""
                match = re.search(r'\((.*?)\)', parametre)
                birim_turu = match.group(1).lower() if match else ""

                # Fire satırını bul
                fire_row_index = -1
                if i + 1 < len(sarfiyat_df):
                    temp_fire_row = sarfiyat_df.iloc[i + 1]
                    if str(temp_fire_row["Parametreler"]).lower().startswith("fire"):
                        fire_row_index = i + 1

                # Toplam Birim Kullanım satırını bul
                toplam_birim_kullanim_row_index = -1
                if fire_row_index != -1 and i + 2 < len(sarfiyat_df): # Fire satırı varsa 2. sonrakine bak
                    temp_toplam_row = sarfiyat_df.iloc[i + 2]
                    if str(temp_toplam_row["Parametreler"]).strip().lower() == "toplam birim kullanım":
                        toplam_birim_kullanim_row_index = i + 2
                elif fire_row_index == -1 and i + 1 < len(sarfiyat_df): # Fire satırı yoksa 1. sonrakine bak
                    temp_toplam_row = sarfiyat_df.iloc[i + 1]
                    if str(temp_toplam_row["Parametreler"]).strip().lower() == "toplam birim kullanım":
                        toplam_birim_kullanim_row_index = i + 1

                toplam_birim_kullanim_row = sarfiyat_df.iloc[toplam_birim_kullanim_row_index] if toplam_birim_kullanim_row_index != -1 else None


                for col in sarfiyat_df.columns:
                    if col not in ["Madde Adı", "Parametreler"] and toplam_birim_kullanim_row is not None and pd.notna(toplam_birim_kullanim_row.get(col)):
                        try:
                            toplam_birim_kullanim = float(toplam_birim_kullanim_row[col])
                            
                            if "kg" in birim_turu or "kilo" in birim_turu:
                                katsayi = 1.0 # Varsayılan katsayı
                                if sarfiyat2_exists and sarfiyat_df_all.equals(sarfiyat2): # Eğer 4. sayfa kullanılıyorsa katsayıyı oradan al
                                    ikinci_birim_row = sarfiyat_df[sarfiyat_df["Madde Adı"] == "2.Br Birim Kullanım"].iloc[0]
                                    if col in ikinci_birim_row and pd.notna(ikinci_birim_row[col]):
                                        katsayi = float(ikinci_birim_row[col])
                                mamul_row[col] = toplam_birim_kullanim * katsayi
                            elif any(x in birim_turu for x in ["litre", "metre", "adet"]):
                                miktar = float(toplam_miktar_row.get(col, 0))
                                mamul_row[col] = toplam_birim_kullanim * miktar
                            else: # Bilinmeyen birim türleri için doğrudan toplam birim kullanımı alınabilir
                                mamul_row[col] = toplam_birim_kullanim
                        except Exception as e:
                            mamul_row[col] = ""
            
                yeni_df_with_mamul.loc[len(yeni_df_with_mamul)] = satir # Birim kullanım satırını ekle

                if fire_row_index != -1:
                    yeni_df_with_mamul.loc[len(yeni_df_with_mamul)] = sarfiyat_df.iloc[fire_row_index] # Fire satırını ekle
                    i += 1
                if toplam_birim_kullanim_row_index != -1:
                    yeni_df_with_mamul.loc[len(yeni_df_with_mamul)] = sarfiyat_df.iloc[toplam_birim_kullanim_row_index] # Toplam Birim Kullanım satırını ekle
                    i += 1

                yeni_df_with_mamul.loc[len(yeni_df_with_mamul)] = mamul_row # Toplam Mamul Kullanımı'nı ekle
                

            else:
                # Diğer tüm satırları aynen aktar
                yeni_df_with_mamul.loc[len(yeni_df_with_mamul)] = satir

            i += 1

        sarfiyat_df = yeni_df_with_mamul.reset_index(drop=True)

        # 1. İlk satır: "Toplam Miktar" satırı
        toplam_miktar_row = sarfiyat_df.iloc[0]

        # 1. "Toplam Mamul Kullanımı" satırlarını filtrele
        mamul_maskesi = sarfiyat_df["Parametreler"] == "Toplam Mamul Kullanımı"

        # "Toplam Mamul Kullanımı Toplam" sütununu oluştur ve her satırın yatay toplamını yaz
        sarfiyat_df.loc[mamul_maskesi, "Toplam Mamul Kullanımı Toplam"] = sarfiyat_df[mamul_maskesi].apply(
            lambda row: sum([
                pd.to_numeric(row[col], errors='coerce') for col in sarfiyat_df.columns
                if col not in ["Madde Adı", "Parametreler", "Gerçekleşen İthalat Miktarı", "Fark", "TEV Durumu", "Toplam Mamul Kullanımı Toplam"]
                and pd.notna(row[col])
            ]),
            axis=1
        )

        for idx, row in sarfiyat_df[mamul_maskesi].iterrows():
            madde_adi = row["Madde Adı"]

            # Birim Kullanım Miktarı satırını bul
            birim_kullanim_satiri = sarfiyat_df[
                (sarfiyat_df["Madde Adı"] == madde_adi) &
                (sarfiyat_df["Parametreler"].str.lower().str.startswith("birim kullanım miktarı"))
            ]

            if birim_kullanim_satiri.empty:
                continue

            parametre = birim_kullanim_satiri.iloc[0]["Parametreler"]
            match = re.search(r'\((.*?)\)', str(parametre))
            birim_turu = match.group(1).lower() if match else ""

            toplam_birim_kullanim_row_found = False
            toplam_birim_kullanim_row = None
            
            # Fire satırı var mı kontrol et
            current_idx_of_birim_kullanim = sarfiyat_df.index.get_loc(idx) - (2 if sarfiyat2_exists and sarfiyat_df_all.equals(sarfiyat2) else 1) # Güncel mamul satırının üstündeki birim kullanım satırının index'i
            
            if current_idx_of_birim_kullanim + 2 < len(sarfiyat_df): # Fire ve Toplam Birim Kullanım satırları olabilir
                if str(sarfiyat_df.iloc[current_idx_of_birim_kullanim + 1]["Parametreler"]).lower().startswith("fire"): # Fire satırı varsa
                    if current_idx_of_birim_kullanim + 2 < len(sarfiyat_df) and str(sarfiyat_df.iloc[current_idx_of_birim_kullanim + 2]["Parametreler"]).strip().lower() == "toplam birim kullanım":
                        toplam_birim_kullanim_row = sarfiyat_df.iloc[current_idx_of_birim_kullanim + 2]
                        toplam_birim_kullanim_row_found = True
                elif str(sarfiyat_df.iloc[current_idx_of_birim_kullanim + 1]["Parametreler"]).strip().lower() == "toplam birim kullanım": # Fire satırı yok ama Toplam Birim Kullanım satırı varsa
                    toplam_birim_kullanim_row = sarfiyat_df.iloc[current_idx_of_birim_kullanim + 1]
                    toplam_birim_kullanim_row_found = True

            if not toplam_birim_kullanim_row_found:
                continue # Toplam birim kullanım satırı bulunamazsa bu maddeyi atla

            mamul_row_calculated = {col: "" for col in sarfiyat_df.columns}
            mamul_row_calculated["Madde Adı"] = madde_adi
            mamul_row_calculated["Parametreler"] = "Toplam Mamul Kullanımı"

            for col in sarfiyat_df.columns:
                if col in ["Madde Adı", "Parametreler"]:
                    continue
                try:
                    toplam_birim_kullanim = float(toplam_birim_kullanim_row.get(col, 0))

                    if "kg" in birim_turu or "kilo" in birim_turu:
                        katsayi = 1.0 # Varsayılan katsayı
                        if sarfiyat2_exists and sarfiyat_df_all.equals(sarfiyat2): # Eğer 4. sayfa kullanılıyorsa katsayıyı oradan al
                            ikinci_birim_row_filtered = sarfiyat_df[sarfiyat_df["Madde Adı"] == "2.Br Birim Kullanım"]
                            if not ikinci_birim_row_filtered.empty and col in ikinci_birim_row_filtered.iloc[0] and pd.notna(ikinci_birim_row_filtered.iloc[0][col]):
                                katsayi = float(ikinci_birim_row_filtered.iloc[0][col])
                        mamul_row_calculated[col] = toplam_birim_kullanim * katsayi

                    elif any(x in birim_turu for x in ["litre", "metre", "adet"]):
                        toplam_miktar = toplam_miktar_row.get(col, 0)
                        mamul_row_calculated[col] = toplam_birim_kullanim * toplam_miktar
                    else:
                        mamul_row_calculated[col] = toplam_birim_kullanim # Varsayılan olarak direkt toplam birim kullanımı

                except Exception as e:
                    mamul_row_calculated[col] = ""
                    
            # Mamul satırını güncelle
            for k, v in mamul_row_calculated.items():
                sarfiyat_df.loc[idx, k] = v

        # Toplam Mamul Kullanımı Toplamını tekrar hesapla
        sarfiyat_df.loc[mamul_maskesi, "Toplam Mamul Kullanımı Toplam"] = sarfiyat_df[mamul_maskesi].apply(
            lambda row: sum(
                pd.to_numeric(row[col], errors='coerce') 
                for col in sarfiyat_df.columns 
                if col not in ["Madde Adı", "Parametreler", "Gerçekleşen İthalat Miktarı", "Fark", "TEV Durumu", "Toplam Mamul Kullanımı Toplam"]
            ) if row["Parametreler"] == "Toplam Mamul Kullanımı" else "", axis=1
        )


        for idx, row in sarfiyat_df[mamul_maskesi].iterrows():
            madde_adi = row["Madde Adı"]
            # ithalat_pivot'tan ilgili Madde Adı'nın Toplam İstatistiki Miktarını bul
            ithalat_miktar_row = ithalat_pivot[ithalat_pivot['Satır Kodu'] == madde_adi]
            if not ithalat_miktar_row.empty:
                sarfiyat_df.loc[idx, "Gerçekleşen İthalat Miktarı"] = ithalat_miktar_row.iloc[0]['Toplam İstatistiki Miktar']
            else:
                sarfiyat_df.loc[idx, "Gerçekleşen İthalat Miktarı"] = 0 
                
        sarfiyat_df.loc[mamul_maskesi, "Fark"] = (
            pd.to_numeric(sarfiyat_df.loc[mamul_maskesi, "Toplam Mamul Kullanımı Toplam"], errors='coerce') -
            pd.to_numeric(sarfiyat_df.loc[mamul_maskesi, "Gerçekleşen İthalat Miktarı"], errors='coerce')
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