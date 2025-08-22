import streamlit as st
import pandas as pd
import re
import numpy as np
from io import BytesIO
from pandas import DataFrame
from openpyxl.utils import column_index_from_string

# Sabitler (neden: magic string ve kolon adlarını tek yerde toplamak)
COL_MADDE = "Madde Adı"
COL_PARAM = "Parametreler"
ROW_TOPLAM_MAMUL = "Toplam Mamul Kullanımı"
ROW_TOPLAM_BIRIM = "Toplam Birim Kullanım"
ROW_FIRE = "Fire"
ROW_TOPLAM_MIKTAR = "Toplam Miktar"
ROW_KULLANILAN_URUN = "Kullanılan Ürün"
ROW_2BR_BIRIM = "2.Br Birim Kullanım"
COL_GERCEK_ITH_MIK = "Gerçekleşen İthalat Miktarı"
COL_FARK = "Fark"
COL_TEV = "TEV Durumu"
COL_TMK_TOPLAM = "Toplam Mamul Kullanımı Toplam"

ZORUNLU_SUTUNLAR_ITHALAT = ["Özel Durum", "Gümrük Vergisi Kalem Rto"]
ZORUNLU_SUTUNLAR_IHRACAT = ["Varış Ülkesi"]

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


def load_excel(uploaded_file):
    try:
        sheet_data = pd.read_excel(uploaded_file, sheet_name=None)
        for sheet_name, df in sheet_data.items():
            required_columns = []
            if "İth" in sheet_name or "Gerç.İth" in sheet_name:
                required_columns = ZORUNLU_SUTUNLAR_ITHALAT
            elif "İhr" in sheet_name or "Gerç.İhr" in sheet_name:
                required_columns = ZORUNLU_SUTUNLAR_IHRACAT

            if required_columns:
                df.columns = df.columns.str.strip()
                if not all(col in df.columns for col in required_columns):
                    st.error(f"'{sheet_name}' sayfasında zorunlu sütunlardan biri eksik. Gerekli sütunlar: {required_columns}")
                    st.stop()
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
    # string rakam filtrelemesi
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
        param = str(row[COL_PARAM]).lower().strip()
        if ("birim kullanım miktarı" in param or "fire" in param or ROW_TOPLAM_BIRIM.lower() in param):
            row_copy = row.copy()
            if pd.isna(row_copy[COL_MADDE]) or row_copy[COL_MADDE] == "":
                row_copy[COL_MADDE] = madde_adi
            blok.append(row_copy)
    return blok


def secili_sarfiyat_sayfasi(s1: pd.DataFrame, s2: pd.DataFrame | None):
    # 4. sayfa varsa onu kullan; yoksa 3. sayfa. Eski mantık s1 üzerinden yanlış karar veriyordu
    if s2 is not None:
        return s2
    return s1


def parse_birim(param_text: str) -> str:
    m = re.search(r"\((.*?)\)", str(param_text or "").lower())
    return m.group(1).strip() if m else ""


def to_float(v) -> float:
    try:
        return float(v)
    except Exception:
        return 0.0


def hesapla_toplam_mamul(
    tbk: float,  # Toplam Birim Kullanım
    birim_turu: str,
    col: str,
    is_fourth_sheet: bool,
    toplam_miktar_row: dict,
    sarfiyat2: pd.DataFrame | None,
) -> float:
    # Senaryoyu tek yerde uygulamak
    if not is_fourth_sheet:
        # 3 sayfa → her zaman dünya toplamıyla çarp
        return tbk * to_float(toplam_miktar_row.get(col, 0))

    # 4 sayfa varsa
    if "kg" in birim_turu or "kilo" in birim_turu:
        # 2.br: sarfiyat2.iloc[1][col]
        katsayi = 0.0
        if sarfiyat2 is not None and 1 < len(sarfiyat2.index) and col in sarfiyat2.columns:
            katsayi = to_float(sarfiyat2.iloc[1][col])
        return tbk * katsayi
    # diğer birimler: dünya toplamıyla çarp
    return tbk * to_float(toplam_miktar_row.get(col, 0))


st.set_page_config(page_title="İthalat ve İhracat Raporları", page_icon="📊", layout="wide")
st.title("İthalat ve İhracat Verilerini Yükleyin")
st.markdown("**İthalat ve İhracat verilerinizi yükleyin ve istediğiniz raporları alın.**")

uploaded_file = st.file_uploader("Excel Dosyasını Yükleyin", type=["xlsx"])

if uploaded_file is not None:
    data = load_excel(uploaded_file)
    if data is None:
        st.stop()

    sheet_names = list(data.keys())
    st.write(f"Yüklenen dosyada şu sheet'ler var: {', '.join(sheet_names)}")

    options = ['Gerç.İth.List.', 'Gerç.İhr.List.', 'Sarfiyat']
    selected_option = st.selectbox("İşlem Yapmak İstediğiniz Veri Tipini Seçin", options)

    if len(sheet_names) >= 3:
        ithalat_df_all = data[sheet_names[0]]
        ihracat_df_all = data[sheet_names[1]]
        sarfiyat1 = data[sheet_names[2]]  # 3. sayfa her zaman var

        sarfiyat2_exists = len(sheet_names) >= 4
        sarfiyat2 = data[sheet_names[3]] if sarfiyat2_exists else None

        ithalat_df = filter_imports(ithalat_df_all)
        ithalat_pivot = (
            ithalat_df.groupby(['Satır Kodu'])['İstatistiki Miktar']
            .sum()
            .reset_index()
            .rename(columns={'İstatistiki Miktar': 'Toplam İstatistiki Miktar'})
        )

        ab_df, non_ab_df = filter_exports(ihracat_df_all, AB_COUNTRIES)
        ab_pivot = (
            ab_df.groupby('Satır Kodu')['İstatistiki Miktar']
            .sum()
            .reset_index()
            .rename(columns={'İstatistiki Miktar': 'Toplam Miktar'})
        )
        non_ab_pivot = (
            non_ab_df.groupby('Satır Kodu')['İstatistiki Miktar']
            .sum()
            .reset_index()
            .rename(columns={'İstatistiki Miktar': 'Toplam Miktar'})
        )
        ozel_df = ihracat_df_all[ihracat_df_all['Varış Ülkesi'].isin(OZEL_ULKELER)]
        kontrol_df = ozel_df.copy() if not ozel_df.empty else pd.DataFrame()

        # Sarfiyat sayfası seçimi (neden: 4. sayfa varsa onu kullan)
        sarfiyat_df_all = secili_sarfiyat_sayfasi(sarfiyat1, sarfiyat2)
        st.info("Kullanılan sarfiyat sayfası: " + (sheet_names[3] if sarfiyat2_exists and sarfiyat_df_all.equals(data[sheet_names[3]]) else sheet_names[2]))

        sarfiyat_df = None
        if sarfiyat_df_all is not None:
            ithalat_kodlari = ithalat_pivot['Satır Kodu'].astype(str).unique()
            dunya_kodlari = non_ab_pivot['Satır Kodu'].astype(str).unique()

            satir_maskesi = sarfiyat_df_all[COL_MADDE].astype(str).isin(ithalat_kodlari)
            kolonlar = [COL_MADDE, COL_PARAM] + [col for col in sarfiyat_df_all.columns if str(col) in dunya_kodlari]
            kolonlar = [col for col in kolonlar if col in sarfiyat_df_all.columns]

            sarfiyat_df_filtered = sarfiyat_df_all[satir_maskesi][kolonlar]
            sarfiyat_df = sarfiyat_df_filtered.merge(
                ithalat_pivot, how="left", left_on=COL_MADDE, right_on="Satır Kodu"
            ).drop(columns=["Satır Kodu"])
            sarfiyat_df.rename(columns={"Toplam İstatistiki Miktar": COL_GERCEK_ITH_MIK}, inplace=True)

        # Dünya toplam miktar haritası
        dunya_pivot_dict = dict(zip(non_ab_pivot['Satır Kodu'].astype(str), non_ab_pivot['Toplam Miktar']))

        # Toplam Miktar satırı
        toplam_miktar_row = {COL_MADDE: "", COL_PARAM: ROW_TOPLAM_MIKTAR}
        for col in [] if sarfiyat_df is None else sarfiyat_df.columns:
            if col in dunya_pivot_dict:
                toplam_miktar_row[col] = dunya_pivot_dict[col]
            elif col not in [COL_MADDE, COL_PARAM, COL_GERCEK_ITH_MIK]:
                toplam_miktar_row[col] = 0

        if sarfiyat_df is not None:
            sarfiyat_df = pd.concat([pd.DataFrame([toplam_miktar_row]), sarfiyat_df], ignore_index=True)

            # Görsel amaçlı başlık satırları (neden: kullanıcı görebilsin)
            urun_adi_satiri_3 = sarfiyat1.iloc[3] if len(sarfiyat1.index) > 3 else pd.Series()
            kullanilan_urun_row_3 = {COL_MADDE: "", COL_PARAM: ROW_KULLANILAN_URUN}
            for col in sarfiyat_df.columns:
                if col not in [COL_MADDE, COL_PARAM] and col in urun_adi_satiri_3.index:
                    kullanilan_urun_row_3[col] = urun_adi_satiri_3[col]

            sarfiyat_df = pd.concat([
                sarfiyat_df.iloc[[0]],
                pd.DataFrame([kullanilan_urun_row_3]),
                sarfiyat_df.iloc[1:]
            ], ignore_index=True)

            if sarfiyat2_exists and sarfiyat_df_all.equals(sarfiyat2):
                urun_adi_satiri_4 = sarfiyat2.iloc[1] if len(sarfiyat2.index) > 1 else pd.Series()
                kullanilan_urun_row_4 = {col: "" for col in sarfiyat_df.columns}
                kullanilan_urun_row_4[COL_MADDE] = ROW_2BR_BIRIM
                kullanilan_urun_row_4[COL_PARAM] = ""
                for col in sarfiyat_df.columns:
                    if col not in [COL_MADDE, COL_PARAM] and col in urun_adi_satiri_4.index:
                        kullanilan_urun_row_4[col] = urun_adi_satiri_4[col]
                sarfiyat_df = pd.concat([
                    sarfiyat_df.iloc[[0, 1]],
                    pd.DataFrame([kullanilan_urun_row_4]),
                    sarfiyat_df.iloc[2:]
                ], ignore_index=True)

            # --- Sarfiyat oluşturma Toplam Mamul Kullanımı Hesabı oldu artık :)))))---
            yeni_sarfiyat_df = pd.DataFrame(columns=sarfiyat_df.columns)
            is_fourth = sarfiyat2_exists and sarfiyat_df_all.equals(sarfiyat2)

            i = 0
            while i < len(sarfiyat_df):
                satir = sarfiyat_df.iloc[i]
                yeni_sarfiyat_df.loc[len(yeni_sarfiyat_df)] = satir

                param_text = str(satir[COL_PARAM]).lower()
                if param_text.startswith("birim kullanım miktarı"):
                    madde_adi = satir[COL_MADDE]
                    birim_turu = parse_birim(satir[COL_PARAM])

                    # (seçilen sayfa)
                    kaynak_sarfiyat = sarfiyat_df_all
                    match_index = kaynak_sarfiyat[
                        (kaynak_sarfiyat[COL_MADDE] == madde_adi)
                        & (kaynak_sarfiyat[COL_PARAM].astype(str).str.lower().str.contains("birim kullanım miktarı"))
                    ].index

                    # Fire ve Toplam Birim Kullanım satırlarını ekle
                    fire_row = None
                    toplam_birim_row = None
                    if not match_index.empty:
                        blok_satirlari = get_madde_blok(kaynak_sarfiyat, match_index[0], madde_adi)
                        for r in blok_satirlari[1:]: 
                            # hangi satır olduğunu algıla :)
                            ptxt = str(r.get(COL_PARAM, "")).strip().lower()
                            yeni_row = {col: r[col] if col in r else "" for col in sarfiyat_df.columns}
                            yeni_sarfiyat_df.loc[len(yeni_sarfiyat_df)] = yeni_row
                            if ptxt.startswith("fire"):
                                fire_row = yeni_row
                            elif ptxt == ROW_TOPLAM_BIRIM.lower():
                                toplam_birim_row = yeni_row

                    # Toplam Mamul Kullanımı satırını hesapla ve ekle
                    mamul_row = {col: "" for col in sarfiyat_df.columns}
                    mamul_row[COL_MADDE] = madde_adi
                    mamul_row[COL_PARAM] = ROW_TOPLAM_MAMUL

                    for col in sarfiyat_df.columns:
                        if col in [COL_MADDE, COL_PARAM, COL_GERCEK_ITH_MIK]:
                            continue
                        tbk = to_float((toplam_birim_row or {}).get(col, satir.get(col))) 
                        mamul_row[col] = hesapla_toplam_mamul(
                            tbk=tbk,
                            birim_turu=birim_turu,
                            col=col,
                            is_fourth_sheet=is_fourth,
                            toplam_miktar_row=toplam_miktar_row,
                            sarfiyat2=sarfiyat2,
                        )

                    yeni_sarfiyat_df.loc[len(yeni_sarfiyat_df)] = mamul_row
                i += 1

            sarfiyat_df = yeni_sarfiyat_df.reset_index(drop=True)

            # TMK toplamını hesapla
            mamul_maskesi = sarfiyat_df[COL_PARAM] == ROW_TOPLAM_MAMUL
            def yatay_toplam(row):
                total = 0.0
                for c in sarfiyat_df.columns:
                    if c in [COL_MADDE, COL_PARAM, COL_GERCEK_ITH_MIK, COL_FARK, COL_TEV, COL_TMK_TOPLAM]:
                        continue
                    total += to_float(row.get(c, 0))
                return total

            sarfiyat_df.loc[mamul_maskesi, COL_TMK_TOPLAM] = sarfiyat_df[mamul_maskesi].apply(yatay_toplam, axis=1)

            # Gerçekleşen İthalat Miktarı (maddeye göre pivot değer)
            for idx, row in sarfiyat_df[mamul_maskesi].iterrows():
                madde_adi = row[COL_MADDE]
                ith_row = ithalat_pivot[ithalat_pivot['Satır Kodu'] == madde_adi]
                sarfiyat_df.loc[idx, COL_GERCEK_ITH_MIK] = (ith_row.iloc[0]['Toplam İstatistiki Miktar'] if not ith_row.empty else 0)

            # Fark ve TEV
            sarfiyat_df.loc[mamul_maskesi, COL_FARK] = (
                pd.to_numeric(sarfiyat_df.loc[mamul_maskesi, COL_TMK_TOPLAM], errors='coerce')
                - pd.to_numeric(sarfiyat_df.loc[mamul_maskesi, COL_GERCEK_ITH_MIK], errors='coerce')
            )
            sarfiyat_df.loc[mamul_maskesi, COL_TEV] = sarfiyat_df.loc[mamul_maskesi, COL_FARK].apply(
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

        # Çıktı
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
        st.download_button(
            "Tüm Verileri İndir",
            data=output_combined,
            file_name="tum_veriler_raporu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
