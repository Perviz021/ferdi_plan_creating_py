import streamlit as st
import pandas as pd
import openpyxl
import os
import re
import time


# Dosya adı temizleme fonksiyonu
def temizle_dosya_adi(dosya_adi):
    temizlenmis_adi = re.sub(r'[\/:*?"<>|]', "_", dosya_adi)
    return temizlenmis_adi


# Veri alanlarını temizleme fonksiyonu
def temizle_hucreler(sheet):
    # Öğrenci bilgileri ve semestr verileri için hücre aralıklarını sıfırlıyoruz
    for row in range(7, 12):  # D7-D15 arası
        sheet[f"D{row}"] = ""
    for row in range(19, 29):  # Payız semestri verileri (A23-H32)
        for col in ["A", "B", "C", "D", "E", "F", "G", "H"]:
            sheet[f"{col}{row}"] = ""
    for row in range(31, 41):  # Yaz semestri verileri (A35-H44)
        for col in ["A", "B", "C", "D", "E", "F", "G", "H"]:
            sheet[f"{col}{row}"] = ""
    sheet["E29"] = ""  # Payız semestr toplamı
    sheet["E41"] = ""  # Yaz semestr toplamı


# Boş satırları silme fonksiyonu
def sil_bos_satirlar(sheet, start_row, end_row):
    for i in range(end_row, start_row - 1, -1):
        if all(
            (
                sheet[f"{col}{i}"].value is None
                or str(sheet[f"{col}{i}"].value).strip() == ""
            )
            for col in ["A", "B", "D", "E", "F", "G", "H"]
        ):
            sheet.delete_rows(i)
            print(f"Satır {i} silindi.")


# Streamlit arayüzü
st.title("Fərdi Tədris Planı")
st.write("Excel faylını yükləyin")

# Dosya yükleme işlemi
uploaded_file = st.file_uploader("Faylı Seçin", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Dosya yüklendiyse işlem başlatılacak
    df = pd.read_excel(uploaded_file)
    pd.set_option("display.max_columns", None)

    # Başlangıç zamanını kaydet
    start_time = time.time()

    # Öğrenci kodunu kullanıcıdan al
    telebe_kodu = st.text_input("Tələbə kodunu daxil edin:")

    # Girilen öğrenci koduna göre filtrele
    filtrelenmis_veri = df[df["Tələbə_kodu"] == telebe_kodu]

    if filtrelenmis_veri.empty:
        st.error(
            f"Daxil edilən tələbə kodu ({telebe_kodu}) ilə uyğun məlumat tapılmadı."
        )
    else:
        # Excel dosyasını işleme başla
        dosya_yolu = "/Users/Acer/Desktop/Avrasiya/ferdi_plan_yarat_js/Tədris_plan_qrafik_bazası_empty1.xlsx"
        workbook = openpyxl.load_workbook(dosya_yolu)
        sheet = workbook.active

        # Verileri temizliyoruz (sadece veri hücrelerini sıfırlıyoruz)
        temizle_hucreler(sheet)

        # Öğrenci bilgilerini D7-D15 hücrelerine aktar
        row = filtrelenmis_veri.iloc[0]
        sheet["C7"] = str(row["Fakültə_kodu"]) + " - " + row["Fakültə_adı"]
        sheet["C8"] = (
            "0"
            + str(row["İxtisasın_şifri"])
            + " / "
            + row["İxtisas_kodu"]
            + " - "
            + row["İxtisas_adı"]
        )
        sheet["C9"] = str(row["İxtisaslaşma_kodu"]) + " - " + row["İxtisaslaşma_adı"]
        sheet["F7"] = row["Təhsilin_səviyyəsi"]
        sheet["D7"] = "Təhsil səviyyəsi:"
        sheet["C10"] = row["Akademik_qrup"]
        sheet["F8"] = row["Proqram_ili"]
        sheet["D8"] = "Proqram ili:"
        sheet["F9"] = row["Tələbənin_təhsil_ili"]
        sheet["D9"] = "Tələbənin təhsil ili:"
        sheet["F10"] = row["Tədris_ili"]
        sheet["D10"] = "Tədris ili:"
        sheet["C11"] = str(row["Tələbə_kodu"]) + " - " + row["Soyadı,_adı_və_ata_adı"]

        # Semestr bilgisine göre dersleri aktar
        payiz_veriler = filtrelenmis_veri[filtrelenmis_veri["Semestr"] == "Payız"]
        yaz_veriler = filtrelenmis_veri[filtrelenmis_veri["Semestr"] == "Yaz"]

        # Payız semestr ders bilgilerini A23-H32 aralıklarına aktar
        payiz_veriler_sirali = payiz_veriler.sort_values(by="Fənnin_semestr_kodu")
        for i, (kod, ders_verisi) in enumerate(
            payiz_veriler_sirali.groupby("Fənnin_semestr_kodu")
        ):
            sheet[f"A{19 + i}"] = kod
            sheet[f"B{19 + i}"] = ders_verisi.iloc[0]["Fənnin_kodu"]
            sheet[f"C{19 + i}"] = ders_verisi.iloc[0]["Fənnin_adı"]
            sheet[f"D{19 + i}"] = ders_verisi.iloc[0]["Kredit_sayı"]
            sheet[f"E{19 + i}"] = ders_verisi.iloc[0]["Kafeda_(fənnin_aid_olduğu)"]
            sheet[f"F{19 + i}"] = ders_verisi.iloc[0]["MüəllimMS"]
            sheet[f"G{19 + i}"] = ders_verisi.iloc[0]["Fənn_qrupuMS"]

        # Yaz semestr ders bilgilerini A35-H44 aralıklarına aktar
        yaz_veriler_sirali = yaz_veriler.sort_values(by="Fənnin_semestr_kodu")
        for i, (kod, ders_verisi) in enumerate(
            yaz_veriler_sirali.groupby("Fənnin_semestr_kodu")
        ):
            sheet[f"A{31 + i}"] = kod
            sheet[f"B{31 + i}"] = ders_verisi.iloc[0]["Fənnin_kodu"]
            sheet[f"C{31 + i}"] = ders_verisi.iloc[0]["Fənnin_adı"]
            sheet[f"D{31 + i}"] = ders_verisi.iloc[0]["Kredit_sayı"]
            sheet[f"E{31 + i}"] = ders_verisi.iloc[0]["Kafeda_(fənnin_aid_olduğu)"]
            sheet[f"F{31 + i}"] = ders_verisi.iloc[0]["MüəllimMS"]
            sheet[f"G{31 + i}"] = ders_verisi.iloc[0]["Fənn_qrupuMS"]

        # Payız semestr toplamını E33 hücresine yazma
        e23_e32_toplam = sum(
            [sheet[f"D{23 + i}"].value or 0 for i in range(len(payiz_veriler_sirali))]
        )
        sheet["D29"] = e23_e32_toplam

        # Yaz semestr toplamını E45 hücresine yazma
        e35_e44_toplam = sum(
            [sheet[f"D{35 + i}"].value or 0 for i in range(len(yaz_veriler_sirali))]
        )
        sheet["D41"] = e35_e44_toplam

        # Yaz semestrindeki boş satırları sil
        sil_bos_satirlar(sheet, 31, 40)
        # Payız semestrindeki boş satırları sil
        sil_bos_satirlar(sheet, 19, 28)

        # Dosya adını temizleyerek oluştur
        dosya_adi = f"{row['Akademik_qrup']}_{row['Tələbə_kodu']}_{row['Soyadı,_adı_və_ata_adı']}.xlsx"
        dosya_adi_temiz = temizle_dosya_adi(dosya_adi)

        # Dosya yolunu oluştur
        dosya_yolu_yeni = f"/Users/Acer/Desktop/Avrasiya/ferdi_plan_yarat_js/fərdi_tədris_stud/{dosya_adi_temiz}"

        # Hedef klasör var mı, kontrol et
        klasor = os.path.dirname(dosya_yolu_yeni)
        if not os.path.exists(klasor):
            os.makedirs(klasor)  # Eğer klasör yoksa, oluştur

        # Dosyayı kaydetme
        workbook.save(dosya_yolu_yeni)

        # Kullanıcıya işlemin tamamlandığını bildirin
        st.success(f"Məlumatlar yadda saxlanıldı: {dosya_yolu_yeni}")

        # Kod çalıştıktan sonra bitiş zamanını kaydet
        end_time = time.time()

        # Geçen süreyi hesapla (saniye cinsinden)
        elapsed_time = end_time - start_time

        # Geçen süreyi yazdır
        st.write(f"Toplam vaxt: {elapsed_time:.2f} saniyə")
        # streamlit run /Users/aytajaskarli/Desktop/Avrasiya/FTP/plan/ftp.py
