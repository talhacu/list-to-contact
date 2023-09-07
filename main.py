import openpyxl

# Giriş dosyası adı ve çıkış dosyası adı
input_file = "dosya_ismi.xlsx"  # Verilerinizi içeren XLSX dosyasının adını girin
output_file = "çıktı_dosyası_ismi.vcf"  # Oluşturulacak VCF dosyasının adını girin

# Excel dosyasını aç ve gerekli sayfayı seç
workbook = openpyxl.load_workbook(input_file)
sheet = workbook.active

# VCF dosyasını oluştur
with open(output_file, "w", encoding="utf-8") as vcf_file:
    for row in sheet.iter_rows(values_only=True):
        name = row[0]  # İsim sütunu
        phone = row[1]  # Telefon numarası sütunu

        vcf_file.write("BEGIN:VCARD\n")
        vcf_file.write("VERSION:3.0\n")
        vcf_file.write(f"N:{name}\n")
        vcf_file.write(f"TEL:{phone}\n")
        vcf_file.write("END:VCARD\n")
