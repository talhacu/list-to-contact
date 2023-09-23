import openpyxl
import tkinter as tk
from tkinter import filedialog
from tkinter import *
from tkinter import messagebox
import time


def convert_excel_to_vcf():
    input_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

    if input_file:
        workbook = openpyxl.load_workbook(input_file)
        sheet = workbook.active

        rows_to_delete = []
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
            if all([cell.value is None for cell in row]):
                rows_to_delete.append(row)

        for row in rows_to_delete:
            sheet.delete_rows(row[0].row)
        workbook.save(input_file)

        output_file = f"{input_file}_KİŞİLER.vcf"
        baslik = kisilerin_basi_entry.get()
        son = kisilerin_sonu_entry.get()

        with open(output_file, "w", encoding="utf-8") as vcf_file:
            for row in sheet.iter_rows(values_only=True):
                name = row[0]
                phone = row[1]


                phone = str(phone).replace(" ", "").zfill(11)
                full_name = f"{baslik} {name} {son}"

                vcf_file.write("BEGIN:VCARD\n")
                vcf_file.write("VERSION:3.0\n")
                vcf_file.write(f"N:{full_name}\n")
                vcf_file.write(f"TEL:{phone}\n")
                vcf_file.write("END:VCARD\n")
        time.sleep(1.5)
        messagebox.showinfo(title="Tamamlandı",message="Liste başarıyla kişiler dosyasına dönüştürüldü")



#tkinter window
root = tk.Tk()
root.title("Listeden Kişiler Dosyası Oluşturma")
root.minsize(height=500,width=500)
root.configure(pady=50)

kisilerin_basi = Label(text="İsimlerin başında ne yazacak? Örnek: OGB15 Talha Bakır")
kisilerin_basi_entry = Entry()
kisilerin_basi.pack()
kisilerin_basi_entry.pack()
kisilerin_sonu = Label(text="İsimlerin sonunda ne yazacak? Örnek: Talha Bakır OG25")
kisilerin_sonu_entry = Entry()
kisilerin_sonu.pack()
kisilerin_sonu_entry.pack()



image = PhotoImage(file="liste.png")
resized = image.subsample(3)
label = Label(root, image=resized, pady=30)
label.place(x=60, y=130)

#sonuç
aciklama = Label(text="Listenizi yukarıdaki şekilde düzenlemelisiniz \n A Sütununda İsim ve Soyisim \n B Sütununda ise başında sıfır olmadan numaralar olmalı")
aciklama.place(x=70, y=280)
convert_button = tk.Button(root, text="Liste Dosyasını Seç ve Dönüştür", command=convert_excel_to_vcf)
convert_button.place(x=120,y=350)
result_label = tk.Label(root, text="")
result_label.place(x=65,y=400)

root.mainloop()
