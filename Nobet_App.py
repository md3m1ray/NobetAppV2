import sqlite3
import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
from datetime import datetime, timedelta
import random
import calendar
from collections import defaultdict

# Veritabanı bağlantısı
conn = sqlite3.connect("nobet_db.sqlite")
cursor = conn.cursor()

# Kişi tablosu oluştur
cursor.execute("""
    CREATE TABLE IF NOT EXISTS Kisiler (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        isim TEXT NOT NULL
    )
""")
cursor.execute("""
    CREATE TABLE IF NOT EXISTS Musaitlik (
        kisi_id INTEGER,
        gun TEXT,
        FOREIGN KEY (kisi_id) REFERENCES Kisiler(id)
    )
""")
conn.commit()


def kisi_ekle():
    isim = entry_isim.get()
    if not isim:
        messagebox.showerror("Hata", "İsim boş olamaz!")
        return

    secili_gunler = [gun for gun, var in gun_durumlari.items() if var.get() == 1]

    if not secili_gunler:
        messagebox.showerror("Hata", "En az bir gün seçmelisiniz!")
        return

    cursor.execute("INSERT INTO Kisiler (isim) VALUES (?)", (isim,))
    kisi_id = cursor.lastrowid

    for gun in secili_gunler:
        cursor.execute("INSERT INTO Musaitlik (kisi_id, gun) VALUES (?, ?)", (kisi_id, gun))

    conn.commit()
    messagebox.showinfo("Başarılı", f"{isim} eklendi.")
    entry_isim.delete(0, tk.END)
    for var in gun_durumlari.values():
        var.set(0)
    kisileri_yukle()


def kisi_guncelle():
    selected = tree.selection()
    if not selected:
        messagebox.showerror("Hata", "Güncellemek için bir kişi seçmelisiniz.")
        return

    kisi_id = tree.item(selected[0], "values")[0]
    isim = entry_isim.get()
    secili_gunler = [gun for gun, var in gun_durumlari.items() if var.get() == 1]

    if not isim:
        messagebox.showerror("Hata", "İsim boş olamaz!")
        return
    if not secili_gunler:
        messagebox.showerror("Hata", "En az bir gün seçmelisiniz!")
        return

    # İsim güncelleme
    cursor.execute("UPDATE Kisiler SET isim = ? WHERE id = ?", (isim, kisi_id))
    # Günleri güncelleme
    cursor.execute("DELETE FROM Musaitlik WHERE kisi_id = ?", (kisi_id,))
    for gun in secili_gunler:
        cursor.execute("INSERT INTO Musaitlik (kisi_id, gun) VALUES (?, ?)", (kisi_id, gun))

    conn.commit()
    messagebox.showinfo("Başarılı", f"{isim} güncellendi.")
    kisileri_yukle()


def kisi_sil():
    selected = tree.selection()
    if not selected:
        messagebox.showerror("Hata", "Silmek için bir kişi seçmelisiniz.")
        return

    kisi_id = tree.item(selected[0], "values")[0]
    cursor.execute("DELETE FROM Kisiler WHERE id = ?", (kisi_id,))
    cursor.execute("DELETE FROM Musaitlik WHERE kisi_id = ?", (kisi_id,))

    conn.commit()
    messagebox.showinfo("Başarılı", "Kişi silindi.")
    kisileri_yukle()


def kisi_sec(event):
    selected = tree.selection()
    if selected:
        kisi_id = tree.item(selected[0], "values")[0]
        isim = tree.item(selected[0], "values")[1]
        entry_isim.delete(0, tk.END)
        entry_isim.insert(0, isim)

        for var in gun_durumlari.values():
            var.set(0)
        cursor.execute("SELECT gun FROM Musaitlik WHERE kisi_id = ?", (kisi_id,))
        gunler = [row[0] for row in cursor.fetchall()]
        for gun in gunler:
            gun_durumlari[gun].set(1)


def kisileri_yukle():
    for item in tree.get_children():
        tree.delete(item)
    cursor.execute("SELECT id, isim FROM Kisiler")
    kisiler = cursor.fetchall()
    for kisi in kisiler:
        cursor.execute("SELECT gun FROM Musaitlik WHERE kisi_id = ?", (kisi[0],))
        gunler = ", ".join([row[0] for row in cursor.fetchall()])
        tree.insert("", "end", values=(kisi[0], kisi[1], gunler))


gunler_map = {
    "Monday": "Pazartesi",
    "Tuesday": "Salı",
    "Wednesday": "Çarşamba",
    "Thursday": "Perşembe",
    "Friday": "Cuma"
}


def aylik_cizelge_olustur():
    try:
        yil = int(entry_yil.get())
        ay = int(entry_ay.get())
    except ValueError:
        messagebox.showerror("Hata", "Geçerli bir yıl ve ay giriniz!")
        return

    gun_sayisi = calendar.monthrange(yil, ay)[1]
    baslangic_tarihi = datetime(yil, ay, 1)
    cizelge = []
    nobet_sayilari = defaultdict(int)
    haftalik_nobetler = set()

    for i in range(gun_sayisi):
        tarih = baslangic_tarihi + timedelta(days=i)
        gun = tarih.strftime('%A')
        gun_turkce = gunler_map.get(gun, None)

        if gun_turkce:
            cursor.execute("""
                SELECT isim FROM Kisiler
                JOIN Musaitlik ON Kisiler.id = Musaitlik.kisi_id
                WHERE Musaitlik.gun = ?
            """, (gun_turkce,))
            uygun_kisiler = [row[0] for row in cursor.fetchall()]

            uygun_kisiler = [kisi for kisi in uygun_kisiler if kisi not in haftalik_nobetler]

            if uygun_kisiler:
                random.shuffle(uygun_kisiler)
                uygun_kisiler.sort(key=lambda x: nobet_sayilari[x])

                nobetci = uygun_kisiler[0]
                cizelge.append([tarih.strftime('%Y-%m-%d'), gun_turkce, nobetci])
                nobet_sayilari[nobetci] += 1
                haftalik_nobetler.add(nobetci)

                if gun == "Friday":
                    haftalik_nobetler.clear()
            else:
                cizelge.append([tarih.strftime('%Y-%m-%d'), gun_turkce, "Nöbetçi Yok"])

    df = pd.DataFrame(cizelge, columns=['Tarih', 'Gün', 'Nöbetçi'])
    df.to_excel("Aylik_Nobet_Cizelgesi.xlsx", index=False)

    # Nöbet sayılarını ek olarak kaydet
    nobet_sayisi_df = pd.DataFrame(list(nobet_sayilari.items()), columns=["Kişi", "Nöbet Sayısı"])
    with pd.ExcelWriter("Aylik_Nobet_Cizelgesi.xlsx", mode="a", engine="openpyxl") as writer:
        nobet_sayisi_df.to_excel(writer, sheet_name="Nöbet Sayıları", index=False)

    messagebox.showinfo("Başarılı",
                        "Aylık nöbet çizelgesi oluşturuldu ve 'Aylik_Nobet_Cizelgesi.xlsx' olarak kaydedildi.")


root = tk.Tk()
root.title("Nöbet Çizelgesi Oluşturma")
root.geometry("500x550")

isim_frame = tk.Frame(root, highlightbackground="green", highlightthickness=1, borderwidth=1)
isim_frame.grid(row=0, column=0, padx=10, pady=10, sticky="w")

isim_label_frame = tk.Frame(isim_frame, highlightbackground="green", highlightthickness=0, borderwidth=0)
isim_label_frame.grid(row=0, column=0, padx=10, pady=5, sticky="w")

tk.Label(isim_label_frame, text="İsim Soyisim:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_isim = tk.Entry(isim_label_frame)
entry_isim.grid(row=0, column=1, padx=5, pady=5, sticky="w")

gun_frame = tk.Frame(isim_frame, highlightbackground="green", highlightthickness=1, borderwidth=0)
gun_frame.grid(row=1, column=0, padx=10, pady=5, sticky="w")

gun_durumlari = {
    "Pazartesi": tk.IntVar(),
    "Salı": tk.IntVar(),
    "Çarşamba": tk.IntVar(),
    "Perşembe": tk.IntVar(),
    "Cuma": tk.IntVar()
}

column = 0
for gun, var in gun_durumlari.items():
    tk.Checkbutton(gun_frame, text=gun, variable=var).grid(row=1, column=column, sticky="nsew")
    column += 1

buton_frame = tk.Frame(isim_frame, highlightbackground="green", highlightthickness=0, borderwidth=0)
buton_frame.grid(row=2, column=0, padx=10, pady=5, sticky="w")

tk.Button(buton_frame, text="Kişi Ekle", fg="white", bg="green", command=kisi_ekle).grid(row=2, column=0, padx=10,
                                                                                         pady=10, sticky="w")
tk.Button(buton_frame, text="Kişi Güncelle", fg="white", bg="blue", command=kisi_guncelle).grid(row=2, padx=10,
                                                                                                column=2, pady=10,
                                                                                                sticky="w")
tk.Button(buton_frame, text="Kişi Sil", fg="white", bg="red", command=kisi_sil).grid(row=2, column=4, padx=10, pady=10,
                                                                                     sticky="w")

liste_frame = tk.Frame(root, highlightbackground="blue", highlightthickness=1, borderwidth=1)
liste_frame.grid(row=3, column=0, padx=10, columnspan=4, pady=5, sticky="w")

tree = ttk.Treeview(liste_frame, columns=("ID", "İsim", "Günler"), show="headings")
tree.heading("ID", text="ID")
tree.heading("İsim", text="İsim Soyisim")
tree.heading("Günler", text="Müsait Günler")
tree.column("ID", width=30)
tree.column("İsim", width=150)
tree.column("Günler", width=250)
tree.grid(row=3, column=1, columnspan=3, pady=10, padx=10, sticky="nsew")
tree.bind("<Double-1>", kisi_sec)

kisileri_yukle()

nobet_frame = tk.Frame(root, highlightbackground="red", highlightthickness=1, borderwidth=1)
nobet_frame.grid(row=5, column=0, columnspan=4, padx=10, pady=5, sticky="w")

takvim_frame = tk.Frame(nobet_frame, highlightbackground="red", highlightthickness=0, borderwidth=0)
takvim_frame.grid(row=5, column=0, columnspan=4, padx=10, pady=5, sticky="w")

tk.Label(takvim_frame, text="Yıl:").grid(row=5, column=0, padx=5, pady=5, sticky="e")
entry_yil = tk.Entry(takvim_frame, width=5)
entry_yil.grid(row=5, column=1, padx=5, pady=5, sticky="w")
entry_yil.insert(0, str(datetime.now().year))

tk.Label(takvim_frame, text="Ay:").grid(row=5, column=2, padx=5, pady=5, sticky="e")
entry_ay = tk.Entry(takvim_frame, width=5)
entry_ay.grid(row=5, column=3, padx=5, pady=5, sticky="w")
entry_ay.insert(0, str(datetime.now().month))

tk.Button(nobet_frame, text="Nöbet Çizelgesi Oluştur", fg="white", bg="orange", command=aylik_cizelge_olustur).grid(
    row=6, column=1, pady=10)

tk.Label(root, text="@md3m1ray", fg="grey").grid(row=10, column=2, padx=5, pady=5, sticky="w")

root.grid_rowconfigure(15, weight=1)
root.grid_columnconfigure(1, weight=1)

root.mainloop()
conn.close()
