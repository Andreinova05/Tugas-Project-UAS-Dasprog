import pandas as pd

# Membaca file Excel
file_path = "template_generate_jadwal (3).xlsx"
excel_file = pd.ExcelFile(file_path)

# Tampilkan nama sheet
print("Sheet yang tersedia:", excel_file.sheet_names)

# Baca isi Sheet1
df = excel_file.parse("Sheet1")

# Tampilkan beberapa baris awal
print("Isi awal data:")
print(df.head())

import pandas as pd

# Baca file
file_path = "template_generate_jadwal (3).xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1")

# Gabung data dua baris
data_bersih = []
for i in range(0, len(df), 2):
    atas = df.iloc[i]
    bawah = df.iloc[i+1] if i+1 < len(df) else None

    kelas = [atas["KELAS"]]
    if bawah is not None and pd.notna(bawah["KELAS"]):
        kelas.append(bawah["KELAS"])

    data_bersih.append({
        "MATA_KULIAH": atas["MATA KULIAH"],
        "DOSEN": atas["DOSEN PENGAJAR"],
        "SMT": atas["SMT"],
        "SKS": atas["SKS"],
        "KELAS": kelas,
        "OFF_ON": atas["Tipe_Kelas"]
    })

df_bersih = pd.DataFrame(data_bersih)

# Simulasi penjadwalan (asumsi 5 hari, 5 jam per hari)
hari = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat"]
jam = ["07:00-09:00", "09:00-11:00", "11:00-13:00", "13:00-15:00", "15:00-17:00"]

# Penjadwalan acak dengan cek konflik
jadwal = []
slot_terpakai = {}

for idx, row in df_bersih.iterrows():
    ketemu_slot = False
    for h in hari:
        for j in jam:
            bentrok = False
            for kls in row["KELAS"]:
                if (h, j, kls) in slot_terpakai:
                    bentrok = True
                    break
            if not bentrok:
                for kls in row["KELAS"]:
                    slot_terpakai[(h, j, kls)] = row["MATA_KULIAH"]
                jadwal.append({
                    "HARI": h,
                    "JAM": j,
                    "MATA_KULIAH": row["MATA_KULIAH"],
                    "DOSEN": row["DOSEN"],
                    "KELAS": ", ".join(row["KELAS"]),
                    "OFF_ON": row["Tipe_Kelas"]
                })
                ketemu_slot = True
                break
        if ketemu_slot:
            break
    if not ketemu_slot:
        jadwal.append({
            "HARI": "BENTROK",
            "JAM": "-",
            "MATA_KULIAH": row["MATA_KULIAH"],
            "DOSEN": row["DOSEN"],
            "KELAS": ", ".join(row["KELAS"]),
            "OFF_ON": row["Tipe_Kelas"]
        })

df_jadwal = pd.DataFrame(jadwal)

# Simpan ke Excel
df_jadwal.to_excel("jadwal_tanpa_bentrok.xlsx", index=False)
print(" Jadwal berhasil disimpan di 'jadwal_tanpa_bentrok.xlsx'")
