import os
import json
import re
import requests
import pandas as pd

# ===============================
# 1. Download file Excel SUS
# ===============================
URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSWDG_1e6OyoSQsKp_Yy4uVAeVJOCrGPVABzv29EgoiNlhfhWV9vqOW2M6mFOvvkGbssfW3PXyH3ybM/pub?output=xlsx"
EXCEL_FILE = "sus_responses.xlsx"

print("Mengunduh file Excel...")
resp = requests.get(URL)
resp.raise_for_status()

with open(EXCEL_FILE, "wb") as f:
    f.write(resp.content)

print("Download selesai:", EXCEL_FILE)

# ===============================
# 2. Baca sheet dbresponden
# ===============================
xls = pd.ExcelFile(EXCEL_FILE)
print("Membaca sheet dbresponden...")
db = pd.read_excel(xls, "dbresponden")

# Pastikan kolom yang dipakai ada
required_cols = ["Nama Aplikasi", "Nama Responden (Participant)", "NIM"]
for col in required_cols:
    if col not in db.columns:
        raise RuntimeError(f"Kolom wajib '{col}' tidak ditemukan di sheet dbresponden")

apps_col = db["Nama Aplikasi"].astype(str).str.strip()
participants = db["Nama Responden (Participant)"].astype(str).str.strip()
nims_raw = db["NIM"]

# ===============================
# Fungsi normalisasi NIM
# - ubah ke string
# - buang spasi
# - hilangkan akhiran '.0' kalau ada
# - kosongkan kalau 'nan'
# ===============================
def normalize_nim(x):
    s = str(x).strip()
    if s.lower() == "nan" or s == "":
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    return s

participants_nim_norm = nims_raw.map(normalize_nim)
participants_norm_name = participants.str.lower()

# ===============================
# 3. Fungsi parsing "Nama Aplikasi"
#    Contoh: "MatchMate (09021282429106  Arkhan Syahputra)"
#    -> app_name="MatchMate", nim="09021282429106", name="Arkhan Syahputra"
# ===============================
def parse_app_full(app_full: str):
    app_full = str(app_full).strip()
    app_name = app_full.split("(", 1)[0].strip()

    m = re.search(r"\((.*)\)", app_full)
    if not m:
        # fallback kalau format aneh
        return app_name, "", app_full

    inside = m.group(1).strip()
    parts = inside.split()
    if len(parts) < 2:
        nim = normalize_nim(parts[0]) if parts else ""
        name = inside
    else:
        nim = normalize_nim(parts[0])
        name = " ".join(parts[1:])
    return app_name, nim, name

# ===============================
# 4. Bangun info pembuat aplikasi,
#    lalu dedup berdasarkan NIM
# ===============================
app_info_map = {}
for app_full in set(apps_col):
    app_name, nim_raw, owner_name = parse_app_full(app_full)
    nim_norm = normalize_nim(nim_raw)
    app_info_map[app_full] = {
        "app": app_name,
        "nim": nim_norm,
        "name": owner_name,
    }

creators_by_nim = {}  # nim_norm -> {name, nim, nim_norm, app}
for info in app_info_map.values():
    nim_norm = info["nim"]
    if not nim_norm:
        # kalau tidak ada NIM, skip dari daftar "creator utama"
        continue

    if nim_norm not in creators_by_nim:
        creators_by_nim[nim_norm] = {
            "name": info["name"],
            "nim": info["nim"],      # versi mentah
            "nim_norm": nim_norm,    # versi normalize
            "app": info["app"],
        }
    # kalau nim_norm sudah ada, biarkan entry pertama sebagai acuan

all_creators = list(creators_by_nim.values())

print(f"Jumlah pembuat aplikasi unik (berdasarkan NIM): {len(all_creators)}")

# ===============================
# 5. Hitung:
#    - berapa kali tiap pembuat isi form (filled) berbasis NIM
#    - aplikasi mana yang belum dia nilai (not_filled)
# ===============================
result = []

# Siapkan set semua nama aplikasi unik (berdasarkan creator)
all_app_names = sorted({c["app"] for c in all_creators})

for creator in all_creators:
    owner_name = creator["name"]
    owner_app = creator["app"]
    owner_nim_norm = creator["nim_norm"]
    owner_name_norm = owner_name.strip().lower()

    # 1) Cari baris di mana dia menjadi responden:
    #    Utama: berdasarkan NIM
    mask_nim = (participants_nim_norm == owner_nim_norm)

    # Fallback: kalau baris berdasarkan NIM 0 (misal NIM responden salah / kosong),
    # pakai pencocokan nama (case-insensitive)
    if not mask_nim.any():
        mask = (participants_norm_name == owner_name_norm)
    else:
        mask = mask_nim

    filled_count = int(mask.sum())

    # 2) Aplikasi yang sudah dia nilai
    rated_apps_full = apps_col[mask].unique()
    rated_app_names = set()
    for af in rated_apps_full:
        info = app_info_map.get(af)
        if info:
            rated_app_names.add(info["app"])
        else:
            # fallback kalau tidak ada di map
            app_name_tmp, _, _ = parse_app_full(af)
            rated_app_names.add(app_name_tmp)

    # 3) Aplikasi yang belum dia nilai (hanya aplikasi milik orang lain)
    not_filled_set = set()
    for other in all_creators:
        if other["nim_norm"] == owner_nim_norm:
            continue  # lewati aplikasinya sendiri
        if other["app"] not in rated_app_names:
            not_filled_set.add(other["app"])

    not_filled_list = sorted(not_filled_set)

    result.append({
        "name": owner_name,          # nama pembuat aplikasi
        "nim": creator["nim"],       # NIM pembuat
        "app": owner_app,            # aplikasi miliknya
        "filled": filled_count,      # berapa kali dia isi form
        "not_filled": not_filled_list,  # list nama aplikasi yang belum dia nilai
    })

# urutkan biar rapi
result = sorted(result, key=lambda x: x["name"].lower())

# ===============================
# 6. Simpan ke data.json (di root)
# ===============================
OUTPUT_FILE = "data.json"
with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f"Berhasil membuat: {OUTPUT_FILE}")
