import os
import json
import re
import requests
import pandas as pd

# 1. Download file Excel dari Google Sheets
URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSWDG_1e6OyoSQsKp_Yy4uVAeVJOCrGPVABzv29EgoiNlhfhWV9vqOW2M6mFOvvkGbssfW3PXyH3ybM/pub?output=xlsx"
EXCEL_FILE = "sus_responses.xlsx"

print("Mengunduh file Excel...")
resp = requests.get(URL)
resp.raise_for_status()

with open(EXCEL_FILE, "wb") as f:
    f.write(resp.content)

print("Download selesai:", EXCEL_FILE)

# 2. Baca sheet dbresponden
xls = pd.ExcelFile(EXCEL_FILE)
print("Membaca sheet dbresponden...")
db = pd.read_excel(xls, "dbresponden")

# -----------------------------
# Siapkan kolom-kolom penting
# -----------------------------
apps_col = db["Nama Aplikasi"].astype(str).str.strip()
participants = db["Nama Responden (Participant)"].astype(str).str.strip()

# Normalisasi nama responden ke lowercase untuk pencocokan
participants_norm = participants.str.lower()

# -----------------------------
# Parsing "Nama Aplikasi" -> app_name, nim, owner_name
# Contoh: "MatchMate (09021282429106  Arkhan Syahputra)"
# -----------------------------
def parse_app_full(app_full: str):
    app_full = app_full.strip()
    # Nama aplikasi = sebelum "("
    app_name = app_full.split("(", 1)[0].strip()

    # Isi dalam kurung
    m = re.search(r"\((.*)\)", app_full)
    if not m:
        return app_name, "", app_full  # fallback

    inside = m.group(1).strip()
    parts = inside.split()
    if len(parts) < 2:
        nim = parts[0] if parts else ""
        owner_name = inside
    else:
        nim = parts[0]
        owner_name = " ".join(parts[1:])
    return app_name, nim, owner_name

# Bangun mapping dari app_full -> (app_name, nim, owner_name)
app_info_map = {}
for app_full in set(apps_col):
    app_name, nim, owner_name = parse_app_full(app_full)
    app_info_map[app_full] = {
        "app": app_name,
        "nim": nim,
        "name": owner_name,
    }

# -----------------------------
# Dedup pembuat aplikasi berdasarkan NIM
# Jika ada 2 baris dengan NIM sama (nama typo), dianggap orang yang sama
# -----------------------------
creators_by_nim = {}  # nim -> {name, nim, app}
for info in app_info_map.values():
    nim = info["nim"]
    if not nim:
        # Kalau tidak ada NIM, bisa di-skip atau ditangani khusus
        continue

    if nim not in creators_by_nim:
        creators_by_nim[nim] = {
            "name": info["name"],
            "nim": nim,
            "app": info["app"],
        }
    else:
        # Kalau sudah ada, kita biarkan yang pertama sebagai "canonical"
        # (bisa juga dibuat log kalau mau cek perbedaan)
        pass

# Daftar semua aplikasi (unik per NIM)
all_apps_list = list(creators_by_nim.values())
all_app_names = sorted({c["app"] for c in all_apps_list})

# -----------------------------
# Untuk tiap pembuat aplikasi:
# - Hitung berapa kali dia mengisi form (sebagai responden)
# - Cari aplikasi mana yang belum dia isi
# -----------------------------
result = []

for nim, creator in creators_by_nim.items():
    owner_name = creator["name"]
    owner_app = creator["app"]

    owner_norm = owner_name.strip().lower()

    # Baris-baris di mana dia muncul sebagai responden
    mask = (participants_norm == owner_norm)
    filled_count = int(mask.sum())

    # Aplikasi yang sudah dia nilai
    rated_apps_full = apps_col[mask].unique()
    rated_app_names = set()
    for af in rated_apps_full:
        info = app_info_map.get(af)
        if info:
            rated_app_names.add(info["app"])

    # Aplikasi yang belum dia nilai:
    # semua aplikasi milik orang lain yang tidak ada di rated_app_names
    not_filled_apps = []
    for app_owner in all_apps_list:
        # Lewatkan aplikasinya sendiri (biasanya tidak perlu menilai aplikasi sendiri)
        if app_owner["nim"] == nim:
            continue
        if app_owner["app"] not in rated_app_names:
            not_filled_apps.append({
                "app": app_owner["app"],
                "creator_name": app_owner["name"],
                "creator_nim": app_owner["nim"],
            })

    result.append({
        "name": owner_name,          # nama pembuat aplikasi
        "nim": nim,
        "app": owner_app,            # aplikasi miliknya
        "filled": filled_count,      # berapa kali dia isi form
        "not_filled_count": len(not_filled_apps),
        "not_filled_apps": not_filled_apps,
    })

# Sort hasil biar rapi (misal berdasarkan nama)
result = sorted(result, key=lambda x: x["name"].lower())

print(f"Jumlah pembuat aplikasi unik (berdasarkan NIM): {len(result)}")

# 5. Simpan ke public/data.json
os.makedirs("public", exist_ok=True)
output_path = os.path.join("public", "data.json")

with open(output_path, "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print("Berhasil membuat:", output_path)
