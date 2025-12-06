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

# Kolom penting
apps_col = db["Nama Aplikasi"].astype(str).str.strip()
participants = db["Nama Responden (Participant)"].astype(str).str.strip()
participants_norm = participants.str.lower()

# ===============================
# 3. Fungsi parsing "Nama Aplikasi"
#    Contoh: "MatchMate (09021282429106  Arkhan Syahputra)"
#    -> app_name="MatchMate", nim="09021282429106", name="Arkhan Syahputra"
# ===============================
def parse_app_full(app_full: str):
    app_full = app_full.strip()
    app_name = app_full.split("(", 1)[0].strip()

    m = re.search(r"\((.*)\)", app_full)
    if not m:
        # fallback kalau format aneh
        return app_name, "", app_full

    inside = m.group(1).strip()
    parts = inside.split()
    if len(parts) < 2:
        nim = parts[0] if parts else ""
        name = inside
    else:
        nim = parts[0]
        name = " ".join(parts[1:])
    return app_name, nim, name

# ===============================
# 4. Bangun info pembuat aplikasi,
#    lalu dedup berdasarkan NIM
# ===============================
app_info_map = {}
for app_full in set(apps_col):
    app_name, nim, owner_name = parse_app_full(app_full)
    app_info_map[app_full] = {
        "app": app_name,
        "nim": nim,
        "name": owner_name,
    }

creators_by_nim = {}  # nim -> {name, nim, app}
for info in app_info_map.values():
    nim = info["nim"]
    if not nim:
        continue  # kalau tidak ada NIM, skip

    if nim not in creators_by_nim:
        creators_by_nim[nim] = {
            "name": info["name"],
            "nim": nim,
            "app": info["app"],
        }
    # kalau nim sudah ada, kita biarkan data pertama sebagai acuan

all_apps_list = list(creators_by_nim.values())

# ===============================
# 5. Hitung:
#    - berapa kali tiap pembuat isi form (filled)
#    - aplikasi mana yang belum dia nilai (not_filled)
# ===============================
result = []

for nim, creator in creators_by_nim.items():
    owner_name = creator["name"]
    owner_app = creator["app"]
    owner_norm = owner_name.strip().lower()

    # baris-baris di mana dia menjadi responden
    mask = (participants_norm == owner_norm)
    filled_count = int(mask.sum())

    # aplikasi yang sudah dia nilai
    rated_apps_full = apps_col[mask].unique()
    rated_app_names = set()
    for af in rated_apps_full:
        info = app_info_map.get(af)
        if info:
            rated_app_names.add(info["app"])

    # aplikasi yang belum dia nilai (milik orang lain)
    not_filled_apps = []
    for app_owner in all_apps_list:
        if app_owner["nim"] == nim:
            continue  # lewati aplikasinya sendiri
        if app_owner["app"] not in rated_app_names:
            not_filled_apps.append(app_owner["app"])

    result.append({
        "name": owner_name,          # nama pembuat aplikasi
        "nim": nim,
        "app": owner_app,            # aplikasi miliknya
        "filled": filled_count,      # berapa kali dia isi form
        "not_filled": not_filled_apps,  # list nama aplikasi yang belum dia nilai
    })

# urutkan biar rapi
result = sorted(result, key=lambda x: x["name"].lower())

print(f"Jumlah pembuat aplikasi unik (NIM): {len(result)}")

# ===============================
# 6. Simpan ke data.json (di root)
# ===============================
OUTPUT_FILE = "data.json"
with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f"Berhasil membuat: {OUTPUT_FILE}")
