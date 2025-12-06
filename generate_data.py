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
nims_raw = db["NIM"]

# ===============================
# Normalisasi NIM
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

db["NIM_norm"] = participants_nim_norm

# ===============================
# 3. Parsing nama aplikasi
# ===============================
def parse_app_full(app_full: str):
    app_full = str(app_full).strip()
    app_name = app_full.split("(", 1)[0].strip()

    m = re.search(r"\((.*)\)", app_full)
    if not m:
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
# 4. Buat map pembuat aplikasi
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

creators_by_nim = {}  # nim_norm -> data pembuat
for info in app_info_map.values():
    nim_norm = info["nim"]
    if not nim_norm:
        continue
    if nim_norm not in creators_by_nim:
        creators_by_nim[nim_norm] = {
            "name": info["name"],
            "nim": info["nim"],
            "nim_norm": nim_norm,
            "app": info["app"],
        }

all_creators = list(creators_by_nim.values())

# ===============================
# 5. Hitung jumlah pengisi aplikasi (NEW!)
# ===============================
# app_full_count_map = berapa banyak baris memilih aplikasi ini
app_full_count_map = apps_col.value_counts().to_dict()

# ===============================
# 6. Hitung aktivitas tiap pembuat
# ===============================
result = []

for creator in all_creators:
    owner_name = creator["name"]
    owner_app = creator["app"]
    owner_nim_norm = creator["nim_norm"]
    owner_name_norm = owner_name.lower().strip()

    # 1) Hitung jumlah dia menjadi responden
    mask_nim = (participants_nim_norm == owner_nim_norm)

    if not mask_nim.any():
        mask = (participants_norm_name == owner_name_norm)
    else:
        mask = mask_nim

    filled_count = int(mask.sum())

    # 2) Ambil aplikasi yang sudah dia nilai
    rated_apps_full = apps_col[mask].unique()

    rated_app_names = set()
    for af in rated_apps_full:
        info = app_info_map.get(af)
        if info:
            rated_app_names.add(info["app"])
        else:
            rated_app_names.add(parse_app_full(af)[0])

    # 3) Hitung aplikasi yang belum dinilai
    not_filled_set = set()
    for other in all_creators:
        if other["nim_norm"] == owner_nim_norm:
            continue
        if other["app"] not in rated_app_names:
            not_filled_set.add(other["app"])

    not_filled_list = sorted(not_filled_set)

    # 4) Hitung jumlah pengisi aplikasi miliknya
    #    Cari app_full yang memiliki owner_nim
    app_fulls_for_owner = [
        af for af, info in app_info_map.items()
        if info["nim"] == owner_nim_norm
    ]

    app_filled_count = sum(app_full_count_map.get(af, 0) for af in app_fulls_for_owner)

    # SIMPAN
    result.append({
        "name": owner_name,
        "nim": creator["nim"],
        "app": owner_app,
        "filled": filled_count,
        "app_filled_count": app_filled_count,  # JUMLAH ORANG MENGISI APLIKASINYA
        "not_filled": not_filled_list
    })

result = sorted(result, key=lambda x: x["name"].lower())

# ===============================
# 7. Hitung NIM bermasalah
# ===============================
nim_issues = []
grouped = db.groupby("Nama Responden (Participant)")

for name, sub in grouped:
    unique_nims = sorted({normalize_nim(n) for n in sub["NIM_norm"].unique() if normalize_nim(n)})
    if len(unique_nims) > 1:
        counts = sub.groupby("NIM_norm").size().to_dict()
        nim_issues.append({
            "name": str(name).strip(),
            "nims": unique_nims,
            "nim_counts": counts,
            "total_rows": int(sub.shape[0])
        })

nim_issues = sorted(nim_issues, key=lambda x: x["name"].lower())

# ===============================
# 8. Save
# ===============================
with open("data.json", "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

with open("nim_issues.json", "w", encoding="utf-8") as f:
    json.dump(nim_issues, f, ensure_ascii=False, indent=2)

print("Berhasil membuat: data.json & nim_issues.json")
