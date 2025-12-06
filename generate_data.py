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
required_cols = ["Nama Aplikasi", "Nama Responden (Participant)", "NIM"]
for col in required_cols:
    if col not in db.columns:
        raise RuntimeError(f"Kolom wajib '{col}' tidak ditemukan di sheet dbresponden")

apps_col = db["Nama Aplikasi"].astype(str).str.strip()
participants = db["Nama Responden (Participant)"].astype(str).str.strip()
nims_raw = db["NIM"]

# ===============================
# Helper: normalisasi NIM
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
# 3. Parsing "Nama Aplikasi"
#    Contoh: "MatchMate (09021282429106  Arkhan Syahputra)"
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
# 4. Map pembuat aplikasi (creator)
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
print(f"Jumlah pembuat aplikasi unik (berdasarkan NIM): {len(all_creators)}")

# ===============================
# 5. Hitung jumlah pengisi tiap aplikasi
# ===============================
# Berapa banyak baris db yang memilih aplikasi (full string)
app_full_count_map = apps_col.value_counts().to_dict()

# ===============================
# 6. Bangun data.json (rekap utama)
# ===============================
result = []

for creator in all_creators:
    owner_name = creator["name"]
    owner_app = creator["app"]
    owner_nim_norm = creator["nim_norm"]
    owner_name_norm = owner_name.lower().strip()

    # 1) Hitung berapa kali dia mengisi form (sebagai responden)
    mask_nim = (participants_nim_norm == owner_nim_norm)
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
            rated_app_names.add(parse_app_full(af)[0])

    # 3) Aplikasi yang belum dia nilai (punya orang lain)
    not_filled_set = set()
    for other in all_creators:
        if other["nim_norm"] == owner_nim_norm:
            continue
        if other["app"] not in rated_app_names:
            not_filled_set.add(other["app"])

    not_filled_list = sorted(not_filled_set)

    # 4) Hitung jumlah orang yang mengisi aplikasi miliknya
    app_fulls_for_owner = [
        af for af, info in app_info_map.items()
        if info["nim"] == owner_nim_norm
    ]
    app_filled_count = sum(app_full_count_map.get(af, 0) for af in app_fulls_for_owner)

    result.append({
        "name": owner_name,
        "nim": creator["nim"],
        "app": owner_app,
        "filled": filled_count,            # dia mengisi form berapa kali
        "app_filled_count": app_filled_count,  # aplikasinya diisi berapa kali
        "not_filled": not_filled_list
    })

result = sorted(result, key=lambda x: x["name"].lower())

# ===============================
# 7. Hitung nim_issues.json
#    (nama responden yang pakai >1 NIM berbeda)
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
print(f"Jumlah nama dengan NIM tidak konsisten: {len(nim_issues)}")

# ===============================
# 8. sus_scores.json
#    SUS per aplikasi + daftar responden
# ===============================
sus_scores = []

if "Jumlah" in db.columns:
    print("Menghitung SUS score per aplikasi...")
    sus_map = {}

    for _, row in db.iterrows():
        app_full = str(row["Nama Aplikasi"]).strip()
        if not app_full:
            continue

        app_name, owner_nim_raw, owner_name = parse_app_full(app_full)
        owner_nim_norm = normalize_nim(owner_nim_raw)

        # SUS raw (0–40) lalu dikali 2.5 -> 0–100
        try:
            raw = float(row["Jumlah"])
        except Exception:
            continue

        if pd.isna(raw):
            continue

        sus_value = raw * 2.5

        respondent_name = str(row["Nama Responden (Participant)"]).strip()
        respondent_nim = normalize_nim(row["NIM"])

        key = (app_name, owner_nim_norm, owner_name)
        if key not in sus_map:
            sus_map[key] = {
                "app": app_name,
                "owner_name": owner_name,
                "owner_nim": owner_nim_norm,
                "scores": [],
                "responses": []
            }

        sus_map[key]["scores"].append(sus_value)
        sus_map[key]["responses"].append({
            "respondent_name": respondent_name,
            "respondent_nim": respondent_nim,
            "sus": sus_value
        })

    for key, entry in sus_map.items():
        scores = entry["scores"]
        if not scores:
            continue
        avg_sus = sum(scores) / len(scores)
        sus_scores.append({
            "app": entry["app"],
            "owner_name": entry["owner_name"],
            "owner_nim": entry["owner_nim"],
            "avg_sus": avg_sus,
            "count": len(scores),
            "responses": entry["responses"]
        })

    sus_scores = sorted(sus_scores, key=lambda x: x["app"].lower())
    print(f"Jumlah aplikasi dengan SUS score: {len(sus_scores)}")
else:
    print("Kolom 'Jumlah' tidak ada di dbresponden, sus_scores.json akan kosong.")

# ===============================
# 9. Simpan semua JSON
# ===============================
with open("data.json", "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

with open("nim_issues.json", "w", encoding="utf-8") as f:
    json.dump(nim_issues, f, ensure_ascii=False, indent=2)

with open("sus_scores.json", "w", encoding="utf-8") as f:
    json.dump(sus_scores, f, ensure_ascii=False, indent=2)

print("Berhasil membuat: data.json, nim_issues.json, sus_scores.json")
