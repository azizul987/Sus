import os
import json
import re
import requests
import pandas as pd
from datetime import datetime, timezone

# ===============================
# GENERATE TIMESTAMP UTC
# ===============================
generated_at = datetime.now(timezone.utc).isoformat()

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
        return s[:-2]
    return s

participants_nim_norm = nims_raw.map(normalize_nim)
participants_norm_name = participants.str.lower()

db["NIM_norm"] = participants_nim_norm

# ===============================
# 3. Parsing app_full → nama app, nim, creator
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
# 4. Map pembuat aplikasi
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

creators_by_nim = {}
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
# 5. Hitung jumlah pengisi tiap aplikasi
# ===============================
app_full_count_map = apps_col.value_counts().to_dict()

# ===============================
# 6. Buat data.json
# ===============================
result = []

for creator in all_creators:
    owner_name = creator["name"]
    owner_app = creator["app"]
    owner_nim_norm = creator["nim_norm"]
    owner_name_norm = owner_name.lower().strip()

    # dia mengisi form berapa kali
    mask_nim = (participants_nim_norm == owner_nim_norm)
    mask = mask_nim if mask_nim.any() else (participants_norm_name == owner_name_norm)
    filled_count = int(mask.sum())

    rated_apps_full = apps_col[mask].unique()
    rated_app_names = set()

    for af in rated_apps_full:
        info = app_info_map.get(af)
        rated_app_names.add(info["app"] if info else parse_app_full(af)[0])

    not_filled = sorted({
        other["app"] for other in all_creators
        if other["nim_norm"] != owner_nim_norm and other["app"] not in rated_app_names
    })

    app_fulls_for_owner = [
        af for af, info in app_info_map.items()
        if info["nim"] == owner_nim_norm
    ]
    app_filled_count = sum(app_full_count_map.get(af, 0) for af in app_fulls_for_owner)

    result.append({
        "name": owner_name,
        "nim": creator["nim"],
        "app": owner_app,
        "filled": filled_count,
        "app_filled_count": app_filled_count,
        "not_filled": not_filled,
        "generated_at": generated_at
    })

result = sorted(result, key=lambda x: x["name"].lower())

# ===============================
# 7. nim_issues.json
# ===============================
nim_issues = []
grouped = db.groupby("Nama Responden (Participant)")

for name, sub in grouped:
    unique_nims = sorted({
        normalize_nim(n) for n in sub["NIM_norm"].unique()
        if normalize_nim(n)
    })
    if len(unique_nims) > 1:
        counts = sub.groupby("NIM_norm").size().to_dict()
        nim_issues.append({
            "name": str(name).strip(),
            "nims": unique_nims,
            "nim_counts": {normalize_nim(k): int(v) for k, v in counts.items()},
            "total_rows": int(sub.shape[0]),
            "generated_at": generated_at
        })

# ===============================
# 8. SUS SCORES + Q1–Q10 + S1–S10
# ===============================

SUS_QUESTIONS = {
    1: "Saya merasa akan sering menggunakan sistem ini",
    2: "Saya merasa sistem ini rumit untuk digunakan",
    3: "Saya merasa sistem ini mudah digunakan",
    4: "Saya membutuhkan bantuan dari orang lain atau teknisi dalam menggunakan sistem ini",
    5: "Saya merasa fitur-fitur sistem ini berjalan dengan semestinya",
    6: "Saya merasa ada banyak hal yang tidak konsisten (tidak serasi pada sistem ini)",
    7: "Saya merasa orang lain akan memahami cara menggunakan sistem ini dengan cepat",
    8: "Saya merasa sistem ini membingungkan",
    9: "Saya merasa tidak ada hambatan dalam menggunakan sistem ini",
    10: "Saya perlu mempelajari banyak hal terlebih dahulu sebelum dapat menggunakan sistem ini dengan baik.  ",
}

question_col_map = {}
for i in range(1, 11):
    qtext = SUS_QUESTIONS[i].strip().lower()
    for col in db.columns:
        if str(col).strip().lower() == qtext:
            question_col_map[i] = col
            break

sus_map = {}

for _, row in db.iterrows():
    app_full = str(row["Nama Aplikasi"]).strip()
    if not app_full:
        continue

    app_name, owner_nim_raw, owner_name = parse_app_full(app_full)
    owner_nim_norm = normalize_nim(owner_nim_raw)

    try:
        jumlah_raw = float(row["Jumlah"])
    except:
        continue

    sus_score = jumlah_raw * 2.5

    respondent_name = str(row["Nama Responden (Participant)"]).strip()
    respondent_nim = normalize_nim(row["NIM"])

    qvals = {}
    qscores = {}

    for qi in range(1, 11):
        col = question_col_map.get(qi)
        if not col:
            qvals[f"q{qi}"] = None
            qscores[f"s{qi}"] = None
            continue

        val = row[col]
        if pd.isna(val):
            qvals[f"q{qi}"] = None
            qscores[f"s{qi}"] = None
            continue

        val = int(val)
        qvals[f"q{qi}"] = val

        if qi % 2 == 1:
            qscores[f"s{qi}"] = val - 1
        else:
            qscores[f"s{qi}"] = 5 - val

    key = (app_name, owner_nim_norm, owner_name)
    if key not in sus_map:
        sus_map[key] = {
            "app": app_name,
            "owner_name": owner_name,
            "owner_nim": owner_nim_norm,
            "scores": [],
            "responses": []
        }

    sus_map[key]["scores"].append(sus_score)
    sus_map[key]["responses"].append({
        "respondent_name": respondent_name,
        "respondent_nim": respondent_nim,
        "jumlah_raw": jumlah_raw,
        "sus": sus_score,
        **qvals,
        **qscores,
        "generated_at": generated_at
    })

sus_scores = []
for key, entry in sus_map.items():
    avg = sum(entry["scores"]) / len(entry["scores"])
    sus_scores.append({
        "app": entry["app"],
        "owner_name": entry["owner_name"],
        "owner_nim": entry["owner_nim"],
        "avg_sus": avg,
        "count": len(entry["scores"]),
        "responses": entry["responses"],
        "generated_at": generated_at
    })

# ===============================
# 9. SIMPAN JSON
# ===============================
json.dump(result, open("data.json","w",encoding="utf-8"), ensure_ascii=False, indent=2)
json.dump(nim_issues, open("nim_issues.json","w",encoding="utf-8"), ensure_ascii=False, indent=2)
json.dump(sus_scores, open("sus_scores.json","w",encoding="utf-8"), ensure_ascii=False, indent=2)

print("Berhasil membuat data.json, nim_issues.json, sus_scores.json")
