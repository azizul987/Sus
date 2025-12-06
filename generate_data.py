import os
import json
import re
import requests
import pandas as pd
from datetime import datetime, timezone

# =====================================================
# Timestamp
# =====================================================
generated_at = datetime.now(timezone.utc).isoformat()

# =====================================================
# 1. Download Excel
# =====================================================
URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSWDG_1e6OyoSQsKp_Yy4uVAeVJOCrGPVABzv29EgoiNlhfhWV9vqOW2M6mFOvvkGbssfW3PXyH3ybM/pub?output=xlsx"
EXCEL_FILE = "sus_responses.xlsx"

resp = requests.get(URL)
resp.raise_for_status()

with open(EXCEL_FILE, "wb") as f:
    f.write(resp.content)

# =====================================================
# 2. Baca sheet dbresponden
# =====================================================
xls = pd.ExcelFile(EXCEL_FILE)
db = pd.read_excel(xls, "dbresponden")

# Kolom wajib
required = ["Nama Aplikasi", "Nama Responden (Participant)", "NIM", "Jumlah"]
for col in required:
    if col not in db.columns:
        raise RuntimeError(f"Kolom wajib '{col}' tidak ditemukan.")

apps_col = db["Nama Aplikasi"].astype(str).str.strip()
participants = db["Nama Responden (Participant)"].astype(str).str.strip()

def normalize_nim(x):
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    return s

db["NIM_norm"] = db["NIM"].map(normalize_nim)
participants_nim_norm = db["NIM_norm"]
participants_norm_name = participants.str.lower()

# =====================================================
# 3. Parsing Nama Aplikasi
# =====================================================
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


# =====================================================
# 4. Pembuat Aplikasi
# =====================================================
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
    if nim_norm and nim_norm not in creators_by_nim:
        creators_by_nim[nim_norm] = {
            "name": info["name"],
            "nim": info["nim"],
            "nim_norm": nim_norm,
            "app": info["app"],
        }

all_creators = list(creators_by_nim.values())


# =====================================================
# 5. Rekap data.json
# =====================================================
app_full_count = apps_col.value_counts().to_dict()

result = []

for creator in all_creators:
    owner_name = creator["name"]
    owner_app = creator["app"]
    owner_nim_norm = creator["nim_norm"]

    mask_nim = (participants_nim_norm == owner_nim_norm)
    mask = mask_nim if mask_nim.any() else (participants_norm_name == owner_name.lower().strip())
    filled_count = int(mask.sum())

    rated_apps_full = apps_col[mask].unique()
    rated_names = {app_info_map.get(af, {}).get("app", parse_app_full(af)[0]) for af in rated_apps_full}

    not_filled = sorted({
        other["app"] for other in all_creators
        if other["nim_norm"] != owner_nim_norm and other["app"] not in rated_names
    })

    owner_fulls = [af for af, info in app_info_map.items() if info["nim"] == owner_nim_norm]
    app_filled_count = sum(app_full_count.get(af, 0) for af in owner_fulls)

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


# =====================================================
# 6. nim_issues.json
# =====================================================
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
            "name": name,
            "nims": unique_nims,
            "nim_counts": {normalize_nim(k): int(v) for k, v in counts.items()},
            "total_rows": int(sub.shape[0]),
            "generated_at": generated_at
        })


# =====================================================
# 7. SUS QUESTIONS + SCORES
# =====================================================
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
    for col in db.columns:
        if col.strip().lower() == SUS_QUESTIONS[i].strip().lower():
            question_col_map[i] = col
            break

sus_map = {}

for _, row in db.iterrows():
    app_full = row["Nama Aplikasi"]
    if not isinstance(app_full, str):
        continue
    app_full = app_full.strip()

    app_name, owner_nim_raw, owner_name = parse_app_full(app_full)
    owner_nim_norm = normalize_nim(owner_nim_raw)

    try:
        jumlah_raw = float(row["Jumlah"])
    except:
        continue

    sus_total = jumlah_raw * 2.5

    respondent_name = row["Nama Responden (Participant)"]
    respondent_nim = normalize_nim(row["NIM"])

    qvals = {}
    svals = {}

    for qi in range(1, 11):
        col = question_col_map.get(qi)
        raw = row[col] if col else None

        if raw is None or pd.isna(raw):
            qvals[f"q{qi}"] = None
            svals[f"s{qi}"] = None
            continue
        
        raw = int(raw)
        qvals[f"q{qi}"] = raw

        if qi % 2 == 1:  # Positive
            svals[f"s{qi}"] = raw - 1
        else:           # Negative
            svals[f"s{qi}"] = 5 - raw

    key = (app_name, owner_nim_norm, owner_name)
    if key not in sus_map:
        sus_map[key] = {
            "app": app_name,
            "owner_name": owner_name,
            "owner_nim": owner_nim_norm,
            "scores": [],
            "responses": []
        }

    sus_map[key]["scores"].append(sus_total)

    sus_map[key]["responses"].append({
        "respondent_name": respondent_name,
        "respondent_nim": respondent_nim,
        "jumlah_raw": jumlah_raw,
        "sus": sus_total,
        **qvals,
        **svals,
        "generated_at": generated_at
    })


# =====================================================
# 8. Final sus_scores.json
# =====================================================
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


# =====================================================
# 9. Save JSON
# =====================================================
json.dump(result, open("data.json","w",encoding="utf-8"), ensure_ascii=False, indent=2)
json.dump(nim_issues, open("nim_issues.json","w",encoding="utf-8"), ensure_ascii=False, indent=2)
json.dump(sus_scores, open("sus_scores.json","w",encoding="utf-8"), ensure_ascii=False, indent=2)

print("JSON generated successfully.")
