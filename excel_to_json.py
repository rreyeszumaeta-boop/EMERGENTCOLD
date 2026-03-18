import json
import openpyxl
from collections import defaultdict

EXCEL_FILE = "BASEDEDATOS.xlsx"
OUTPUT_FILE = "dashboard-data.json"

def normalize_plant(name):
    if name is None:
        return None
    n = str(name).strip().lower()
    if "duran" in n:
        return "duran"
    if "caucedo" in n:
        return "caucedo"
    if "fp4" in n:
        return "fp4"
    return n

def sheet_to_rows(ws):
    rows = list(ws.iter_rows(values_only=True))
    rows = [row for row in rows if any(cell is not None for cell in row)]
    if not rows:
        return []

    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    data_rows = rows[1:]

    result = []
    for row in data_rows:
        item = {}
        for i, key in enumerate(header):
            if not key:
                continue
            item[key] = row[i] if i < len(row) else None
        result.append(item)
    return result

wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)

daily_rows = sheet_to_rows(wb["daily_generation"])
monthly_rows = sheet_to_rows(wb["monthly_ytd"])
irr_rows = sheet_to_rows(wb["irradiation"])
alarm_rows = sheet_to_rows(wb["alarms"])
meta_rows = sheet_to_rows(wb["meta"])

monthlyData = {}
for row in daily_rows:
    month_key = str(row["month_key"]).strip()
    day_label = str(row["day_label"]).strip()

    if month_key not in monthlyData:
        monthlyData[month_key] = {
            "label": month_key,
            "days": [],
            "generation": {
                "duran": [],
                "caucedo": [],
                "fp4": []
            },
            "contractual": {
                "duran": [],
                "caucedo": [],
                "fp4": []
            }
        }

    monthlyData[month_key]["days"].append(day_label)
    monthlyData[month_key]["generation"]["duran"].append(float(row.get("duran_generation") or 0))
    monthlyData[month_key]["generation"]["caucedo"].append(float(row.get("caucedo_generation") or 0))
    monthlyData[month_key]["generation"]["fp4"].append(float(row.get("fp4_generation") or 0))

    monthlyData[month_key]["contractual"]["duran"].append(float(row.get("duran_contractual") or 0))
    monthlyData[month_key]["contractual"]["caucedo"].append(float(row.get("caucedo_contractual") or 0))
    monthlyData[month_key]["contractual"]["fp4"].append(float(row.get("fp4_contractual") or 0))

ytdMonthlySummary = {}
for row in monthly_rows:
    month_key = str(row["month_key"]).strip()
    ytdMonthlySummary[month_key] = {
        "duran": {
            "com": float(row.get("duran_com") or 0),
            "real": float(row.get("duran_real") or 0)
        },
        "caucedo": {
            "com": float(row.get("caucedo_com") or 0),
            "real": float(row.get("caucedo_real") or 0)
        },
        "fp4": {
            "com": float(row.get("fp4_com") or 0),
            "real": float(row.get("fp4_real") or 0)
        }
    }

irradiation = {}
for row in irr_rows:
    month_key = str(row["month_key"]).strip()
    irradiation[month_key] = {
        "duran_real": float(row.get("duran_real") or 0),
        "duran_contractual": float(row.get("duran_contractual") or 0),
        "fp4_real": float(row.get("fp4_real") or 0),
        "fp4_contractual": float(row.get("fp4_contractual") or 0)
    }

alarms = []
for row in alarm_rows:
    alarms.append({
        "month_key": str(row.get("month_key") or "").strip(),
        "plant": normalize_plant(row.get("plant")),
        "code": row.get("code"),
        "hours": float(row.get("hours") or 0),
        "loss_kwh": float(row.get("loss_kwh") or 0)
    })

meta = {}
for row in meta_rows:
    plant = normalize_plant(row.get("plant"))
    if not plant:
        continue
    meta[plant] = {
        "capacity_kwp": float(row.get("capacity_kwp") or 0)
    }

dashboard_data = {
    "monthlyData": monthlyData,
    "ytdMonthlySummary": ytdMonthlySummary,
    "irradiation": irradiation,
    "alarms": alarms,
    "meta": meta
}

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(dashboard_data, f, ensure_ascii=False, indent=2)

print(f"Archivo generado: {OUTPUT_FILE}")