"""Rebuild the POL Locations PWA from source Excel files.

Run from the pol-locations-pwa folder:  python3 build_app.py
Sources (in parent folder): POL LOCS.xlsx (sheet 124 LOCS), tankstk DD.MM.YYYY.xlsx (sheet tankstk),
TT_Loading_Summary_*.xlsx (sheet Summary)
Standing corrections (per Tarun):
  1. Androth Terminal (4253): use file figure; fall back to 310 KL only if absent from tankstk
  2. Ambala Terminal (1122): no MS tankage (removed from products, totals and maintenance list)
  3. Avg TT loading/day, Monthly Rake Unloading and ATF TTs Handling/Day rounded to whole numbers
  4. Tank-wise details show no status remarks (Operative / Under Receipt) -- handled in template
"""
import openpyxl, json, datetime, collections, os, re

SRC_LOCS = "../POL LOCS.xlsx"
SRC_TANK = "../tankstk 01.10.2026.xlsx"
SRC_TT = "../TT LOADING APR SEP 2026.XLSX"   # raw month-wise TT loading report
TT_MONTHS = {"202604", "202605", "202606", "202607", "202608", "202609"}
TT_PERIOD = "Apr–Sep 2026"
ASOF = "01.10.2026"
REGION = {"NR": "Northern Region", "ER": "Eastern Region", "WR": "Western Region", "SR": "Southern Region"}

def clean(v):
    if isinstance(v, str):
        v = v.strip()
        return v if v else None
    return v

# ---- master ----
wb = openpyxl.load_workbook(SRC_LOCS, read_only=True, data_only=True)
ws = wb["124 LOCS"]
locs = []
for r in ws.iter_rows(min_row=3, values_only=True):
    if r[1] is None:
        continue
    r = [clean(x) for x in r]
    comm = r[11]
    comm_str = comm.strftime("%d %b %Y") if isinstance(comm, datetime.datetime) else (str(comm) if comm else None)
    locs.append({
        "plant": r[1], "region": REGION.get(r[2], r[2]), "soCode": r[3], "so": r[4], "name": r[5],
        "district": r[6], "stateCode": r[7], "state": r[8], "type": r[9], "area": r[10],
        "commissioned": comm_str, "age": r[12], "shift": r[13], "lat": r[14], "lon": r[15],
        "autoTT": r[16], "autoTFMS": r[17], "autoManual": r[18], "automated": r[19], "smart": r[20],
        "receipt": r[21], "delivery": r[22], "vendorTTES": r[23], "vendorTFMS": r[24],
        "prodType": r[25], "prodHandled": r[26], "tlf": r[27],
        "throughput": round(r[28]) if isinstance(r[28], (int, float)) else r[28],
        "ebms": r[29],
        "avgTT": round(r[31]) if isinstance(r[31], (int, float)) else r[31],
        "rake": round(r[32]) if isinstance(r[32], (int, float)) else r[32],
        "atfDay": round(r[33]) if isinstance(r[33], (int, float)) else r[33]})
wb.close()
assert len(locs) == 124, f"expected 124 locations, got {len(locs)}"

# ---- tankage (Material Name col J, Tankage col L, Tank col K, status col U) ----
wb = openpyxl.load_workbook(SRC_TANK, read_only=True, data_only=True)
ws = wb["tankstk"]
tk = collections.defaultdict(lambda: collections.defaultdict(lambda: [0.0, 0, []]))
maint = collections.defaultdict(list)   # plant -> tanks under maintenance (col M > 0)
for r in ws.iter_rows(min_row=2, values_only=True):
    if r[1] is None or r[9] is None:
        continue
    m = r[12]
    if isinstance(m, (int, float)) and m > 0:
        maint[r[1]].append({"k": str(r[10]).strip() if r[10] is not None else "–",
                            "m": str(r[9]).strip(),
                            "c": round(r[11] or 0),
                            "u": round(m)})
    e = tk[r[1]][str(r[9]).strip()]
    e[0] += r[11] or 0
    e[1] += 1
    e[2].append({"k": str(r[10]).strip() if r[10] is not None else "–",
                 "t": round(r[11] or 0),
                 "s": clean(r[20])})
wb.close()

for l in locs:
    mats = tk.get(l["plant"], {})
    prods = sorted(((m, round(v[0]), v[1], v[2]) for m, v in mats.items() if v[0] > 0), key=lambda x: -x[1])
    l["products"] = [{"m": m, "t": t, "n": n,
                      "tanks": sorted(tanks, key=lambda x: -x["t"])} for m, t, n, tanks in prods]
    l["tankageTotal"] = round(sum(v[0] for v in mats.values()))
    l["tankCount"] = sum(v[1] for v in mats.values())
    l["maint"] = sorted(maint.get(l["plant"], []), key=lambda x: -x["c"])

# ---- TT loading (material group wise) from raw month-wise report ----
# Columns: A State Office, B Plant, C Month (YYYYMM), D Material Group, E Total Volume KL,
#          G Total No. of TTs, K No. of working days. Total rows (Month/Plant/State Office) excluded.
# Avg vol/month = total volume / months with loading; avg vol/day and TTs/day = totals / working days.
wb = openpyxl.load_workbook(SRC_TT, read_only=True, data_only=True)
ws = wb[wb.sheetnames[0]]
agg = collections.defaultdict(lambda: [0.0, 0.0, 0.0, 0])   # (plant, group) -> vol, tts, days, months
for r in ws.iter_rows(min_row=2, values_only=True):
    plant, month, grp = (str(r[1]).strip() if r[1] is not None else ""), str(r[2]).strip(), (r[3] or "")
    if not plant or month not in TT_MONTHS or "Total" in str(grp):
        continue
    vol = float(r[4] or 0); tts = float(r[6] or 0); days = float(r[10] or 0)
    if vol <= 0 and tts <= 0:
        continue
    e = agg[(plant, str(grp).strip())]
    e[0] += vol; e[1] += tts; e[2] += days; e[3] += 1
wb.close()
ttl = collections.defaultdict(list)
for (plant, grp), (vol, tts, days, n) in agg.items():
    ttl[plant].append({"g": grp, "vm": vol / n, "vd": vol / days if days else 0,
                       "td": tts / days if days else 0})

for l in locs:
    rows = ttl.get(str(l["plant"]), [])
    l["ttLoading"] = sorted(rows, key=lambda x: -x["vm"])

# ---- manual corrections ----
# Androth (4253): 310 KL override applies only when absent from tankstk file
a = next(l for l in locs if l["plant"] == 4253)
if a["tankageTotal"] == 0:
    a["tankageTotal"] = 310
    a["manualNote"] = True

amb = next(l for l in locs if l["plant"] == 1122)        # Ambala: no MS (standing correction)
removed = [p for p in amb["products"] if p["m"].startswith("MS")]
amb["products"] = [p for p in amb["products"] if not p["m"].startswith("MS")]
amb["tankageTotal"] -= sum(p["t"] for p in removed)
amb["tankCount"] -= sum(p["n"] for p in removed)
amb["maint"] = [t for t in amb["maint"] if not t["m"].startswith("MS")]

# ---- write ----
json.dump(locs, open("data.json", "w"), ensure_ascii=False)
tpl = open("app_template.html", encoding="utf-8").read()
open("index.html", "w", encoding="utf-8").write(
    tpl.replace("__DATA__", json.dumps(locs, ensure_ascii=False, separators=(",", ":")))
       .replace("__ASOF__", ASOF).replace("__TTPERIOD__", TT_PERIOD))

# bump service worker cache version
sw = open("sw.js", encoding="utf-8").read()
m = re.search(r'pol-locations-v(\d+)', sw)
new = f"pol-locations-v{int(m.group(1)) + 1}"
open("sw.js", "w", encoding="utf-8").write(sw.replace(m.group(0), new))

print("locations:", len(locs))
print("grand total tankage:", sum(l["tankageTotal"] for l in locs), "KL")
print("no tankage:", [(l["plant"], l["name"]) for l in locs if l["tankageTotal"] == 0])
print("Ambala:", amb["tankageTotal"], "KL", [(p["m"], p["t"]) for p in amb["products"]])
print("tanks under maintenance:", sum(len(l["maint"]) for l in locs),
      "at", sum(1 for l in locs if l["maint"]), "locations")
print("sw cache:", new)
