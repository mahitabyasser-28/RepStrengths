# =====================================================================
# SALES REP STRENGTHS — SOLD UNITS BY PROJECT × UNIT TYPE (Colab)
# Source: ONLY the master reservations sheet (tab FY26). No Salesforce.
# Views: OVERALL this year + each QUARTER (Q1, Q2, Q3, Q4)
# ---------------------------------------------------------------------
# WHAT IT DOES
#   1. READS the master sheet (SOURCE, tab FY26 — never written to), copies its
#      values into your working sheet (TARGET, tab "FY26 (extract)") and reads
#      the data from there. Every run takes a fresh full copy, so anything added
#      or changed in the master is picked up on the next run.
#
#   2. LEAVES OUT (as if the rows were not there):
#        • Contracted = "Terminated Won"
#        • Team Lead = KSA - Speakol, Oman - Speakol, CIC   (EXCLUDE_TEAM_LEADS)
#        • Mr.Mahmoud's Referral ticked TRUE                  (EXCLUDE_REFERRALS)
#      Every left-out row is listed in the Excel sheet "Left Out" with the reason.
#      No filter on the client's country.
#
#   3. TOTAL SOLD = every remaining unit, broken down by the Contracted column:
#        Yes · No · Terminated · (blank / other)
#      WIN RATE = (Yes + No) ÷ (Yes + No + Terminated)
#      → a Terminated unit is still SOLD, but counts as LOST for the rep.
#
#   4. For every rep, overall and per quarter: units sold in EVERY project and
#      unit type, their share of the rep's sales (= strength), and the win rate.
#      • Quarter = the sheet's Quarter column (falls back to R. Date).
#      • A unit shared by two reps ("A and B") counts HALF for each
#        (SPLIT_SHARED), so team / company totals equal the sheet's row count.
#      • Rep → Team Leader = latest team leader written on their reservations;
#        Team Leader → Sales Manager = latest Sales Manager written with that TL.
#
# USAGE (Google Colab) — just run this cell. Colab asks once for Google sign-in.
#   Output: Rep_Strengths_<Day><Mon>.html + .xlsx (downloaded automatically)
#   No Google access? Set SOURCE_URL to a local .xlsx / .csv path instead.
#
# ON GITHUB ACTIONS: run_github.py runs this file unchanged and makes the Colab
#   sign-in use the Google service account (secret GOOGLE_SERVICE_ACCOUNT_JSON).
# =====================================================================

# !pip -q install gspread

import json, os, re, time
from datetime import datetime, date, timedelta, timezone
from collections import defaultdict, Counter
import pandas as pd

T0 = time.time()
def phase(msg): print(f"\n⏱ [{time.time()-T0:4.0f}s] {msg}")

# ------------------------- CONFIG (change per run) -------------------------
TZ = timezone(timedelta(hours=3))

# ◀◀ THE SPREADSHEETS
#   SOURCE = the master reservations sheet → only READ, never written to.
#   TARGET = your working sheet → each run copies the source tab into it, then the report reads from it.
SOURCE_URL = "https://docs.google.com/spreadsheets/d/1IyJegYDSyLPMywAxv5beHtjxip2uPn4v8pUvpjmv8Lc/edit?gid=1692628846"
SOURCE_TABS = ["FY26"]                     # ◀◀ tab name(s) to copy
SOURCE_GIDS = []                           # ◀◀ or tab id(s) (the number after gid= in the link) · both empty = every tab
TARGET_URL = "https://docs.google.com/spreadsheets/d/1CI5l-BZeuaqew-VXzxwjyyqdJyWIINz0XA0o1aw-Ydk/edit"
TARGET_TAB_SUFFIX = " (extract)"

# ◀◀ ROWS LEFT OUT (as if they were not in the sheet)
EXCLUDE_TEAM_LEADS = ["KSA - Speakol", "Oman - Speakol", "CIC"]   # exact Team Lead values
EXCLUDE_CONTRACTED = ["Terminated Won"]                           # Contracted values
EXCLUDE_REFERRALS = True                                          # Mr.Mahmoud's Referral = TRUE

# ◀◀ CONTRACTED → bucket. Win rate = (Yes + No) ÷ (Yes + No + Terminated)
CONTRACTED_YES = ["yes"]
CONTRACTED_NO = ["no"]
CONTRACTED_TERMINATED = ["terminated"]
#   Values are matched loosely: case, extra spaces, dashes and brackets are ignored, so
#   "YES", "Yes ", "Terminated-Won", "terminated (won)" all work. Anything else → "Not set"
#   (counted in Total Sold, not in the win rate) and printed in the console.
CONTRACTED_COLUMN = ""                     # ◀◀ exact header of the Contracted column · "" = find it automatically

# ◀◀ LAUNCH DAYS — False = launch peaks count like any other day (default) · True = leave them out
EXCLUDE_LAUNCH_DAYS = False
LAUNCH_WINDOWS = [
    ("2026-01-10", "2026-01-11"), ("2026-03-17", "2026-03-18"), ("2026-04-21", "2026-04-21"),
    ("2026-04-30", "2026-04-30"), ("2026-06-07", "2026-06-11"), ("2026-07-06", "2026-07-06"),
    ("2026-07-15", "2026-07-15"), ("2026-07-27", "2026-07-27"), ("2026-08-02", "2026-08-03"),
    ("2026-08-17", "2026-08-17"), ("2026-08-31", "2026-08-31"),
]

SPLIT_SHARED = True                        # ◀◀ "A and B" → 0.5 each · False = 1 each (totals then exceed the row count)
TL_SALES_MANAGER_OVERRIDE = {"Mohamed Karam": "Ahmed Sarwat (direct)"}   # TLs with no Sales Manager in the sheet
NAME_ALIASES = {"karim moghazy": "Kareem Moghazy"}                       # sheet spelling → one name
PROJECT_ALIASES = {"South Med": "SouthMed", "SharmBay": "Sharm Bay"}
CHECK_REPS = []                            # ◀◀ e.g. ["Mai Taki"] → prints every one of their rows and what happened to it

NO_PROJECT, NO_TYPE, NO_TL, NO_SM, NO_REP = "(no project)", "(type not set)", "(no team leader)", "(no sales manager)", "(no rep)"
MIN_DECIDED = 5                            # best project needs at least this many Yes + No + Terminated units
MIN_CLOSER = 20                            # headline "best win rate" needs at least this many
CURRENCY = "EGP"
FILE_TAG = "Rep_Strengths"
AUTO_DOWNLOAD = True
# ---------------------------------------------------------------------------


def norm(s): return re.sub(r"\s+", " ", str(s or "").strip()).lower()
def normtl(s): return re.sub(r"\s*-\s*", "-", norm(s))           # "KSA - Speakol" == "KSA-Speakol"
def pct(a, b): return (a / b) if b else None
def alias(n):
    n = re.sub(r"\s+", " ", str(n or "").strip())
    return NAME_ALIASES.get(n.lower(), n)
def money(v):
    s = re.sub(r"[^\d.]", "", str(v or ""))
    try: return float(s) if s else 0.0
    except ValueError: return 0.0
def parse_date(v):
    s = str(v or "").strip()
    if not s: return None
    if re.fullmatch(r"\d{5}(\.0+)?", s):                           # Sheets serial (e.g. 46180 = 7 Jun 2026)
        return date(1899, 12, 30) + timedelta(days=int(float(s)))
    for f in ("%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d-%b-%Y", "%m/%d/%y"):
        try: return datetime.strptime(s, f).date()
        except ValueError: pass
    return None
def split_names(v):
    return [alias(x) for x in re.split(r"\s+and\s+|\s*&\s*|\s*/\s*", str(v or "").strip(), flags=re.I) if x.strip()]
def clean_name(n): return "" if not n or n.strip().upper() in ("NONE", "N/A", "-", "#N/A") else n


# ============================ 1. READ THE SHEET ============================
COLS = {   # our name → how to recognise the header (first match wins)
    "ref":   lambda h: "referral" in h,
    "sm":    lambda h: h in ("sales manger", "sales manager"),
    "tl":    lambda h: h in ("team lead", "team leader"),
    "rep":   lambda h: h in ("sales", "sales rep", "sales person"),
    "client": lambda h: h.startswith("client name"),
    "proj":  lambda h: h == "project",
    "type":  lambda h: h == "type",
    "resv":  lambda h: h.replace(".", "").replace(" ", "") in ("resrvno", "reservationno", "reservationnumber"),
    "unit":  lambda h: h in ("old ro", "unit", "unit code"),
    "value": lambda h: h == "value",
    "chan":  lambda h: h == "channel",
    "src":   lambda h: h == "source",
    "country": lambda h: h == "country",
    "date":  lambda h: h in ("r. date", "r.date", "reservation date", "date"),
    "quarter": lambda h: h == "quarter",
}

def map_cols(header):
    hs = [norm(h) for h in header]
    out = {}
    for k, test in COLS.items():
        j = next((i for i, h in enumerate(hs) if test(h)), None)
        if j is not None: out[k] = j
    return out

def sheet_key(url):
    m = re.search(r"/d/([a-zA-Z0-9-_]+)", url or "")
    return m.group(1) if m else None

def extract_to_target():
    """Read the SOURCE sheet (read-only) → write the values into the TARGET sheet → read them back from the target."""
    from google.colab import auth
    auth.authenticate_user()
    import gspread
    from google.auth import default
    creds, _ = default()
    gc = gspread.authorize(creds)
    src_key, tgt_key = sheet_key(SOURCE_URL), sheet_key(TARGET_URL)
    if not src_key or not tgt_key: raise SystemExit("❌ SOURCE_URL / TARGET_URL is not a Google Sheets link")
    if src_key == tgt_key: raise SystemExit("❌ SOURCE and TARGET are the same sheet — refusing to write to the source")

    src = gc.open_by_key(src_key)                      # ── SOURCE: only get_all_values() is called on it
    want = {norm(t) for t in SOURCE_TABS}
    all_tabs = src.worksheets()
    tabs = [ws for ws in all_tabs if (not want and not SOURCE_GIDS) or norm(ws.title) in want or ws.id in SOURCE_GIDS]
    missing = want - {norm(ws.title) for ws in all_tabs}
    if missing: print(f"   ⚠️ tab(s) not found: {sorted(missing)} — tabs in the sheet: {[ws.title for ws in all_tabs]}")
    if not tabs: raise SystemExit(f"❌ none of the tabs {SOURCE_TABS or SOURCE_GIDS} are in the source sheet")
    data = [(ws.title, ws.get_all_values()) for ws in tabs]
    print(f"   📖 read from source “{src.title}” (not modified): " + ", ".join(f"“{t}” {len(v):,} rows" for t, v in data))

    tgt = gc.open_by_key(tgt_key)                      # ── TARGET: cleared and refilled every run
    existing = {ws.title: ws for ws in tgt.worksheets()}
    out = []
    for title, vals in data:
        name = (title + TARGET_TAB_SUFFIX)[:100]
        ncols = max((len(r) for r in vals), default=1)
        vals = [r + [""] * (ncols - len(r)) for r in vals]
        ws = existing.get(name) or tgt.add_worksheet(title=name, rows=max(len(vals), 1), cols=max(ncols, 1))
        ws.clear()
        ws.resize(rows=max(len(vals), 1), cols=max(ncols, 1))
        for i in range(0, len(vals), 5000):
            ws.update(range_name=f"A{i+1}", values=vals[i:i+5000], value_input_option="RAW")
        out.append((name, ws))
    info = existing.get("_refresh_log") or tgt.add_worksheet(title="_refresh_log", rows=200, cols=4)
    if not info.get_all_values(): info.append_row(["Refreshed (Cairo)", "Source", "Tabs", "Rows"])
    info.append_row([datetime.now(TZ).strftime("%Y-%m-%d %H:%M"), src.title, ", ".join(n for n, _ in out),
                     sum(len(v) for _, v in data)])
    print(f"   ✍️ written to target “{tgt.title}”: " + ", ".join(f"“{n}”" for n, _ in out))
    return [(name, ws.get_all_values()) for name, ws in out]

def col_letter(j):
    s = ""
    j += 1
    while j: j, r = divmod(j - 1, 26); s = chr(65 + r) + s
    return s

def contract_key(v):
    """'Terminated - Won ' → 'terminated won'"""
    return re.sub(r"\s+", " ", re.sub(r"[^a-z]+", " ", str(v or "").lower())).strip()

def contracted(v):
    c = contract_key(v)
    if not c: return "unset"
    if c in {contract_key(x) for x in EXCLUDE_CONTRACTED} or ("terminat" in c and "won" in c): return "excluded"
    if c in {contract_key(x) for x in CONTRACTED_TERMINATED} or c.startswith("terminat"): return "term"
    if c in {contract_key(x) for x in CONTRACTED_YES} or c in ("y", "contracted"): return "yes"
    if c in {contract_key(x) for x in CONTRACTED_NO} or c in ("n", "not contracted"): return "no"
    return "unset"

def find_contract_col(header, body, skip=()):
    """Pick the Contracted column: CONTRACTED_COLUMN if set, else the column (preferably named '…contract…')
    whose filled cells are mostly Yes / No / Terminated / Terminated Won."""
    hs = [norm(h) for h in header]
    if CONTRACTED_COLUMN:
        j = next((i for i, h in enumerate(hs) if h == norm(CONTRACTED_COLUMN)), None)
        if j is None: print(f"   ⚠️ CONTRACTED_COLUMN “{CONTRACTED_COLUMN}” not found — searching automatically")
        else: return j, "set in CONTRACTED_COLUMN"
    best = None
    for j in range(len(header)):
        if j in skip: continue                              # never the referral / date / name columns
        vals = [str(r[j]).strip() for r in body if j < len(r) and str(r[j]).strip()]
        if not vals: continue
        score = sum(contracted(v) != "unset" for v in vals) / len(vals)
        named = "contract" in hs[j]
        if not named and not any(contract_key(v).startswith("terminat") for v in vals): continue   # unnamed column must contain Terminated
        if score >= 0.5 or (named and score > 0):
            key = (score >= 0.5, named, score, len(vals))
            if best is None or key > best[0]: best = (key, j)
    if best: return best[1], f"{best[0][2]:.0%} of filled cells are Yes / No / Terminated"
    named = [i for i, h in enumerate(hs) if "contract" in h]
    if named: return named[0], "matched by header name only — values not recognised"
    print("   ⚠️ no Contracted column found — every unit will be 'Not set'. Set CONTRACTED_COLUMN to its header.")
    return None, ""

def read_sheet():
    """→ list of raw row dicts (one per reservation row)."""
    if str(SOURCE_URL).startswith("http"):
        tabs = extract_to_target()
    else:
        path, tabs = SOURCE_URL, []
        if path.lower().endswith(".csv"):
            tabs.append((os.path.basename(path), pd.read_csv(path, dtype=str, header=None, keep_default_na=False).values.tolist()))
        else:
            for name, df in pd.read_excel(path, sheet_name=None, dtype=str, header=None).items():
                if SOURCE_TABS and norm(name) not in {norm(t) for t in SOURCE_TABS}: continue
                tabs.append((name, df.fillna("").values.tolist()))
    raw = []
    for name, vals in tabs:
        hdr_i = next((i for i, r in enumerate(vals[:10]) if {"rep", "proj", "type"} <= set(map_cols(r))), None)
        if hdr_i is None: print(f"   ↷ tab “{name}” skipped (no reservation columns)"); continue
        cm = map_cols(vals[hdr_i])
        j, why = find_contract_col(vals[hdr_i], vals[hdr_i + 1:], skip=set(cm.values()))
        if j is not None:
            cm["contract"] = j
            seen = Counter(str(r[j]).strip() or "(blank)" for r in vals[hdr_i + 1:] if j < len(r) and any(str(x).strip() for x in r))
            print(f"   📌 Contracted = column {col_letter(j)} “{vals[hdr_i][j]}” ({why}) · values: "
                  + ", ".join(f"“{k}” ×{v}" for k, v in seen.most_common(8)))
        need = [k for k in ("contract", "tl", "ref", "quarter", "date") if k not in cm]
        if need: print(f"   ⚠️ tab “{name}”: column(s) not found: {need} — headers are {[h for h in vals[hdr_i] if h][:30]}")
        n0 = len(raw)
        for i, r in enumerate(vals[hdr_i + 1:], start=hdr_i + 2):
            g = lambda k: (str(r[cm[k]]).strip() if k in cm and cm[k] < len(r) else "")
            if not any(str(x).strip() for x in r): continue                    # fully empty row
            if not (g("rep") or g("proj") or g("resv") or g("client")): continue
            raw.append({k: g(k) for k in [*COLS, "contract"]} | {"tab": name, "row": i})
        print(f"   ✅ tab “{name}”: {len(raw)-n0:,} reservation rows")
    return raw


# ============================ 2. BUILD ============================
def quarter_of(r, d):
    q = norm(r.get("quarter")).upper()
    if q in ("Q1", "Q2", "Q3", "Q4"): return q
    return f"Q{(d.month - 1) // 3 + 1}" if d else None

def build(raw):
    excl_tl = {normtl(x) for x in EXCLUDE_TEAM_LEADS}
    ldays = set()
    if EXCLUDE_LAUNCH_DAYS:
        for a, b in LAUNCH_WINDOWS:
            x, e = date.fromisoformat(a), date.fromisoformat(b)
            while x <= e: ldays.add(x); x += timedelta(days=1)
    rows, left_out, odd_contract, no_q = [], [], Counter(), 0
    for r in raw:
        d = parse_date(r["date"])
        q = quarter_of(r, d)
        reps = [x for x in (clean_name(n) for n in split_names(r["rep"])) if x] or [NO_REP]
        w = (1 / len(reps)) if SPLIT_SHARED else 1
        st = contracted(r["contract"])
        pj = PROJECT_ALIASES.get(r["proj"].strip(), r["proj"].strip()) or NO_PROJECT
        base = {"rep": None, "proj": pj, "type": r["type"].strip() or NO_TYPE, "q": q or "—", "date": d.isoformat() if d else r["date"],
                "resv": re.sub(r"\D", "", r["resv"]), "client": r["client"], "unit": r["unit"], "status": st,
                "contract_raw": r["contract"], "amt": money(r["value"]), "tl_raw": r["tl"], "sm_raw": r["sm"],
                "chan": r["chan"], "src": r["src"], "country": r["country"], "row": r["row"], "tab": r["tab"]}
        reason = None
        if any(normtl(t) in excl_tl for t in (split_names(r["tl"]) or [r["tl"]])) or normtl(r["tl"]) in excl_tl:
            reason = f"Team Lead: {r['tl']}"
        elif st == "excluded": reason = f"Contracted: {r['contract']}"
        elif EXCLUDE_REFERRALS and r["ref"].strip().upper() in ("TRUE", "1", "YES"): reason = "Mr. Mahmoud's Referral"
        elif d and d in ldays: reason = "Launch day"
        elif not q: reason = "No quarter / unreadable date"; no_q += 1
        for pos, rep in enumerate(reps):
            row = base | {"rep": rep, "pos": pos, "w": w, "amt": base["amt"] * w}
            if reason: left_out.append(row | {"reason": reason})
            else: rows.append(row)
        if not reason and st == "unset": odd_contract[r["contract"] or "(blank)"] += 1

    # --- team: rep → latest TL; TL → latest Sales Manager written with that TL
    rep_tl, rep_sm, tl_sm = {}, {}, {}
    sm_names = {norm(x) for r in rows for x in split_names(r["sm_raw"])}
    for r in sorted(rows, key=lambda r: (r["date"], r["row"])):
        tls, sms = split_names(r["tl_raw"]), split_names(r["sm_raw"])
        i = r["pos"]                                   # shared row: "A and B" ↔ "TL1 and TL2"
        tl = clean_name(tls[i] if i < len(tls) else (tls[0] if tls else ""))
        sm = clean_name(sms[i] if i < len(sms) else (sms[0] if sms else ""))
        if tl and not sm and norm(tl) in sm_names: sm, tl = tl, ""
        if tl: rep_tl[r["rep"]] = tl
        else: rep_tl.setdefault(r["rep"], NO_TL)
        if sm: rep_sm[r["rep"]] = sm
        if tl and sm: tl_sm[tl] = sm
    for tl, sm in TL_SALES_MANAGER_OVERRIDE.items(): tl_sm[alias(tl)] = sm
    team = {rep: (tl_sm.get(tl) or rep_sm.get(rep) or NO_SM, tl) for rep, tl in rep_tl.items()}
    for r in rows + left_out: r["sm"], r["tl"] = team.get(r["rep"], (clean_name(r["sm_raw"]) or NO_SM, clean_name(r["tl_raw"]) or NO_TL))

    # --- console summary
    n_rows = len({(r["tab"], r["row"]) for r in rows})
    units = lambda st: sum(r["w"] for r in rows if r["status"] == st)
    print(f"   ✅ counted {n_rows:,} reservation rows = {sum(r['w'] for r in rows):,.1f} units (excluding left out)")
    print(f"   Σ  including left out: {sum(r['w'] for r in rows) + sum(x['w'] for x in left_out):,.1f} units "
          f"({len({(x['tab'], x['row']) for x in left_out}):,} left-out rows)")
    print(f"      Yes {units('yes'):,.1f} · No {units('no'):,.1f} · Terminated {units('term'):,.1f} · Not set {units('unset'):,.1f}")
    lo_rows = defaultdict(set)
    for x in left_out: lo_rows[x["reason"]].add((x["tab"], x["row"]))
    for why, s in sorted(lo_rows.items(), key=lambda kv: -len(kv[1])): print(f"   ✂️ left out {len(s):,} rows — {why}")
    if odd_contract:
        print(f"   ⚠️ Contracted values not recognised (counted in Total Sold, not in the win rate): "
              + ", ".join(f"“{k}” ×{v}" for k, v in odd_contract.most_common(10)))
    for q in ("Q1", "Q2", "Q3", "Q4"):
        s = [r for r in rows if r["q"] == q]
        if s:
            y, n, t = (sum(r["w"] for r in s if r["status"] == k) for k in ("yes", "no", "term"))
            print(f"   {q}: {sum(r['w'] for r in s):,.1f} sold · win rate {pct(y + n, y + n + t) or 0:.0%}")

    for name in CHECK_REPS:
        nm = norm(alias(name)); print(f"\n   🔍 {name}:")
        for r in sorted([r for r in rows if norm(r["rep"]) == nm], key=lambda r: r["date"]):
            print(f"      ✅ counted   {r['date']}  {r['q']}  {r['resv']:>9}  {r['proj']} / {r['type']}  · {r['contract_raw'] or '(blank)'}  · units {r['w']:g}")
        for x in sorted([x for x in left_out if norm(x["rep"]) == nm], key=lambda x: x["date"]):
            print(f"      ✂️ left out  {x['date']}  {x['q']}  {x['resv']:>9}  {x['proj']} / {x['type']}  → {x['reason']}")
    return rows, left_out, team


# ============================ 3. AGGREGATE ============================
def aggregate(rows):
    """→ reps {rep: summary}, combos [rep×project×type], rps [rep×project], pts [project×type]"""
    def blank(**kw): return {**kw, "sold": 0.0, "yes": 0.0, "no": 0.0, "term": 0.0, "unset": 0.0, "amt": 0.0}
    def add(a, r):
        a["sold"] += r["w"]; a[r["status"]] += r["w"]; a["amt"] += r["amt"]
    def rate(a): a["win_rate"] = pct(a["yes"] + a["no"], a["yes"] + a["no"] + a["term"]); return a
    combo, rp, pt, reps = {}, {}, {}, {}
    for r in rows:
        add(combo.setdefault((r["rep"], r["proj"], r["type"]), blank(rep=r["rep"], tl=r["tl"], sm=r["sm"], proj=r["proj"], type=r["type"])), r)
        add(rp.setdefault((r["rep"], r["proj"]), blank(rep=r["rep"], tl=r["tl"], sm=r["sm"], proj=r["proj"])), r)
        p = pt.setdefault((r["proj"], r["type"]), blank(proj=r["proj"], type=r["type"], _reps=set())); add(p, r); p["_reps"].add(r["rep"])
        add(reps.setdefault(r["rep"], blank(rep=r["rep"], tl=r["tl"], sm=r["sm"])), r)
    for d in (combo, rp, pt, reps):
        for a in d.values(): rate(a)
    for c in combo.values(): c["share"] = pct(c["sold"], reps[c["rep"]]["sold"])
    for p in pt.values(): p["reps"] = len(p.pop("_reps"))
    for r in reps.values():
        mine = [c for c in combo.values() if c["rep"] == r["rep"]]
        top = sorted(mine, key=lambda c: (-c["sold"], -c["amt"]))[:1]
        types = Counter()
        for c in mine: types[c["type"]] += c["sold"]
        tt = types.most_common(1)
        projs = [p for p in rp.values() if p["rep"] == r["rep"] and p["proj"] != NO_PROJECT]
        best = sorted([p for p in projs if p["yes"] + p["no"] + p["term"] >= MIN_DECIDED],
                      key=lambda p: (-(p["win_rate"] or 0), -p["sold"]))[:1]
        weak = sorted([p for p in projs if p["term"] > 0], key=lambda p: (-p["term"], p["win_rate"] or 0))[:1]
        r.update({"top_combo": f"{top[0]['proj']} / {top[0]['type']}" if top else None,
                  "top_share": pct(top[0]["sold"], r["sold"]) if top else None,
                  "top_type": tt[0][0] if tt else None, "top_type_share": pct(tt[0][1], r["sold"]) if tt else None,
                  "best_proj": best[0]["proj"] if best else None, "best_proj_rate": best[0]["win_rate"] if best else None,
                  "weak_proj": weak[0]["proj"] if weak else None, "weak_proj_term": weak[0]["term"] if weak else None})
    return reps, list(combo.values()), list(rp.values()), list(pt.values())

def by_period(rows):
    out = {"Overall": aggregate(rows)}
    for q in ("Q1", "Q2", "Q3", "Q4"):
        sub = [r for r in rows if r["q"] == q]
        if sub: out[q] = aggregate(sub)
    return out


# ============================ 4. EXCEL ============================
def write_excel(path, rows, left_out, gen):
    phase("3) Writing Excel…")
    P = by_period(rows); periods = list(P)
    reps_all = P["Overall"][0]; qn = {q: P[q][0] for q in periods if q != "Overall"}
    porder = {p: i for i, p in enumerate(periods)}
    inP = lambda x, p: p == "Overall" or x["q"] == p
    rep_lo = lambda rep, p: sum(x["w"] for x in left_out if x["rep"] == rep and not x["reason"].startswith("Team Lead") and inP(x, p))
    def counts(a, prefix=""):
        return {f"{prefix}Total Sold": a["sold"], f"{prefix}Contracted: Yes": a["yes"], f"{prefix}Contracted: No": a["no"],
                f"{prefix}Terminated": a["term"], f"{prefix}Contracted Not Set": a["unset"], f"{prefix}Win Rate": a["win_rate"]}
    rep_df = pd.DataFrame([{
        "Sales Manager": r["sm"], "Team Leader": r["tl"], "Sales Rep": r["rep"], **counts(r),
        "Left Out": rep_lo(r["rep"], "Overall"), "Total Incl. Left Out": r["sold"] + rep_lo(r["rep"], "Overall"),
        **{f"Sold {q}": qn[q].get(r["rep"], {}).get("sold", 0) for q in qn},
        **{f"Win Rate {q}": qn[q].get(r["rep"], {}).get("win_rate") for q in qn},
        f"Value Sold ({CURRENCY})": r["amt"], "#1 Strength (Project / Type)": r["top_combo"], "#1 Share Of Sales": r["top_share"],
        "Main Unit Type": r["top_type"], "Type Share Of Sales": r["top_type_share"],
        "Best Project (Win Rate)": r["best_proj"], "Best Project Win Rate": r["best_proj_rate"],
        "Weak Spot (Most Terminated)": r["weak_proj"], "Terminated There": r["weak_proj_term"]} for r in reps_all.values()]
    ).sort_values(["Sales Manager", "Team Leader", "Total Sold"], ascending=[True, True, False], ignore_index=True)
    rq_df = pd.DataFrame([{"Period": p, "Sales Manager": r["sm"], "Team Leader": r["tl"], "Sales Rep": r["rep"], **counts(r),
                           "Left Out": rep_lo(r["rep"], p), "Total Incl. Left Out": r["sold"] + rep_lo(r["rep"], p),
                           f"Value Sold ({CURRENCY})": r["amt"], "#1 Strength (Project / Type)": r["top_combo"],
                           "#1 Share Of Sales": r["top_share"], "Main Unit Type": r["top_type"],
                           "Best Project (Win Rate)": r["best_proj"], "Best Project Win Rate": r["best_proj_rate"],
                           "Weak Spot (Most Terminated)": r["weak_proj"]} for p in periods for r in P[p][0].values()])
    rq_df = rq_df.assign(_o=rq_df["Period"].map(porder)).sort_values(["Sales Rep", "_o"], ignore_index=True).drop(columns="_o")
    combo_df = pd.DataFrame([{"Period": p, "Sales Manager": c["sm"], "Team Leader": c["tl"], "Sales Rep": c["rep"],
                              "Project": c["proj"], "Unit Type": c["type"], **counts(c), "Share Of Rep's Sales": c["share"],
                              f"Value Sold ({CURRENCY})": c["amt"]} for p in periods for c in P[p][1]])
    combo_df = combo_df.assign(_o=combo_df["Period"].map(porder)).sort_values(["Sales Rep", "_o", "Total Sold"], ascending=[True, True, False], ignore_index=True).drop(columns="_o")
    rp_df = pd.DataFrame([{"Period": p, "Sales Manager": x["sm"], "Team Leader": x["tl"], "Sales Rep": x["rep"], "Project": x["proj"],
                           **counts(x), f"Value Sold ({CURRENCY})": x["amt"]} for p in periods for x in P[p][2]])
    rp_df = rp_df.assign(_o=rp_df["Period"].map(porder)).sort_values(["Sales Rep", "_o", "Total Sold"], ascending=[True, True, False], ignore_index=True).drop(columns="_o")
    pt_rows = []
    for p in periods:
        lo_pt = Counter()
        for x in left_out:
            if inP(x, p): lo_pt[(x["proj"], x["type"])] += x["w"]
        seen = set()
        for x in P[p][3]:
            k = (x["proj"], x["type"]); seen.add(k)
            pt_rows.append({"Period": p, "Project": x["proj"], "Unit Type": x["type"], **counts(x), "Left Out": lo_pt[k],
                            "Total Incl. Left Out": x["sold"] + lo_pt[k], "Reps Who Sold It": x["reps"], f"Value Sold ({CURRENCY})": x["amt"]})
        for k, v in lo_pt.items():                     # combinations that only exist among the left-out rows
            if k not in seen:
                pt_rows.append({"Period": p, "Project": k[0], "Unit Type": k[1], "Total Sold": 0, "Contracted: Yes": 0, "Contracted: No": 0,
                                "Terminated": 0, "Contracted Not Set": 0, "Win Rate": None, "Left Out": v, "Total Incl. Left Out": v,
                                "Reps Who Sold It": 0, f"Value Sold ({CURRENCY})": 0})
    pt_df = pd.DataFrame(pt_rows)
    pt_df = pt_df.assign(_o=pt_df["Period"].map(porder)).sort_values(["_o", "Total Sold"], ascending=[True, False], ignore_index=True).drop(columns="_o")
    tl_rows = []
    for p in periods:
        agg = defaultdict(lambda: {"reps": 0, "sold": 0.0, "yes": 0.0, "no": 0.0, "term": 0.0, "unset": 0.0, "amt": 0.0})
        for r in P[p][0].values():
            a = agg[(r["sm"], r["tl"])]; a["reps"] += 1
            for k in ("sold", "yes", "no", "term", "unset", "amt"): a[k] += r[k]
            a["lo"] = a.get("lo", 0) + rep_lo(r["rep"], p)
        for (sm, tl), a in agg.items():
            a["win_rate"] = pct(a["yes"] + a["no"], a["yes"] + a["no"] + a["term"])
            tl_rows.append({"Period": p, "Sales Manager": sm, "Team Leader": tl, "Reps": a["reps"], **counts(a),
                            "Left Out": a.get("lo", 0), "Total Incl. Left Out": a["sold"] + a.get("lo", 0),
                            f"Value Sold ({CURRENCY})": a["amt"]})
    tl_df = pd.DataFrame(tl_rows)
    tl_df = tl_df.assign(_o=tl_df["Period"].map(porder)).sort_values(["_o", "Total Sold"], ascending=[True, False], ignore_index=True).drop(columns="_o")
    st_name = {"yes": "Yes", "no": "No", "term": "Terminated", "unset": "Not set", "excluded": "Terminated Won"}
    lo_df = pd.DataFrame([{"Reason": x["reason"], "Sheet Row": x["row"], "Quarter": x["q"], "Reservation Date": x["date"],
                           "Sales Rep": x["rep"], "Team Lead (sheet)": x["tl_raw"], "Reservation No.": x["resv"], "Client": x["client"],
                           "Project": x["proj"], "Unit Type": x["type"], "Contracted": x["contract_raw"], "Units (Share)": x["w"],
                           f"Value ({CURRENCY})": x["amt"]} for x in left_out])
    if len(lo_df): lo_df = lo_df.sort_values(["Reason", "Sheet Row"], ignore_index=True)
    all_df = pd.DataFrame([{"Sheet Row": r["row"], "Quarter": r["q"], "Reservation Date": r["date"], "Sales Manager": r["sm"],
                            "Team Leader": r["tl"], "Sales Rep": r["rep"], "Units (Share)": r["w"], "Contracted": r["contract_raw"],
                            "Contracted Bucket": st_name[r["status"]], "Reservation No.": r["resv"], "Client": r["client"],
                            "Project": r["proj"], "Unit Type": r["type"], f"Value ({CURRENCY})": r["amt"],
                            "Channel": r["chan"], "Source": r["src"], "Country": r["country"]} for r in rows]
                          ).sort_values(["Sheet Row", "Sales Rep"], ignore_index=True)
    reasons = sorted({x["reason"] for x in left_out})
    sum_rows = []
    for p in periods:
        sold = sum(r["w"] for r in rows if inP(r, p))
        lo = {why: sum(x["w"] for x in left_out if x["reason"] == why and inP(x, p)) for why in reasons}
        y, n, t = (sum(r["w"] for r in rows if inP(r, p) and r["status"] == k) for k in ("yes", "no", "term"))
        sum_rows.append({"Period": p, "Total Sold (Excluding Left Out)": sold, "Contracted: Yes": y, "Contracted: No": n, "Terminated": t,
                         "Win Rate": pct(y + n, y + n + t), **{f"Left Out — {k}": v for k, v in lo.items()},
                         "Left Out (All)": sum(lo.values()), "Total Sold (Including Left Out)": sold + sum(lo.values())})
    sum_df = pd.DataFrame(sum_rows)
    readme = pd.DataFrame([
        ["Source", f"Only the master reservations sheet, tab(s) {', '.join(SOURCE_TABS)} (copied read-only into the working sheet). Nothing from Salesforce."],
        ["Generated", gen],
        ["Total Sold (Excluding Left Out)", "Every unit left after the exclusions below, whatever its Contracted state. Used for all strengths and win rates."],
        ["Total Sold (Including Left Out)", "Total Sold + the left-out units (Terminated Won, excluded team leads, referrals): every unit in the sheet. Shown for reference only."],
        ["Contracted: Yes / No / Terminated", "The Contracted column. Terminated = still sold, but counted as lost for the rep."],
        ["Contracted Not Set", "Blank or unrecognised Contracted value: counted in Total Sold, not in the win rate."],
        ["Win Rate", "(Yes + No) ÷ (Yes + No + Terminated)."],
        ["Left out", f"Contracted = {', '.join(EXCLUDE_CONTRACTED)}; Team Lead = {', '.join(EXCLUDE_TEAM_LEADS)}"
                     + ("; Mr. Mahmoud's Referral = TRUE" if EXCLUDE_REFERRALS else "") + ("; launch days" if EXCLUDE_LAUNCH_DAYS else "")
                     + ". As if those rows were not there. Listed in sheet 'Left Out'."],
        ["Client country", "Not filtered."],
        ["Quarter", "The sheet's Quarter column (R. Date when it is empty)."],
        ["Shared units", "A unit sold by 'A and B' counts 0.5 for each rep, so totals equal the number of sheet rows." if SPLIT_SHARED
                         else "A shared unit counts 1 for each rep."],
        ["Team", "Rep → latest Team Leader on their reservations; Team Leader → latest Sales Manager written with that TL."],
        ["Best Project", f"Highest win rate among projects with at least {MIN_DECIDED} Yes + No + Terminated units."],
        ["Weak Spot", "Project with the most Terminated units."],
        ["All Counted Rows", "Every sheet row used, with its Sheet Row number so you can find it in the master."]], columns=["Item", "Meaning"])
    sheets = [("1 Read Me", readme), ("Summary", sum_df), ("2 Rep Strengths", rep_df), ("3 Rep by Quarter", rq_df), ("4 Rep x Project x Type", combo_df),
              ("5 Rep x Project", rp_df), ("6 Project x Type", pt_df), ("7 Team Leaders", tl_df), ("8 Left Out", lo_df),
              ("9 All Counted Rows", all_df)]
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, d in sheets: d.to_excel(xw, sheet_name=name, index=False)
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.comments import Comment
    CAPS = [
        ("Win Rate", "(Yes + No) ÷ (Yes + No + Terminated) in that period."),
        ("Total Incl. Left Out", "Total Sold + left-out units (Terminated Won, referrals; on project / type rows also CIC / KSA / Oman). Reference only."),
        ("Including Left Out", "Every unit in the sheet: counted + left out."),
        ("Excluding Left Out", "Units counted in the report (left-out rows removed)."),
        ("Left Out", "Units left out (Terminated Won, excluded team leads, Mr. Mahmoud's referrals)."),
        ("Total Sold", "All units sold (Yes + No + Terminated + Not set). 0.5 when shared with another rep."),
        ("Sold Q", "All units sold in that quarter."),
        ("Contracted: Yes", "Units whose Contracted column is Yes."), ("Contracted: No", "Units whose Contracted column is No."),
        ("Contracted Not Set", "Blank / unrecognised Contracted value — counted in Total Sold, not in the win rate."),
        ("Terminated There", "Terminated units in the weak-spot project."),
        ("Terminated", "Units whose Contracted column is Terminated — still sold, counted as lost for the rep."),
        ("Value", f"Total value of the units, {CURRENCY}."), ("#1 Share", "Share of the rep's units in their #1 project / type."),
        ("#1 Strength", "The project / unit type the rep sold most."), ("Type Share", "Share of the rep's units that are their main type."),
        ("Main Unit Type", "The unit type the rep sold most."), ("Best Project Win Rate", "Win rate in the best project."),
        ("Best Project", f"Highest win rate (min {MIN_DECIDED} decided units)."), ("Weak Spot", "Project with the most Terminated units."),
        ("Share Of Rep", "Units in this project / type ÷ all the rep's units."), ("Units (Share)", "1 = own unit · 0.5 = shared."),
        ("Reps Who Sold It", "Number of reps with at least one unit here."), ("Sheet Row", "Row number in the master sheet tab."),
        ("Period", "Overall = whole year · Q1 Jan–Mar · Q2 Apr–Jun · Q3 Jul–Sep · Q4 Oct–Dec.")]
    wb = openpyxl.load_workbook(path)
    for name, d in sheets:
        ws = wb[name]
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="0D6670")
            c.alignment = Alignment(wrap_text=True, vertical="center")
            cap = next((m for k, m in CAPS if k.lower() in str(c.value or "").lower()), None)
            if cap: c.comment = Comment(cap, "Rep strengths"); c.comment.width, c.comment.height = 260, 70
        for j, col in enumerate(d.columns, start=1):
            L = get_column_letter(j)
            width = max([len(str(col))] + [len(str(v)) for v in d[col].head(500)]) + 2
            ws.column_dimensions[L].width = min(max(width, 10), 110 if name == "1 Read Me" else 42)
            fmt = "0%" if ("Rate" in col or "Share" in col) else ("#,##0" if ("Value" in col or "Amount" in col) else None)
            if fmt:
                for row in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                    for c in row: c.number_format = fmt
        ws.freeze_panes = "A2"
        if ws.max_row > 1: ws.auto_filter.ref = ws.dimensions
    wb.save(path)


# ============================ 5. HTML ============================
def render(rows, left_out, team, meta):
    combo = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0, 0.0, 0.0])        # sold, yes, no, term, unset, amt
    idx = {"yes": 1, "no": 2, "term": 3, "unset": 4}
    for r in rows:
        c = combo[(r["rep"], r["proj"], r["type"], r["q"])]
        c[0] += r["w"]; c[idx[r["status"]]] += r["w"]; c[5] += r["amt"]
    used = {r["rep"] for r in rows} | {x["rep"] for x in left_out if not x["reason"].startswith("Team Lead")}
    lbl = {"yes": "Yes", "no": "No", "term": "Terminated", "unset": "Not set"}
    data = {
        "meta": meta,
        "reps": [{"r": rep, "sm": sm, "tl": tl} for rep, (sm, tl) in sorted(team.items()) if rep in used]
                + [{"r": rep, "sm": x["sm"], "tl": x["tl"]} for rep, x in sorted({x["rep"]: x for x in left_out
                   if x["rep"] not in team and not x["reason"].startswith("Team Lead")}.items())],
        "c": [[k[0], k[1], k[2], k[3]] + [round(v, 2) for v in vals[:5]] + [round(vals[5])] for k, vals in combo.items()],
        "opps": [{"n": f"{r['resv']} · {r['client']}".strip(" ·"), "a": r["chan"] + (" · " + r["src"] if r["src"] else ""),
                  "r": r["rep"], "p": r["proj"], "t": r["type"], "x": lbl[r["status"]], "q": r["q"], "cd": r["date"],
                  "amt": r["amt"], "u": r["w"], "row": r["row"]} for r in rows],
        "xo": [{"n": f"{x['resv']} · {x['client']}".strip(" ·"), "a": x["chan"], "r": x["rep"], "p": x["proj"], "t": x["type"],
                "x": x["reason"], "q": x["q"], "cd": x["date"], "amt": x["amt"], "u": x["w"], "row": x["row"],
                "tlx": x["reason"].startswith("Team Lead")} for x in left_out],
    }
    js = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    return ('<!doctype html><html lang="en"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">'
            '</head><body>' + TEMPLATE.replace("__DATA__", js) + '</body></html>')


TEMPLATE = r"""<title>Sales Rep Strengths</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Instrument+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap">
<style>
:root{
  --bg:#f5f6f4; --panel:#ffffff; --ink:#18201f; --muted:#5d6a68; --line:#dfe4e2; --soft:#eef1ef;
  --accent:#0d6670; --accent-soft:#dcefee; --ok:#2f8a4e; --ok-soft:#dff1e4; --no:#3b6fb6; --no-soft:#e1ebf7;
  --warn:#b7791f; --warn-soft:#f8ecd3; --bad:#c2412d; --bad-soft:#f8e0db; --unset:#8a9491; --unset-soft:#eceeed;
  --sans:"Instrument Sans",system-ui,-apple-system,"Segoe UI",sans-serif; --mono:"IBM Plex Mono",ui-monospace,Menlo,Consolas,monospace;
  --shadow:0 1px 2px rgba(20,30,30,.06),0 4px 16px rgba(20,30,30,.05);
}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){
  --bg:#111615; --panel:#1a2120; --ink:#e7ecea; --muted:#9aa8a5; --line:#2c3735; --soft:#222b29;
  --accent:#4fb3bb; --accent-soft:#1d3a3c; --ok:#5cc07c; --ok-soft:#1d3426; --no:#7aa7e6; --no-soft:#1c2a3d;
  --warn:#e0a84a; --warn-soft:#3a2f19; --bad:#ec7a65; --bad-soft:#3d221d; --unset:#8d9996; --unset-soft:#262e2c;
  --shadow:none; color-scheme:dark}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 var(--sans);padding:0 max(16px,env(safe-area-inset-right,0px)) 48px max(16px,env(safe-area-inset-left,0px))}
.wrap{max-width:1400px;margin:0 auto;display:flex;flex-direction:column;gap:18px}
.num{font-family:var(--mono);font-variant-numeric:tabular-nums}
header.top{padding-top:24px}
h1{font-size:26px;line-height:1.15;margin:0;font-weight:700;letter-spacing:-.01em}
.sub{color:var(--muted);font-size:13px;margin-top:4px}
.take{font-size:15px;margin:10px 0 0;max-width:90ch}
.take b{font-weight:600}
.filters{position:sticky;top:0;z-index:20;background:var(--bg);padding:10px 0;border-bottom:1px solid var(--line);display:flex;flex-wrap:wrap;gap:8px;align-items:flex-end}
.filters label,.seg{display:flex;flex-direction:column;gap:2px;font-size:12px;color:var(--muted)}
.seg .btns{display:flex;border:1px solid var(--line);border-radius:6px;overflow:hidden}
.seg .btns button{border:0;border-radius:0;border-right:1px solid var(--line)} .seg .btns button:last-child{border-right:0}
select,input[type=search]{font:13px var(--sans);color:var(--ink);background:var(--panel);border:1px solid var(--line);border-radius:6px;padding:6px 8px;max-width:220px}
button{font:500 13px var(--sans);cursor:pointer;border-radius:6px;border:1px solid var(--line);background:var(--panel);color:var(--ink);padding:6px 10px}
button.on{background:var(--accent);border-color:var(--accent);color:var(--panel)}
button:disabled{opacity:.45;cursor:not-allowed}
select:focus-visible,input:focus-visible,button:focus-visible,[tabindex]:focus-visible{outline:2px solid var(--accent);outline-offset:1px}
.chips{display:flex;flex-wrap:wrap;gap:6px;font-size:12px;color:var(--muted);align-items:center}
.chip{background:var(--accent-soft);color:var(--ink);border-radius:999px;padding:2px 10px}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:10px;box-shadow:var(--shadow);padding:16px;min-width:0}
.panel h2{font-size:16px;margin:0;font-weight:600}
.phead{display:flex;flex-wrap:wrap;gap:10px;justify-content:space-between;align-items:center}
.note{font-size:12.5px;color:var(--muted);margin:4px 0 0;max-width:100ch}
.qrow{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px}
.qt{text-align:left;display:flex;flex-direction:column;gap:2px;padding:12px 14px;border-radius:10px;background:var(--panel);box-shadow:var(--shadow)}
.qt .l{font-size:12px;color:var(--muted)} .qt .v{font:500 22px/1.15 var(--mono)} .qt .s{font-size:12px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.qt.on{background:var(--accent-soft);box-shadow:inset 0 0 0 1px var(--accent)} .qt.on .l,.qt.on .s{color:var(--ink)}
.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px}
.kpi.main{border-color:var(--accent)} .kpi.alt .v{color:var(--muted)}
.kpi{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:12px 14px;text-align:left;display:flex;flex-direction:column;gap:2px;box-shadow:var(--shadow)}
button.kpi:hover{border-color:var(--accent)}
.kpi .l{font-size:12px;color:var(--muted);display:flex;align-items:center;gap:6px} .kpi .v{font:500 24px/1.15 var(--mono)} .kpi .s{font-size:11.5px;color:var(--muted)}
.dot{width:9px;height:9px;border-radius:2px;display:inline-block}
.d-yes{background:var(--ok)} .d-no{background:var(--no)} .d-term{background:var(--bad)} .d-unset{background:var(--unset)}
.stack{display:flex;height:12px;border-radius:4px;overflow:hidden;background:var(--soft);margin-top:10px}
.stack span{display:block;height:100%}
.legend{display:flex;flex-wrap:wrap;gap:6px 18px;font-size:12px;color:var(--muted);margin-top:8px;align-items:center}
.legend b{color:var(--ink);font-weight:600}
.tbl-wrap{overflow:auto;margin-top:12px}
table{border-collapse:collapse;width:100%;font-size:13px}
th{font-size:12px;color:var(--muted);font-weight:600;text-align:left;padding:8px;border-bottom:1px solid var(--line);white-space:nowrap;background:var(--panel);position:sticky;top:0;z-index:1;vertical-align:bottom}
td{padding:7px 8px;border-bottom:1px solid var(--line);white-space:nowrap}
tfoot td{font-weight:600;border-top:2px solid var(--line);background:var(--soft);position:sticky;bottom:0}
.r{text-align:right}
th.sort{cursor:pointer} th.sort:hover{color:var(--ink)}
th .cap{display:block;font-weight:400;font-size:11px;line-height:1.25;color:var(--muted);white-space:normal;min-width:64px;max-width:150px;margin-top:2px}
th.r .cap{margin-left:auto}
th .arr{font-family:var(--mono);opacity:.45;margin-left:3px} th.act .arr{opacity:1;color:var(--accent)}
th.qc,td.qc{background:var(--soft)} th.qc.cur,td.qc.cur{background:var(--accent-soft);font-weight:600}
td.c-yes{color:var(--ok)} td.c-no{color:var(--no)} td.c-term{color:var(--bad)} td.c-unset{color:var(--unset)}
.click tbody tr{cursor:pointer} .click tbody tr:hover td{background:var(--soft)}
tr.pick td{font-weight:600} tr.pick td:first-child{box-shadow:inset 3px 0 0 var(--accent)}
.heat th.c{text-align:center;white-space:normal;min-width:74px;max-width:110px;line-height:1.25}
.heat th.c small{display:block;font-weight:400}
.heat th.c .cap{margin:2px auto 0}
.heat td.cell{text-align:center;font-family:var(--mono);font-size:12px;cursor:pointer;border:2px solid var(--panel);border-radius:4px;min-width:64px}
.heat td.cell:hover{outline:2px solid var(--accent)}
.heat td.nm{cursor:pointer;position:sticky;left:0;background:var(--panel);z-index:1} .heat td.nm:hover{color:var(--accent);text-decoration:underline}
.heat th:first-child{position:sticky;left:0;z-index:2}
.sortbar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-top:10px;font-size:12px;color:var(--muted)}
.grid3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:18px;margin-top:14px}
.barrow{display:grid;grid-template-columns:minmax(0,42%) 1fr;gap:10px;align-items:center;padding:4px 0;font-size:13px;cursor:pointer;border-radius:6px}
.barrow:hover{background:var(--soft)}
.barrow .lb{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:right}
.barrow .tr{position:relative;display:flex;align-items:center;gap:8px}
.barrow .b{height:14px;border-radius:3px;background:var(--accent);min-width:2px}
.barrow .ref{position:absolute;top:-4px;bottom:-4px;border-left:1px dashed var(--muted)}
.prof-head{display:flex;flex-wrap:wrap;justify-content:space-between;gap:10px;align-items:baseline}
.prof-name{font-size:22px;font-weight:700}
.tag{display:inline-block;font-size:11.5px;padding:1px 8px;border-radius:999px}
.t-Yes{background:var(--ok-soft);color:var(--ok)} .t-No{background:var(--no-soft);color:var(--no)} .t-Terminated{background:var(--bad-soft);color:var(--bad)}
.t-Not{background:var(--unset-soft);color:var(--unset)} .t-ex{background:var(--warn-soft);color:var(--warn)}
.drawer-bg{position:fixed;inset:0;background:rgba(10,16,16,.35);z-index:40}
.drawer{position:fixed;top:0;right:0;bottom:0;width:min(1080px,100%);background:var(--panel);z-index:41;border-left:1px solid var(--line);display:flex;flex-direction:column;padding-top:env(safe-area-inset-top,0px)}
.drawer header{display:flex;justify-content:space-between;gap:12px;padding:16px 18px;border-bottom:1px solid var(--line)}
.drawer h3{margin:0;font-size:18px}
.drawer .body{overflow:auto;padding:0 18px 18px;flex:1}
.foot{font-size:12px;color:var(--muted);line-height:1.6}
.empty{padding:24px;text-align:center;color:var(--muted)}
@media (max-width:1100px){.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media (max-width:1000px){.qrow{grid-template-columns:repeat(3,minmax(0,1fr))}.grid3{grid-template-columns:1fr}}
@media (max-width:560px){.kpis,.qrow{grid-template-columns:repeat(2,minmax(0,1fr))}.filters{position:static}select,input[type=search]{max-width:none;width:100%}.filters label{flex:1 1 45%}}
</style>

<div class="wrap">
  <header class="top">
    <h1>Sales rep strengths by project and unit type</h1>
    <div class="sub" id="sub"></div>
    <p class="take" id="take"></p>
  </header>

  <div class="filters">
    <div class="seg"><span>Period</span><div class="btns" id="f-period"></div></div>
    <label>Sales manager<select id="f-sm"></select></label>
    <label>Team leader<select id="f-tl"></select></label>
    <label>Sales rep<select id="f-rep"></select></label>
    <button id="reset" type="button">Clear filters</button>
    <div class="chips" id="chips"></div>
  </div>

  <div><p class="note" style="margin:0 0 6px">Each tile: <b>units sold excluding left-out deals</b> · the same <b>including left out</b> · <b>win rate</b> = (Yes + No) ÷ (Yes + No + Terminated) · the <b>project / unit type sold most</b>. Click a tile to switch the whole page to that period.</p>
  <section class="qrow" id="qrow" aria-label="Overall and quarters"></section></div>
  <div class="panel" id="warn" hidden style="border-color:var(--warn);background:var(--warn-soft)"></div>

  <section>
    <div class="kpis" id="kpis"></div>
    <div class="stack" id="stack" title="Total sold split by Contracted"></div>
    <div class="legend" id="stack-legend"></div>
  </section>

  <section class="panel" id="profile" hidden></section>

  <section class="panel">
    <div class="phead"><h2 id="heat-h">Where each rep's sales come from</h2>
      <div class="seg"><span>Columns</span><div class="btns"><button type="button" data-hc="all">Every project / type</button><button type="button" data-hc="10">Top 10</button></div></div></div>
    <div class="legend"><span><b>Sold</b> = rep's units in the period</span><span><b>%</b> in a cell = share of that rep's units in that project / type</span><span>Darker = bigger share</span><span>Click any header to sort · click a cell for the reservations</span></div>
    <div class="tbl-wrap" style="max-height:640px"><table class="heat" id="heat"></table></div>
  </section>

  <section class="panel">
    <h2 id="pt-h">Every project and unit type</h2>
    <p class="note">All project / unit type combinations in this selection with their Contracted breakdown. The total row matches both totals above. Click a row to see the reservations.</p>
    <div class="tbl-wrap" style="max-height:560px"><table class="click" id="pt"></table></div>
  </section>

  <section class="panel">
    <h2 id="tbl-h">Strengths by rep</h2>
    <p class="note" id="tbl-note"></p>
    <div class="sortbar"><input type="search" id="q" placeholder="Search reps, projects or types">
      <label>Sort by <select id="sortby"></select></label><button type="button" id="sortdir"></button></div>
    <div class="tbl-wrap" style="max-height:640px"><table class="click" id="tbl"></table></div>
  </section>

  <p class="foot" id="foot"></p>
</div>
<div id="drawer-root"></div>

<script>
const D = __DATA__, META = D.meta, CUR = META.currency, QS = ["Q1","Q2","Q3","Q4"];
const C = D.c.map(([r,p,t,q,w,y,n,te,u,amt]) => ({r,p,t,q,w,y,n,te,u,amt}));
const REP = new Map(D.reps.map(x => [x.r, x]));
const F = {sm:"", tl:"", rep:"", period:"all"};
let SORT = {k:"w", d:-1}, PSORT = {k:"w", d:-1}, HSORT = {k:"w", d:-1}, HCOLS = "all", Q = "";
const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fmtN = n => (n || 0).toLocaleString("en-US",{maximumFractionDigits:1});
const pct = v => v == null ? "—" : Math.round(v * 100) + "%";
function fmtM(v){ if(!v) return "0"; const a=Math.abs(v);
  if(a>=1e9) return (v/1e9).toFixed(2)+"B"; if(a>=1e6) return (v/1e6).toFixed(a>=1e8?0:1)+"M";
  if(a>=1e3) return (v/1e3).toFixed(0)+"K"; return String(Math.round(v)); }
const fmtD = s => /^\d{4}-\d{2}-\d{2}$/.test(s||"") ? new Date(s+"T00:00:00").toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}) : (s||"—");
const uniq = a => [...new Set(a)].sort((x,y)=>x.localeCompare(y));
const PL = p => p==="all" ? "Overall this year" : p+" ("+META.qlabel[p]+")";
const inP = (q,p) => (p ?? F.period)==="all" || q===(p ?? F.period);
const winr = o => (o.y+o.n+o.te) ? (o.y+o.n)/(o.y+o.n+o.te) : null;
const acc = () => ({w:0,y:0,n:0,te:0,u:0,amt:0});
function addTo(a,c){ a.w+=c.w; a.y+=c.y; a.n+=c.n; a.te+=c.te; a.u+=c.u; a.amt+=c.amt; return a; }

/* ---------- scope ---------- */
function teamReps(){ return D.reps.filter(x => (!F.sm || x.sm===F.sm) && (!F.tl || x.tl===F.tl)); }
function scopeNames(){ const t=teamReps(); return new Set((F.rep ? t.filter(x=>x.r===F.rep) : t).map(x=>x.r)); }

/* ---------- aggregation ---------- */
function stats(names, p){
  const reps=new Map();
  for(const c of C){
    if(!names.has(c.r) || !inP(c.q,p)) continue;
    let s=reps.get(c.r); if(!s){ s={r:c.r,...acc(),cm:new Map(),pm:new Map(),tm:new Map()}; reps.set(c.r,s); }
    addTo(s,c);
    const ck=c.p+"|"+c.t; s.cm.set(ck, addTo(s.cm.get(ck)||{p:c.p,t:c.t,...acc()}, c));
    s.pm.set(c.p, addTo(s.pm.get(c.p)||{p:c.p,...acc()}, c));
    s.tm.set(c.t,(s.tm.get(c.t)||0)+c.w);
  }
  for(const s of reps.values()){
    const info=REP.get(s.r)||{}; s.tl=info.tl; s.sm=info.sm; s.wr=winr(s);
    const top=[...s.cm.values()].filter(x=>x.w>0).sort((a,b)=>b.w-a.w||b.amt-a.amt)[0];
    s.tc = top ? top.p+" / "+top.t : null; s.ts = top && s.w ? top.w/s.w : null;
    const tt=[...s.tm.entries()].sort((a,b)=>b[1]-a[1])[0]; s.tt = tt?tt[0]:null;
    const projs=[...s.pm.values()].filter(x=>x.p!=="(no project)");
    const best=projs.filter(x=>x.y+x.n+x.te>=META.min_decided).sort((a,b)=>winr(b)-winr(a)||b.w-a.w)[0];
    s.bp=best?best.p:null; s.bpr=best?winr(best):null;
    const weak=projs.filter(x=>x.te>0).sort((a,b)=>b.te-a.te)[0]; s.wp=weak?weak.p:null; s.wpt=weak?weak.te:null;
  }
  return reps;
}
// left-out rows in scope: company-wide = all of them (incl. CIC / KSA / Oman); filtered = those reps' own left-out rows
function loList(names, p){ const all=!F.sm&&!F.tl&&!F.rep; return D.xo.filter(o=>inP(o.q,p)&&(all||(!o.tlx&&names.has(o.r)))); }
const units = l => l.reduce((a,o)=>a+(o.u||0),0);
function totals(names, p){ const a=acc(); for(const c of C) if(names.has(c.r)&&inP(c.q,p)) addTo(a,c); a.wr=winr(a); return a; }
function combos(names, p){
  const m=new Map(); for(const c of C){ if(!names.has(c.r)||!inP(c.q,p)) continue; const k=c.p+"|"+c.t;
    const x=m.get(k)||{k,p:c.p,t:c.t,...acc(),reps:new Set()}; addTo(x,c); x.reps.add(c.r); m.set(k,x); }
  for(const x of m.values()){ x.wr=winr(x); x.nr=x.reps.size; }
  return [...m.values()];
}
const topCombo = (names,p) => combos(names,p).sort((a,b)=>b.w-a.w)[0];

/* ---------- generic sortable table ---------- */
function headCell(c, st, attr){
  const [k,l,cap,n,f,cls]=c;
  return '<th class="sort'+(n?' r':'')+(cls?' '+cls:'')+(st.k===k?' act':'')+'" '+attr+'="'+esc(k)+'" title="Click to sort">'+l+
    '<span class="arr">'+(st.k===k?(st.d<0?'↓':'↑'):'↕')+'</span>'+(cap?'<span class="cap">'+esc(cap)+'</span>':'')+'</th>';
}
function cellVal(c, v){ const [k,l,cap,n,f]=c; return v==null?'—':(f?f(v):(n?fmtN(v):v)); }
function sortRows(list, st){
  return list.sort((a,b)=>{ const x=a[st.k], y=b[st.k]; if(x==null&&y==null) return 0; if(x==null) return 1; if(y==null) return -1;
    return st.d*(typeof x==="string"?x.localeCompare(y):x-y) || (a.r||a.p||"").localeCompare(b.r||b.p||""); });
}
const TXT=new Set(["r","tl","sm","tc","tt","bp","wp","p","t"]);
const STATUS_CLS={y:"c-yes",n:"c-no",te:"c-term",u:"c-unset"};

/* ---------- filters ---------- */
function fill(id, k, vals, all){
  const el=$(id); if(F[k] && !vals.includes(F[k])) F[k]="";
  el.innerHTML='<option value="">'+all+'</option>'+vals.map(v=>'<option'+(v===F[k]?' selected':'')+' value="'+esc(v)+'">'+esc(v)+'</option>').join("");
}
function renderFilters(){
  fill("f-sm","sm",uniq(D.reps.map(x=>x.sm)),"All sales managers");
  fill("f-tl","tl",uniq(D.reps.filter(x=>!F.sm||x.sm===F.sm).map(x=>x.tl)),"All team leaders");
  fill("f-rep","rep",uniq(teamReps().map(x=>x.r)),"All reps");
  $("f-period").innerHTML=["all",...QS].map(p=>'<button type="button" data-p="'+p+'" class="'+(F.period===p?'on':'')+'"'+
    (p!=="all"&&!META.quarters.includes(p)?' disabled':'')+'>'+(p==="all"?"Overall":p)+'</button>').join("");
  const ch=[PL(F.period)]; if(F.sm) ch.push(F.sm); if(F.tl) ch.push(F.tl); if(F.rep) ch.push(F.rep); if(Q) ch.push('search “'+Q+'”');
  $("chips").innerHTML='Showing: '+ch.map(c=>'<span class="chip">'+esc(c)+'</span>').join("");
  document.querySelectorAll("[data-hc]").forEach(b=>b.classList.toggle("on", b.dataset.hc===HCOLS));
}

/* ---------- quarter strip + KPIs ---------- */
function renderTop(){
  const names=scopeNames();
  $("qrow").innerHTML=["all",...QS].map(p=>{
    const t=totals(names,p), tc=topCombo(names,p), dis=p!=="all"&&!META.quarters.includes(p);
    return '<button type="button" class="qt'+(F.period===p?' on':'')+'" data-p="'+p+'"'+(dis?' disabled':'')+'><span class="l">'+(p==="all"?"Overall this year":p+" · "+META.qlabel[p])+
      '</span><span class="v">'+fmtN(t.w)+' sold</span><span class="s">'+fmtN(t.w+units(loList(names,p)))+' incl. left out · win '+pct(t.wr)+' · '+(tc?esc(tc.p+" / "+tc.t):"—")+'</span></button>';
  }).join("");
  const t=totals(names), st=stats(names), xs=loList(names), xu=units(xs);
  const scope=F.rep||F.tl&&(F.tl+"'s team")||F.sm&&(F.sm+"'s teams")||"Company-wide";
  const tc=topCombo(names);
  const best=[...st.values()].filter(s=>s.y+s.n+s.te>=META.min_closer&&s.wr!=null).sort((a,b)=>b.wr-a.wr||b.w-a.w)[0];
  $("take").innerHTML = esc(scope)+", "+esc(PL(F.period).toLowerCase())+": <b>"+fmtN(t.w)+"</b> units sold, win rate <b>"+pct(t.wr)+"</b>. "+
    (tc?"The most-sold combination is <b>"+esc(tc.p+" / "+tc.t)+"</b> with <b>"+fmtN(tc.w)+"</b> units.":"")+
    (!F.rep&&best?" Among reps with "+META.min_closer+"+ decided units, the highest win rate is <b>"+esc(best.r)+"</b> at <b>"+pct(best.wr)+"</b>.":"");
  const share=v=>t.w?" · "+pct(v/t.w)+" of sold":"";
  const tile=(lab,v,s,x,dot,cls)=>(x?'<button type="button" class="kpi'+(cls?' '+cls:'')+'" data-x="'+x+'">':'<div class="kpi">')+'<span class="l">'+(dot?'<span class="dot '+dot+'"></span>':'')+lab+'</span><span class="v">'+v+'</span><span class="s">'+s+'</span>'+(x?'</button>':'</div>');
  const why=Object.entries(xs.reduce((m,o)=>(m[o.x]=(m[o.x]||0)+o.u,m),{})).map(([k,v])=>fmtN(v)+" "+k).join(" · ");
  $("kpis").innerHTML =
    tile("Total sold — excluding left out", fmtN(t.w), "the units this report measures (½ when shared) · click for the list", "all", null, "main") +
    tile("Total sold — including left out", fmtN(t.w+xu), "= "+fmtN(t.w)+" counted + "+fmtN(xu)+" left out"+(!F.sm&&!F.tl&&!F.rep?" (incl. CIC / KSA / Oman)":"")+" · click for the list", "incl", null, "alt") +
    tile("Contracted: Yes", fmtN(t.y), "Contracted = Yes"+share(t.y), "Yes", "d-yes") +
    tile("Contracted: No", fmtN(t.n), "Contracted = No"+share(t.n), "No", "d-no") +
    tile("Terminated", fmtN(t.te), "still sold, counted as lost"+share(t.te), "Terminated", "d-term") +
    tile("Win rate", pct(t.wr), "(Yes + No) ÷ (Yes + No + Terminated)") +
    tile("Value sold", fmtM(t.amt), "total value of all units sold, "+CUR+" (B = billion, M = million)") +
    tile("Left out", fmtN(xu), (why||"nothing")+" · not in strengths or win rate", "ex");
  const seg=(v,cls)=>t.w&&v?'<span class="'+cls+'" style="width:'+(100*v/t.w)+'%"></span>':'';
  $("stack").innerHTML=seg(t.y,"d-yes")+seg(t.n,"d-no")+seg(t.te,"d-term")+seg(t.u,"d-unset");
  $("stack").title="Total sold excluding left out, split by Contracted";
  $("stack-legend").innerHTML='<span><span class="dot d-yes"></span> Yes '+fmtN(t.y)+'</span><span><span class="dot d-no"></span> No '+fmtN(t.n)+'</span><span><span class="dot d-term"></span> Terminated '+fmtN(t.te)+'</span>'+
    (t.u?'<span><span class="dot d-unset"></span> Contracted not set '+fmtN(t.u)+' (in total, not in win rate)</span>':'')+'<span>= '+fmtN(t.w)+' total sold</span>';
}

/* ---------- profile ---------- */
function bars(items, ref){
  const max=Math.max(1e-9,...items.map(d=>d.v));
  return items.map(d=>'<div class="barrow" tabindex="0" data-dp="'+esc(d.p||"")+'" data-dt="'+esc(d.t||"")+'"'+(d.q?' data-p="'+d.q+'"':'')+'><span class="lb" title="'+esc(d.label)+'">'+esc(d.label)+'</span><span class="tr">'+
    '<span class="b" style="width:'+(68*d.v/(ref!==undefined?1:max))+'%"></span><span class="num">'+esc(d.text)+'</span>'+
    (ref!=null?'<span class="ref" style="left:'+(68*ref)+'%"></span>':'')+'</span></div>').join("") || '<div class="empty">Nothing in this period.</div>';
}
function renderProfile(){
  const el=$("profile"); if(!F.rep){ el.hidden=true; return; } el.hidden=false;
  const info=REP.get(F.rep)||{}, s=stats(new Set([F.rep])).get(F.rep), one=new Set([F.rep]);
  const head='<div class="prof-head"><div><div class="prof-name">'+esc(F.rep)+'</div><div class="sub">'+esc(info.tl)+' · '+esc(info.sm)+' · '+esc(PL(F.period))+'</div></div><button type="button" id="clear-rep">Back to all reps</button></div>';
  const qb=QS.filter(q=>META.quarters.includes(q)).map(q=>{ const t=totals(one,q); return {q,label:q,v:t.w,text:fmtN(t.w)+" sold · win "+pct(t.wr)}; });
  if(!s){ el.innerHTML=head+'<p class="take">No units in '+esc(PL(F.period).toLowerCase())+'.</p><div style="max-width:420px">'+bars(qb)+'</div>'; return; }
  const mix=[...s.cm.values()].filter(c=>c.w>0).sort((a,b)=>b.w-a.w).map(c=>({p:c.p,t:c.t,label:c.p+" / "+c.t,v:c.w,text:fmtN(c.w)+" · "+pct(c.w/s.w)+" · win "+pct(winr(c))}));
  const rate=[...s.pm.values()].filter(x=>x.y+x.n+x.te>0).map(x=>({p:x.p,label:x.p,v:winr(x),text:pct(winr(x))+" ("+fmtN(x.y+x.n)+" of "+fmtN(x.y+x.n+x.te)+")"})).sort((a,b)=>b.v-a.v);
  el.innerHTML=head+
    '<p class="take"><b>'+fmtN(s.w)+'</b> sold — Yes <b>'+fmtN(s.y)+'</b>, No <b>'+fmtN(s.n)+'</b>, Terminated <b>'+fmtN(s.te)+'</b>'+(s.u?', not set '+fmtN(s.u):'')+
    ' — win rate <b>'+pct(s.wr)+'</b>. Strongest in <b>'+esc(s.tc||"—")+'</b> ('+pct(s.ts)+' of their units), main unit type <b>'+esc(s.tt||"—")+'</b>. '+
    (s.bp?'Best win rate in <b>'+esc(s.bp)+'</b> ('+pct(s.bpr)+'). ':'')+(s.wp?'Most terminated in <b>'+esc(s.wp)+'</b> ('+fmtN(s.wpt)+').':'')+'</p>'+
    '<div class="grid3"><div><h2>Sales mix</h2><p class="note">Every project / unit type: units · share of their units · win rate.</p>'+bars(mix)+'</div>'+
    '<div><h2>Win rate by project</h2><p class="note">(Yes + No) ÷ (Yes + No + Terminated). Dashed line = their overall ('+pct(s.wr)+').</p>'+bars(rate,s.wr)+'</div>'+
    '<div><h2>By quarter</h2><p class="note">Units sold and win rate per quarter. Click a quarter to switch.</p>'+bars(qb)+'</div></div>';
}

/* ---------- heatmap ---------- */
function renderHeat(){
  const names=scopeNames(), st=stats(names), s=[...st.values()].filter(x=>x.w>0);
  let cols=combos(names).filter(x=>x.w>0).sort((a,b)=>b.w-a.w);
  if(HCOLS!=="all") cols=cols.slice(0,10);
  const share=(r,k)=>{ const c=r.cm.get(k); return c&&c.w?c.w/r.w:0; };
  if(HSORT.k!=="r"&&HSORT.k!=="w"&&!cols.some(c=>c.k===HSORT.k)) HSORT={k:"w",d:-1};
  s.sort((a,b)=> HSORT.k==="r" ? HSORT.d*a.r.localeCompare(b.r) : (HSORT.k==="w" ? HSORT.d*(a.w-b.w) : HSORT.d*(share(a,HSORT.k)-share(b,HSORT.k))) || b.w-a.w || a.r.localeCompare(b.r));
  $("heat-h").textContent="Where each rep's sales come from — "+PL(F.period);
  if(!s.length){ $("heat").innerHTML='<tr><td class="empty">No units sold in this selection.</td></tr>'; return; }
  const th=(k,lab,cap,cls)=>'<th class="sort'+(cls?' '+cls:'')+(HSORT.k===k?' act':'')+'" data-hk="'+esc(k)+'" title="Click to sort">'+lab+'<span class="arr">'+(HSORT.k===k?(HSORT.d<0?'↓':'↑'):'↕')+'</span><span class="cap">'+cap+'</span></th>';
  let h='<thead><tr>'+th("r","Rep","sales rep")+th("w","Sold","units in period","r")+
    cols.map(c=>th(c.k,'<small>'+esc(c.p)+'</small>'+esc(c.t),fmtN(c.w)+" units · % of rep","c")).join("")+'</tr></thead><tbody>';
  for(const r of s){
    h+='<tr'+(r.r===F.rep?' class="pick"':'')+'><td class="nm" tabindex="0" data-rep="'+esc(r.r)+'">'+esc(r.r)+'</td><td class="r num">'+fmtN(r.w)+'</td>';
    for(const col of cols){
      const c=r.cm.get(col.k), sh=c&&c.w?c.w/r.w:0;
      const bg=sh?'background:color-mix(in srgb, var(--accent) '+Math.round(12+88*sh)+'%, transparent);color:'+(sh>.45?'var(--panel)':'var(--ink)'):'';
      h+='<td class="cell" tabindex="0" style="'+bg+'" data-cr="'+esc(r.r)+'" data-cp="'+esc(col.p)+'" data-ct="'+esc(col.t)+'" title="'+
        (c?esc(r.r+" · "+col.p+" / "+col.t+"\nSold "+fmtN(c.w)+" ("+pct(sh)+" of their units)\nYes "+fmtN(c.y)+" · No "+fmtN(c.n)+" · Terminated "+fmtN(c.te)+"\nWin rate "+pct(winr(c))):"No units")+'">'+(sh>=.005?pct(sh):"")+'</td>';
    }
    h+='</tr>';
  }
  $("heat").innerHTML=h+'</tbody>';
}

/* ---------- every project × type ---------- */
const PCOLS=[["p","Project","project"],["t","Unit type","unit type"],["w","Total sold","excluding left out",1],["wi","Incl. left out","+ left-out units",1],["y","Yes","contracted",1,null,"",""],["n","No","not contracted yet",1],
  ["te","Terminated","counted as lost",1],["u","Not set","no Contracted value",1],["wr","Win rate","(Yes+No) ÷ (Yes+No+Term.)",1,pct],
  ["amt","Value sold",CUR,1,fmtM],["nr","Reps","reps who sold it",1]];
function renderPT(){
  const names=scopeNames(), t=totals(names), lo=loList(names), base=combos(names), by=new Map(base.map(x=>[x.k,x]));
  for(const o of lo){ const k=o.p+"|"+o.t; if(!by.has(k)){ const x={k,p:o.p,t:o.t,...acc(),wr:null,nr:0}; by.set(k,x); base.push(x); } by.get(k).lo=(by.get(k).lo||0)+o.u; }
  for(const x of base) x.wi=x.w+(x.lo||0);
  const list=sortRows(base,PSORT);
  $("pt-h").textContent="Every project and unit type — "+PL(F.period)+" ("+list.length+")";
  const cls=k=>STATUS_CLS[k]||"";
  $("pt").innerHTML='<thead><tr>'+PCOLS.map(c=>headCell(c,PSORT,"data-pk")).join("")+'</tr></thead><tbody>'+
    (list.map(x=>'<tr data-pp="'+esc(x.p)+'" data-pt="'+esc(x.t)+'">'+PCOLS.map(c=>'<td class="'+(c[3]?'r num ':'')+cls(c[0])+'">'+esc(cellVal(c,x[c[0]]))+'</td>').join("")+'</tr>').join("")||'<tr><td class="empty" colspan="11">No units.</td></tr>')+
    '</tbody><tfoot><tr><td>Total</td><td></td><td class="r num">'+fmtN(t.w)+'</td><td class="r num">'+fmtN(t.w+units(lo))+'</td><td class="r num c-yes">'+fmtN(t.y)+'</td><td class="r num c-no">'+fmtN(t.n)+'</td><td class="r num c-term">'+fmtN(t.te)+'</td><td class="r num c-unset">'+fmtN(t.u)+'</td><td class="r num">'+pct(t.wr)+'</td><td class="r num">'+fmtM(t.amt)+'</td><td class="r num">'+fmtN(names.size)+'</td></tr></tfoot>';
}

/* ---------- strengths table ---------- */
const TCOLS=[
  ["r","Rep","sales rep"],["tl","Team leader","latest team leader on their reservations"],["sm","Sales manager","manager of that team leader"],
  ["w","Total sold","excluding left out (½ when shared)",1],["wi","Incl. left out","+ their Terminated Won / referral units",1],
  ["Q1","Q1","sold Jan–Mar",1,null,"qc"],["Q2","Q2","sold Apr–Jun",1,null,"qc"],["Q3","Q3","sold Jul–Sep",1,null,"qc"],["Q4","Q4","sold Oct–Dec",1,null,"qc"],
  ["y","Yes","contracted",1],["n","No","not contracted yet",1],["te","Terminated","still sold, counted as lost",1],["u","Not set","no Contracted value",1],
  ["wr","Win rate","(Yes+No) ÷ (Yes+No+Terminated)",1,pct],["amt","Value sold","total value, "+CUR,1,fmtM],
  ["tc","#1 strength","project / type they sell most"],["ts","Share of sales","% of their units in their #1",1,pct],["tt","Main unit type","type they sell most"],
  ["bp","Best project","highest win rate (min "+META.min_decided+" decided)"],["bpr","Its win rate","win rate there",1,pct],
  ["wp","Weak spot","most terminated units"],["wpt","Terminated there","terminated units there",1]];
function renderTable(){
  $("tbl-h").textContent="Strengths by rep — "+PL(F.period);
  $("tbl-note").textContent="Everything follows the period you picked, except the Q1–Q4 columns, which always show each quarter so you can see the trend. Click any column header (or use “Sort by”) to sort; click again to reverse. Click a row to open the rep's profile.";
  $("sortby").innerHTML=TCOLS.map(([k,l])=>'<option value="'+k+'"'+(SORT.k===k?' selected':'')+'>'+esc(l)+'</option>').join("");
  $("sortdir").textContent = SORT.d<0 ? (TXT.has(SORT.k)?"Z → A":"Highest first") : (TXT.has(SORT.k)?"A → Z":"Lowest first");
  const names=scopeNames(), st=stats(names);
  const qn={}; for(const q of QS){ qn[q]=new Map(); for(const c of C) if(names.has(c.r)&&c.q===q) qn[q].set(c.r,(qn[q].get(c.r)||0)+c.w); }
  const q=Q.toLowerCase();
  const lor=new Map(); for(const o of D.xo) if(!o.tlx&&names.has(o.r)&&inP(o.q)) lor.set(o.r,(lor.get(o.r)||0)+o.u);
  for(const [r,u] of lor) if(!st.has(r)){ const info=REP.get(r)||{}; st.set(r,{r,tl:info.tl,sm:info.sm,...acc(),wr:null,tc:null,ts:null,tt:null,bp:null,bpr:null,wp:null,wpt:null}); }
  const list=sortRows([...st.values()].map(s=>{ const o={...s}; for(const k of QS) o[k]=qn[k].get(s.r)||0; o.wi=s.w+(lor.get(s.r)||0); return o; })
    .filter(r=>!q||[r.r,r.tl,r.sm,r.tc,r.tt,r.bp,r.wp].some(x=>x&&x.toLowerCase().includes(q))), SORT);
  const sum=k=>list.reduce((a,r)=>a+(r[k]||0),0), y=sum("y"), n=sum("n"), te=sum("te");
  const tot={r:"Total ("+list.length+" reps)",w:sum("w"),wi:sum("wi"),Q1:sum("Q1"),Q2:sum("Q2"),Q3:sum("Q3"),Q4:sum("Q4"),y,n,te,u:sum("u"),wr:(y+n+te)?(y+n)/(y+n+te):null,amt:sum("amt")};
  const qcls=c=>c[5]==="qc"?' qc'+(F.period===c[0]?' cur':''):'';
  $("tbl").innerHTML='<thead><tr>'+TCOLS.map(c=>headCell([c[0],c[1],c[2],c[3],c[4],qcls(c).trim()],SORT,"data-k")).join("")+'</tr></thead><tbody>'+
    (list.map(r=>'<tr data-rep="'+esc(r.r)+'"'+(r.r===F.rep?' class="pick"':'')+'>'+TCOLS.map(c=>'<td class="'+(c[3]?'r num':'')+qcls(c)+' '+(STATUS_CLS[c[0]]||"")+'">'+esc(cellVal(c,r[c[0]]))+'</td>').join("")+'</tr>').join("") ||
     '<tr><td class="empty" colspan="'+TCOLS.length+'">No reps with units in this selection.</td></tr>')+'</tbody>'+
    '<tfoot><tr>'+TCOLS.map(c=>'<td class="'+(c[3]?'r num':'')+qcls(c)+'">'+(c[0] in tot?esc(cellVal(c,tot[c[0]])):'')+'</td>').join("")+'</tr></tfoot>';
}

/* ---------- drawer ---------- */
let DR=null;
function openDeals(title, list){
  const ord={Yes:0,No:1,Terminated:2,"Not set":3};
  list=[...list].sort((a,b)=>(ord[a.x]??9)-(ord[b.x]??9)||String(b.cd).localeCompare(a.cd));
  DR={title}; document.body.style.overflow="hidden";
  const units=list.reduce((a,o)=>a+(o.u||1),0);
  const tag=x=>'<span class="tag '+(x in ord?'t-'+x.split(" ")[0]:'t-ex')+'">'+esc(x)+'</span>';
  $("drawer-root").innerHTML='<div class="drawer-bg" id="dbg"></div><aside class="drawer" role="dialog" aria-modal="true" aria-label="Reservations"><header><div><h3>'+esc(title)+'</h3><div class="sub">'+
    esc(PL(F.period))+' · '+fmtN(list.length)+' rows · '+fmtN(units)+' units</div></div><button type="button" id="dclose">Close</button></header><div class="body"><div class="tbl-wrap" style="margin-top:0"><table><thead><tr>'+
    '<th class="r">Sheet row<span class="cap">row in the master tab</span></th><th>Quarter</th><th>Date</th><th>Reservation · client</th><th>Rep</th><th>Project</th><th>Unit type</th><th>Contracted</th><th class="r">Units<span class="cap">0.5 = shared</span></th><th class="r">Value<span class="cap">'+CUR+'</span></th></tr></thead><tbody>'+
    (list.map(o=>'<tr><td class="r num">'+esc(o.row)+'</td><td>'+esc(o.q)+'</td><td>'+fmtD(o.cd)+'</td><td>'+esc(o.n)+(o.a?'<div class="note" style="margin:0">'+esc(o.a)+'</div>':'')+'</td><td>'+esc(o.r)+'</td><td>'+esc(o.p)+'</td><td>'+esc(o.t)+
      '</td><td>'+tag(o.x)+'</td><td class="r num">'+fmtN(o.u)+'</td><td class="r num">'+(o.amt?fmtM(o.amt):'—')+'</td></tr>').join("")||'<tr><td class="empty" colspan="10">Nothing here.</td></tr>')+
    '</tbody></table></div></div></aside>';
  $("dclose").focus();
}
function closeDrawer(){ DR=null; $("drawer-root").innerHTML=""; document.body.style.overflow=""; }

/* ---------- events ---------- */
function render(){ renderFilters(); renderTop(); renderProfile(); renderHeat(); renderPT(); renderTable(); }
function setPeriod(p){ if(p!=="all"&&!META.quarters.includes(p)) return; F.period=p; render(); }
function pickRep(r){
  if(F.rep===r){ F.rep=""; return render(); }
  const x=REP.get(r); F.rep=r; if(x){ F.tl=x.tl; F.sm=x.sm; }
  render(); $("profile").scrollIntoView({behavior:"smooth",block:"start"});
}
function flip(st,k){ return st.k===k ? {k,d:-st.d} : {k,d:TXT.has(k)?1:-1}; }
document.addEventListener("click", e=>{
  const t=e.target.closest("button,td.cell,td.nm,th.sort,tr[data-rep],tr[data-pp],.barrow");
  if(!t||t.disabled) return;
  if(t.id==="dbg"||t.id==="dclose") return closeDrawer();
  if(t.id==="reset"){ F.sm=F.tl=F.rep=""; Q=""; $("q").value=""; return render(); }
  if(t.id==="clear-rep"){ F.rep=""; return render(); }
  if(t.dataset.hc){ HCOLS=t.dataset.hc; renderFilters(); return renderHeat(); }
  if(t.dataset.p) return setPeriod(t.dataset.p);
  const names=scopeNames(), who=F.rep||F.tl||F.sm||"All reps", mine=D.opps.filter(o=>names.has(o.r)&&inP(o.q));
  if(t.dataset.x==="ex") return openDeals(who+" — left out", loList(names));
  if(t.dataset.x==="incl") return openDeals(who+" — all units including left out", mine.concat(loList(names)));
  if(t.dataset.x==="all") return openDeals(who+" — all units sold", mine);
  if(t.dataset.x) return openDeals(who+" — Contracted: "+t.dataset.x, mine.filter(o=>o.x===t.dataset.x));
  if(t.matches("td.cell")){ const r=t.dataset.cr,p=t.dataset.cp,ty=t.dataset.ct; return openDeals(r+" — "+p+" / "+ty, mine.filter(o=>o.r===r&&o.p===p&&o.t===ty)); }
  if(t.matches("tr[data-pp]")){ const p=t.dataset.pp,ty=t.dataset.pt; return openDeals(who+" — "+p+" / "+ty, mine.filter(o=>o.p===p&&o.t===ty)); }
  if(t.matches(".barrow")){ const p=t.dataset.dp,ty=t.dataset.dt; return openDeals(F.rep+" — "+p+(ty?" / "+ty:""), mine.filter(o=>o.r===F.rep&&o.p===p&&(!ty||o.t===ty))); }
  if(t.matches("td.nm")) return pickRep(t.dataset.rep);
  if(t.matches("th.sort")&&t.dataset.hk!==undefined){ const k=t.dataset.hk; HSORT = HSORT.k===k?{k,d:-HSORT.d}:{k,d:k==="r"?1:-1}; return renderHeat(); }
  if(t.matches("th.sort")&&t.dataset.pk!==undefined){ PSORT=flip(PSORT,t.dataset.pk); return renderPT(); }
  if(t.matches("th.sort")){ SORT=flip(SORT,t.dataset.k); return renderTable(); }
  if(t.id==="sortdir"){ SORT.d=-SORT.d; return renderTable(); }
  if(t.matches("tr[data-rep]")) return pickRep(t.dataset.rep);
});
document.addEventListener("keydown", e=>{
  if(e.key==="Escape"&&DR) return closeDrawer();
  if((e.key==="Enter"||e.key===" ")&&e.target.matches("td.cell,td.nm,.barrow")){ e.preventDefault(); e.target.click(); }
});
$("f-sm").addEventListener("change",e=>{ F.sm=e.target.value; F.tl=""; F.rep=""; render(); });
$("f-tl").addEventListener("change",e=>{ F.tl=e.target.value; F.rep=""; if(F.tl){ const x=D.reps.find(y=>y.tl===F.tl); if(x&&!F.sm) F.sm=x.sm; } render(); });
$("f-rep").addEventListener("change",e=>{ if(e.target.value) pickRep(e.target.value); else { F.rep=""; render(); } });
$("sortby").addEventListener("change",e=>{ const k=e.target.value; SORT={k,d:TXT.has(k)?1:-1}; renderTable(); });
$("q").addEventListener("input",e=>{ Q=e.target.value.trim(); renderFilters(); renderTable(); });

$("sub").textContent = "Source: master reservations sheet ("+META.source+") only · generated "+META.generated+" Cairo";
$("foot").innerHTML = "<b>Left out, as if not in the sheet:</b> "+esc(META.left_out_rules)+". "+
  "Rows from the excluded team leads ("+esc(META.excluded_tls)+", "+fmtN(META.tl_excluded_rows)+" rows) only appear in the company-wide “including left out” total. No filter on the client's country. "+
  "<b>How it is measured.</b> Total sold = every remaining unit. Win rate = (Yes + No) ÷ (Yes + No + Terminated); Terminated is still sold but counts as lost for the rep; a blank Contracted value counts in Total sold but not in the win rate. "+
  "Quarters come from the sheet's Quarter column. A unit shared by two reps counts half for each, so totals equal the number of rows. "+
  "Each rep sits under the latest team leader on their reservations, and each team leader under the latest sales manager written with them. Amounts in "+CUR+".";
if(META.unset_share>0.2){ const w=$("warn"); w.hidden=false;
  w.innerHTML="<b>Check the Contracted column.</b> "+Math.round(META.unset_share*100)+"% of units have a Contracted value the script did not recognise, so they are not in the win rate. Values seen: "+
    META.unset_values.map(([v,n])=>"“"+esc(v)+"” ×"+n).join(", ")+". Fix them in the master sheet or set CONTRACTED_COLUMN / CONTRACTED_YES / CONTRACTED_NO in the script."; }
render();
</script>
"""


# ============================ 6. MAIN ============================
def ordinal(n): return f"{n}{'th' if 10 <= n % 100 <= 20 else {1:'st',2:'nd',3:'rd'}.get(n % 10,'th')}"

def main():
    phase("1) Reading the master sheet → copying into the working sheet…")
    raw = read_sheet()
    if not raw: print("❌ no reservation rows found — check SOURCE_URL / SOURCE_TABS"); return
    phase("2) Applying exclusions and the Contracted breakdown…")
    rows, left_out, team = build(raw)

    today = datetime.now(TZ).date()
    year = today.year
    quarters = [q for q in ("Q1", "Q2", "Q3", "Q4") if any(r["q"] == q for r in rows)]
    qlabel = {"Q1": "Jan–Mar", "Q2": "Apr–Jun", "Q3": "Jul–Sep", "Q4": "Oct–Dec"}
    gen = datetime.now(TZ).strftime("%Y-%m-%d %H:%M")
    rules = [f"Contracted = {', '.join(EXCLUDE_CONTRACTED)}"] + (["Mr. Mahmoud's Referral ticked"] if EXCLUDE_REFERRALS else []) \
            + (["units sold on launch days"] if EXCLUDE_LAUNCH_DAYS else [])
    meta = {"quarters": quarters, "qlabel": qlabel, "source": ", ".join(SOURCE_TABS) or "all tabs", "generated": gen,
            "currency": CURRENCY, "min_decided": MIN_DECIDED, "min_closer": MIN_CLOSER,
            "left_out_rules": "; ".join(rules), "excluded_tls": ", ".join(EXCLUDE_TEAM_LEADS),
            "tl_excluded_rows": len({(x["tab"], x["row"]) for x in left_out if x["reason"].startswith("Team Lead")}),
            "unset_values": Counter((r["contract_raw"] or "(blank)") for r in rows if r["status"] == "unset").most_common(6),
            "unset_share": pct(sum(r["w"] for r in rows if r["status"] == "unset"), sum(r["w"] for r in rows)) or 0}
    tag = f"{FILE_TAG}_{ordinal(today.day)}{today:%b}"
    html_out, xlsx_out = f"{tag}.html", f"{tag}.xlsx"
    with open(html_out, "w", encoding="utf-8") as f:
        f.write(render(rows, left_out, team, meta))
    write_excel(xlsx_out, rows, left_out, gen)

    reps = aggregate(rows)[0]
    print(f"\n📊 {len(reps)} reps · {sum(r['sold'] for r in reps.values()):,.1f} units sold")
    top = sorted(reps.values(), key=lambda r: -r["sold"])[:25]
    P = by_period(rows)
    print(pd.DataFrame([{"Rep": r["rep"], "TL": r["tl"], "Sold": round(r["sold"], 1),
                         **{q: round(P[q][0].get(r["rep"], {}).get("sold", 0), 1) for q in P if q != "Overall"},
                         "Yes": round(r["yes"], 1), "No": round(r["no"], 1), "Term.": round(r["term"], 1),
                         "Win %": f"{r['win_rate']:.0%}" if r["win_rate"] is not None else "—",
                         "#1 Strength": r["top_combo"] or "—"} for r in top]).to_string(index=False))
    print(f"\n📄 {os.path.abspath(html_out)}\n📄 {os.path.abspath(xlsx_out)}\n✅ Total runtime: {time.time()-T0:,.0f}s")
    if AUTO_DOWNLOAD:
        try:
            from google.colab import files
            files.download(html_out); files.download(xlsx_out)
        except Exception:
            pass


main()
