"""
Competitor Intelligence — full market scrape for every nearby self-storage property.

Flow:
  1. Pick Yardi area self-storage export
  2. Enter a market name (e.g. "Portland, Oregon")
  3. For each facility, paste their pricing page URL — first facility is YOUR property
  4. Batch-scrape all URLs in parallel
  5. Write a standalone Excel intelligence report

Output: Portland, Oregon Competitor Intelligence YYYY-MM-DD.xlsx
  Our Position  your rates vs. the market at a glance — rank, median, % off
  Summary       full rate matrix with color-coded competitive gaps
  Market Stats  median / avg / min / max per unit size and type
  All Units     flat table of every unit extracted
  [Facility]    one sheet per competitor with full unit detail + vs. your rate
"""

import re
import statistics
import sys
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from scrape import scrape_urls

# ── Style constants ───────────────────────────────────────────────────────────

DARK_FILL     = PatternFill("solid", fgColor="1F3864")   # navy — headers
ACCENT_FILL   = PatternFill("solid", fgColor="2E75B6")   # blue — sub-headers
LIGHT_FILL    = PatternFill("solid", fgColor="D6E4F0")   # light blue — alternating
OUR_FILL      = PatternFill("solid", fgColor="1F3864")   # navy — our property row
MEDIAN_FILL   = PatternFill("solid", fgColor="F2F2F2")   # gray — median row
CHEAPER_FILL  = PatternFill("solid", fgColor="E2EFDA")   # green — they're cheaper
PRICIER_FILL  = PatternFill("solid", fgColor="FCE4D6")   # red — they're more expensive
GREEN_FILL    = PatternFill("solid", fgColor="C6EFCE")   # strong green — we're winning
RED_FILL      = PatternFill("solid", fgColor="FFC7CE")   # strong red — we're exposed

WHITE_FONT    = Font(color="FFFFFF", bold=True)
BOLD_FONT     = Font(bold=True)
GRAY_FONT     = Font(color="666666", italic=True)
CENTER        = Alignment(horizontal="center", vertical="center")
CENTER_WRAP   = Alignment(horizontal="center", vertical="center", wrap_text=True)

SUMMARY_SIZES = ["5x5", "5x10", "5x15", "10x10", "10x15", "10x20", "10x25", "10x30"]
UNIT_TYPES    = ["1st", "1st CC", "Drive-Up", "Drive Up CC", "2nd", "2nd CC"]


# ── Unit classification ───────────────────────────────────────────────────────

def _blob(unit):
    return f"{unit.get('features') or ''} {unit.get('type') or ''}".lower()

def _is_cc(unit):
    return any(t in _blob(unit) for t in ["climate", " cc", "cc ", "cool", "heated"])

def _is_drive_up(unit):
    return any(t in _blob(unit) for t in ["drive-up", "drive up", "driveup"])

def _is_vehicle(unit):
    return any(t in f"{_blob(unit)} {unit.get('size') or ''}".lower()
               for t in ["rv", "boat", "vehicle", "parking"])

def _floor(unit):
    upper = ["2nd floor", "second floor", "upper level", "upper floor",
             "elevator", "3rd floor", "4th floor", "5th floor", "stairs"]
    return "2nd" if any(t in _blob(unit) for t in upper) else "1st"

def classify_unit(unit):
    if _is_vehicle(unit):       return "Vehicle"
    if _is_drive_up(unit):      return "Drive Up CC" if _is_cc(unit) else "Drive-Up"
    cc, fl = _is_cc(unit), _floor(unit)
    if fl == "2nd":             return "2nd CC" if cc else "2nd"
    return "1st CC" if cc else "1st"

def normalize_size(s):
    if not s:
        return None
    s = str(s).lower().replace("'", "").replace('"', "").replace(" ", "")
    m = re.match(r"(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)", s)
    if not m:
        return s
    a, b = float(m.group(1)), float(m.group(2))
    lo, hi = (a, b) if a <= b else (b, a)
    lo = int(lo) if lo == int(lo) else lo
    hi = int(hi) if hi == int(hi) else hi
    return f"{lo}x{hi}"

def _is_locker(unit):
    return "locker" in f"{unit.get('type') or ''} {unit.get('size') or ''}".lower()


# ── Rate calculations ─────────────────────────────────────────────────────────

def effective_move_in_rate(rate, promo_months, admin_fee=0.0, lock_fee=0.0):
    """What the tenant pays on move-in day: month-1 rent (after promo) + admin fee + lock fee."""
    if not rate:
        return None
    month1_rent = rate * max(0.0, 1.0 - min(promo_months or 0.0, 1.0))
    return round(month1_rent + (admin_fee or 0) + (lock_fee or 0), 2)

def effective_yearly_rate(rate, promo_months, admin_fee=0.0, lock_fee=0.0):
    """Total 12-month cost as a monthly equivalent: ((rate×12) − promo savings + admin + lock) / 12."""
    if not rate:
        return None
    promo_savings = (promo_months or 0.0) * rate
    total = (rate * 12) - promo_savings + (admin_fee or 0) + (lock_fee or 0)
    return round(total / 12, 2)

def sqft_from_size(size_norm):
    """Parse square footage from a normalized size string, e.g. '5x10' → 50."""
    if not size_norm:
        return None
    m = re.match(r"(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)", str(size_norm))
    return round(float(m.group(1)) * float(m.group(2)), 1) if m else None


# ── Yardi import ──────────────────────────────────────────────────────────────

def _yardi_col(df, *candidates):
    norm = {c.strip().lower(): c for c in df.columns}
    for c in candidates:
        if c.strip().lower() in norm:
            return norm[c.strip().lower()]
    return None

def parse_yardi(path):
    df = pd.read_excel(path, sheet_name=0, header=0)
    col_name     = _yardi_col(df, "Property Name")
    col_address  = _yardi_col(df, "Address")
    col_city     = _yardi_col(df, "City")
    col_state    = _yardi_col(df, "State")
    col_zip      = _yardi_col(df, "ZIP")
    col_phone    = _yardi_col(df, "Phone Number")
    col_nrsf_a   = _yardi_col(df, "Actual Rentable SqFt")
    col_nrsf_e   = _yardi_col(df, "Estimated Rentable SqFt[1]", " Estimated Rentable SqFt[1]")
    col_year     = _yardi_col(df, "Completion Year")
    col_status   = _yardi_col(df, "Property Status")
    col_distance = _yardi_col(df, "Distance (miles)")

    facilities = []
    for _, row in df.iterrows():
        name = row[col_name] if col_name else None
        if pd.isna(name):
            continue
        city  = str(row[col_city]).strip()  if col_city  and pd.notna(row[col_city])  else ""
        state = str(row[col_state]).strip() if col_state and pd.notna(row[col_state]) else ""
        zip_s = str(row[col_zip]).strip()   if col_zip   and pd.notna(row[col_zip])   else ""
        addr  = str(row[col_address]).strip() if col_address and pd.notna(row[col_address]) else ""
        dist  = float(row[col_distance]) if col_distance and pd.notna(row[col_distance]) else 9999

        nrsf_raw = (row[col_nrsf_a] if col_nrsf_a and pd.notna(row[col_nrsf_a]) else
                    row[col_nrsf_e] if col_nrsf_e and pd.notna(row[col_nrsf_e]) else None)
        nrsf = int(round(float(nrsf_raw))) if nrsf_raw is not None and pd.notna(nrsf_raw) else None

        year = None
        yr_raw = row[col_year] if col_year else None
        if yr_raw is not None and pd.notna(yr_raw):
            year = yr_raw.year if hasattr(yr_raw, "year") else (int(yr_raw) if str(yr_raw).isdigit() else None)

        phone_raw = row[col_phone] if col_phone else None
        phone = None
        if phone_raw is not None and pd.notna(phone_raw):
            digits = re.sub(r"\D", "", str(phone_raw))
            phone = f"({digits[:3]}) {digits[3:6]}-{digits[6:]}" if len(digits) == 10 else digits

        status = str(row[col_status]).strip() if col_status and pd.notna(row[col_status]) else ""

        facilities.append({
            "name":     str(name).strip(),
            "address":  f"{addr}, {city}, {state} {zip_s}".strip(", ") if addr else f"{city}, {state}".strip(),
            "city":     city,
            "state":    state,
            "phone":    phone,
            "distance": dist,
            "dist_str": f"{dist:.2f} mi" if dist < 9999 else "",
            "nrsf":     nrsf,
            "status":   status,
            "year":     year,
            "url":      None,
        })

    facilities.sort(key=lambda f: f["distance"])
    return facilities


# ── URL collection ────────────────────────────────────────────────────────────

def collect_urls(facilities):
    print("\n  Collecting URLs...")
    for i, fac in enumerate(facilities):
        is_ours = (i == 0)
        label   = "YOUR PROPERTY" if is_ours else f"Competitor {i}"
        entered = simpledialog.askstring(
            f"{'★ ' if is_ours else ''}URL — {label}",
            f"{'★ YOUR PROPERTY ★' if is_ours else 'Competitor'}\n\n"
            f"Facility:  {fac['name']}\n"
            f"Address:   {fac['address']}\n"
            f"Distance:  {fac.get('dist_str', '') or '—'}\n\n"
            "Paste their storage pricing/units page URL.\n"
            "(Leave blank to skip.)",
        )
        if entered and entered.strip().lower().startswith(("http://", "https://")):
            fac["url"] = entered.strip()
            print(f"    {'★' if is_ours else '✓'} {fac['name']}")
        else:
            print(f"    — {fac['name']} skipped")


# ── Scraping + enrichment ─────────────────────────────────────────────────────

def enrich_units(units):
    for u in units:
        u["size_norm"] = normalize_size(u.get("size"))
        u["unit_type"] = classify_unit(u)
        rate  = u.get("in_store_rate") or u.get("web_rate")
        promo = u.get("promo_months")
        admin = u.get("admin_fee") or 0.0
        lock  = u.get("lock_fee") or 0.0
        ins   = u.get("insurance_monthly") or 0.0
        u["effective_move_in"] = effective_move_in_rate(rate, promo, admin, lock)
        u["effective_yearly"]  = effective_yearly_rate(rate, promo, admin, lock)
        u["sqft"]              = sqft_from_size(u.get("size_norm"))
        u["rate_per_sqft"]     = round(rate / u["sqft"], 2) if rate and u["sqft"] else None
    return [u for u in units if not _is_locker(u)]

def scrape_all(facilities):
    to_scrape = [f for f in facilities if f.get("url")]
    if not to_scrape:
        return []

    print(f"\n  Scraping {len(to_scrape)} URLs in parallel...")
    batch = scrape_urls([f["url"] for f in to_scrape])

    results = []
    for i, fac in enumerate(facilities):
        if not fac.get("url"):
            continue
        url = fac["url"]
        match = (batch.get(url)
                 or batch.get(url.rstrip("/") + "/")
                 or batch.get(url.rstrip("/")))
        if not match or not match[2]:
            print(f"    ✗ {fac['name']} — no units returned")
            continue
        scraped_name, admin_fee, units = match
        units = enrich_units(units)
        is_ours = (i == 0)
        print(f"    {'★' if is_ours else '✓'} {fac['name']} ({len(units)} units)")
        results.append({
            "facility":     fac,
            "display_name": scraped_name or fac["name"],
            "admin_fee":    admin_fee or 0.0,
            "units":        units,
            "is_ours":      is_ours,
        })

    return results


# ── Helpers ───────────────────────────────────────────────────────────────────

def _header_row(ws, row, values, fill, font, col_start=1):
    for i, val in enumerate(values, col_start):
        c = ws.cell(row=row, column=i, value=val)
        c.fill = fill
        c.font = font
        c.alignment = CENTER_WRAP

def _autofit(ws, min_w=8, max_w=50):
    for col_cells in ws.columns:
        letter, best = None, 0
        for cell in col_cells:
            letter = get_column_letter(cell.column)
            try:
                best = max(best, len(str(cell.value or "")))
            except Exception:
                pass
        if letter:
            ws.column_dimensions[letter].width = min(max_w, max(min_w, best + 2))

def _safe_sheet_name(name, used):
    name = re.sub(r'[\\/:*?"<>|]', "", name)[:31].strip()
    base = name
    for i in range(1, 100):
        candidate = base if i == 1 else f"{base[:28]} {i}"
        if candidate not in used:
            used.add(candidate)
            return candidate
    return name

def get_our_rates(our_scraped):
    """Returns {size_norm: cheapest_rate} for our property."""
    rates = {}
    if not our_scraped:
        return rates
    for u in our_scraped["units"]:
        size = u.get("size_norm")
        rate = u.get("in_store_rate") or u.get("web_rate")
        if size and rate:
            if size not in rates or rate < rates[size]:
                rates[size] = rate
    return rates


# ── Sheet: Our Position ───────────────────────────────────────────────────────

def write_our_position_sheet(wb, scraped, our_rates, market_name, today):
    """
    Executive summary — your rates vs. the market for every unit size.
    One row per size: Your Rate | # Cheaper | Market Min | Market Median | Market Max | Rank | vs. Median
    """
    ws = wb.create_sheet("Our Position")

    ws.merge_cells("A1:I1")
    ws["A1"] = f"Our Competitive Position  —  {market_name}  |  {today}"
    ws["A1"].font = Font(bold=True, size=13, color="1F3864")
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:L2")
    ws["A2"] = (
        "vs. Median: negative ($) = we are cheaper (green = competitive), positive = we are more expensive (red = exposure).  "
        "Rank 1 = cheapest in market.  Z-Score = SDs from market mean (green = below mean, red = above).  "
        "Blank row = we don't offer that size."
    )
    ws["A2"].font = Font(italic=True, color="666666", size=9)
    ws.row_dimensions[2].height = 16

    # Col layout: Size | Sqft | Our Rate | Our $/sqft | Mkt Median | Mkt Median $/sqft |
    #             vs. Median ($) | vs. Median (%) | Rank (1=cheapest) | Mkt Std Dev | Z-Score | Position
    headers = [
        "Unit Size", "Sqft",
        "Our Rate", "Our $/sqft",
        "Mkt Median", "Mkt Median\n$/sqft",
        "vs. Median ($)\n− = we're cheaper", "vs. Median (%)",
        "Our Rank\n(1 = cheapest)",
        "Mkt Std Dev", "Z-Score", "Position",
    ]
    _header_row(ws, 3, headers, DARK_FILL, WHITE_FONT)
    ws.row_dimensions[3].height = 40

    competitors = [s for s in scraped if not s["is_ours"]]

    row = 4
    for size in SUMMARY_SIZES:
        norm    = normalize_size(size)
        sf      = sqft_from_size(norm)
        our_rate = our_rates.get(norm)
        our_per_sqft = round(our_rate / sf, 2) if our_rate and sf else None

        comp_rates = []
        for s in competitors:
            for u in s["units"]:
                if u.get("size_norm") == norm and u.get("unit_type") != "Vehicle":
                    r = u.get("in_store_rate") or u.get("web_rate")
                    if r:
                        comp_rates.append(r)
                        break  # one (cheapest) per facility

        if not our_rate and not comp_rates:
            continue

        mkt_median  = round(statistics.median(comp_rates), 2)           if comp_rates else None
        mkt_avg     = round(sum(comp_rates) / len(comp_rates), 2)       if comp_rates else None
        mkt_stdev   = round(statistics.stdev(comp_rates), 2)            if len(comp_rates) >= 2 else None
        mkt_med_psf = round(mkt_median / sf, 2)                         if mkt_median and sf else None

        rank_str         = None
        vs_median_dollar = None
        vs_median_pct    = None
        z_score          = None
        position_str     = None

        if our_rate and comp_rates:
            all_rates = sorted(comp_rates + [our_rate])
            rank      = all_rates.index(our_rate) + 1
            rank_str  = f"{rank} of {len(all_rates)}"

            if mkt_median:
                vs_median_dollar = round(our_rate - mkt_median, 2)
                vs_median_pct    = round((our_rate - mkt_median) / mkt_median * 100, 1)

            if mkt_avg and mkt_stdev and mkt_stdev > 0:
                z_score = round((our_rate - mkt_avg) / mkt_stdev, 2)

            if vs_median_pct is not None:
                apct = abs(vs_median_pct)
                direction = "above" if vs_median_pct > 0 else "below"
                if apct < 3:
                    position_str = "At market median"
                elif apct < 10:
                    position_str = f"Slightly {direction} median ({vs_median_pct:+.1f}%)"
                elif apct < 20:
                    position_str = f"Moderately {direction} median ({vs_median_pct:+.1f}%)"
                elif apct < 35:
                    position_str = f"Well {direction} median ({vs_median_pct:+.1f}%)"
                else:
                    position_str = f"Significantly {direction} median ({vs_median_pct:+.1f}%)"

        vals = [
            size, sf,
            our_rate, our_per_sqft,
            mkt_median, mkt_med_psf,
            vs_median_dollar, vs_median_pct,
            rank_str,
            mkt_stdev, z_score, position_str,
        ]

        for col_idx, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=col_idx, value=val)
            c.alignment = CENTER

        # vs. Median, Rank, Position: color by vs. median sign
        if vs_median_dollar is not None:
            med_fill = GREEN_FILL if vs_median_dollar <= 0 else RED_FILL
            for col in (7, 8, 9, 12):
                ws.cell(row=row, column=col).fill = med_fill

        # Z-score: independent coloring
        if z_score is not None:
            ws.cell(row=row, column=11).fill = GREEN_FILL if z_score <= 0 else RED_FILL

        if vs_median_pct is not None:
            ws.cell(row=row, column=8).value = f"{vs_median_pct:+.1f}%"

        row += 1

    _autofit(ws)
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["C"].width = 11
    ws.column_dimensions["E"].width = 11
    ws.column_dimensions["G"].width = 18
    ws.column_dimensions["L"].width = 32


# ── Sheet: Summary ────────────────────────────────────────────────────────────

def write_summary_sheet(wb, scraped, our_rates, market_name, today):
    """
    Matrix of cheapest rate per facility × unit size.
    Our property row is pinned at top in navy.
    Competitor cells are color-coded: green = cheaper than us, red = more expensive.
    """
    ws = wb.create_sheet("Summary")
    ws.freeze_panes = "C4"

    ws.merge_cells(f"A1:{get_column_letter(2 + len(SUMMARY_SIZES))}1")
    ws["A1"] = f"{market_name}  —  Competitor Rate Summary  |  {today}"
    ws["A1"].font = Font(bold=True, size=13, color="1F3864")
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 24

    # Legend row
    ws.merge_cells(f"A2:{get_column_letter(2 + len(SUMMARY_SIZES))}2")
    ws["A2"] = (
        "Rates shown are cheapest in-store rate for any unit of that size.  "
        "Green cell = competitor is cheaper than us for that size.  "
        "Red cell = competitor is more expensive."
    )
    ws["A2"].font = Font(italic=True, color="666666", size=9)
    ws.row_dimensions[2].height = 14

    _header_row(ws, 3, ["Facility", "Distance (mi)", "Website"] + SUMMARY_SIZES, DARK_FILL, WHITE_FONT)
    ws.row_dimensions[3].height = 28

    # Sort: our property first, then competitors by distance
    ours   = [s for s in scraped if s["is_ours"]]
    others = [s for s in scraped if not s["is_ours"]]
    ordered = ours + others

    data_start = 4
    for row_idx, s in enumerate(ordered, data_start):
        is_ours = s["is_ours"]
        units   = s["units"]

        dist_raw  = s["facility"].get("distance")
        dist_num  = round(dist_raw, 4) if dist_raw and dist_raw < 9999 else None
        url = s["facility"].get("url", "") or ""
        name_cell = ws.cell(row=row_idx, column=1, value=s["display_name"])
        dist_cell = ws.cell(row=row_idx, column=2, value=0.0 if is_ours else dist_num)
        url_cell  = ws.cell(row=row_idx, column=3, value=url)
        if url:
            url_cell.hyperlink = url
            url_cell.font = Font(color="FFFFFF" if is_ours else "0563C1", underline="single",
                                 bold=is_ours)

        if is_ours:
            for col in range(1, 4 + len(SUMMARY_SIZES)):
                c = ws.cell(row=row_idx, column=col)
                c.fill = OUR_FILL
                if col != 3:  # URL cell font already set above
                    c.font = WHITE_FONT
        else:
            name_cell.font = BOLD_FONT

        for col_idx, size in enumerate(SUMMARY_SIZES, 4):
            norm = normalize_size(size)
            matching = [u for u in units
                        if u.get("size_norm") == norm and u.get("unit_type") != "Vehicle"]
            price = None
            if matching:
                in_store = [u["in_store_rate"] for u in matching if u.get("in_store_rate")]
                web      = [u["web_rate"]      for u in matching if u.get("web_rate")]
                prices   = in_store or web
                price    = round(min(prices), 2) if prices else None

            c = ws.cell(row=row_idx, column=col_idx, value=price)
            c.alignment = CENTER

            if is_ours:
                c.fill = OUR_FILL
                c.font = WHITE_FONT
            elif price is not None:
                our_rt = our_rates.get(norm)
                if our_rt:
                    c.fill = CHEAPER_FILL if price < our_rt else PRICIER_FILL

    # Market median row
    last_data = data_start + len(ordered) - 1
    median_row = last_data + 2
    med_label = ws.cell(row=median_row, column=1, value="MARKET MEDIAN")
    med_label.font = Font(bold=True, italic=True)
    med_label.fill = MEDIAN_FILL

    comp_rows = [data_start + i for i, s in enumerate(ordered) if not s["is_ours"]]
    for col_idx in range(4, 4 + len(SUMMARY_SIZES)):
        col_letter = get_column_letter(col_idx)
        if comp_rows:
            addrs = ",".join(f"{col_letter}{r}" for r in comp_rows)
            formula = f"=IFERROR(MEDIAN({addrs}),\"\")"
        else:
            formula = ""
        c = ws.cell(row=median_row, column=col_idx, value=formula)
        c.fill = MEDIAN_FILL
        c.font = Font(bold=True, italic=True)
        c.alignment = CENTER

    ws.cell(row=median_row, column=2).fill = MEDIAN_FILL
    ws.cell(row=median_row, column=3).fill = MEDIAN_FILL

    _autofit(ws)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 13


# ── Sheet: Market Stats ───────────────────────────────────────────────────────

def write_market_stats_sheet(wb, scraped):
    ws = wb.create_sheet("Market Stats")
    competitors = [s for s in scraped if not s["is_ours"]]

    headers = ["Unit Size", "Unit Type", "# Facilities", "Avg Rate", "Median", "Min", "Max", "Spread"]
    _header_row(ws, 1, headers, DARK_FILL, WHITE_FONT)
    ws.row_dimensions[1].height = 25

    row = 2
    last_size = None
    for size in SUMMARY_SIZES:
        norm = normalize_size(size)
        for utype in UNIT_TYPES:
            rates = []
            for s in competitors:
                for u in s["units"]:
                    if u.get("size_norm") == norm and u.get("unit_type") == utype:
                        rate = u.get("in_store_rate") or u.get("web_rate")
                        if rate:
                            rates.append(rate)
            if not rates:
                continue
            fill = LIGHT_FILL if size != last_size and last_size is not None else None
            last_size = size
            for col_idx, val in enumerate([
                size, utype, len(rates),
                round(sum(rates) / len(rates), 2),
                round(statistics.median(rates), 2),
                round(min(rates), 2),
                round(max(rates), 2),
                round(max(rates) - min(rates), 2),
            ], 1):
                c = ws.cell(row=row, column=col_idx, value=val)
                if fill:
                    c.fill = fill
            row += 1

    _autofit(ws)


# ── Sheet: All Units ──────────────────────────────────────────────────────────

def write_all_units_sheet(wb, scraped, today):
    ws = wb.create_sheet("All Units")

    headers = [
        "Survey Date", "Facility", "Our Property?", "Distance (mi)", "Distance Label",
        "Size", "Size (Norm)", "Unit Type", "Features",
        "Web Rate", "In-Store Rate", "Promo", "Promo Months",
        "Admin Fee", "Insurance/Mo", "Lock Fee", "Auto-Pay Discount",
        "Eff. Move-In", "Eff. Yearly Rate",
    ]
    _header_row(ws, 1, headers, DARK_FILL, WHITE_FONT)
    ws.row_dimensions[1].height = 25

    row = 2
    for s in scraped:
        dist_raw = s["facility"].get("distance")
        dist_num = round(dist_raw, 4) if dist_raw and dist_raw < 9999 else None
        for u in s["units"]:
            for col_idx, val in enumerate([
                today,
                s["display_name"],
                "Yes" if s["is_ours"] else "",
                dist_num,
                s["facility"].get("dist_str", ""),
                u.get("size"),
                u.get("size_norm"),
                u.get("unit_type"),
                u.get("features"),
                u.get("web_rate"),
                u.get("in_store_rate"),
                u.get("promo_description"),
                u.get("promo_months"),
                u.get("admin_fee"),
                u.get("insurance_monthly"),
                u.get("lock_fee"),
                u.get("auto_pay_discount"),
                u.get("effective_move_in"),
                u.get("effective_yearly"),
            ], 1):
                ws.cell(row=row, column=col_idx, value=val)
            row += 1

    _autofit(ws)
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["H"].width = 45
    ws.column_dimensions["K"].width = 35


# ── Sheet: Per facility ───────────────────────────────────────────────────────

def write_facility_sheet(wb, s, our_rates, today, used_names):
    sheet_name = _safe_sheet_name(s["display_name"], used_names)
    ws = wb.create_sheet(sheet_name)
    fac = s["facility"]
    is_ours = s["is_ours"]

    # Header block
    ws["A1"] = ("★ " if is_ours else "") + s["display_name"] + (" (OUR PROPERTY)" if is_ours else "")
    ws["A1"].font = Font(bold=True, size=12, color="1F3864")
    ws["A2"] = fac.get("address", "")
    ws["A2"].font = Font(color="444444")

    nrsf_str = f"{fac['nrsf']:,}" if fac.get("nrsf") else "N/A"
    ws["A3"] = (
        f"Distance: {fac.get('dist_str', 'N/A') or 'N/A'}  |  NRSF: {nrsf_str}  |  "
        f"Year Built: {fac.get('year', 'N/A')}  |  Status: {fac.get('status', 'N/A')}  |  "
        f"Surveyed: {today}"
    )
    ws["A3"].font = Font(italic=True, color="666666", size=9)

    url = fac.get("url", "")
    if url:
        ws["A4"] = url
        ws["A4"].hyperlink = url
        ws["A4"].font = Font(color="0563C1", underline="single", size=9)
        info_row = 5
    else:
        info_row = 4

    if s.get("admin_fee"):
        ws[f"A{info_row}"] = f"Admin Fee: ${s['admin_fee']:.0f}"
        ws[f"A{info_row}"].font = Font(italic=True, color="444444")
        info_row += 1

    header_row = info_row + 1
    headers = [
        "Size", "Sqft", "Unit Type", "Features",
        "Web Rate", "In-Store Rate", "$/sqft",
        "Promo", "Promo Months",
        "Admin Fee", "Insurance/Mo", "Lock Fee", "Auto-Pay Disc.",
        "Eff. Move-In", "Eff. Yearly Rate",
    ]
    if not is_ours:
        headers += ["vs. Our Rate", "vs. Our $/sqft"]

    _header_row(ws, header_row, headers, ACCENT_FILL if not is_ours else DARK_FILL, WHITE_FONT)
    ws.row_dimensions[header_row].height = 22

    row = header_row + 1
    for u in s["units"]:
        comp_rate = u.get("in_store_rate") or u.get("web_rate")
        vals = [
            u.get("size"),
            u.get("sqft"),
            u.get("unit_type"),
            u.get("features"),
            u.get("web_rate"),
            u.get("in_store_rate"),
            u.get("rate_per_sqft"),
            u.get("promo_description"),
            u.get("promo_months"),
            u.get("admin_fee"),
            u.get("insurance_monthly"),
            u.get("lock_fee"),
            u.get("auto_pay_discount"),
            u.get("effective_move_in"),
            u.get("effective_yearly"),
        ]

        vs_our = vs_our_psf = None
        if not is_ours:
            norm = u.get("size_norm")
            sf   = u.get("sqft")
            our_rt = our_rates.get(norm) if norm else None
            if our_rt and comp_rate:
                vs_our = round(comp_rate - our_rt, 2)
                our_psf = round(our_rt / sf, 2) if sf else None
                if our_psf and u.get("rate_per_sqft"):
                    vs_our_psf = round(u["rate_per_sqft"] - our_psf, 2)
            vals += [vs_our, vs_our_psf]

        for col_idx, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=col_idx, value=val)
            if col_idx in (5, 6, 7, 14, 15):
                c.alignment = CENTER

        if not is_ours and vs_our is not None:
            fill = CHEAPER_FILL if vs_our < 0 else PRICIER_FILL
            for col in range(len(vals) - 1, len(vals) + 1):
                ws.cell(row=row, column=col).fill = fill
                ws.cell(row=row, column=col).alignment = CENTER

        row += 1

    # Add a note explaining vs. Our Rate
    if not is_ours and not our_rates:
        ws.cell(row=row + 1, column=len(headers)).value = "* Our property URL not entered — no comparison available"
        ws.cell(row=row + 1, column=len(headers)).font = GRAY_FONT

    _autofit(ws)
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["F"].width = 35


# ── Report builder ────────────────────────────────────────────────────────────

def build_report(scraped, out_path, market_name, today):
    wb = Workbook()
    wb.remove(wb.active)

    our_scraped = next((s for s in scraped if s["is_ours"]), None)
    our_rates   = get_our_rates(our_scraped)

    used_names = set()

    write_our_position_sheet(wb, scraped, our_rates, market_name, today)
    used_names.add("Our Position")

    write_summary_sheet(wb, scraped, our_rates, market_name, today)
    used_names.add("Summary")

    write_market_stats_sheet(wb, scraped)
    used_names.add("Market Stats")

    write_all_units_sheet(wb, scraped, today)
    used_names.add("All Units")

    # Our property sheet first, then competitors
    ours   = [s for s in scraped if s["is_ours"]]
    others = [s for s in scraped if not s["is_ours"]]
    for s in ours + others:
        write_facility_sheet(wb, s, our_rates, today, used_names)

    wb.save(out_path)
    print(f"\n  Saved: {out_path}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    root = tk.Tk()
    root.withdraw()

    # 1. Pick Yardi export
    yardi_path = filedialog.askopenfilename(
        title="Select Yardi area self-storage export",
        filetypes=[("Excel files", "*.xlsx *.xlsm")],
    )
    if not yardi_path:
        sys.exit(0)

    # 2. Market name
    market_name = simpledialog.askstring(
        "Market Name",
        "Enter the market name for this report.\n\nExamples:\n"
        "  Portland, Oregon\n  Austin, TX — North\n  Denver Metro",
    )
    if not market_name or not market_name.strip():
        sys.exit(0)
    market_name = market_name.strip()

    # 3. Load Yardi
    print(f"Loading: {Path(yardi_path).name}")
    facilities = parse_yardi(yardi_path)
    if not facilities:
        messagebox.showerror("Yardi Error", "No facilities found in the selected file.")
        sys.exit(1)
    print(f"  Found {len(facilities)} facilities.")
    print(f"  First facility (YOUR PROPERTY): {facilities[0]['name']}")

    # 4. Collect URLs — first prompt is labeled as our property
    messagebox.showinfo(
        "Enter URLs",
        f"You will be prompted for a pricing page URL for each of the "
        f"{len(facilities)} facilities.\n\n"
        f"The first facility is YOUR PROPERTY:\n  {facilities[0]['name']}\n\n"
        "Leave blank to skip any facility.",
    )
    collect_urls(facilities)

    ready = [f for f in facilities if f.get("url")]
    skip  = [f for f in facilities if not f.get("url")]
    if not ready:
        messagebox.showerror("No URLs", "No URLs were entered. Nothing to scrape.")
        sys.exit(1)

    # 5. Confirm
    skip_msg = f"Skipping {len(skip)} (no URL entered).\n" if skip else ""
    if not messagebox.askyesno(
        "Ready to Scrape",
        f"Ready to scrape {len(ready)} facilities.\n{skip_msg}"
        "\nThis usually takes 30–90 seconds. Continue?",
    ):
        sys.exit(0)

    # 6. Scrape
    scraped = scrape_all(facilities)
    if not scraped:
        messagebox.showerror("No Data", "No pricing data was returned. Check your URLs and try again.")
        sys.exit(1)

    if not any(s["is_ours"] for s in scraped):
        messagebox.showwarning(
            "Our Property Missing",
            f"The scrape for {facilities[0]['name']} returned no data.\n\n"
            "The report will be generated without comparative analysis.\n"
            "Check the URL and re-run to get full comparisons.",
        )

    # 7. Build report
    today = date.today().isoformat()
    safe_market = re.sub(r'[\\/:*?"<>|]', "", market_name)
    out_path = Path(yardi_path).parent / f"{safe_market} Competitor Intelligence {today}.xlsx"
    build_report(scraped, out_path, market_name, today)

    messagebox.showinfo(
        "Done",
        f"Report saved:\n{out_path.name}\n\n"
        f"Facilities scraped:   {len(scraped)}\n"
        f"Total units captured: {sum(len(s['units']) for s in scraped)}\n"
        f"Survey date: {today}",
    )


if __name__ == "__main__":
    main()
