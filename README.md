# Competitor Analysis

Competitor intelligence tool for self-storage market analysis. Scrapes live unit pricing from every nearby competitor and generates a standalone Excel report with comparative analysis against your own property.

---

## How It Works

1. Export a Yardi area self-storage report for your target market
2. Double-click **`Run Competitor Intelligence.bat`**
3. Enter the market name (e.g. "Portland, Oregon")
4. Paste a pricing page URL for each competitor — first entry is **your property**
5. The tool batch-scrapes all URLs simultaneously and writes the report

Total runtime is typically 30–90 seconds regardless of competitor count.

---

## Output

**`[Market] Competitor Intelligence YYYY-MM-DD.xlsx`** saved next to the Yardi file.

| Sheet | Contents |
|---|---|
| **Our Position** | Executive summary — your rank, Z-score, and position vs. market for each unit size |
| **Summary** | Full rate matrix. Your row in navy. Competitor cells color-coded green (cheaper) / red (more expensive) |
| **Market Stats** | Avg / median / min / max / spread per unit size and type |
| **All Units** | Flat table of every unit scraped across all facilities |
| **[Facility]** | One sheet per competitor with full unit detail and vs. your rate column |

---

## Metrics Explained

**Effective Move-In Rate** — what the tenant actually pays on day 1:
> Month 1 rent (after promo, max = free) + admin fee + lock fee + first month insurance

**Effective Yearly Rate** — total 12-month cost as a monthly equivalent:
> ((Rate × 12) − promo savings + admin fee + lock fee + insurance × 12) ÷ 12

**Z-Score** — how many standard deviations your rate sits above or below the market mean:
> Z = (Your Rate − Market Mean) ÷ Market Std Dev
> - |Z| < 0.5 → at market average
> - |Z| 0.5–1.0 → slightly above/below
> - |Z| 1.0–2.0 → moderate outlier
> - |Z| > 2.0 → significant outlier

---

## Setup

Requires Python 3.12+ and the packages in `requirements.txt`:

```
pip install -r requirements.txt
```
