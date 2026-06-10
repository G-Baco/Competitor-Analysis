"""
Firecrawl v2 wrapper — extracts storage unit pricing from facility URLs.
Uses batch scraping so all URLs are processed in parallel by Firecrawl.
"""

import time
import requests

FIRECRAWL_API_KEY = "fc-7b63860795ad40e88979d919ebc10f18"
FIRECRAWL_BASE = "https://api.firecrawl.dev/v2"

EXTRACTION_PROMPT = (
    "Extract all storage units listed on this page. For each unit: "
    "1) Capture the web/online rate (the lower discounted rate shown as the online price) AND the in-store rate "
    "(the higher non-discounted rate). If only one price is shown, capture it as in_store_rate. "
    "2) If a promotional offer exists, express it as 'promo_months_free_equivalent' — the number of equivalent "
    "free months the promo represents. Examples: 'first month free' = 1.0, 'first month 50% off' = 0.5, "
    "'50% off first 2 months' = 1.0, 'half off first month' = 0.5, '$50 off' = calculate as $50 / web_rate. "
    "3) Also capture the raw promo text exactly as shown on the page in 'promo_description'. "
    "4) If an admin fee or one-time move-in fee is mentioned anywhere on the page, capture it as 'admin_fee'. "
    "If no admin fee is mentioned, omit the field. "
    "5) Capture 'features' as a comma-separated string of all amenities/attributes for the unit. "
    "Always include floor access level using EXACTLY one of these two values: '1st floor' or '2nd floor+'. "
    "'1st floor' means ground-level access — no stairs or elevator required. "
    "'2nd floor+' means any upper-floor access — requires elevator or stairs, regardless of the actual floor number "
    "(2nd, 3rd, 4th, 5th floor units all = '2nd floor+'). "
    "Indicators of upper floor: elevator access, stairs required, floor number above 1, 'upper level'. "
    "Indicators of ground floor: 'ground level', 'drive-up', '1st floor', no floor mention (default to 1st floor). "
    "Also include climate controlled, drive-up, interior, and any other relevant attributes. "
    "6) At the facility level (not per unit): if a required insurance or protection plan is mentioned, "
    "capture the lowest available monthly cost as 'insurance_monthly'. If optional, still capture it. "
    "If no insurance is mentioned, omit. "
    "7) At the facility level: if a lock purchase or lock fee is required at move-in, capture the cost as 'lock_fee'. "
    "If not mentioned, omit. "
    "8) At the facility level: if an auto-pay or autopay discount is offered, capture the monthly savings amount "
    "as 'auto_pay_discount' (e.g. '$5 off with autopay' = 5.0). If not mentioned, omit."
)

EXTRACTION_SCHEMA = {
    "type": "object",
    "properties": {
        "facility_name": {"type": "string"},
        "admin_fee": {"type": "number"},
        "insurance_monthly": {"type": "number"},
        "lock_fee": {"type": "number"},
        "auto_pay_discount": {"type": "number"},
        "units": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "size": {"type": "string"},
                    "type": {"type": "string"},
                    "features": {"type": "string"},
                    "web_rate": {"type": "number"},
                    "in_store_rate": {"type": "number"},
                    "promo_description": {"type": "string"},
                    "promo_months_free_equivalent": {"type": "number"},
                    "availability": {"type": "number"},
                },
            },
        },
    },
}

HEADERS = {
    "Authorization": f"Bearer {FIRECRAWL_API_KEY}",
    "Content-Type": "application/json",
}


def effective_rate(rate, promo_months, admin_fee=0.0):
    """((rate × 12) - promo_savings + admin_fee) / 12"""
    if not rate:
        return None
    savings = (promo_months or 0.0) * rate
    return round(((rate * 12) - savings + admin_fee) / 12, 2)


def _normalize_units(raw_units, admin_fee, insurance_monthly=0.0, lock_fee=0.0, auto_pay_discount=0.0):
    units = []
    for u in (raw_units or []):
        web_rate = u.get("web_rate")
        in_store_rate = u.get("in_store_rate")
        promo_months = u.get("promo_months_free_equivalent") or 0.0
        units.append({
            "size": u.get("size"),
            "type": u.get("type"),
            "features": u.get("features"),
            "web_rate": web_rate,
            "in_store_rate": in_store_rate,
            "promo_description": u.get("promo_description"),
            "promo_months": promo_months if promo_months else None,
            "admin_fee": admin_fee if admin_fee else None,
            "insurance_monthly": insurance_monthly if insurance_monthly else None,
            "lock_fee": lock_fee if lock_fee else None,
            "auto_pay_discount": auto_pay_discount if auto_pay_discount else None,
            "effective_rate": effective_rate(in_store_rate or web_rate, promo_months, admin_fee),
            "availability": u.get("availability"),
        })
    return units


def scrape_urls(urls):
    """Batch scrape a list of URLs. Returns dict of {url: (facility_name, admin_fee, units)}."""

    # Submit batch job
    payload = {
        "urls": urls,
        "formats": [{"type": "json", "schema": EXTRACTION_SCHEMA, "prompt": EXTRACTION_PROMPT}],
        "maxAge": 0,  # always fetch fresh — no cached pricing
    }
    resp = requests.post(f"{FIRECRAWL_BASE}/batch/scrape", headers=HEADERS, json=payload, timeout=30)
    resp.raise_for_status()
    batch_id = resp.json()["id"]
    status_url = f"{FIRECRAWL_BASE}/batch/scrape/{batch_id}"

    print(f"  Batch job {batch_id} submitted for {len(urls)} URLs...")

    # Poll until complete
    while True:
        time.sleep(3)
        status_resp = requests.get(status_url, headers=HEADERS, timeout=30)
        status_resp.raise_for_status()
        data = status_resp.json()
        completed = data.get("completed", 0)
        total = data.get("total", len(urls))
        print(f"  Progress: {completed}/{total}")
        if data.get("status") == "completed":
            break
        if completed > 0 and completed >= total:
            break

    # Parse results — keyed by sourceURL
    results = {}
    for item in data.get("data", []):
        source_url = item.get("metadata", {}).get("sourceURL", "")
        extract = (item.get("json") or {})
        facility_name = extract.get("facility_name")
        admin_fee = extract.get("admin_fee") or 0.0
        insurance_monthly = extract.get("insurance_monthly") or 0.0
        lock_fee = extract.get("lock_fee") or 0.0
        auto_pay_discount = extract.get("auto_pay_discount") or 0.0
        units = _normalize_units(extract.get("units"), admin_fee, insurance_monthly, lock_fee, auto_pay_discount)
        results[source_url] = (facility_name, admin_fee, units)

    return results


def scrape_url(url):
    """Scrape a single URL. Convenience wrapper around scrape_urls."""
    results = scrape_urls([url])
    # Match by URL regardless of trailing slash
    for key, val in results.items():
        if key.rstrip("/") == url.rstrip("/"):
            return val
    # Fallback: return first result
    if results:
        return next(iter(results.values()))
    return None, 0.0, []
