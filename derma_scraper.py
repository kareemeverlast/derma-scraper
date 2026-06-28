"""
Derma Scope — Gulf Dermatology Clinic Scraper
==============================================

A one-shot Streamlit app that pulls dermatology clinics across the GCC
(Saudi Arabia, UAE, Qatar, Kuwait, Bahrain, Oman) from the
Google Places API (New) and exports them to Excel / CSV.

Why this design:
  * rating + count only  -> Text Search (New) returns everything we need in ONE
    call via a field mask, so we never pay for a separate Place Details call.
  * Gulf population sits in ~25 metros -> we search a curated city list with a
    radius bias instead of gridding empty desert.
  * Bilingual keywords (EN + AR) -> clinics list under both.

Run:
    pip install streamlit requests pandas openpyxl
    streamlit run derma_scraper.py

You need a Google Cloud project with **Places API (New)** enabled and an API key.
Set a budget alert / quota cap in the console so a runaway loop can't surprise you.
"""

import os
import re
import time
from io import BytesIO
from datetime import datetime, timezone

import requests
import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------- #
# Static data
# --------------------------------------------------------------------------- #

PLACES_ENDPOINT = "https://places.googleapis.com/v1/places:searchText"

# Text Search (New) is billed per request. Used only to estimate run cost.
COST_PER_CALL_USD = 0.032

# Only the fields we actually need -> keeps the billing field-tier low.
FIELD_MASK = ",".join([
    "places.id",
    "places.displayName",
    "places.formattedAddress",
    "places.location",
    "places.rating",
    "places.userRatingCount",
    "places.googleMapsUri",
    "places.websiteUri",
    "places.internationalPhoneNumber",
    "places.types",
    "places.businessStatus",
    "nextPageToken",
])

# Region -> countries we cover. Display name : ISO-2 code (used as Places regionCode).
# Cities for each are pulled from the geonamescache package at runtime.
# Grouped Gulf -> wider Middle East -> North Africa (some countries span groups).
REGION_COUNTRIES = {
    # --- Gulf States ---
    "Saudi Arabia": "SA",
    "United Arab Emirates": "AE",
    "Qatar": "QA",
    "Kuwait": "KW",
    "Bahrain": "BH",
    "Oman": "OM",
    # --- Wider Middle East ---
    "Egypt": "EG",
    "Jordan": "JO",
    "Lebanon": "LB",
    "Syria": "SY",
    "Iraq": "IQ",
    "Iran": "IR",
    "Yemen": "YE",
    "Palestine": "PS",
    "Israel": "IL",
    "Turkey": "TR",
    "Cyprus": "CY",
    # --- North Africa ---
    "Morocco": "MA",
    "Algeria": "DZ",
    "Tunisia": "TN",
    "Libya": "LY",
    "Sudan": "SD",
    "Mauritania": "MR",
    "Western Sahara": "EH",
}


@st.cache_data(show_spinner=False)
def build_region_catalogue() -> dict:
    """
    Build {country: {"region": iso2, "cities": {city: (lat, lng)}}} for every
    country in REGION_COUNTRIES, using the offline geonamescache dataset.
    Cities are ordered most-populous-first; duplicate names keep the largest.
    """
    import geonamescache
    gc = geonamescache.GeonamesCache()

    by_cc = {}
    for c in gc.get_cities().values():
        by_cc.setdefault(c.get("countrycode"), []).append(c)

    catalogue = {}
    for name, iso in REGION_COUNTRIES.items():
        items = sorted(by_cc.get(iso, []),
                       key=lambda x: x.get("population", 0) or 0, reverse=True)
        cities = {}
        for it in items:
            cn = it["name"]
            if cn not in cities:  # first (most populous) wins on duplicate name
                cities[cn] = (float(it["latitude"]), float(it["longitude"]))
        catalogue[name] = {"region": iso, "cities": cities}
    return catalogue

DEFAULT_KEYWORDS_EN = [
    "dermatology clinic",
    "dermatologist",
    "skin clinic",
    "skin care clinic",
    "cosmetic dermatology",
    "laser skin clinic",
    "dermatology center",
]

DEFAULT_KEYWORDS_AR = [
    "عيادة جلدية",
    "دكتور جلدية",
    "طبيب جلدية",
    "جلدية وتجميل",
    "مركز جلدية",
    "عيادة تجميل وليزر",
    "أمراض جلدية",
]


# --------------------------------------------------------------------------- #
# Core logic (no Streamlit calls in here -> unit-testable)
# --------------------------------------------------------------------------- #

def is_arabic(text: str) -> bool:
    """True if the string contains any Arabic-range character."""
    return any("\u0600" <= ch <= "\u06FF" for ch in text)


def parse_place(p: dict, city: str, country: str) -> dict:
    """Flatten one Places API result into a clean row."""
    loc = p.get("location", {}) or {}
    return {
        "place_id": p.get("id", ""),
        "name": (p.get("displayName") or {}).get("text", ""),
        "types": ", ".join(p.get("types", []) or []),
        "address": p.get("formattedAddress", ""),
        "city": city,
        "country": country,
        "lat": loc.get("latitude"),
        "lng": loc.get("longitude"),
        "maps_url": p.get("googleMapsUri", ""),
        "website": p.get("websiteUri", ""),
        "phone": p.get("internationalPhoneNumber", ""),
        "email": "",  # filled later (best-effort) from the clinic website
        "rating": p.get("rating"),
        "review_count": p.get("userRatingCount"),
        "business_status": p.get("businessStatus", ""),
        "scraped_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }


def build_selections(city_selections: dict, catalogue: dict) -> list:
    """{country: [cities]} + catalogue -> [(country, region, city, lat, lng), ...]"""
    out = []
    for country, cities in city_selections.items():
        region = catalogue[country]["region"]
        for city in cities:
            lat, lng = catalogue[country]["cities"][city]
            out.append((country, region, city, lat, lng))
    return out


def merge_catalogue(base: dict, custom: dict) -> dict:
    """Base catalogue + user-added countries/cities (custom wins on overlap)."""
    import copy
    data = copy.deepcopy(base)
    for country, info in (custom or {}).items():
        if country not in data:
            data[country] = {"region": info.get("region", ""), "cities": {}}
        if info.get("region"):
            data[country]["region"] = info["region"]
        data[country]["cities"].update(info.get("cities", {}))
    return data


def _err_detail(exc: requests.HTTPError) -> str:
    """Pull Google's human-readable error message out of an HTTPError."""
    try:
        return exc.response.json().get("error", {}).get("message", str(exc))
    except Exception:
        return str(exc)


def search_text(session, api_key, query, lat, lng, radius_m, region_code,
                page_token=None, timeout=30):
    """One Text Search (New) request. Returns (places_list, next_page_token)."""
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": FIELD_MASK,
    }
    body = {
        "textQuery": query,
        "languageCode": "ar" if is_arabic(query) else "en",
        "locationBias": {
            "circle": {
                "center": {"latitude": lat, "longitude": lng},
                "radius": float(radius_m),
            }
        },
    }
    if region_code:
        body["regionCode"] = region_code
    if page_token:
        body["pageToken"] = page_token

    resp = session.post(PLACES_ENDPOINT, headers=headers, json=body, timeout=timeout)
    resp.raise_for_status()
    data = resp.json()
    return data.get("places", []), data.get("nextPageToken")


def search_with_retry(session, api_key, query, lat, lng, radius_m, region_code,
                      page_token, delay):
    """search_text with a single back-off retry on 429 (rate limit)."""
    try:
        return search_text(session, api_key, query, lat, lng, radius_m,
                           region_code, page_token)
    except requests.HTTPError as exc:
        if exc.response is not None and exc.response.status_code == 429:
            time.sleep(max(delay * 4, 2.0))
            return search_text(session, api_key, query, lat, lng, radius_m,
                               region_code, page_token)
        raise


def run_scrape(api_key, selections, keywords, radius_m, max_pages, delay,
               progress_cb=None, session=None, fetch_emails=False, email_timeout=12):
    """
    Loop over (city x keyword), paginate, dedup on place_id.

    Returns (results_dict, errors_list). results_dict is {place_id: row}.
    progress_cb(fraction: float, message: str) is called after each query.
    """
    session = session or requests.Session()
    results, errors = {}, []
    api_calls = 0  # actual billable Places requests that succeeded
    total = max(len(selections) * len(keywords), 1)
    done = 0

    for (country, region, city, lat, lng) in selections:
        for kw in keywords:
            query = f"{kw} {city}"
            page_token = None
            for _ in range(max_pages):
                try:
                    places, page_token = search_with_retry(
                        session, api_key, query, lat, lng, radius_m,
                        region, page_token, delay
                    )
                    api_calls += 1
                except requests.HTTPError as exc:
                    errors.append(f"{query}: HTTP {exc.response.status_code if exc.response else '?'} — {_err_detail(exc)}")
                    break
                except Exception as exc:  # network / timeout / parse
                    errors.append(f"{query}: {exc}")
                    break

                for p in places:
                    pid = p.get("id")
                    if pid and pid not in results:
                        results[pid] = parse_place(p, city, country)

                if not page_token:
                    break
                time.sleep(delay)  # small pause before next page

            done += 1
            if progress_cb:
                progress_cb(done / total, f"{city} · {kw} — {len(results)} unique so far")
            time.sleep(delay)

    # --- Optional second phase: best-effort email lookup from websites ----- #
    if fetch_emails and results:
        web = requests.Session()
        rows = [r for r in results.values() if r.get("website")]
        for i, row in enumerate(rows, 1):
            row["email"] = fetch_emails_for_site(web, row["website"], email_timeout)
            if progress_cb:
                progress_cb(min(i / len(rows), 1.0),
                            f"Emails: {i}/{len(rows)} sites checked")

    return results, errors, api_calls


# --- Email extraction (clinic websites; Places API has no email field) ----- #

_UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")

_EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
_MAILTO_RE = re.compile(r"mailto:([^\"'>?\s]+)", re.I)

# Skip tracking libs, image sprites (logo@2x.png), placeholders, etc.
_EMAIL_JUNK = ("@sentry", "@wix", "wixpress", "@example", "@domain", "@email.",
               "@2x", "@3x", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg")

# Pages to probe, in order; stops as soon as any email is found.
_CONTACT_PATHS = ["", "/contact", "/contact-us", "/contactus", "/about", "/ar/contact"]


def extract_emails_from_html(html: str) -> set:
    """Pull clean email addresses out of a page's HTML."""
    found = set()
    for m in _MAILTO_RE.findall(html):
        found.add(m.split("?")[0].strip().lower())
    for m in _EMAIL_RE.findall(html):
        found.add(m.strip().lower())
    return {
        e for e in found
        if e and e.count("@") == 1 and len(e) <= 100
        and not any(j in e for j in _EMAIL_JUNK)
    }


def fetch_emails_for_site(session, website: str, timeout: int = 12) -> str:
    """Best-effort: probe a site's homepage + common contact pages for emails."""
    if not website:
        return ""
    base = website.rstrip("/")
    for path in _CONTACT_PATHS:
        url = base + path if path else base
        try:
            r = session.get(url, timeout=timeout,
                            headers={"User-Agent": _UA}, allow_redirects=True)
        except Exception:
            continue
        if r.status_code == 200 and r.text:
            emails = extract_emails_from_html(r.text)
            if emails:
                return ", ".join(sorted(emails))
    return ""


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Clinics")
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# Streamlit UI
# --------------------------------------------------------------------------- #

def main():
    st.set_page_config(page_title="Derma Scope — Clinic Scraper",
                       page_icon="🩺", layout="wide")
    st.title("🩺 Derma Scope — Dermatology Clinic Scraper")
    st.caption("Pulls dermatology clinics from Google Places API (New) across the "
               "countries you choose. One-time export to Excel / CSV.")

    with st.sidebar:
        st.header("⚙️ Settings")
        api_key = st.text_input(
            "Google Places API key", type="password",
            value=os.environ.get("GOOGLE_MAPS_API_KEY", ""),
            help="Needs **Places API (New)** enabled in Google Cloud.",
        )
        radius_km = st.slider("Search radius per city (km)", 5, 50, 30)
        max_pages = st.slider("Max pages per query (×20 results)", 1, 3, 3)
        delay = st.slider("Delay between calls (sec)", 0.0, 2.0, 0.3, 0.1,
                          help="A little pause keeps you under the QPS limit.")
        st.divider()
        fetch_emails = st.checkbox(
            "🔎 Also find emails (slower)",
            help="Visits each clinic's website to pull a public email. "
                 "Best-effort — many sites won't list one. Roughly doubles runtime.",
        )

    # --- 1. Regions -------------------------------------------------------- #
    st.subheader("1 · Choose regions")

    # Add-your-own UI (writes to session_state, which survives reruns)
    with st.expander("➕ Add a country or city that isn't listed"):
        st.caption("Tip: get coordinates from Google Maps — right-click a spot → "
                   "click the lat, lng at the top to copy.")

        st.markdown("**Add a country**")
        a1, a2, a3 = st.columns([3, 2, 1.4])
        nc_name = a1.text_input("Country name", key="nc_name")
        nc_region = a2.text_input("ISO code (e.g. EG, SA, JO)", key="nc_region", max_chars=2)
        if a3.button("Add country", key="nc_btn"):
            if nc_name.strip():
                cd = st.session_state.setdefault("custom_data", {})
                cd.setdefault(nc_name.strip(), {"region": nc_region.strip().upper(), "cities": {}})
                st.rerun()

        st.markdown("**Add a city**")
        known = list(merge_catalogue(build_region_catalogue(), st.session_state.get("custom_data", {})).keys())
        c1, c2, c3, c4, c5 = st.columns([2, 2, 1.3, 1.3, 1.2])
        ct_country = c1.selectbox("Country", known, key="nc_city_country")
        ct_city = c2.text_input("City name", key="nc_city_name")
        ct_lat = c3.number_input("Lat", value=0.0, format="%.4f", key="nc_city_lat")
        ct_lng = c4.number_input("Lng", value=0.0, format="%.4f", key="nc_city_lng")
        if c5.button("Add city", key="nc_city_btn"):
            if ct_country and ct_city.strip():
                cd = st.session_state.setdefault("custom_data", {})
                region = merge_catalogue(build_region_catalogue(), cd)[ct_country]["region"]
                entry = cd.setdefault(ct_country, {"region": region, "cities": {}})
                entry["cities"][ct_city.strip()] = (float(ct_lat), float(ct_lng))
                st.rerun()

        if st.session_state.get("custom_data"):
            if st.button("🗑️ Clear my added countries/cities", key="clear_custom"):
                st.session_state["custom_data"] = {}
                st.rerun()

    catalogue = merge_catalogue(build_region_catalogue(), st.session_state.get("custom_data", {}))
    countries = st.multiselect(
        "Countries", list(catalogue.keys()),
        default=[c for c in ["Egypt"] if c in catalogue],
        help="Grouped Gulf → wider Middle East → North Africa.",
    )

    # Quick-fill helpers — these set the per-country selections explicitly,
    # which avoids accidentally running every city in the region at once.
    if countries:
        st.caption("Quick-fill cities for the selected countries:")
        q1, q2, q3, q4 = st.columns(4)
        top_n = q4.number_input("Top N", 1, 100, 10, label_visibility="collapsed")
        if q1.button(f"⭐ Top {int(top_n)} per country"):
            for c in countries:
                st.session_state[f"cities_{c}"] = list(catalogue[c]["cities"])[:int(top_n)]
            st.rerun()
        if q2.button("🌍 ALL cities (heavy)"):
            for c in countries:
                st.session_state[f"cities_{c}"] = list(catalogue[c]["cities"])
            st.rerun()
        if q3.button("✖️ Clear cities"):
            for c in countries:
                st.session_state[f"cities_{c}"] = []
            st.rerun()

    city_selections = {}
    for c in countries:
        opts = list(catalogue[c]["cities"].keys())
        # seed a sensible default (top 10) the first time a country appears
        if f"cities_{c}" not in st.session_state:
            st.session_state[f"cities_{c}"] = opts[:10]
        city_selections[c] = st.multiselect(
            f"{c} — cities ({len(opts)} available)", opts, key=f"cities_{c}"
        )

    # --- 2. Keywords ------------------------------------------------------- #
    st.subheader("2 · Search keywords")
    col_en, col_ar = st.columns(2)
    en_text = col_en.text_area("English (one per line)",
                               "\n".join(DEFAULT_KEYWORDS_EN), height=190)
    ar_text = col_ar.text_area("Arabic (one per line)",
                               "\n".join(DEFAULT_KEYWORDS_AR), height=190)
    keywords = [k.strip() for k in (en_text + "\n" + ar_text).splitlines() if k.strip()]

    # --- Estimate + Run ---------------------------------------------------- #
    selections = build_selections(city_selections, catalogue)
    est_calls = len(selections) * len(keywords) * max_pages
    st.subheader("3 · Run")
    st.info(f"{len(selections)} cities × {len(keywords)} keywords × up to {max_pages} pages "
            f"≈ **{est_calls}** API calls max (fewer in practice — most queries return <60 results).")

    can_run = bool(api_key) and bool(selections) and bool(keywords)
    if not can_run:
        st.warning("Add an API key, at least one city, and at least one keyword to enable the run.")

    if st.button("🚀 Start scraping", type="primary", disabled=not can_run):
        progress = st.progress(0.0)
        status = st.empty()

        def cb(frac, msg):
            progress.progress(min(frac, 1.0))
            status.write(msg)

        with st.spinner("Scraping Google Places…"):
            results, errors, api_calls = run_scrape(
                api_key, selections, keywords, radius_km * 1000,
                max_pages, delay, progress_cb=cb, fetch_emails=fetch_emails,
            )
        st.session_state["df"] = pd.DataFrame(list(results.values()))
        st.session_state["errors"] = errors
        st.session_state["api_calls"] = api_calls

    # --- Results ----------------------------------------------------------- #
    if "df" in st.session_state:
        df = st.session_state["df"]
        st.success(f"✅ Done — {len(df)} unique clinics.")

        calls = st.session_state.get("api_calls", 0)
        cost = calls * COST_PER_CALL_USD
        st.markdown(
            f"💸 **Cost of this run: ~\\${cost:.2f}** · "
            f"{calls} Places API calls @ \\${COST_PER_CALL_USD:.3f} each · "
            f"email lookups are free *(counts against your ~\\$200/month free credit)*."
        )

        errs = st.session_state.get("errors") or []
        if errs:
            with st.expander(f"⚠️ {len(errs)} warnings / errors"):
                for e in errs:
                    st.write("•", e)

        if not df.empty:
            st.dataframe(df, use_container_width=True, hide_index=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            d1, d2 = st.columns(2)
            d1.download_button(
                "⬇️ Download Excel", to_excel_bytes(df),
                f"derma_clinics_{ts}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            # utf-8-sig so Arabic renders correctly when opened in Excel
            d2.download_button(
                "⬇️ Download CSV", df.to_csv(index=False).encode("utf-8-sig"),
                f"derma_clinics_{ts}.csv", "text/csv",
            )


if __name__ == "__main__":
    main()