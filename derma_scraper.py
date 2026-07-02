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

import storage

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
    """Flatten one Places API result into a clean row.

    The commented-out fields below are intentionally excluded from the export.
    To bring any back, un-comment its line — it's already requested in
    FIELD_MASK, so the column will populate automatically (no extra cost, since
    rating/phone/website already set the billing tier). If you re-enable lat/lng,
    also un-comment the `loc = ...` line just above the return.
    """
    # loc = p.get("location", {}) or {}              # needed only for lat/lng
    return {
        "place_id": p.get("id", ""),  # kept internally for dedup + resume; not exported
        "name": (p.get("displayName") or {}).get("text", ""),
        # "types": ", ".join(p.get("types", []) or []),   # removed
        "address": p.get("formattedAddress", ""),
        "city": city,
        "country": country,
        # "lat": loc.get("latitude"),                # removed
        # "lng": loc.get("longitude"),               # removed
        # "maps_url": p.get("googleMapsUri", ""),    # removed
        "website": p.get("websiteUri", ""),
        "phone": p.get("internationalPhoneNumber", ""),
        "email": "",  # filled later (best-effort) from the clinic website
        "rating": p.get("rating"),
        "review_count": p.get("userRatingCount"),
        # "business_status": p.get("businessStatus", ""),  # removed
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


def _request_with_retry(fn, *, retries=3, base_delay=1.0, max_backoff=20.0,
                        retryable_statuses=(429, 500, 502, 503, 504)):
    """Run fn() with exponential-backoff retries on rate limits, 5xx, and
    transient network errors (timeouts, dropped connections). Any other
    HTTPError, or the final attempt's exception, propagates to the caller."""
    last_exc = None
    for attempt in range(retries + 1):
        try:
            return fn()
        except requests.HTTPError as exc:
            status = exc.response.status_code if exc.response is not None else None
            if status not in retryable_statuses or attempt == retries:
                raise
            last_exc = exc
        except (requests.Timeout, requests.ConnectionError) as exc:
            if attempt == retries:
                raise
            last_exc = exc
        time.sleep(min(base_delay * (2 ** attempt), max_backoff))
    raise last_exc


def search_with_retry(session, api_key, query, lat, lng, radius_m, region_code,
                      page_token, delay):
    """search_text with exponential-backoff retries on 429/5xx/network errors."""
    return _request_with_retry(
        lambda: search_text(session, api_key, query, lat, lng, radius_m,
                            region_code, page_token)
    )


def run_scrape(api_key, selections, keywords, radius_m, max_pages, delay, run_id,
               progress_cb=None, session=None, fetch_emails=False, email_timeout=12,
               completed_queries=None, seen_place_ids=None):
    """
    Loop over (city x keyword), paginate, dedup on place_id.

    Every checkpoint below is written to storage.py immediately, not batched
    to the end — so a crash, timeout, or closed tab loses at most the one
    in-flight query or website, never the whole run. A (city, keyword) pair
    is only marked complete once its pagination finishes without error; on
    any failure it's left pending so a later Resume retries exactly that
    pair instead of silently skipping it forever.

    completed_queries / seen_place_ids: pass the sets loaded from storage.py
    when resuming a run, so already-done work is skipped. Pass None (or
    empty sets) for a fresh run.

    Returns errors_list. Results and the running api_calls count live in
    storage.py, not in a returned dict — the caller reads them back via
    storage.get_run_clinics() / storage.get_run().
    progress_cb(fraction: float, message: str) is called after each query.
    """
    session = session or requests.Session()
    completed_queries = set(completed_queries or ())
    seen_place_ids = set(seen_place_ids or ())
    errors = []
    total = max(len(selections) * len(keywords), 1)
    done = len(completed_queries)

    for (country, region, city, lat, lng) in selections:
        for kw in keywords:
            if (city, kw) in completed_queries:
                continue

            query = f"{kw} {city}"
            page_token = None
            new_rows = []
            query_failed = False
            query_calls = 0
            for _ in range(max_pages):
                try:
                    places, page_token = search_with_retry(
                        session, api_key, query, lat, lng, radius_m,
                        region, page_token, delay
                    )
                    query_calls += 1
                except requests.HTTPError as exc:
                    msg = f"{query}: HTTP {exc.response.status_code if exc.response else '?'} — {_err_detail(exc)}"
                    errors.append(msg)
                    storage.log_error(run_id, msg)
                    query_failed = True
                    break
                except Exception as exc:  # network / timeout / parse
                    msg = f"{query}: {exc}"
                    errors.append(msg)
                    storage.log_error(run_id, msg)
                    query_failed = True
                    break

                for p in places:
                    pid = p.get("id")
                    if pid and pid not in seen_place_ids:
                        seen_place_ids.add(pid)
                        new_rows.append(parse_place(p, city, country))

                if not page_token:
                    break
                time.sleep(delay)  # small pause before next page

            if new_rows:
                storage.save_clinics_batch(run_id, new_rows)
            if not query_failed:
                storage.mark_query_complete(run_id, city, kw)
                completed_queries.add((city, kw))
            storage.touch_run(run_id, api_calls_delta=query_calls)

            done += 1
            if progress_cb:
                progress_cb(done / total, f"{city} · {kw} — {len(seen_place_ids)} unique so far")
            time.sleep(delay)

    # --- Optional second phase: best-effort email lookup from websites ----- #
    if fetch_emails:
        web = requests.Session()
        pending = storage.get_pending_email_rows(run_id)
        for i, row in enumerate(pending, 1):
            try:
                email, reachable = fetch_emails_for_site(web, row["website"], email_timeout)
            except Exception as exc:  # unexpected parsing bug etc — don't lose the whole loop
                msg = f"{row['website']}: {exc}"
                errors.append(msg)
                storage.log_error(run_id, msg)
                email, reachable = "", True  # treat as checked so it isn't retried forever

            if reachable:
                storage.update_email_result(run_id, row["place_id"], email)
            else:
                msg = f"{row['website']}: unreachable, will retry on resume"
                errors.append(msg)
                storage.log_error(run_id, msg)

            if progress_cb:
                progress_cb(min(i / len(pending), 1.0) if pending else 1.0,
                            f"Emails: {i}/{len(pending)} sites checked")

    return errors


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


def fetch_emails_for_site(session, website: str, timeout: int = 12):
    """Best-effort: probe a site's homepage + common contact pages for emails.

    Returns (email, reachable). reachable is False only when every single
    contact path failed to connect at all (timeout, DNS, connection reset —
    e.g. the employee's internet dropped), so the caller can leave the site
    unchecked and retry it later instead of wrongly recording "no email
    found" for a site that was never actually reached.
    """
    if not website:
        return "", True
    base = website.rstrip("/")
    any_reachable = False
    for path in _CONTACT_PATHS:
        url = base + path if path else base
        try:
            r = _request_with_retry(
                lambda: session.get(url, timeout=timeout,
                                    headers={"User-Agent": _UA}, allow_redirects=True),
                retries=1, base_delay=1.5,
            )
        except Exception:
            continue
        any_reachable = True
        if r.status_code == 200 and r.text:
            emails = extract_emails_from_html(r.text)
            if emails:
                return ", ".join(sorted(emails)), True
    return "", any_reachable


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

    storage.init_db()

    # --- Previous / in-progress runs ---------------------------------------- #
    # Every clinic and email is saved as it's found (see run_scrape), so a
    # dropped connection, timeout, or closed tab never loses scraped data —
    # it's always recoverable here.
    resumable = storage.list_resumable_runs()
    if resumable:
        with st.expander(f"📂 {len(resumable)} in-progress run(s) found — resume instead of starting over?",
                         expanded=not st.session_state.get("active_run_id")):
            for r in resumable:
                emails_total = r["emails_done"] + r["emails_pending"]
                line = (f"**{r['label']}** — started by *{r['owner'] or 'unknown'}* · "
                       f"{r['clinics_found']} clinics found · "
                       f"{r['queries_done']}/{r['total_queries']} queries done")
                if emails_total:
                    line += f" · {r['emails_done']}/{emails_total} emails checked"
                line += f" · last updated {r['updated_at']}"
                st.write(line)
                b1, b2 = st.columns([1, 1])
                if b1.button("📂 Load this run", key=f"load_{r['id']}"):
                    st.session_state["active_run_id"] = r["id"]
                    st.rerun()
                if b2.button("🗑️ Cancel (keep data)", key=f"cancel_{r['id']}"):
                    storage.cancel_run(r["id"])
                    if st.session_state.get("active_run_id") == r["id"]:
                        del st.session_state["active_run_id"]
                    st.rerun()

    if st.session_state.get("active_run_id"):
        st.info(f"📌 Run #{st.session_state['active_run_id']} is loaded — click **Resume scraping** "
                "below to continue it with its original city/keyword selections, or scroll down to "
                "view/download what's saved so far.")
        if st.button("➕ Start a fresh run instead"):
            del st.session_state["active_run_id"]
            st.rerun()

    with st.sidebar:
        st.header("⚙️ Settings")
        api_key = st.text_input(
            "Google Places API key", type="password",
            value=os.environ.get("GOOGLE_MAPS_API_KEY", ""),
            help="Needs **Places API (New)** enabled in Google Cloud.",
        )
        owner = st.text_input(
            "Started by (your name)",
            help="Shown in the in-progress runs list above so you (and others "
                 "sharing this app) can find your own runs.",
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

    active_run_id = st.session_state.get("active_run_id")
    active_run = storage.get_run(active_run_id) if active_run_id else None
    if active_run is None and active_run_id:
        del st.session_state["active_run_id"]

    # Only treat this as a resume while the loaded run is still unfinished —
    # once it's completed, the button should start a fresh run from whatever
    # the sidebar/form are currently set to, not silently no-op on old params.
    resuming = bool(active_run) and active_run["status"] == "running"
    can_run = bool(api_key) and (resuming or (bool(selections) and bool(keywords)))
    if not can_run:
        st.warning("Add an API key, at least one city, and at least one keyword to enable the run.")

    button_label = "▶️ Resume scraping" if resuming else "🚀 Start scraping"
    if st.button(button_label, type="primary", disabled=not can_run):
        progress = st.progress(0.0)
        status = st.empty()

        def cb(frac, msg):
            progress.progress(min(frac, 1.0))
            status.write(msg)

        if resuming:
            run_id = active_run_id
            p = active_run["params"]
            run_selections, run_keywords = p["selections"], p["keywords"]
            run_radius_m, run_max_pages, run_delay = p["radius_m"], p["max_pages"], p["delay"]
            run_email_timeout = p.get("email_timeout", 12)
            # Frozen from the original run, not the sidebar's current, unkeyed
            # checkbox state — that resets to unchecked on a fresh session (the
            # exact case Resume exists for), which would otherwise silently skip
            # the rest of the email step and mark the run complete regardless.
            run_fetch_emails = p.get("fetch_emails", False)
            completed = storage.get_completed_queries(run_id)
            seen_ids = storage.get_existing_place_ids(run_id)
        else:
            run_selections, run_keywords = selections, keywords
            run_radius_m, run_max_pages, run_delay = radius_km * 1000, max_pages, delay
            run_email_timeout = 12
            run_fetch_emails = fetch_emails
            label = (f"{len(selections)} cities × {len(keywords)} keywords — "
                    f"{datetime.now().strftime('%Y-%m-%d %H:%M')}")
            params = {
                "selections": run_selections, "keywords": run_keywords,
                "radius_m": run_radius_m, "max_pages": run_max_pages,
                "delay": run_delay, "email_timeout": run_email_timeout,
                "fetch_emails": run_fetch_emails,
            }
            run_id = storage.create_run(owner, label, params,
                                        len(run_selections) * len(run_keywords))
            st.session_state["active_run_id"] = run_id
            completed, seen_ids = set(), set()

        with st.spinner("Scraping Google Places…"):
            run_scrape(
                api_key, run_selections, run_keywords, run_radius_m,
                run_max_pages, run_delay, run_id, progress_cb=cb,
                fetch_emails=run_fetch_emails, email_timeout=run_email_timeout,
                completed_queries=completed, seen_place_ids=seen_ids,
            )

        total_queries = len(run_selections) * len(run_keywords)
        all_queries_done = len(storage.get_completed_queries(run_id)) >= total_queries
        emails_done = not run_fetch_emails or not storage.get_pending_email_rows(run_id)
        if all_queries_done and emails_done:
            storage.mark_run_completed(run_id)

    # --- Results ----------------------------------------------------------- #
    active_run_id = st.session_state.get("active_run_id")
    run = storage.get_run(active_run_id) if active_run_id else None
    if run is None and active_run_id:
        del st.session_state["active_run_id"]

    if run:
        df = pd.DataFrame(storage.get_run_clinics(active_run_id))
        completed_n = len(storage.get_completed_queries(active_run_id))
        status_label = "✅ Done" if run["status"] == "completed" else "⏳ In progress"
        st.success(f"{status_label} — {len(df)} unique clinics · "
                  f"{completed_n}/{run['total_queries']} queries done.")

        calls = run["api_calls"]
        cost = calls * COST_PER_CALL_USD
        st.markdown(
            f"💸 **Cost so far: ~\\${cost:.2f}** · "
            f"{calls} Places API calls @ \\${COST_PER_CALL_USD:.3f} each · "
            f"email lookups are free *(counts against your ~\\$200/month free credit)*."
        )

        errs = storage.get_run_errors(active_run_id)
        if errs:
            with st.expander(f"⚠️ {len(errs)} warnings / errors"):
                for e in errs:
                    st.write("•", e)

        if not df.empty:
            all_cols = list(df.columns)
            show_cols = st.multiselect(
                "Columns to include (table + downloads)",
                all_cols, default=all_cols, key="show_cols",
                help="Untick any column to drop it from the table and the exports.",
            )
            cols = [c for c in all_cols if c in show_cols] or all_cols
            view = df[cols]

            st.dataframe(view, use_container_width=True, hide_index=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            d1, d2 = st.columns(2)
            d1.download_button(
                "⬇️ Download Excel", to_excel_bytes(view),
                f"derma_clinics_{ts}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            # utf-8-sig so Arabic renders correctly when opened in Excel
            d2.download_button(
                "⬇️ Download CSV", view.to_csv(index=False).encode("utf-8-sig"),
                f"derma_clinics_{ts}.csv", "text/csv",
            )


if __name__ == "__main__":
    main()