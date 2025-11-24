import os
import time
import logging
from typing import Dict, Any, Optional
from urllib.parse import urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup



GOOGLE_PLACES_API_KEY = os.environ.get("GOOGLE_PLACES_API_KEY")

INPUT_EXCEL_PATH = "Book3.xlsx"
OUTPUT_EXCEL_PATH = "Book3_enriched.xlsx"

WIKIDATA_API_BASE = "https://www.wikidata.org/w/api.php"


ENABLE_GOOGLE = True
ENABLE_WIKIDATA = True
ENABLE_WEBSITE_SCRAPING = True

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

API_SLEEP_SECONDS = 0.3




def normalize_company_key(name: str, country_code: Optional[str]) -> str:
    if not isinstance(name, str):
        name = ""
    name_clean = name.strip().lower()
    country_clean = (country_code or "").strip().lower()
    return f"{name_clean}__{country_clean}"


def get_address_component(components, type_name, return_short=False):
    if not components:
        return None
    for comp in components:
        types = comp.get("types", [])
        if type_name in types:
            return comp.get("short_name" if return_short else "long_name")
    return None


def extract_domain(url: str):
    if not isinstance(url, str) or not url:
        return None
    if not (url.startswith("http://") or url.startswith("https://")):
        url = "http://" + url
    try:
        netloc = urlparse(url).netloc
        if netloc.startswith("www."):
            netloc = netloc[4:]
        return netloc or None
    except Exception:
        return None


def extract_tld(domain: str):
    if not isinstance(domain, str) or not domain:
        return None
    parts = domain.split(".")
    if len(parts) < 2:
        return None
    return parts[-1]


def is_empty(val) -> bool:
    return pd.isna(val) or val == ""


def clean_for_excel(value):
    """
    Remove characters that Excel (openpyxl) considers illegal.
    Keep normal whitespace (tab, newline, carriage return) and printable chars.
    """
    if not isinstance(value, str):
        return value
    return "".join(
        ch for ch in value
        if ch in ("\t", "\n", "\r") or ord(ch) >= 32
    )


# -----------------------------
# 2. GOOGLE PLACES API CLIENT
# -----------------------------

def google_places_details(place_id: str) -> Dict[str, Any]:
    if not GOOGLE_PLACES_API_KEY or not ENABLE_GOOGLE:
        return {}

    url = "https://maps.googleapis.com/maps/api/place/details/json"
    params = {
        "place_id": place_id,
        "key": GOOGLE_PLACES_API_KEY,
        "fields": "name,formatted_address,formatted_phone_number,website,address_components,geometry"
    }
    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        logging.error(f"Google Place Details failed for '{place_id}': {e}")
        return {}


def google_places_search(company_name: str,
                         country: Optional[str],
                         city: Optional[str]) -> Dict[str, Any]:
    if not GOOGLE_PLACES_API_KEY or not ENABLE_GOOGLE:
        return {}

    query_parts = []
    if isinstance(company_name, str) and company_name.strip():
        query_parts.append(company_name.strip())
    if isinstance(city, str) and city.strip():
        query_parts.append(city.strip())
    if isinstance(country, str) and country.strip():
        query_parts.append(country.strip())

    if not query_parts:
        return {}

    query = ", ".join(query_parts)
    params = {"query": query, "key": GOOGLE_PLACES_API_KEY}
    url = "https://maps.googleapis.com/maps/api/place/textsearch/json"

    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
    except Exception as e:
        logging.error(f"Google Places request failed for '{company_name}': {e}")
        return {}

    data = resp.json()
    results = data.get("results", [])
    if not results:
        return {}

    top = results[0]
    place_id = top.get("place_id")

    enriched = {
        "name": top.get("name"),
        "formatted_address": top.get("formatted_address"),
        "website": None,
        "phone": None,
        "country_long": None,
        "country_short": None,
        "region": None,
        "city": None,
        "postcode": None,
        "street": None,
        "street_number": None,
        "lat": None,
        "lng": None,
    }

    if not place_id:
        return enriched

    details = google_places_details(place_id)
    result = details.get("result", {}) if isinstance(details, dict) else {}

    enriched["website"] = result.get("website")
    enriched["phone"] = result.get("formatted_phone_number")
    enriched["name"] = result.get("name") or enriched["name"]

    addr_components = result.get("address_components", [])
    geometry = result.get("geometry", {})

    enriched["country_long"] = get_address_component(addr_components, "country", return_short=False)
    enriched["country_short"] = get_address_component(addr_components, "country", return_short=True)
    enriched["region"] = get_address_component(addr_components, "administrative_area_level_1", return_short=False)
    city_locality = get_address_component(addr_components, "locality", return_short=False)
    city_postal_town = get_address_component(addr_components, "postal_town", return_short=False)
    enriched["city"] = city_locality or city_postal_town
    enriched["postcode"] = get_address_component(addr_components, "postal_code", return_short=False)
    enriched["street"] = get_address_component(addr_components, "route", return_short=False)
    enriched["street_number"] = get_address_component(addr_components, "street_number", return_short=False)

    location = geometry.get("location", {})
    enriched["lat"] = location.get("lat")
    enriched["lng"] = location.get("lng")

    return enriched




def wikidata_basic(company_name: str) -> Dict[str, Any]:
    if not company_name or not ENABLE_WIKIDATA:
        return {}

    params = {
        "action": "wbsearchentities",
        "search": company_name,
        "language": "en",
        "format": "json",
        "type": "item",
        "limit": 1,
    }
    headers = {"User-Agent": "CompanyEnrichmentScript/1.0 (contact: example@example.com)"}

    try:
        resp = requests.get(WIKIDATA_API_BASE, params=params, headers=headers, timeout=10)
        resp.raise_for_status()
    except Exception as e:
        logging.error(f"Wikidata search failed for '{company_name}': {e}")
        return {}

    data = resp.json()
    results = data.get("search", [])
    if not results:
        return {}

    qid = results[0].get("id")
    if not qid:
        return {}

    params = {
        "action": "wbgetentities",
        "ids": qid,
        "format": "json",
        "languages": "en",
        "props": "descriptions|claims",
    }

    try:
        resp = requests.get(WIKIDATA_API_BASE, params=params, headers=headers, timeout=10)
        resp.raise_for_status()
    except Exception as e:
        logging.error(f"Wikidata entity fetch failed for '{company_name}' ({qid}): {e}")
        return {}

    data = resp.json()
    entity = data.get("entities", {}).get(qid, {})
    claims = entity.get("claims", {})
    desc_en = entity.get("descriptions", {}).get("en", {}).get("value")

    def get_first_claim(prop):
        try:
            return claims[prop][0]["mainsnak"]["datavalue"]["value"]
        except Exception:
            return None

    year_founded = None
    v_inception = get_first_claim("P571")
    if isinstance(v_inception, dict):
        time_str = v_inception.get("time")
        if isinstance(time_str, str) and len(time_str) >= 5 and time_str[1:5].isdigit():
            year_founded = int(time_str[1:5])

    employee_count = None
    v_employees = get_first_claim("P1128")
    if isinstance(v_employees, dict):
        amount = v_employees.get("amount")
        if isinstance(amount, str):
            try:
                employee_count = int(float(amount))
            except Exception:
                pass

    return {
        "year_founded": year_founded,
        "employee_count": employee_count,
        "wd_description": desc_en,
    }


# -----------------------------
# 4. WEBSITE SCRAPER
# -----------------------------

def scrape_website_for_contacts(url: str) -> Dict[str, Optional[str]]:
    if not ENABLE_WEBSITE_SCRAPING:
        return {}

    if not isinstance(url, str) or not url:
        return {}

    if not url.startswith("http://") and not url.startswith("https://"):
        url = "http://" + url

    try:
        resp = requests.get(url, timeout=5)
    except Exception as e:
        logging.warning(f"Website scrape failed for '{url}': {e}")
        return {}

    if resp.status_code >= 400:
        logging.warning(f"Website scrape failed for '{url}': HTTP {resp.status_code}")
        return {}

    html = resp.text
    soup = BeautifulSoup(html, "html.parser")

    result = {
        "email": None,
        "facebook_url": None,
        "linkedin_url": None,
        "twitter_url": None,
        "instagram_url": None,
        "youtube_url": None,
        "meta_description": None,
    }

    meta_desc = soup.find("meta", attrs={"name": "description"})
    if meta_desc and meta_desc.get("content"):
        result["meta_description"] = meta_desc["content"].strip()

    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href.startswith("mailto:") and not result["email"]:
            email = href.split("mailto:")[1].split("?")[0]
            result["email"] = email

        if "facebook.com" in href and not result["facebook_url"]:
            result["facebook_url"] = href
        elif "linkedin.com" in href and not result["linkedin_url"]:
            result["linkedin_url"] = href
        elif ("twitter.com" in href or "x.com" in href) and not result["twitter_url"]:
            result["twitter_url"] = href
        elif "instagram.com" in href and not result["instagram_url"]:
            result["instagram_url"] = href
        elif "youtube.com" in href and not result["youtube_url"]:
            result["youtube_url"] = href

    return result


# -----------------------------
# 5. MAIN PIPELINE
# -----------------------------

def run_enrichment():
    logging.info(f"Loading input from {INPUT_EXCEL_PATH}")
    df = pd.read_excel(INPUT_EXCEL_PATH)

    text_cols = [
        "company_name",
        "main_country",
        "main_country_code",
        "main_region",
        "main_city",
        "main_postcode",
        "main_street",
        "main_street_number",
        "primary_phone",
        "website_url",
        "website_domain",
        "website_tld",
        "generated_description",
        "generated_business_tags",
        "short_description",
        "long_description",
        "business_tags",
        "email",
        "emails",
        "facebook_url",
        "facebook",
        "linkedin_url",
        "linkedin",
        "twitter_url",
        "twitter",
        "instagram_url",
        "instagram",
        "youtube_url",
        "youtube",
    ]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype("object")

    numeric_cols = ["year_founded", "employee_count"]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    def resolve_col(candidates):
        for c in candidates:
            if c in df.columns:
                return c
        return None

    email_col = resolve_col(["email", "emails"])
    facebook_col = resolve_col(["facebook_url", "facebook"])
    linkedin_col = resolve_col(["linkedin_url", "linkedin"])
    twitter_col = resolve_col(["twitter_url", "twitter"])
    instagram_col = resolve_col(["instagram_url", "instagram"])
    youtube_col = resolve_col(["youtube_url", "youtube"])

    short_desc_col = resolve_col(["generated_description", "short_description"])
    long_desc_col = resolve_col(["long_description"])

    gp_cache: Dict[str, Dict[str, Any]] = {}
    wd_cache: Dict[str, Dict[str, Any]] = {}
    website_cache: Dict[str, Dict[str, Any]] = {}

    for idx, row in df.iterrows():
        company_name = row.get("input_company_name")
        country = row.get("input_main_country")
        country_code = row.get("input_main_country_code")
        city = row.get("input_main_city")

        key = normalize_company_key(company_name, country_code)
        if not key:
            continue

        def maybe_fill(col, value):
            if col not in df.columns:
                return
            if value is None or value == "":
                return
            current = df.at[idx, col]
            if is_empty(current):
                clean_val = clean_for_excel(str(value))
                if df[col].dtype != "object":
                    df[col] = df[col].astype("object")
                df.at[idx, col] = clean_val

        # ---- Google ----
        if ENABLE_GOOGLE:
            if key not in gp_cache:
                gp_data = google_places_search(company_name, country, city)
                gp_cache[key] = gp_data
                time.sleep(API_SLEEP_SECONDS)
            else:
                gp_data = gp_cache[key]
        else:
            gp_data = {}

        if gp_data:
            maybe_fill("company_name", gp_data.get("name"))
            maybe_fill("main_country", gp_data.get("country_long"))
            maybe_fill("main_country_code", gp_data.get("country_short"))
            maybe_fill("main_region", gp_data.get("region"))
            maybe_fill("main_city", gp_data.get("city"))
            maybe_fill("main_postcode", gp_data.get("postcode"))
            maybe_fill("main_street", gp_data.get("street"))
            maybe_fill("main_street_number", gp_data.get("street_number"))
            maybe_fill("main_latitude", gp_data.get("lat"))
            maybe_fill("main_longitude", gp_data.get("lng"))
            maybe_fill("primary_phone", gp_data.get("phone"))

            website = gp_data.get("website")
            if website:
                maybe_fill("website_url", website)
                domain = extract_domain(website)
                if domain:
                    maybe_fill("website_domain", domain)
                    tld = extract_tld(domain)
                    if tld:
                        maybe_fill("website_tld", tld)

        # ---- Wikidata ----
        if ENABLE_WIKIDATA:
            if key not in wd_cache:
                wd_data = wikidata_basic(company_name)
                wd_cache[key] = wd_data
                time.sleep(API_SLEEP_SECONDS)
            else:
                wd_data = wd_cache[key]
        else:
            wd_data = {}

        if wd_data:
            yf = wd_data.get("year_founded")
            if "year_founded" in df.columns and yf is not None and is_empty(df.at[idx, "year_founded"]):
                df.at[idx, "year_founded"] = yf

            ec = wd_data.get("employee_count")
            if "employee_count" in df.columns and ec is not None and is_empty(df.at[idx, "employee_count"]):
                df.at[idx, "employee_count"] = ec

            desc = wd_data.get("wd_description")
            if desc:
                desc_clean = clean_for_excel(desc)
                if short_desc_col and is_empty(df.at[idx, short_desc_col]):
                    df.at[idx, short_desc_col] = desc_clean
                elif long_desc_col and is_empty(df.at[idx, long_desc_col]):
                    df.at[idx, long_desc_col] = desc_clean

        # ---- Website scraping ----
        website_val = df.at[idx, "website_url"] if "website_url" in df.columns else None
        if ENABLE_WEBSITE_SCRAPING and isinstance(website_val, str) and website_val:
            web_key = extract_domain(website_val) or website_val
            if web_key not in website_cache:
                contacts = scrape_website_for_contacts(website_val)
                website_cache[web_key] = contacts
                time.sleep(API_SLEEP_SECONDS)
            else:
                contacts = website_cache[web_key]

            if contacts:
                if email_col:
                    maybe_fill(email_col, contacts.get("email"))
                if facebook_col:
                    maybe_fill(facebook_col, contacts.get("facebook_url"))
                if linkedin_col:
                    maybe_fill(linkedin_col, contacts.get("linkedin_url"))
                if twitter_col:
                    maybe_fill(twitter_col, contacts.get("twitter_url"))
                if instagram_col:
                    maybe_fill(instagram_col, contacts.get("instagram_url"))
                if youtube_col:
                    maybe_fill(youtube_col, contacts.get("youtube_url"))

                meta_desc = contacts.get("meta_description")
                if meta_desc:
                    meta_desc_clean = clean_for_excel(meta_desc)
                    if short_desc_col and is_empty(df.at[idx, short_desc_col]):
                        df.at[idx, short_desc_col] = meta_desc_clean
                    elif long_desc_col and is_empty(df.at[idx, long_desc_col]):
                        df.at[idx, long_desc_col] = meta_desc_clean

    logging.info(f"Saving enriched output to {OUTPUT_EXCEL_PATH}")
    df.to_excel(OUTPUT_EXCEL_PATH, index=False)
    logging.info("All done 🎉")


if __name__ == "__main__":
    run_enrichment()
