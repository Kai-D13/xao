import json
import os
import time
import openpyxl
import requests
from flask import Flask, jsonify, render_template, Response

app = Flask(__name__)

SERVICES_KEY = "e3464a9335a846e985861bdf43fd8700201a93af28006a40"
TILEMAP_KEY  = "06fadcaa43886a1b8a3fd81709a1f9723bb3e25d1010554b"
EXCEL_PATH   = os.path.join(os.path.dirname(__file__), "check_failed.xlsx")
CACHE_PATH   = os.path.join(os.path.dirname(__file__), "geocache.json")

# 2 Hub NVCT
HUB_CARRIERS = {
    "NVCT Hub Di Linh - Lâm Đồng - Child",
    "NVCT Hub Bảo Lộc_Child",
}

HUB_LOCATIONS = [
    {
        "name": "NVCT Hub Di Linh - Lâm Đồng - Child",
        "lat": 11.572849213854973,
        "lng": 108.04066512374126,
    },
    {
        "name": "NVCT Hub Bảo Lộc_Child",
        "lat": 11.541409083328261,
        "lng": 107.82235962024197,
    },
]


def load_cache():
    if os.path.exists(CACHE_PATH):
        with open(CACHE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(cache):
    with open(CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def geocode_address(address, cache):
    if address in cache:
        return cache[address]

    try:
        # Step 1: search/v3 to get ref_id
        r1 = requests.get(
            "https://maps.vietmap.vn/api/search/v3",
            params={"apikey": SERVICES_KEY, "text": address},
            timeout=10,
        )
        r1.raise_for_status()
        results = r1.json()
        if not isinstance(results, list) or len(results) == 0:
            print(f"  [WARN] No search result for: {address}")
            return None

        ref_id = results[0].get("ref_id")
        if not ref_id:
            print(f"  [WARN] No ref_id for: {address}")
            return None

        # Step 2: place/v3 to get lat/lng
        r2 = requests.get(
            "https://maps.vietmap.vn/api/place/v3",
            params={"apikey": SERVICES_KEY, "refid": ref_id},
            timeout=10,
        )
        r2.raise_for_status()
        place = r2.json()
        lat = place.get("lat")
        lng = place.get("lng")

        if lat and lng:
            coord = {"lat": float(lat), "lng": float(lng)}
            cache[address] = coord
            save_cache(cache)
            return coord

        print(f"  [WARN] No coords in place response for: {address}")
    except Exception as e:
        print(f"  [ERROR] Geocode failed for '{address}': {e}")

    return None


def load_data():
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active

    # Aggregate: key = (carrier_name, full_address), value = list of total_order_ward
    agg = {}
    for r in range(2, ws.max_row + 1):
        carrier = ws.cell(r, 1).value
        province = ws.cell(r, 3).value
        district = ws.cell(r, 4).value
        ward = ws.cell(r, 5).value
        total_order = ws.cell(r, 6).value
        full_address = ws.cell(r, 7).value

        if not full_address or not carrier:
            continue

        key = (carrier, full_address)
        if key not in agg:
            agg[key] = {
                "carrier_name": carrier,
                "full_address": full_address,
                "province_name": province,
                "district_name": district,
                "ward_name": ward,
                "orders": [],
            }
        if total_order is not None:
            agg[key]["orders"].append(float(total_order))

    return list(agg.values())


@app.route("/api/mapstyle")
def map_style():
    """Fetch VietMap style.json and patch glyphs URL to one that has Open Sans fonts."""
    r = requests.get(
        f"https://maps.vietmap.vn/api/maps/light/style.json?apikey={TILEMAP_KEY}",
        timeout=10,
    )
    style = r.json()
    # Replace glyphs URL — OpenMapTiles CDN has Open Sans, Roboto, Noto Sans all available
    style["glyphs"] = "https://fonts.openmaptiles.org/{fontstack}/{range}.pbf"
    return Response(
        json.dumps(style),
        content_type="application/json",
        headers={"Access-Control-Allow-Origin": "*"},
    )


@app.route("/")
def index():
    return render_template("index.html", tilemap_key=TILEMAP_KEY)


@app.route("/api/points")
def api_points():
    cache = load_cache()
    rows = load_data()
    points = []

    for item in rows:
        address = item["full_address"]
        coord = geocode_address(address, cache)
        if not coord:
            print(f"  [SKIP] No coord for: {address}")
            continue

        orders = item["orders"]
        avg_order = round(sum(orders) / len(orders), 1) if orders else 0
        is_hub = item["carrier_name"] in HUB_CARRIERS

        points.append({
            "lat": coord["lat"],
            "lng": coord["lng"],
            "carrier_name": item["carrier_name"],
            "full_address": address,
            "province_name": item["province_name"],
            "district_name": item["district_name"],
            "ward_name": item["ward_name"],
            "avg_order": avg_order,
            "is_hub_carrier": is_hub,
            "color": "#2563EB" if is_hub else "#DC2626",
        })
        time.sleep(0.15)  # Rate limit

    return jsonify({
        "points": points,
        "hubs": HUB_LOCATIONS,
    })


@app.route("/api/geocode-status")
def geocode_status():
    cache = load_cache()
    rows = load_data()
    total = len(rows)
    cached = sum(1 for r in rows if r["full_address"] in cache)
    return jsonify({"total": total, "geocoded": cached, "pending": total - cached})


if __name__ == "__main__":
    print("Starting mapping server at http://localhost:5000")
    app.run(debug=True, port=5000)
