"""
NL 511 Road Conditions & Events Extractor
==========================================
Fetches road conditions (winter roads) and traffic events from the
511nl.ca API (IBI511 platform) and exports to CSV, Excel, and KML.

SETUP:
  pip install requests pandas openpyxl

USAGE:
  python nl511_extract.py

OUTPUT:
  nl511_road_conditions.csv
  nl511_events.csv
  nl511_data.xlsx         (both datasets as separate sheets)
  nl511_roads.kml         (colour-coded map, open in Google Earth / Maps)
"""

import requests
import pandas as pd
import json
import sys
import os
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime

# ─────────────────────────────────────────────
#  CONFIGURATION — paste your API key here
# ─────────────────────────────────────────────
API_KEY = os.environ.get("NL511_API_KEY", "")          # <-- replace with your key
BASE_URL = "https://511nl.ca/api/v2/get"

# Output file paths (saved alongside this script)
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
OUT_EXCEL    = os.path.join(SCRIPT_DIR, "nl511_data.xlsx")
OUT_COND_CSV = os.path.join(SCRIPT_DIR, "nl511_road_conditions.csv")
OUT_EVT_CSV  = os.path.join(SCRIPT_DIR, "nl511_events.csv")
OUT_KML      = os.path.join(SCRIPT_DIR, "nl511_roads.kml")

# ─────────────────────────────────────────────
#  KML / POLYLINE HELPERS
# ─────────────────────────────────────────────

def decode_polyline(encoded: str) -> list:
    """
    Decode a Google Maps encoded polyline string into a list of (lat, lng) tuples.
    Pure Python — no extra library needed.
    """
    coords, index, lat, lng = [], 0, 0, 0
    while index < len(encoded):
        for is_lng in (False, True):
            result, shift = 0, 0
            while True:
                b = ord(encoded[index]) - 63
                index += 1
                result |= (b & 0x1F) << shift
                shift += 5
                if b < 0x20:
                    break
            delta = ~(result >> 1) if result & 1 else result >> 1
            if is_lng:
                lng += delta
            else:
                lat += delta
        coords.append((lat / 1e5, lng / 1e5))
    return coords


# KML colour map — AABBGGRR format (KML uses BGR, not RGB)
# Keyed on substrings that appear in the road_condition value (case-insensitive)
CONDITION_COLOURS = {
    "closed":               "ff0000ff",   # red      — most specific first
    "partly covered":       "ff00aaff",   # amber    — must come before "covered"
    "covered snow packed":  "ff0055ff",   # orange
    "compact snow":         "ff00ccff",   # yellow
    "icy":                  "ff0000cc",   # dark red
    "bare wet":             "ff44cc00",   # lime green
    "bare dry":             "ff000000",   # black
}
DEFAULT_COLOUR = "ff888888"   # grey for anything unrecognised


def condition_colour(condition: str) -> str:
    if not condition:
        return DEFAULT_COLOUR
    lower = condition.lower()
    for keyword, colour in CONDITION_COLOURS.items():
        if keyword in lower:
            return colour
    return DEFAULT_COLOUR


def build_kml(df_main: pd.DataFrame, df_poly: pd.DataFrame, label: str) -> ET.Element:
    """
    Build a KML <Document> element from a main dataframe and its polyline dataframe.
    Segments are colour-coded by road_condition (or event_type for events).
    """
    doc = ET.Element("Document")
    ET.SubElement(doc, "name").text = f"NL 511 — {label}"
    ET.SubElement(doc, "description").text = (
        f"Exported {datetime.now().strftime('%Y-%m-%d %H:%M')} from 511nl.ca"
    )

    # Write one Style per unique colour so the KML stays small
    used_colours = set(CONDITION_COLOURS.values()) | {DEFAULT_COLOUR}
    for colour in used_colours:
        style = ET.SubElement(doc, "Style", id=f"s_{colour}")
        line  = ET.SubElement(style, "LineStyle")
        ET.SubElement(line, "color").text = colour
        ET.SubElement(line, "width").text = "3"

    # Merge main data with polylines on id
    if df_poly.empty or "encoded_polyline" not in df_poly.columns:
        return doc

    merged = df_main.merge(df_poly, on="id", how="inner")

    condition_col = "road_condition" if "road_condition" in merged.columns else "event_type"

    for _, row in merged.iterrows():
        encoded = row.get("encoded_polyline", "")
        if not encoded:
            continue

        coords = decode_polyline(str(encoded))
        if not coords:
            continue

        condition  = str(row.get(condition_col, "") or "")
        colour     = condition_colour(condition)
        area       = str(row.get("area", "") or "")
        location   = str(row.get("location_description", "") or "")
        roadway    = str(row.get("roadway_name", "") or "")
        visibility = str(row.get("visibility", "") or "")
        updated    = str(row.get("last_updated", "") or "")

        placemark = ET.SubElement(doc, "Placemark")
        ET.SubElement(placemark, "name").text = roadway or location
        ET.SubElement(placemark, "styleUrl").text = f"#s_{colour}"

        desc_lines = [
            f"<b>Condition:</b> {condition}",
            f"<b>Visibility:</b> {visibility}",
            f"<b>Location:</b> {location}",
            f"<b>Area:</b> {area}",
            f"<b>Updated:</b> {updated}",
        ]
        if "secondary_conditions" in row and row["secondary_conditions"]:
            desc_lines.insert(1, f"<b>Secondary:</b> {row['secondary_conditions']}")
        ET.SubElement(placemark, "description").text = "<br/>".join(desc_lines)

        linestring = ET.SubElement(placemark, "LineString")
        ET.SubElement(linestring, "tessellate").text = "1"
        ET.SubElement(linestring, "coordinates").text = " ".join(
            f"{lng},{lat},0" for lat, lng in coords
        )

    return doc


def export_kml(df_cond, df_cond_poly, df_evt, df_evt_poly):
    """Write a single KML file with road conditions and events as separate folders."""
    kml_root = ET.Element("kml", xmlns="http://www.opengis.net/kml/2.2")
    master   = ET.SubElement(kml_root, "Document")
    ET.SubElement(master, "name").text = "NL 511 Road Data"

    # Roads folder
    roads_folder = ET.SubElement(master, "Folder")
    ET.SubElement(roads_folder, "name").text = "Road Conditions"
    roads_doc = build_kml(df_cond, df_cond_poly, "Road Conditions")
    for child in list(roads_doc):
        roads_folder.append(child)

    # Events folder
    events_folder = ET.SubElement(master, "Folder")
    ET.SubElement(events_folder, "name").text = "Events & Incidents"
    events_doc = build_kml(df_evt, df_evt_poly, "Events & Incidents")
    for child in list(events_doc):
        events_folder.append(child)

    # Pretty-print
    raw = ET.tostring(kml_root, encoding="unicode")
    pretty = minidom.parseString(raw).toprettyxml(indent="  ", encoding="utf-8")
    with open(OUT_KML, "wb") as f:
        f.write(pretty)
    print(f"  ✔  KML map              → {OUT_KML}")

# ─────────────────────────────────────────────
#  API HELPERS
# ─────────────────────────────────────────────

def fetch(endpoint: str, extra_params: dict = None) -> list:
    """
    Fetch a JSON list from a 511nl.ca endpoint.
    Tries v2 first, falls back to v3 if 404.
    """
    params = {"key": API_KEY}
    if extra_params:
        params.update(extra_params)

    for version in ["v2", "v3"]:
        url = f"https://511nl.ca/api/{version}/get/{endpoint}"
        try:
            resp = requests.get(url, params=params, timeout=30)
            if resp.status_code == 200:
                data = resp.json()
                # Some endpoints wrap in a list, some return dict; normalise
                if isinstance(data, list):
                    return data
                if isinstance(data, dict):
                    # Look for a list inside common wrapper keys
                    for key in ("Items", "items", "results", "data"):
                        if key in data and isinstance(data[key], list):
                            return data[key]
                    return [data]
            elif resp.status_code == 404:
                continue          # try next version
            elif resp.status_code == 401:
                print(f"[ERROR] 401 Unauthorized — check your API key.")
                sys.exit(1)
            else:
                print(f"[WARN]  {url} → HTTP {resp.status_code}")
        except requests.exceptions.RequestException as e:
            print(f"[ERROR] Could not reach {url}: {e}")
    return []


# ─────────────────────────────────────────────
#  ROAD CONDITIONS (WINTER ROADS)
# ─────────────────────────────────────────────

def normalise_condition(raw: dict) -> dict:
    """Flatten one road-condition record into a clean dict."""
    def get(*keys):
        """Try multiple key variants (camelCase, PascalCase, hyphenated)."""
        for k in keys:
            if k in raw:
                return raw[k]
        return None

    # Field names confirmed from NL511 API docs (note spaces in multi-word names)
    secondary = get("Secondary Conditions", "SecondaryConditions",
                    "Secondary-Conditions", "secondaryConditions")
    if isinstance(secondary, list):
        secondary = " | ".join(str(s) for s in secondary if s)
    elif secondary == []:
        secondary = None

    last_updated_raw = get("LastUpdated", "lastUpdated", "UpdatedAt", "updated_at")
    last_updated = parse_timestamp(last_updated_raw)

    # "Primary Condition" and "Visibility" are two distinct fields per the API docs.
    # Primary Condition = road surface condition (e.g. Bare & Dry, Compact Snow)
    # Visibility        = driving visibility rating (e.g. Good, Fair)
    road_condition = get("Primary Condition", "PrimaryCondition",
                         "Primary-Condition", "primaryCondition")
    visibility     = get("Visibility", "visibility")

    return {
        "id":                  get("Id", "id"),
        "roadway_name":        get("RoadwayName", "Roadway", "roadwayName"),
        "area":                get("AreaName", "Area", "areaName"),
        "location_description":get("LocationDescription", "Description", "locationDescription"),
        "road_condition":      road_condition,
        "secondary_conditions":secondary,
        "visibility":          visibility,
        "last_updated":        last_updated,
        # Polyline stored separately — excluded from main sheets to keep output readable
        "_encoded_polyline":   get("EncodedPolyline", "encodedPolyline", "Polyline"),
    }


# ─────────────────────────────────────────────
#  EVENTS / INCIDENTS
# ─────────────────────────────────────────────

def normalise_event(raw: dict) -> dict:
    """Flatten one event/incident record into a clean dict."""
    def get(*keys):
        for k in keys:
            if k in raw:
                return raw[k]
        return None

    start = parse_timestamp(get("StartDate", "startDate", "Start", "start"))
    end   = parse_timestamp(get("EndDate",   "endDate",   "End",   "end"))

    return {
        "id":                  get("Id", "id", "EventId"),
        "event_type":          get("EventType", "eventType", "Type", "type"),
        "status":              get("Status", "status"),
        "severity":            get("Severity", "severity"),
        "roadway_name":        get("RoadwayName", "Roadway", "roadwayName"),
        "direction":           get("DirectionOfTravel", "Direction", "direction",
                                   "TravelDirection", "travelDirection"),
        "area":                get("AreaName", "Area", "areaName"),
        "location_description":get("LocationDescription", "Description", "locationDescription"),
        "description":         get("Description", "description", "Comment", "comment"),
        "lanes_affected":      get("LanesAffected", "lanesAffected", "Lanes", "LanesAffectedCount"),
        "start_date":          start,
        "end_date":            end,
        # Polyline stored separately — excluded from main sheets to keep output readable
        "_encoded_polyline":   get("EncodedPolyline", "encodedPolyline", "Polyline"),
    }


# ─────────────────────────────────────────────
#  UTILITY
# ─────────────────────────────────────────────

def parse_timestamp(value) -> str:
    """Try to return a readable timestamp string from a variety of formats."""
    if value is None:
        return None
    if isinstance(value, str):
        # Already ISO-ish
        return value.replace("T", " ").replace("Z", "").strip()
    if isinstance(value, (int, float)):
        # Unix epoch (milliseconds or seconds)
        if value > 1e10:
            value /= 1000
        try:
            return datetime.utcfromtimestamp(value).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return str(value)
    return str(value)


def to_dataframe(records: list, normaliser):
    """
    Returns (df_main, df_polylines).
    df_main    — all readable columns, no polyline clutter
    df_polylines — id + encoded_polyline, for mapping use
    """
    if not records:
        return pd.DataFrame(), pd.DataFrame()

    rows = [normaliser(r) for r in records]
    df = pd.DataFrame(rows)

    # Split out the polyline column (prefixed with _ so it's easy to find)
    poly_cols = [c for c in df.columns if c.startswith("_")]
    main_cols = [c for c in df.columns if not c.startswith("_")]

    df_main = df[main_cols].copy()
    df_main.dropna(axis=1, how="all", inplace=True)

    if poly_cols and "id" in df_main.columns:
        df_poly = df[["id"] + poly_cols].copy()
        df_poly.columns = ["id"] + [c.lstrip("_") for c in poly_cols]
        df_poly.dropna(subset=[c.lstrip("_") for c in poly_cols], how="all", inplace=True)
    else:
        df_poly = pd.DataFrame()

    return df_main, df_poly


# ─────────────────────────────────────────────
#  EXPORT
# ─────────────────────────────────────────────

def export(df_cond, df_cond_poly, df_evt, df_evt_poly):
    from openpyxl.styles import Font, PatternFill, Alignment
    run_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # ── CSV (main readable data only, no polylines) ───────────────
    if not df_cond.empty:
        df_cond.to_csv(OUT_COND_CSV, index=False, encoding="utf-8-sig")
        print(f"  ✔  Road conditions CSV  → {OUT_COND_CSV}  ({len(df_cond)} rows)")
    else:
        print("  ⚠  No road condition data to export.")

    if not df_evt.empty:
        df_evt.to_csv(OUT_EVT_CSV, index=False, encoding="utf-8-sig")
        print(f"  ✔  Events CSV           → {OUT_EVT_CSV}  ({len(df_evt)} rows)")
    else:
        print("  ⚠  No event data to export.")

    # ── Excel workbook (multi-sheet) ─────────────────────────────
    with pd.ExcelWriter(OUT_EXCEL, engine="openpyxl") as writer:

        def write_sheet(df: pd.DataFrame, sheet_name: str, title_suffix: str = ""):
            if df.empty:
                pd.DataFrame({"Note": [f"No data retrieved — {run_time}"]}).to_excel(
                    writer, sheet_name=sheet_name, index=False
                )
                return
            df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
            ws = writer.sheets[sheet_name]

            # Title row
            title = f"NL 511 — {sheet_name}{title_suffix} — extracted {run_time}"
            ws.cell(1, 1, title).font = Font(bold=True, size=12)

            # Style header row (row 2)
            header_fill = PatternFill("solid", fgColor="2E5E8E")
            header_font = Font(bold=True, color="FFFFFF")
            for cell in ws[2]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")

            # Auto-width columns (cap at 55 chars)
            for col_cells in ws.columns:
                max_len = max(
                    (len(str(cell.value or "")) for cell in col_cells),
                    default=10
                )
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 55)

            ws.freeze_panes = "A3"

        write_sheet(df_cond, "Road Conditions")
        write_sheet(df_evt,  "Events & Incidents")

        # Polyline sheets — for mapping / GIS use
        if not df_cond_poly.empty:
            write_sheet(df_cond_poly, "Polylines - Roads",
                        " (map geometry — use with Google Maps or GIS)")
        if not df_evt_poly.empty:
            write_sheet(df_evt_poly, "Polylines - Events",
                        " (map geometry — use with Google Maps or GIS)")

        # Metadata sheet
        meta = pd.DataFrame({
            "Field": ["Extracted at (UTC)", "Road condition rows", "Event rows",
                      "Source — conditions", "Source — events",
                      "Polyline note"],
            "Value": [run_time, len(df_cond), len(df_evt),
                      "https://511nl.ca/api/v2/get/winterroads",
                      "https://511nl.ca/api/v2/get/events",
                      "Encoded polylines are Google Maps format. "
                      "Decode at: https://developers.google.com/maps/documentation/utilities/polylineutility"],
        })
        meta.to_excel(writer, sheet_name="Metadata", index=False)
        ws_meta = writer.sheets["Metadata"]
        for col_cells in ws_meta.columns:
            ws_meta.column_dimensions[col_cells[0].column_letter].width = 70

    print(f"  ✔  Excel workbook       → {OUT_EXCEL}")


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────

def main():
    if API_KEY == "YOUR_API_KEY_HERE":
        print("[ERROR] Please open nl511_extract.py and replace YOUR_API_KEY_HERE "
              "with your actual API key from 511nl.ca")
        sys.exit(1)

    print("NL 511 Extractor — starting…\n")

    # ── Fetch ────────────────────────────────
    print("→ Fetching road conditions (winterroads)…")
    raw_cond = fetch("winterroads")
    print(f"   {len(raw_cond)} segment(s) received.")

    print("→ Fetching events & incidents…")
    raw_evt = fetch("events")
    print(f"   {len(raw_evt)} event(s) received.\n")

    # Debug: show field names + values (excluding polyline so they're not buried)
    if raw_cond:
        print("── Road condition fields returned by API ─────────")
        sample = {k: v for k, v in raw_cond[0].items()
                  if "polyline" not in k.lower() and "encoded" not in k.lower()}
        print(json.dumps(sample, indent=2, default=str))
        print()
    if raw_evt:
        print("── Event fields returned by API ──────────────────")
        sample = {k: v for k, v in raw_evt[0].items()
                  if "polyline" not in k.lower() and "encoded" not in k.lower()}
        print(json.dumps(sample, indent=2, default=str))
        print()

    # ── Normalise ────────────────────────────
    df_cond, df_cond_poly = to_dataframe(raw_cond, normalise_condition)
    df_evt,  df_evt_poly  = to_dataframe(raw_evt,  normalise_event)

    # ── Export ───────────────────────────────
    print("→ Exporting…")
    export(df_cond, df_cond_poly, df_evt, df_evt_poly)
    export_kml(df_cond, df_cond_poly, df_evt, df_evt_poly)

    print("\nDone! Open nl511_data.xlsx to browse the data, or nl511_roads.kml in Google Earth.\n")

    # ── Quick summary ────────────────────────
    if not df_cond.empty and "road_condition" in df_cond.columns:
        print("── Primary condition breakdown ───────────────────")
        print(df_cond["road_condition"].value_counts().to_string())
        print()
    if not df_cond.empty and "visibility" in df_cond.columns:
        print("── Visibility breakdown ──────────────────────────")
        print(df_cond["visibility"].value_counts().to_string())
        print()
    if not df_evt.empty and "event_type" in df_evt.columns:
        print("── Event type breakdown ──────────────────────────")
        print(df_evt["event_type"].value_counts().to_string())


if __name__ == "__main__":
    main()
