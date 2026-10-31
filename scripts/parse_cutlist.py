import os
import re
import json
import time

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATE_FILE = os.path.join(BASE_DIR, "bazar_factory_saw_state.json")
DIAGRAMS_DIR = os.path.join(BASE_DIR, "public", "diagrams")

os.makedirs(DIAGRAMS_DIR, exist_ok=True)

# High-contrast standalone SVG stylesheet embedded in every diagram
EMBEDDED_SVG_CSS = """
<style>
  svg.packing {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    background-color: #0b1325;
    border-radius: 8px;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.5);
    user-select: none;
    overflow: visible;
  }
  .bin-outer { fill: #131d33; stroke: #38bdf8; stroke-width: 1.5px; }
  .bin-inner { opacity: 0.95; }
  .bin-trimming { fill: none; stroke: #f43f5e; stroke-width: 1px; stroke-dasharray: 4,4; }
  .bin-dimension { fill: #38bdf8; font-size: 14px; font-weight: 800; }
  
  .item-outer { fill: #0f172a; stroke: #334155; stroke-width: 1px; }
  .item-projection-shape { stroke-width: 1.2px; fill-opacity: 0.92; }
  .item-label { fill: #0f172a; font-weight: 900; font-size: 26px; pointer-events: none; }
  .item-dimension { fill: #0f172a; font-weight: 800; font-size: 13px; }
  
  .leftover-dimension { fill: #94a3b8; font-weight: 700; font-size: 12px; }
  
  .cut-outer { fill: #000000; opacity: 0.35; }
  .cut-inner { fill: #f59e0b; }
  .cut-primary .cut-inner { fill: #ef4444; stroke: #ef4444; stroke-width: 1.5px; }
  .cut-trimming .cut-inner { fill: #64748b; }
</style>
"""

# Sources: 18mm MDF Core Cutlist and 8mm Back Panel Cutlist
SOURCE_CONFIGS = [
    {
        "id": "18mm",
        "thickness": 18,
        "default_material": "18mm MDF Core",
        "file_candidates": [
            os.path.join(BASE_DIR, "src", "data", "cutlist.html"),
            os.path.join(BASE_DIR, "public", "cutlist.html"),
            os.path.join(BASE_DIR, "data", "cutlist.html")
        ]
    },
    {
        "id": "8mm",
        "thickness": 8,
        "default_material": "Back Panel (8mm)",
        "file_candidates": [
            os.path.join(BASE_DIR, "src", "data", "cutlist2.html"),
            os.path.join(BASE_DIR, "public", "cutlist2.html"),
            os.path.join(BASE_DIR, "data", "cutlist2.html")
        ]
    }
]

def clean_html_text(text):
    if not text:
        return ""
    clean = re.sub(r'<[^>]+>', ' ', text)
    clean = clean.replace('&nbsp;', ' ').replace('&amp;', '&').replace('&quot;', '"')
    clean = re.sub(r'color:\s*#[0-9a-fA-F]+;\"?>?\s*', '', clean)
    clean = re.sub(r'^[^\w]+', '', clean)
    return re.sub(r'\s+', ' ', clean).strip()

def extract_global_material_summary():
    """Extract OpenCutList overall project rough length, area, volume & lipping totals"""
    main_html_path = os.path.join(BASE_DIR, "src", "data", "cutlist.html")
    if not os.path.exists(main_html_path):
        return []

    try:
        with open(main_html_path, "r", encoding="utf-8", errors="ignore") as f:
            header_sample = f.read(250000)

        matches = re.findall(r'<tr[^>]*id=\"ladb_group_[^\"]*_summary\"[^>]*>(.*?)<\/tr>', header_sample, re.DOTALL)
        summaries = []
        for m in matches:
            tds = re.findall(r'<td[^>]*>(.*?)<\/td>', m, re.DOTALL)
            row_vals = [re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', td)).strip() for td in tds]
            if len(row_vals) >= 6:
                mat_name = clean_html_text(row_vals[1])
                summaries.append({
                    "material": mat_name,
                    "parts_qty": int(re.sub(r'[^\d]', '', row_vals[2]) or 0),
                    "rough_length": row_vals[3],
                    "rough_area": row_vals[4],
                    "rough_volume": row_vals[5]
                })
        return summaries
    except Exception as e:
        print("[SUMMARY EXTRACTION ERROR]", e)
        return []

def parse_file(html_path, thickness, default_mat, id_prefix):
    if not os.path.exists(html_path):
        return []

    file_size = os.path.getsize(html_path)
    print(f"[PARSER] Parsing {id_prefix} from: {html_path} ({file_size} bytes)")
    t0 = time.time()

    with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
        html = f.read()

    tables = re.findall(r'<table[^>]*class="[^"]*ladb-cutlist-packing-group[^"]*"[^>]*>(.*?)<\/table>', html, re.DOTALL)
    print(f"[PARSER] Found {len(tables)} packing group tables in {id_prefix}.")

    sheets = []
    for t_html in tables:
        m_range = re.search(r'<strong>\s*(\d+)(?:\s*\.\.\.\s*(\d+))?\s*\/\s*(\d+)\s*<\/strong>', t_html)
        if not m_range:
            continue

        start_s = int(m_range.group(1))
        end_s = int(m_range.group(2)) if m_range.group(2) else start_s
        total_s = int(m_range.group(3))

        m_count = re.search(r'<div class="ladb-heading-count[^"]*">x&nbsp;(\d+)<\/div>', t_html)
        repeat_count = int(m_count.group(1)) if m_count else (end_s - start_s + 1)

        m_mat = re.search(r'<div class="ladb-heading">.*?([^\-]+)\s*-\s*<i[^>]*><\/i>\s*<em>([^<]+)<\/em>', t_html)
        mat_raw = m_mat.group(1).strip() if m_mat else default_mat
        mat_desc = clean_html_text(mat_raw)
        if not mat_desc:
            mat_desc = default_mat
        sheet_dim = m_mat.group(2).strip() if m_mat else "2800 mm x 2070 mm"

        m_eff = re.search(r'Efficiency\s*:\s*<\/strong>\s*([\d\.]+)%', t_html)
        eff = m_eff.group(1) + "%" if m_eff else "90.0%"

        # Cut stats: Length of Cuts, Grain, Parts count
        m_len = re.search(r'Length of Cuts\s*:\s*<\/strong>\s*([^<]+)', t_html)
        cut_length_m = m_len.group(1).strip() if m_len else None

        m_grain = re.search(r'(Grain follows\s+[^\.<&]+)', t_html)
        grain_spec = m_grain.group(1).strip() if m_grain else "Grain follows length"

        cuts_summary = re.findall(r'(\d+\s*mm)\s*<i[^>]*class="ladb-opencutlist-icon-horizontal-cut-bottom"', t_html)
        rip_cuts = cuts_summary if cuts_summary else []

        # Extract SVG Cutting Pattern Diagram and embed self-contained styling
        diagram_url = None
        m_svg = re.search(r'(<svg[^>]*class="[^"]*packing[^"]*"[^>]*>.*?<\/svg>)', t_html, re.DOTALL)
        if not m_svg:
            m_svg = re.search(r'(<svg[^>]*>.*?<\/svg>)', t_html, re.DOTALL)

        sheet_id = f"{id_prefix}_S{start_s}_{end_s}"
        if m_svg:
            svg_content = m_svg.group(1)

            # Clean style attribute on svg tag
            svg_content = re.sub(r'style="max-height:\s*[^"]+"', 'style="width: 100%; height: auto; max-height: 520px; display: block; margin: 0 auto; background: #0b1325;"', svg_content)
            
            # Inject standalone CSS inside <defs> or right after <svg ...>
            if '<defs>' in svg_content:
                svg_content = svg_content.replace('<defs>', '<defs>' + EMBEDDED_SVG_CSS, 1)
            else:
                svg_content = re.sub(r'(<svg[^>]*>)', r'\1' + EMBEDDED_SVG_CSS, svg_content, 1)

            svg_path = os.path.join(DIAGRAMS_DIR, f"{sheet_id}.svg")
            try:
                with open(svg_path, 'w', encoding='utf-8') as svg_f:
                    svg_f.write(svg_content)
                diagram_url = f"/diagrams/{sheet_id}.svg"
            except Exception as svg_err:
                print(f"[SVG SAVE WARNING] {sheet_id}: {svg_err}")

        parts = []
        pdf_sets_found = set()
        rows = re.findall(r'<tr[^>]*class="[^"]*ladb-cutlist-row[^"]*"[^>]*>(.*?)<\/tr>', t_html, re.DOTALL)

        for r in rows:
            m_code = re.search(r'<td class="ladb-cutlist-value[^"]*">\s*<span[^>]*>([^<]+)<\/span>', r)
            code = m_code.group(1).strip() if m_code else ""

            m_name = re.search(r'<span class="ladb-cutlist-part-name">([^<]+)<\/span>', r)
            name = m_name.group(1).strip() if m_name else ""

            m_core = re.search(r'<div class="ladb-part-description">([^<]+)<\/div>', r)
            core = m_core.group(1).strip() if m_core else ""

            m_qty = re.search(r'<td class="text-center">\s*(\d+)\s*<\/td>', r)
            qty_per_sheet = int(m_qty.group(1)) if m_qty else 1

            dims = re.findall(r'<span class="ladb-flex-grow-1">\s*(\d+)\s*mm\s*<\/span>', r)
            l = int(dims[0]) if len(dims) > 0 else 0
            w = int(dims[1]) if len(dims) > 1 else 0

            pdf_num = "PDF 5"
            m_pdf = re.search(r'PDF\s*([1-5])', name, re.I)
            if m_pdf:
                pdf_num = f"PDF {m_pdf.group(1)}"
            pdf_sets_found.add(pdf_num)

            villa = "STANDARD 40"
            m_v = re.search(r'-V(\d{2})-', name) or re.search(r'V(\d{2})', name)
            if m_v:
                villa = f"Villa {m_v.group(1)}"

            cab = "CABINET"
            if "CB-03-SF" in name or "CB03SF" in name: cab = "Bedroom 03 (CB-03-SF)"
            elif "CB-02-SF" in name or "CB02SF" in name: cab = "Bedroom 02 (CB-02-SF)"
            elif "SVMB" in name: cab = "Master Bed (SVMB)"
            elif "SCMB" in name: cab = "Side Cabinet (SCMB)"
            elif "CB-01" in name: cab = "Bedroom 01 (CB-01)"
            elif "DRW" in name: cab = "Drawers (DRW)"

            # Lip determination
            if thickness == 8 or "BACK PANEL" in name.upper():
                lip = "Raw / Grooved (No Lip)"
            elif "FACE" in name.upper() or "DRAWER" in name.upper():
                lip = "2mm 4-Sides"
            else:
                lip = "2mm Front"

            parts.append({
                "code": code,
                "name": name,
                "pdf_set": pdf_num,
                "core": core,
                "qty_per_sheet": qty_per_sheet,
                "total_batch_qty": qty_per_sheet * repeat_count,
                "l": l, "w": w, "t": thickness,
                "cab": cab,
                "villa": villa,
                "lip": lip
            })

        # Expand sheet repeat groups into individual sheet records for full MES granular tracking
        for sheet_num in range(start_s, end_s + 1):
            sheet_id = f"{id_prefix}_S{sheet_num}"
            sheet_title = f"{id_prefix.upper()} Sheet {sheet_num}"

            # Copy parts with quantity adjusted for 1 single sheet
            sheet_parts = []
            for p in parts:
                p_copy = dict(p)
                p_copy["total_batch_qty"] = p["qty_per_sheet"] # 1 sheet
                sheet_parts.append(p_copy)

            sheets.append({
                "sheet_id": sheet_id,
                "sheet_title": sheet_title,
                "thickness": thickness,
                "start_sheet": sheet_num,
                "end_sheet": sheet_num,
                "repeat_count": 1,
                "total_sheets": total_s,
                "material": mat_desc,
                "sheet_dim": sheet_dim,
                "efficiency": eff,
                "cut_length": cut_length_m or "28.0 m",
                "grain": grain_spec,
                "rip_cuts": rip_cuts,
                "diagram_url": diagram_url,
                "pdf_sets": sorted(list(pdf_sets_found)),
                "parts": sheet_parts,
                "assigned_saw": "Saw 1",
                "assigned_team": "Team A (Ahmed)",
                "status": "QUEUED",
                "queued_at": int(time.time()),
                "start_time": None,
                "cut_time": None,
                "duration_min": None
            })

    print(f"[PARSER] Successfully parsed {len(sheets)} sheets for {id_prefix} in {time.time() - t0:.2f}s")
    return sheets

def parse_all_cutlists():
    all_sheets = []
    for cfg in SOURCE_CONFIGS:
        found_file = None
        for candidate in cfg["file_candidates"]:
            if os.path.exists(candidate):
                found_file = candidate
                break
        if found_file:
            parsed = parse_file(found_file, cfg["thickness"], cfg["default_material"], cfg["id"])
            all_sheets.extend(parsed)
        else:
            print(f"[PARSER] Notice: No file found for {cfg['id']}")
    return all_sheets

def sync_state():
    parsed_sheets = parse_all_cutlists()
    global_material_summary = extract_global_material_summary()

    saws_status = {
        "Saw 1": {"status": "ONLINE", "team": "Team A (Ahmed)", "assigned_sheets": 0, "cut_sheets": 0},
        "Saw 2": {"status": "IN_REPAIR", "team": "Team B (Under Repair)", "assigned_sheets": 0, "cut_sheets": 0}
    }
    audit = []
    overtime = []

    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
                saws_status = saved.get("saws_status", saws_status)
                saved_sheets = {s["sheet_id"]: s for s in saved.get("sheets", [])}
                for s in parsed_sheets:
                    if s["sheet_id"] in saved_sheets:
                        prev = saved_sheets[s["sheet_id"]]
                        s["status"] = prev.get("status", "QUEUED")
                        s["assigned_saw"] = prev.get("assigned_saw", "Saw 1")
                        s["assigned_team"] = prev.get("assigned_team", "Team A (Ahmed)")
                        s["queued_at"] = prev.get("queued_at", s.get("queued_at"))
                        s["start_time"] = prev.get("start_time")
                        s["cut_time"] = prev.get("cut_time")
                        s["duration_min"] = prev.get("duration_min")
                audit = saved.get("audit", [])
                overtime = saved.get("overtime", [])
        except Exception as e:
            print("[SYNC ERROR]", e)

    data = {
        "saws_status": saws_status,
        "sheets": parsed_sheets,
        "material_summary": global_material_summary,
        "audit": audit,
        "overtime": overtime,
        "savedAt": time.strftime("%Y-%m-%d %H:%M:%S")
    }

    with open(STATE_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)
    print(f"[PARSER] State saved to {STATE_FILE} with {len(parsed_sheets)} total sheets and diagrams extracted.")

if __name__ == '__main__':
    sync_state()
