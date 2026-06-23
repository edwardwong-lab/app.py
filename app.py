import streamlit as st
import pandas as pd
import io
import zipfile
import re
from datetime import datetime
import traceback

st.set_page_config(page_title="Price Label Generator", page_icon="🏷️", layout="wide")

try:
    import fitz
except ImportError:
    st.error("PyMuPDF not installed. Run: pip install pymupdf")
    st.stop()

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ── helpers ───────────────────────────────────────────────────────────────
def norm(v):
    if v is None: return None
    s = str(v).strip()
    if not s or s.lower() in ("nan","none",""): return None
    if re.match(r'^\d+\.0$', s): s = s[:-2]
    return s.lower()

def safe_name(s):
    s = str(s).strip()
    s = re.sub(r'[\\/*?:"<>|]', '', s)
    return s.strip()

def auto_col(df, hints_list):
    cl = {c.lower().strip(): c for c in df.columns}
    for hints in hints_list:
        for h in hints:
            if h in cl: return cl[h]
        for h in hints:
            for k, v in cl.items():
                if h in k: return v
    return list(df.columns)[0]

def index_pdf(pdf_bytes):
    """Returns {norm_article_code: [page_indices]} skipping LS pages."""
    art_map = {}
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    n = doc.page_count
    for i, page in enumerate(doc):
        text = page.get_text("text")
        if re.search(r'\b\d{6,10}LS\b', text):
            continue
        for pat in [r'\b(\d{6,10})\b', r'(?:article|art|code|no|item)[^\d]*(\d{5,10})']:
            for m in re.finditer(pat, text, re.IGNORECASE):
                c = norm(m.group(1))
                if c: art_map.setdefault(c, []).append(i)
    doc.close()
    return art_map, n

def load_df(raw, fname):
    if fname.lower().endswith(".csv"):
        for enc in ["utf-8","latin-1","cp1252"]:
            try: return pd.read_csv(io.BytesIO(raw), dtype=str, encoding=enc)
            except: pass
    return pd.read_excel(io.BytesIO(raw), dtype=str)

def build_store_pdf(article_codes, zone_map, master_map, pdf_zone_bytes, pdf_master_bytes):
    """
    For each article code:
      - Use zone PDF page if exists (priority)
      - Else use master PDF page
    Returns merged PDF bytes with no duplicate articles.
    """
    out_doc = fitz.open()
    seen_articles = set()

    # Pass 1: zone PDF (priority)
    zone_pages = []
    for art in article_codes:
        if art in seen_articles: continue
        if art in zone_map:
            zone_pages.extend(zone_map[art])
            seen_articles.add(art)

    if zone_pages and pdf_zone_bytes:
        src = fitz.open(stream=pdf_zone_bytes, filetype="pdf")
        for p in sorted(set(zone_pages)):
            out_doc.insert_pdf(src, from_page=p, to_page=p)
        src.close()

    # Pass 2: master PDF fallback for remaining articles
    master_pages = []
    for art in article_codes:
        if art in seen_articles: continue
        if art in master_map:
            master_pages.extend(master_map[art])
            seen_articles.add(art)

    if master_pages and pdf_master_bytes:
        src = fitz.open(stream=pdf_master_bytes, filetype="pdf")
        for p in sorted(set(master_pages)):
            out_doc.insert_pdf(src, from_page=p, to_page=p)
        src.close()

    if len(out_doc) == 0:
        out_doc.close()
        return None, 0, len(seen_articles)

    buf = io.BytesIO()
    out_doc.save(buf)
    n = len(out_doc)
    out_doc.close()
    return buf.getvalue(), n, len(seen_articles)

# ── session state ─────────────────────────────────────────────────────────
# Clean up stale keys from older versions
for _stale in ["pdf_zone", "soh_pbo", "soh_npbo"]:
    if _stale in st.session_state:
        del st.session_state[_stale]
for k, v in {
    "pdf_pbo":      {"bytes": None, "name": None, "art_map": None, "n_pages": 0},
    "pdf_npbo":     {"bytes": None, "name": None, "art_map": None, "n_pages": 0},
    "pdf_zones":    [],   # list of {name, bytes, art_map, n_pages, zone_tag}
    "stock_df":     None,
    "col_zone":     None, "col_aam": None, "col_store": None,
    "col_article":  None, "col_soh": None, "col_category": None,
    "pbo_cats":     ["Medicine"],
    "zip_bytes":    None, "rep_bytes": None, "store_rows": [],
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── styles ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
*{font-family:'Inter',sans-serif;}
.hdr{background:linear-gradient(135deg,#1B4F72,#2E86AB);color:white;padding:1.8rem 2rem;border-radius:12px;margin-bottom:1.5rem;}
.hdr h1{font-size:1.5rem;font-weight:700;margin:0;}
.hdr p{font-size:0.87rem;opacity:0.85;margin:0.3rem 0 0;}
.card{background:white;border:1.5px solid #E2E8F0;border-radius:10px;padding:1.3rem;margin-bottom:1rem;}
.sec{font-size:0.68rem;font-weight:700;letter-spacing:1px;text-transform:uppercase;background:#1B4F72;color:white;padding:2px 9px;border-radius:20px;}
.badge-pbo{background:#DCFCE7;color:#166534;padding:2px 8px;border-radius:10px;font-size:0.75rem;font-weight:600;}
.badge-npbo{background:#FEF9C3;color:#854D0E;padding:2px 8px;border-radius:10px;font-size:0.75rem;font-weight:600;}
.badge-zone{background:#EDE9FE;color:#5B21B6;padding:2px 8px;border-radius:10px;font-size:0.75rem;font-weight:600;}
.pdf-row{background:#F8FAFC;border:1px solid #E2E8F0;border-radius:8px;padding:0.8rem 1rem;margin-bottom:0.5rem;display:flex;align-items:center;gap:10px;}
.stbl{width:100%;border-collapse:collapse;font-size:0.82rem;margin-top:0.5rem;}
.stbl th{background:#1B4F72;color:white;padding:7px 10px;text-align:left;}
.stbl td{padding:6px 10px;border-bottom:1px solid #E2E8F0;}
.stbl tr:hover td{background:#F8FAFC;}
.fname{font-family:monospace;background:#F1F5F9;padding:2px 5px;border-radius:4px;font-size:0.78rem;color:#1B4F72;}
.logic-box{background:#F0F9FF;border:1.5px solid #BAE6FD;border-radius:8px;padding:1rem;font-size:0.85rem;line-height:1.7;}
</style>""", unsafe_allow_html=True)

st.markdown("""
<div class="hdr">
  <h1>🏷️ Price Label Filter & Store PDF Generator</h1>
  <p>3 PDFs (Master PBO · Master NPBO · Zone Mixed) + 1 SOH file → store PDFs split by PBO/NPBO with zone price priority</p>
</div>""", unsafe_allow_html=True)

# Logic explanation
st.markdown("""
<div class="logic-box">
  <b>📌 How it works:</b><br>
  1. Each article in SOH is classified: <b>Category = Medicine → PBO</b>, everything else → <b>NPBO</b><br>
  2. Label source priority per article: <b>Zone PDF first</b> (zone price) → fallback to <b>Master PBO</b> or <b>Master NPBO</b><br>
  3. No duplicate labels — each article appears only once per store output file<br>
  4. Output: one ZIP with files named <code>Zone - AAM - Store - PBO.pdf</code> or <code>NPBO.pdf</code>
</div><br>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════
# STEP 1 — UPLOAD 3 PDFs
# ══════════════════════════════════════════════════════════════════════════
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<span class="sec">Step 1</span>', unsafe_allow_html=True)
st.subheader("Upload Price Label PDFs")

c1, c2, c3 = st.columns(3)

with c1:
    st.markdown('<span class="badge-pbo">PBO</span> &nbsp;<b>Master PBO PDF</b>', unsafe_allow_html=True)
    st.caption("Items without zone price — PBO (Medicine)")
    f1 = st.file_uploader("Master PBO PDF", type=["pdf"], key="fu_pbo", label_visibility="collapsed")
    if f1:
        raw = f1.read()
        st.session_state.pdf_pbo = {"bytes": raw, "name": f1.name, "art_map": None, "n_pages": 0}
        st.session_state.zip_bytes = None
    if st.session_state.pdf_pbo["name"]:
        idx_info = f"✅ {st.session_state.pdf_pbo['n_pages']:,} pages indexed" if st.session_state.pdf_pbo["art_map"] else "⏳ Not indexed yet"
        st.success(f"📄 {st.session_state.pdf_pbo['name']} · {idx_info}")

with c2:
    st.markdown('<span class="badge-npbo">NPBO</span> &nbsp;<b>Master NPBO PDF</b>', unsafe_allow_html=True)
    st.caption("Items without zone price — NPBO (non-Medicine)")
    f2 = st.file_uploader("Master NPBO PDF", type=["pdf"], key="fu_npbo", label_visibility="collapsed")
    if f2:
        raw = f2.read()
        st.session_state.pdf_npbo = {"bytes": raw, "name": f2.name, "art_map": None, "n_pages": 0}
        st.session_state.zip_bytes = None
    if st.session_state.pdf_npbo["name"]:
        idx_info = f"✅ {st.session_state.pdf_npbo['n_pages']:,} pages indexed" if st.session_state.pdf_npbo["art_map"] else "⏳ Not indexed yet"
        st.success(f"📄 {st.session_state.pdf_npbo['name']} · {idx_info}")

with c3:
    st.markdown('<span class="badge-zone">ZONE</span> &nbsp;<b>Zone Price PDFs (Mixed, one per zone)</b>', unsafe_allow_html=True)
    st.caption("Upload all zone PDFs — app auto-detects zone name from filename")
    f3_list = st.file_uploader("Zone PDFs", type=["pdf"], key="fu_zone",
                                accept_multiple_files=True, label_visibility="collapsed")
    if f3_list:
        existing = {s["name"]: s for s in st.session_state.pdf_zones}
        new_zones = []
        for f3 in f3_list:
            if f3.name in existing:
                new_zones.append(existing[f3.name])
            else:
                # Auto-detect zone tag from filename
                # e.g. "ZonePrice_Zone2.pdf" → "Zone2"
                # e.g. "zone_price_north.pdf" → "north"
                fname_clean = re.sub(r'\.pdf$', '', f3.name, flags=re.IGNORECASE)
                # Remove common prefixes
                for prefix in ["zoneprice","zone_price","zone price","price_zone","pricezone","zone"]:
                    fname_clean = re.sub(rf'^{prefix}[\s_\-]*', '', fname_clean, flags=re.IGNORECASE)
                zone_tag = fname_clean.strip(" _-") or f3.name
                new_zones.append({
                    "name": f3.name, "bytes": f3.read(),
                    "art_map": None, "n_pages": 0,
                    "zone_tag": zone_tag,
                })
        st.session_state.pdf_zones = new_zones
        st.session_state.zip_bytes = None

    if st.session_state.pdf_zones:
        for zslot in st.session_state.pdf_zones:
            idx_info = f"✅ {zslot['n_pages']:,} pages" if zslot["art_map"] else "⏳ Not indexed"
            st.markdown(f"📄 `{zslot['name']}` → zone tag: **{zslot['zone_tag']}** · {idx_info}")

st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════
# STEP 2 — UPLOAD SOH FILE
# ══════════════════════════════════════════════════════════════════════════
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<span class="sec">Step 2</span>', unsafe_allow_html=True)
st.subheader("Upload SOH Stock File")
st.caption("One file containing all stores — with Zone, AAM, Store, Article Code, SOH, and Category columns.")

sf = st.file_uploader("📊 SOH File (Excel or CSV)", type=["xlsx","xls","csv"], key="fu_stock")
if sf:
    raw2 = sf.read()
    try:
        df = load_df(raw2, sf.name)
        st.session_state.stock_df    = df
        st.session_state.col_zone    = auto_col(df, [["zone","region","area","district","territory"]])
        st.session_state.col_aam     = auto_col(df, [["aam","area manager","assistant area manager"]])
        st.session_state.col_store   = auto_col(df, [["store","outlet","site","branch","store name","store code"]])
        st.session_state.col_article = auto_col(df, [["article code","articlecode","article","material","sku","item code","barcode"]])
        st.session_state.col_soh     = auto_col(df, [["soh","stock on hand","balance qty","qty","quantity","stock","balance"]])
        st.session_state.col_category= auto_col(df, [["category","cat","type","item type","product type","class","group"]])
        st.session_state.zip_bytes   = None
        st.success(f"✅ {sf.name} — {len(df):,} rows · {len(df.columns)} columns")
    except Exception as e:
        st.error(f"Cannot read file: {e}")

st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════
# STEP 3 — COLUMN MAPPING + CATEGORY RULE
# ══════════════════════════════════════════════════════════════════════════
if st.session_state.stock_df is not None:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<span class="sec">Step 3</span>', unsafe_allow_html=True)
    st.subheader("Column Mapping & PBO/NPBO Rule")

    df   = st.session_state.stock_df
    cols = list(df.columns)

    st.markdown("**Column Mapping**")
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    with c1: col_zone     = st.selectbox("🗺️ Zone",     cols, index=cols.index(st.session_state.col_zone)     if st.session_state.col_zone     in cols else 0, key="m_zone")
    with c2: col_aam      = st.selectbox("👤 AAM",      cols, index=cols.index(st.session_state.col_aam)      if st.session_state.col_aam      in cols else 0, key="m_aam")
    with c3: col_store    = st.selectbox("🏪 Store",    cols, index=cols.index(st.session_state.col_store)    if st.session_state.col_store    in cols else 0, key="m_store")
    with c4: col_article  = st.selectbox("🔢 Article",  cols, index=cols.index(st.session_state.col_article)  if st.session_state.col_article  in cols else 0, key="m_article")
    with c5: col_soh      = st.selectbox("📦 SOH",      cols, index=cols.index(st.session_state.col_soh)      if st.session_state.col_soh      in cols else 0, key="m_soh")
    with c6: col_category = st.selectbox("🏷️ Category", cols, index=cols.index(st.session_state.col_category) if st.session_state.col_category in cols else 0, key="m_cat")

    st.session_state.col_zone     = col_zone
    st.session_state.col_aam      = col_aam
    st.session_state.col_store    = col_store
    st.session_state.col_article  = col_article
    st.session_state.col_soh      = col_soh
    st.session_state.col_category = col_category

    st.markdown("---")

    # PBO category selector
    st.markdown("**PBO Category Rule**")
    st.caption("Select which category values = PBO (Medicine). Everything else = NPBO.")

    all_cats = sorted([c for c in df[col_category].dropna().astype(str).str.strip().unique()
                       if c and c.lower() not in ("nan","none","")])

    # Default: pre-select "Medicine" if present
    default_pbo = [c for c in all_cats if c.lower() == "medicine"]

    pbo_cats = st.multiselect(
        "Categories classified as PBO",
        options=all_cats,
        default=default_pbo if default_pbo else [],
        help="Articles with these categories will go into the PBO output file. All others → NPBO."
    )
    st.session_state.pbo_cats = pbo_cats

    if pbo_cats:
        st.markdown(f'<span class="badge-pbo">PBO</span> &nbsp; {", ".join(pbo_cats)}', unsafe_allow_html=True)
        non_pbo = [c for c in all_cats if c not in pbo_cats]
        if non_pbo:
            st.markdown(f'<span class="badge-npbo">NPBO</span> &nbsp; {", ".join(non_pbo[:10])}{"..." if len(non_pbo)>10 else ""}', unsafe_allow_html=True)
    else:
        st.warning("⚠️ No PBO categories selected — all articles will go to NPBO.")

    with st.expander("👁️ Preview first 5 rows"):
        st.dataframe(df[[col_zone, col_aam, col_store, col_article, col_soh, col_category]].head(), use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════
    # STEP 4 — FILTERS
    # ══════════════════════════════════════════════════════════════════════
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<span class="sec">Step 4</span>', unsafe_allow_html=True)
    st.subheader("Filters")

    df2 = st.session_state.stock_df

    # Zone multi-select
    st.markdown("**🗺️ Zone Selection**")
    all_zones = sorted([z for z in df2[col_zone].dropna().astype(str).str.strip().unique()
                        if z and z.lower() not in ("nan","none","")])
    selected_zones = st.multiselect(
        "Select zones (leave empty = all zones)",
        options=all_zones, default=[],
        placeholder="Select zones or leave empty for all..."
    )
    st.info(f"📌 {'Zones: ' + ', '.join(selected_zones) if selected_zones else 'All zones will be generated.'}")

    st.markdown("---")

    # Store selection
    st.markdown("**🏪 Store Selection**")
    all_stores = sorted([s for s in df2[col_store].dropna().astype(str).str.strip().unique()
                         if s and s.lower() not in ("nan","none","")])
    sc1, sc2 = st.columns([3,1])
    with sc1:
        store_search = st.text_input("🔍 Search stores", placeholder="Type to filter...", key="srch_s")
    with sc2:
        st.markdown("<br>", unsafe_allow_html=True)
        sel_all = st.checkbox("✅ Select all", value=True, key="chk_all")

    visible = [s for s in all_stores if store_search.lower() in s.lower()] if store_search else all_stores
    st.caption(f"Showing {len(visible)} of {len(all_stores)} stores")
    checked = []
    cbs = st.columns(4)
    for i, s in enumerate(visible):
        with cbs[i%4]:
            if st.checkbox(s, value=sel_all, key=f"cb_{s}"):
                checked.append(s)
    selected_stores = all_stores if (sel_all and not store_search) else checked
    st.info(f"📌 **{len(selected_stores)}** of **{len(all_stores)}** stores selected.")

    st.markdown("---")

    # SOH filter
    st.markdown("**📦 SOH Filter**")
    filter_soh = not st.checkbox(
        "Include all matched articles regardless of SOH (default: SOH > 0 only)", value=False)
    st.info("✅ Only SOH > 0 included." if filter_soh else "⚠️ All matched articles included.")

    st.markdown('</div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════
    # STEP 5 — GENERATE
    # ══════════════════════════════════════════════════════════════════════
    has_pbo_pdf  = st.session_state.pdf_pbo["bytes"]  is not None
    has_npbo_pdf = st.session_state.pdf_npbo["bytes"] is not None
    has_zone_pdf = len(st.session_state.pdf_zones) > 0
    has_any_pdf  = has_pbo_pdf or has_npbo_pdf or has_zone_pdf

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<span class="sec">Step 5</span>', unsafe_allow_html=True)
    st.subheader("Generate")

    if not has_any_pdf:
        st.warning("⬆️ Upload at least one PDF in Step 1.")
    elif not selected_stores:
        st.warning("⬆️ Select at least one store in Step 4.")
    else:
        zone_label = f"Zones: {', '.join(selected_zones)}" if selected_zones else "All Zones"
        run = st.button(
            f"⚡ Generate Store PDFs  ·  {zone_label}  ·  {len(selected_stores)} stores",
            type="primary", use_container_width=True, key="btn_run"
        )

        progress_bar = st.progress(0)
        status_box   = st.empty()

        if run:
            log_lines = []
            def log(msg):
                log_lines.append(msg)
                status_box.info("\n\n".join(log_lines[-5:]))

            try:
                # ── Phase 1: Index PDFs ───────────────────────────────
                log("📖 Indexing PDFs...")
                # Index master PDFs
                total_pdfs = (1 if st.session_state.pdf_pbo["bytes"] else 0) +                              (1 if st.session_state.pdf_npbo["bytes"] else 0) +                              len(st.session_state.pdf_zones)
                done = 0

                for slot_key, label in [("pdf_pbo","Master PBO"),("pdf_npbo","Master NPBO")]:
                    slot = st.session_state[slot_key]
                    if not slot["bytes"]: continue
                    log(f"📖 Indexing {label}: {slot['name']}")
                    am, npages = index_pdf(slot["bytes"])
                    st.session_state[slot_key]["art_map"] = am
                    st.session_state[slot_key]["n_pages"] = npages
                    done += 1
                    progress_bar.progress(done / max(total_pdfs,1) * 0.25)
                    log(f"   ✅ {npages:,} pages · {len(am):,} unique codes")

                # Index all zone PDFs
                for zi, zslot in enumerate(st.session_state.pdf_zones):
                    log(f"📖 Indexing Zone PDF [{zi+1}/{len(st.session_state.pdf_zones)}]: {zslot['name']} (zone: {zslot['zone_tag']})")
                    am, npages = index_pdf(zslot["bytes"])
                    st.session_state.pdf_zones[zi]["art_map"] = am
                    st.session_state.pdf_zones[zi]["n_pages"] = npages
                    done += 1
                    progress_bar.progress(done / max(total_pdfs,1) * 0.25)
                    log(f"   ✅ {npages:,} pages · {len(am):,} unique codes · zone: {zslot['zone_tag']}")

                progress_bar.progress(0.25)

                # Build zone lookup: {zone_tag_lower: (art_map, bytes)}
                zone_lookup = {}
                for zslot in st.session_state.pdf_zones:
                    if zslot["art_map"]:
                        zone_lookup[zslot["zone_tag"].lower()] = (zslot["art_map"], zslot["bytes"])

                pbo_map    = st.session_state.pdf_pbo["art_map"]  or {}
                npbo_map   = st.session_state.pdf_npbo["art_map"] or {}
                pbo_bytes  = st.session_state.pdf_pbo["bytes"]
                npbo_bytes = st.session_state.pdf_npbo["bytes"]

                all_zone_codes = set()
                for am, _ in zone_lookup.values():
                    all_zone_codes.update(am.keys())
                all_pdf_codes = all_zone_codes | set(pbo_map.keys()) | set(npbo_map.keys())
                log(f"Total unique codes across all PDFs: {len(all_pdf_codes):,} ({len(zone_lookup)} zone PDFs)")

                # ── Phase 2: Prepare SOH data ─────────────────────────
                log("📊 Processing SOH file...")
                df = st.session_state.stock_df.copy()
                df["_zone"]  = df[col_zone].astype(str).str.strip()
                df["_aam"]   = df[col_aam].astype(str).str.strip()
                df["_store"] = df[col_store].astype(str).str.strip()
                df["_art"]   = df[col_article].apply(norm)
                df["_soh"]   = pd.to_numeric(df[col_soh], errors="coerce").fillna(0)
                df["_cat"]   = df[col_category].astype(str).str.strip()

                # Step A: Classify PBO vs NPBO by category
                pbo_cats_lower = [c.lower() for c in (pbo_cats or [])]
                df["_version"] = df["_cat"].apply(
                    lambda c: "PBO" if c.lower() in pbo_cats_lower else "NPBO"
                )

                # Step B: Clean blanks
                df = df[
                    df["_art"].notna() &
                    (df["_store"].str.len() > 0) & (df["_store"] != "nan") &
                    (df["_aam"].str.len() > 0)   & (df["_aam"]   != "nan") &
                    (df["_zone"].str.len() > 0)  & (df["_zone"]  != "nan")
                ]

                # Step C: Filter by selected zones first
                if selected_zones:
                    df = df[df["_zone"].isin(selected_zones)]
                    log(f"Zone filter: {', '.join(selected_zones)} → {len(df):,} records")
                else:
                    log(f"All zones → {len(df):,} records")

                # Step D: Filter by selected stores
                df = df[df["_store"].isin(selected_stores)]
                log(f"Store filter: {df['_store'].nunique()} stores → {len(df):,} records")

                # Step E: Filter SOH > 0
                if filter_soh:
                    df = df[df["_soh"] > 0]
                    log(f"SOH > 0 filter → {len(df):,} records")

                # Step F: Build zone PDF lookup with robust matching
                # Normalize both zone names from SOH and zone tags from filenames
                def normalize_zone(s):
                    """Lowercase, remove spaces/dashes, strip 'zone' prefix for comparison."""
                    s = str(s).lower().strip()
                    s = re.sub(r'[\s\-_]', '', s)   # remove spaces, dashes, underscores
                    return s

                zone_lookup_norm = {}  # {normalized_tag: (art_map, bytes, original_tag)}
                for zslot in st.session_state.pdf_zones:
                    if zslot["art_map"]:
                        nk = normalize_zone(zslot["zone_tag"])
                        zone_lookup_norm[nk] = (zslot["art_map"], zslot["bytes"], zslot["zone_tag"])

                def find_zone_pdf(zone_name):
                    """Find the matching zone PDF art_map and bytes for a given zone name."""
                    nz = normalize_zone(zone_name)
                    # Exact match first
                    if nz in zone_lookup_norm:
                        return zone_lookup_norm[nz][0], zone_lookup_norm[nz][1]
                    # Partial match: SOH zone contains tag or tag contains SOH zone
                    for nk, (zam, zbyt, orig) in zone_lookup_norm.items():
                        if nk in nz or nz in nk:
                            return zam, zbyt
                    return {}, None  # No zone PDF for this zone → use master only

                log(f"Zone PDFs indexed: {list(zone_lookup_norm.keys())}")
                log(f"Zones in SOH: {sorted(df['_zone'].unique())[:10]}")

                progress_bar.progress(0.35)

                # ── Phase 3: Generate one PBO + one NPBO per store ────
                # Each store gets 2 files: PBO and NPBO
                # Logic per file:
                #   1. Get all articles for this store+version (from SOH, SOH>0, correct category)
                #   2. For each article, find label page:
                #      a. Check zone PDF for this store's zone → use if found (priority)
                #      b. Else check master PBO or NPBO PDF → use if found
                #      c. Else → missing
                #   3. No duplicates — each article appears once

                combos = df[["_zone","_aam","_store"]].drop_duplicates().sort_values(["_zone","_aam","_store"])
                total  = len(combos) * 2
                log(f"📦 Generating {len(combos)} stores × 2 versions = {total} PDFs...")

                zip_buf    = io.BytesIO()
                store_rows = []
                job_idx    = 0

                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for _, combo in combos.iterrows():
                        zone_v  = combo["_zone"]
                        aam_v   = combo["_aam"]
                        store_v = combo["_store"]

                        # Get the zone PDF for this store's zone
                        zone_am, zone_byt = find_zone_pdf(zone_v)

                        # Full group for this store
                        grp = df[
                            (df["_zone"]==zone_v) &
                            (df["_aam"]==aam_v) &
                            (df["_store"]==store_v)
                        ]

                        for version in ["PBO", "NPBO"]:
                            job_idx += 1
                            pct = 0.35 + (job_idx / max(total,1)) * 0.58
                            progress_bar.progress(min(pct, 0.93))
                            log(f"[{job_idx}/{total}] {zone_v} · {aam_v} · {store_v} · {version}")

                            # Articles for this version only
                            grp_ver   = grp[grp["_version"] == version]
                            art_codes = set(grp_ver["_art"].dropna().unique())

                            if not art_codes:
                                store_rows.append({
                                    "Zone": zone_v, "AAM": aam_v, "Store": store_v,
                                    "Version": version, "SOH Articles": 0,
                                    "Zone Price Labels": 0, "Master Labels": 0,
                                    "Total Labels": 0, "Missing": 0,
                                    "Output File": f"— (no {version} articles in SOH)"
                                })
                                continue

                            # Master map for this version
                            master_map = pbo_map  if version == "PBO"  else npbo_map
                            master_byt = pbo_bytes if version == "PBO" else npbo_bytes

                            # ── Assign each article to a source (no duplicates) ──
                            out_doc      = fitz.open()
                            seen_arts    = set()   # track which articles already added
                            zone_arts    = set()
                            master_arts  = set()
                            missing_arts = set()

                            # Pass 1: Zone PDF (priority)
                            zone_pages = []
                            for art in sorted(art_codes):
                                if art in seen_arts: continue
                                if art in zone_am:
                                    pages = zone_am[art]
                                    zone_pages.extend(pages)
                                    seen_arts.add(art)
                                    zone_arts.add(art)

                            if zone_pages and zone_byt:
                                src = fitz.open(stream=zone_byt, filetype="pdf")
                                for p in sorted(set(zone_pages)):
                                    out_doc.insert_pdf(src, from_page=p, to_page=p)
                                src.close()

                            # Pass 2: Master PDF fallback
                            master_pages = []
                            for art in sorted(art_codes):
                                if art in seen_arts: continue
                                if art in master_map:
                                    pages = master_map[art]
                                    master_pages.extend(pages)
                                    seen_arts.add(art)
                                    master_arts.add(art)

                            if master_pages and master_byt:
                                src = fitz.open(stream=master_byt, filetype="pdf")
                                for p in sorted(set(master_pages)):
                                    out_doc.insert_pdf(src, from_page=p, to_page=p)
                                src.close()

                            # Pass 3: Identify missing
                            for art in art_codes:
                                if art not in seen_arts:
                                    missing_arts.add(art)

                            fname = f"{safe_name(zone_v)} - {safe_name(aam_v)} - {safe_name(store_v)} - {version}.pdf"

                            if len(out_doc) > 0:
                                buf = io.BytesIO()
                                out_doc.save(buf)
                                zf.writestr(fname, buf.getvalue())
                                n_labels = len(out_doc)
                            else:
                                n_labels = 0
                            out_doc.close()

                            store_rows.append({
                                "Zone":              zone_v,
                                "AAM":               aam_v,
                                "Store":             store_v,
                                "Version":           version,
                                "SOH Articles":      len(art_codes),
                                "Zone Price Labels": len(zone_arts),
                                "Master Labels":     len(master_arts),
                                "Total Labels":      n_labels,
                                "Missing":           len(missing_arts),
                                "Output File":       fname if n_labels > 0 else "— (no matches)",
                            })

                # ── Phase 4: Validation report ────────────────────────
                progress_bar.progress(0.94)
                log("📊 Creating validation report...")
                rep_buf = io.BytesIO()
                if HAS_OPENPYXL:
                    with pd.ExcelWriter(rep_buf, engine="openpyxl") as wr:
                        # Summary
                        pbo_files  = sum(1 for r in store_rows if r["Version"]=="PBO"  and r["Total Labels"]>0)
                        npbo_files = sum(1 for r in store_rows if r["Version"]=="NPBO" and r["Total Labels"]>0)
                        summary = {
                            "Zones Generated":      ", ".join(selected_zones) if selected_zones else "All",
                            "Stores Generated":     len(selected_stores),
                            "PBO Categories":       ", ".join(pbo_cats) if pbo_cats else "None",
                            "SOH Filter":           "SOH > 0" if filter_soh else "All records",
                            "Master PBO PDF":       st.session_state.pdf_pbo["name"] or "Not uploaded",
                            "Master NPBO PDF":      st.session_state.pdf_npbo["name"] or "Not uploaded",
                            "Zone PDFs":            ", ".join(z["name"] for z in st.session_state.pdf_zones) or "None",
                            "PBO Output Files":     pbo_files,
                            "NPBO Output Files":    npbo_files,
                            "Total Output Files":   pbo_files + npbo_files,
                            "Generated At":         datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        }
                        pd.DataFrame(list(summary.items()), columns=["Metric","Value"]) \
                          .to_excel(wr, sheet_name="Summary", index=False)

                        # All store results
                        pd.DataFrame(store_rows).to_excel(wr, sheet_name="Store Summary", index=False)

                        # PBO only
                        pbo_rows = [r for r in store_rows if r["Version"]=="PBO"]
                        if pbo_rows:
                            pd.DataFrame(pbo_rows).to_excel(wr, sheet_name="PBO Results", index=False)

                        # NPBO only
                        npbo_rows = [r for r in store_rows if r["Version"]=="NPBO"]
                        if npbo_rows:
                            pd.DataFrame(npbo_rows).to_excel(wr, sheet_name="NPBO Results", index=False)

                        # Missing articles
                        miss = [r for r in store_rows if r["Missing"] > 0]
                        if miss:
                            pd.DataFrame(miss)[["Zone","AAM","Store","Version","SOH Articles","Missing"]] \
                              .to_excel(wr, sheet_name="Missing Articles", index=False)

                st.session_state.zip_bytes  = zip_buf.getvalue()
                st.session_state.rep_bytes  = rep_buf.getvalue()
                st.session_state.store_rows = store_rows
                progress_bar.progress(1.0)

                pbo_count  = sum(1 for r in store_rows if r["Version"]=="PBO"  and r["Total Labels"]>0)
                npbo_count = sum(1 for r in store_rows if r["Version"]=="NPBO" and r["Total Labels"]>0)
                log(f"✅ Done! {pbo_count} PBO + {npbo_count} NPBO PDFs generated.")

            except Exception as e:
                st.error(f"❌ Error: {e}")
                st.code(traceback.format_exc())

    st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════
# DOWNLOAD
# ══════════════════════════════════════════════════════════════════════════
if st.session_state.zip_bytes:
    st.markdown("---")
    st.success("✅ Files ready to download!")

    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "📦 Download ZIP (All Store PDFs)",
            data=st.session_state.zip_bytes,
            file_name=f"Store_Price_Labels_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip",
            use_container_width=True, type="primary", key="dl_zip")
    with d2:
        if st.session_state.rep_bytes and len(st.session_state.rep_bytes) > 0:
            st.download_button(
                "📊 Download Validation Report",
                data=st.session_state.rep_bytes,
                file_name=f"validation_report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key="dl_rep")

    if st.session_state.store_rows:
        with st.expander("📋 View results"):
            rows = st.session_state.store_rows
            sc1, sc2 = st.columns([3,1])
            with sc1:
                srch = st.text_input("🔍 Search by Zone / AAM / Store", key="srch_res")
            with sc2:
                ver_filter = st.multiselect("Version", ["PBO","NPBO"], default=["PBO","NPBO"], key="vf")
            if srch:
                rows = [r for r in rows if any(srch.lower() in str(r[k]).lower() for k in ["Zone","AAM","Store"])]
            if ver_filter:
                rows = [r for r in rows if r["Version"] in ver_filter]

            rh = "".join(
                f"<tr>"
                f"<td>{r['Zone']}</td><td>{r['AAM']}</td><td>{r['Store']}</td>"
                f"<td><span class=\"{'badge-pbo' if r['Version']=='PBO' else 'badge-npbo'}\">{r['Version']}</span></td>"
                f"<td style='text-align:center'>{r['Articles']}</td>"
                f"<td style='text-align:center;color:#5B21B6'>{r['Zone Price Labels']}</td>"
                f"<td style='text-align:center;color:#1E40AF'>{r['Master Labels']}</td>"
                f"<td style='text-align:center;color:#166534'><b>{r['Total Labels']}</b></td>"
                f"<td style='text-align:center;color:#991B1B'>{r['Missing']}</td>"
                f"<td><span class='fname'>{r['Output File']}</span></td>"
                f"</tr>"
                for r in rows
            )
            st.markdown(
                f'<table class="stbl"><thead><tr>'
                f'<th>Zone</th><th>AAM</th><th>Store</th><th>Ver</th>'
                f'<th>Articles</th><th>Zone Labels</th><th>Master Labels</th>'
                f'<th>Total</th><th>Missing</th><th>Output File</th>'
                f'</tr></thead><tbody>{rh}</tbody></table>',
                unsafe_allow_html=True)

st.markdown("---")
st.caption("Price Label Filter & Store PDF Generator — No data retained after session ends.")
