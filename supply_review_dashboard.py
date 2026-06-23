"""
Supply Review MDS Dashboard — Cummins India
"""

import streamlit as st
import pandas as pd
import datetime
import io
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Supply Review Dashboard", page_icon="⚙️", layout="wide")

st.markdown("""
<style>
.stApp { background-color: #FFFFFF; }
.block-container { padding-top: 1.2rem; padding-bottom: 2rem; }

/* sidebar */
div[data-testid="stSidebar"] { background: #f7f7f7; border-right: 1px solid #ddd; }

/* title */
.dash-title {
    border-left: 4px solid #C00000;
    padding: 6px 0 6px 14px;
    margin-bottom: 18px;
}
.dash-title h1 { font-size: 1.15rem; font-weight: 700; margin: 0; color: #1A1A1A; }
.dash-title p  { font-size: 0.75rem; color: #666; margin: 2px 0 0; }

/* summary cards */
.summary-cards { display: flex; gap: 12px; margin-bottom: 18px; flex-wrap: wrap; }
.card {
    border: 1px solid #e0e0e0;
    border-top: 3px solid #C00000;
    border-radius: 3px;
    padding: 10px 16px;
    min-width: 140px;
    background: #fff;
}
.card .c-label { font-size: 0.68rem; color: #888; text-transform: uppercase; letter-spacing: .4px; }
.card .c-val   { font-size: 1.25rem; font-weight: 700; color: #1A1A1A; line-height: 1.2; }
.card .c-sub   { font-size: 0.72rem; color: #666; }

/* main segment table */
table.seg-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.80rem;
    font-family: 'Segoe UI', Arial, sans-serif;
}
table.seg-table thead tr:first-child th {
    background: #1A1A1A;
    color: #fff;
    padding: 7px 10px;
    text-align: center;
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: .3px;
    border-right: 1px solid #333;
    white-space: nowrap;
}
table.seg-table thead tr:first-child th.left { text-align: left; }
table.seg-table thead tr:second-child th,
table.seg-table thead tr + tr th {
    background: #3a3a3a;
    color: #ddd;
    padding: 5px 10px;
    text-align: center;
    font-size: 0.72rem;
    font-weight: 500;
    border-right: 1px solid #444;
    white-space: nowrap;
}
table.seg-table tbody td {
    padding: 6px 10px;
    border-bottom: 1px solid #ebebeb;
    text-align: center;
    color: #1A1A1A;
    font-size: 0.80rem;
    white-space: nowrap;
    border-right: 1px solid #f0f0f0;
}
table.seg-table tbody td.left { text-align: left; font-weight: 600; }
table.seg-table tbody tr:hover td { background: #fafafa; }
table.seg-table tbody tr.total-row td {
    background: #f5f5f5;
    font-weight: 700;
    border-top: 2px solid #C00000;
    color: #1A1A1A;
}
.delta-pos  { color: #1e6e1e; font-weight: 600; }
.delta-neg  { color: #C00000; font-weight: 600; }
.delta-zero { color: #aaa; }

/* section divider */
.section-title {
    font-size: 0.78rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .6px;
    color: #C00000;
    border-bottom: 1px solid #e0e0e0;
    padding-bottom: 4px;
    margin: 20px 0 10px;
}

/* item table */
table.item-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.78rem;
    font-family: 'Segoe UI', Arial, sans-serif;
}
table.item-table thead tr:first-child th {
    background: #2c2c2c;
    color: #fff;
    padding: 6px 10px;
    text-align: center;
    font-size: 0.73rem;
    font-weight: 600;
    border-right: 1px solid #444;
    white-space: nowrap;
}
table.item-table thead tr:first-child th.left { text-align: left; }
table.item-table thead tr + tr th {
    background: #404040;
    color: #ccc;
    padding: 5px 10px;
    font-size: 0.70rem;
    font-weight: 500;
    text-align: center;
    border-right: 1px solid #555;
    white-space: nowrap;
}
table.item-table tbody td {
    padding: 5px 10px;
    border-bottom: 1px solid #ececec;
    text-align: center;
    font-size: 0.78rem;
    color: #1A1A1A;
    white-space: nowrap;
    border-right: 1px solid #f5f5f5;
}
table.item-table tbody td.left { text-align: left; }
table.item-table tbody tr:nth-child(even) td { background: #fafafa; }
table.item-table tbody tr:hover td { background: #f0f4f8; }
</style>
""", unsafe_allow_html=True)

# ── helpers ──────────────────────────────────────────────────────────────────
def excel_date(serial):
    try:
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(serial))
    except Exception:
        return None

def safe(val):
    if val is None or not isinstance(val, (int, float)):
        return 0.0
    return float(val)

def fmt_q(v):
    return "–" if v == 0 else str(int(round(v)))

def fmt_v(v):
    return "–" if v == 0 else f"₹{v:,.2f} Cr"

def delta_html(d):
    if d == 0:
        return '<span class="delta-zero">–</span>'
    sign = "+" if d > 0 else ""
    cls = "delta-pos" if d > 0 else "delta-neg"
    return f'<span class="{cls}">{sign}{int(round(d))}</span>'

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Export Data')
    return output.getvalue()

@st.cache_data(show_spinner=False)
def parse_file(uploaded_bytes, filename):
    import io, pyxlsb
    buf = io.BytesIO(uploaded_bytes)
    with pyxlsb.open_workbook(buf) as wb:
        if 'MDS' not in wb.sheets:
            return None, None
        with wb.get_sheet('MDS') as ws:
            all_rows = [[c.v for c in row] for row in ws.rows()]

    header_row_idx = next(
        (i for i, row in enumerate(all_rows[:15]) if 'Plant' in row and 'Segment' in row),
        None
    )
    if header_row_idx is None:
        return None, None

    header = all_rows[header_row_idx]

    bl_indices = [i for i, v in enumerate(header) if v == 'B/L']
    month_cols = []
    for bl_idx in bl_indices:
        after = [(i, v) for i, v in enumerate(header) if i > bl_idx and isinstance(v, float) and v > 40000]
        if len(after) >= 3:
            month_cols = [after[0][0], after[1][0], after[2][0]]
            break
    if not month_cols:
        return None, None

    months = []
    for ci in month_cols:
        d = excel_date(header[ci])
        months.append(d.strftime("%b %Y") if d else f"col{ci}")

    dem_q_total = next((i for i, v in enumerate(header) if v == 'Q2 Dem (Q)'), None)
    dem_v_total = next((i for i, v in enumerate(header) if v == 'Q2 Dem (V)'), None)
    sup_q_total = next((i for i, v in enumerate(header) if v == 'Q2 Sup (Q)'), None)
    sup_v_total = next((i for i, v in enumerate(header) if v == 'Q2 Sup (V)'), None)
    if None in (dem_q_total, dem_v_total, sup_q_total, sup_v_total):
        return None, None

    dem_q_cols = [dem_q_total - 3, dem_q_total - 2, dem_q_total - 1]
    dem_v_cols = [dem_v_total - 3, dem_v_total - 2, dem_v_total - 1]
    sup_q_cols = [sup_q_total - 3, sup_q_total - 2, sup_q_total - 1]
    sup_v_cols = [sup_v_total - 3, sup_v_total - 2, sup_v_total - 1]

    records = []
    for row in all_rows[header_row_idx + 1:]:
        if len(row) <= max(sup_v_cols):
            continue
        if not row[0] or not isinstance(row[0], str):
            continue
        seg = row[7]
        if not seg or not isinstance(seg, str) or seg in ('Segment', 'RAIL'):
            continue
        seg = seg.strip()
        if seg.upper() == 'RAIL':
            seg = 'Rail'
        records.append({
            'segment':  seg,
            'item':     str(row[1]) if row[1] else '',
            'customer': str(row[8]) if row[8] else '',
            'category': str(row[2]) if row[2] else '',
            'plant':    str(row[0]),
            'dem_m1_q': safe(row[dem_q_cols[0]]), 'dem_m2_q': safe(row[dem_q_cols[1]]), 'dem_m3_q': safe(row[dem_q_cols[2]]),
            'dem_q':    safe(row[dem_q_total]),
            'dem_m1_v': safe(row[dem_v_cols[0]]), 'dem_m2_v': safe(row[dem_v_cols[1]]), 'dem_m3_v': safe(row[dem_v_cols[2]]),
            'dem_v':    safe(row[dem_v_total]),
            'sup_m1_q': safe(row[sup_q_cols[0]]), 'sup_m2_q': safe(row[sup_q_cols[1]]), 'sup_m3_q': safe(row[sup_q_cols[2]]),
            'sup_q':    safe(row[sup_q_total]),
            'sup_m1_v': safe(row[sup_v_cols[0]]), 'sup_m2_v': safe(row[sup_v_cols[1]]), 'sup_m3_v': safe(row[sup_v_cols[2]]),
            'sup_v':    safe(row[sup_v_total]),
        })

    return pd.DataFrame(records), months

# ── sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Supply Review MDS")
    st.markdown("---")
    st.markdown("**Upload Files**")

    f1 = st.file_uploader("M1 — First MDS",  type=["xlsb"], key="f1")
    f2 = st.file_uploader("M2 — Mid-Quarter", type=["xlsb"], key="f2")
    f3 = st.file_uploader("M3 — End of Quarter", type=["xlsb"], key="f3")

    st.markdown("---")
    view_mode = st.radio("View", ["Demand vs Supply", "Supply Plan Detail"], index=0)

# ── parse files ───────────────────────────────────────────────────────────────
file_data = {}
for label, f in [("M1", f1), ("M2", f2), ("M3", f3)]:
    if f:
        df, months = parse_file(f.read(), f.name)
        if df is not None:
            file_data[label] = {"df": df, "months": months}

# ── title ─────────────────────────────────────────────────────────────────────
loaded_labels = [k for k in ["M1","M2","M3"] if k in file_data]
sub = ", ".join(
    f"{k}: {file_data[k]['months'][0]}–{file_data[k]['months'][2]}"
    for k in loaded_labels
) if loaded_labels else "Upload M1 / M2 / M3 files to begin"

st.markdown(f"""
<div class="dash-title">
  <h1>Supply Review — MDS Dashboard</h1>
  <p>{sub}</p>
</div>
""", unsafe_allow_html=True)

if not file_data:
    st.info("Upload supply review .xlsb files using the sidebar.")
    st.stop()

# ── collect segments ──────────────────────────────────────────────────────────
all_segments = sorted(set(
    s for v in file_data.values()
    for s in v["df"]["segment"].unique()
    if s and s != 'Export'
))

active_segments = st.multiselect(
    "Segments", options=all_segments, default=all_segments, key="seg_filter"
)
if not active_segments:
    active_segments = all_segments

# ── summary cards ─────────────────────────────────────────────────────────────
cols = st.columns(len(file_data) * 2)
ci = 0
for k in ["M1","M2","M3"]:
    if k not in file_data:
        continue
    df = file_data[k]["df"]
    months = file_data[k]["months"]
    filt = df[df['segment'].isin(active_segments)]
    with cols[ci]:
        st.markdown(f"""<div class="card">
          <div class="c-label">{k} Demand</div>
          <div class="c-val">{int(filt['dem_q'].sum())}</div>
          <div class="c-sub">{months[0]} – {months[2]}</div>
        </div>""", unsafe_allow_html=True)
    with cols[ci+1]:
        st.markdown(f"""<div class="card">
          <div class="c-label">{k} Supply Plan</div>
          <div class="c-val">{int(filt['sup_q'].sum())}</div>
          <div class="c-sub">{fmt_v(filt['sup_v'].sum())}</div>
        </div>""", unsafe_allow_html=True)
    ci += 2

st.markdown("---")

# ── segment-level aggregates ──────────────────────────────────────────────────
def seg_agg(k):
    df = file_data[k]["df"]
    return df[df['segment'].isin(active_segments)].groupby('segment').agg(
        dem_q=('dem_q','sum'), dem_v=('dem_v','sum'),
        sup_q=('sup_q','sum'), sup_v=('sup_v','sum'),
        sup_m1_q=('sup_m1_q','sum'), sup_m2_q=('sup_m2_q','sum'), sup_m3_q=('sup_m3_q','sum'),
        sup_m1_v=('sup_m1_v','sum'), sup_m2_v=('sup_m2_v','sum'), sup_m3_v=('sup_m3_v','sum'),
    ).reset_index()

aggs = {k: seg_agg(k) for k in loaded_labels}

export_aggs = aggs[loaded_labels[0]].copy()
export_aggs.columns = [f"{loaded_labels[0]}_{col}" if col != 'segment' else col for col in export_aggs.columns]
for k in loaded_labels[1:]:
    temp_df = aggs[k].copy()
    temp_df.columns = [f"{k}_{col}" if col != 'segment' else col for col in temp_df.columns]
    export_aggs = export_aggs.merge(temp_df, on='segment', how='outer')

# ── main segment table ────────────────────────────────────────────────────────
if view_mode == "Demand vs Supply":
    h = '<table class="seg-table"><thead><tr>'
    h += '<th class="left" rowspan="2">Segment</th>'
    for k in loaded_labels:
        months = file_data[k]["months"]
        h += f'<th colspan="3">{k} &nbsp;({months[0]} – {months[2]})</th>'
    if "M1" in loaded_labels and "M2" in loaded_labels:
        h += '<th colspan="2">Δ M2 vs M1</th>'
    if "M2" in loaded_labels and "M3" in loaded_labels:
        h += '<th colspan="2">Δ M3 vs M2</th>'
    h += '</tr><tr>'
    for _ in loaded_labels:
        h += '<th>Demand (Q)</th><th>Supply (Q)</th><th>Gap (Q)</th>'
    if "M1" in loaded_labels and "M2" in loaded_labels:
        h += '<th>Supply Δ</th><th>Demand Δ</th>'
    if "M2" in loaded_labels and "M3" in loaded_labels:
        h += '<th>Supply Δ</th><th>Demand Δ</th>'
    h += '</tr></thead><tbody>'

    totals = {k: {'d':0,'s':0} for k in loaded_labels}
    rows_html = ''
    for seg in active_segments:
        rd = {}
        seg_total_activity = 0
        
        for k in loaded_labels:
            a = aggs[k]
            r = a[a['segment']==seg]
            rd[k] = {'d': r['dem_q'].sum(), 's': r['sup_q'].sum()}
            seg_total_activity += (rd[k]['d'] + rd[k]['s'])
            
        if seg_total_activity == 0:
            continue
            
        for k in loaded_labels:
            totals[k]['d'] += rd[k]['d']
            totals[k]['s'] += rd[k]['s']
            
        cells = ''
        for k in loaded_labels:
            d, s = rd[k]['d'], rd[k]['s']
            cells += f'<td>{fmt_q(d)}</td><td>{fmt_q(s)}</td><td>{delta_html(s-d)}</td>'
        if "M1" in loaded_labels and "M2" in loaded_labels:
            cells += f'<td>{delta_html(rd["M2"]["s"]-rd["M1"]["s"])}</td><td>{delta_html(rd["M2"]["d"]-rd["M1"]["d"])}</td>'
        if "M2" in loaded_labels and "M3" in loaded_labels:
            cells += f'<td>{delta_html(rd["M3"]["s"]-rd["M2"]["s"])}</td><td>{delta_html(rd["M3"]["d"]-rd["M2"]["d"])}</td>'
        rows_html += f'<tr><td class="left">{seg}</td>{cells}</tr>'

    t_cells = ''
    for k in loaded_labels:
        td, ts = totals[k]['d'], totals[k]['s']
        t_cells += f'<td>{fmt_q(td)}</td><td>{fmt_q(ts)}</td><td>{delta_html(ts-td)}</td>'
    if "M1" in loaded_labels and "M2" in loaded_labels:
        t_cells += f'<td>{delta_html(totals["M2"]["s"]-totals["M1"]["s"])}</td><td>{delta_html(totals["M2"]["d"]-totals["M1"]["d"])}</td>'
    if "M2" in loaded_labels and "M3" in loaded_labels:
        t_cells += f'<td>{delta_html(totals["M3"]["s"]-totals["M2"]["s"])}</td><td>{delta_html(totals["M3"]["d"]-totals["M2"]["d"])}</td>'
    rows_html += f'<tr class="total-row"><td class="left">TOTAL</td>{t_cells}</tr>'

    st.markdown(h + rows_html + '</tbody></table>', unsafe_allow_html=True)
    st.download_button(label="Export Segment Table to Excel", data=to_excel(export_aggs), file_name="segment_summary.xlsx")

else:
    h = '<table class="seg-table"><thead><tr><th class="left" rowspan="2">Segment</th>'
    for k in loaded_labels:
        h += f'<th colspan="6">{k}</th>'
    h += '</tr><tr>'
    for k in loaded_labels:
        for m in file_data[k]["months"]:
            h += f'<th>{m} (Q)</th>'
        for m in file_data[k]["months"]:
            h += f'<th>{m} (₹ Cr)</th>'
    h += '</tr></thead><tbody>'
    
    rows_html = ''
    for seg in active_segments:
        seg_total_activity = 0
        for k in loaded_labels:
            a = aggs[k]
            r = a[a['segment']==seg]
            seg_total_activity += r['sup_m1_q'].sum() + r['sup_m2_q'].sum() + r['sup_m3_q'].sum() + r['sup_m1_v'].sum() + r['sup_m2_v'].sum() + r['sup_m3_v'].sum()
            
        if seg_total_activity == 0:
            continue
            
        cells = ''
        for k in loaded_labels:
            a = aggs[k]
            r = a[a['segment']==seg]
            for col in ['sup_m1_q','sup_m2_q','sup_m3_q']:
                cells += f'<td>{fmt_q(r[col].sum())}</td>'
            for col in ['sup_m1_v','sup_m2_v','sup_m3_v']:
                v = r[col].sum()
                cells += f'<td>{"–" if v==0 else f"{v:.2f}"}</td>'
        rows_html += f'<tr><td class="left">{seg}</td>{cells}</tr>'
        
    st.markdown(h + rows_html + '</tbody></table>', unsafe_allow_html=True)
    st.download_button(label="Export Segment Plan to Excel", data=to_excel(export_aggs), file_name="segment_plan.xlsx")

# ── drilldown ─────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Item Detail</div>', unsafe_allow_html=True)

drill_col1, drill_col_p, drill_col2, drill_col3, drill_col4 = st.columns([1.5, 1.5, 1.5, 1.5, 2])
with drill_col1:
    seg_choice = st.multiselect("Segment", options=active_segments, key="seg_select")

segs_to_use = seg_choice if seg_choice else active_segments

with drill_col_p:
    plant_options = []
    for k in loaded_labels:
        plants = file_data[k]["df"][file_data[k]["df"]["segment"].isin(segs_to_use)]["plant"].unique().tolist()
        plant_options += [p for p in plants if p not in plant_options and p]
    plant_choice = st.multiselect("Plant", options=sorted(plant_options), key="plant_select")

with drill_col2:
    cat_options = []
    for k in loaded_labels:
        cats = file_data[k]["df"][file_data[k]["df"]["segment"].isin(segs_to_use)]["category"].unique().tolist()
        cat_options += [c for c in cats if c not in cat_options and c]
    cat_choice = st.multiselect("Category", options=sorted(cat_options), key="cat_select")

with drill_col3:
    cust_options = []
    for k in loaded_labels:
        custs = file_data[k]["df"][file_data[k]["df"]["segment"].isin(segs_to_use)]["customer"].unique().tolist()
        cust_options += [c for c in custs if c not in cust_options and c]
    cust_choice = st.multiselect("Customer", options=sorted(cust_options), key="cust_select")

with drill_col4:
    item_search = st.text_input("Search Item No", placeholder="e.g. A061E426", key="item_search")

dfs_to_merge = []
for k in loaded_labels:
    df_seg = file_data[k]["df"]
    df_seg = df_seg[df_seg['segment'].isin(segs_to_use)].copy()
    
    # Updated groupby to include 'plant'
    df_seg = df_seg.groupby(['item','customer','category', 'plant']).agg(
        dem_q=('dem_q','sum'), sup_q=('sup_q','sum'), sup_v=('sup_v','sum'),
        sup_m1_q=('sup_m1_q','sum'), sup_m2_q=('sup_m2_q','sum'), sup_m3_q=('sup_m3_q','sum'),
    ).reset_index()
    df_seg.columns = ['item','customer','category', 'plant',
                      f'{k}_dem_q', f'{k}_sup_q', f'{k}_sup_v',
                      f'{k}_sup_m1', f'{k}_sup_m2', f'{k}_sup_m3']
    dfs_to_merge.append(df_seg)

if not dfs_to_merge:
    st.stop()

merged = dfs_to_merge[0]
for extra in dfs_to_merge[1:]:
    merged = merged.merge(extra, on=['item','customer','category', 'plant'], how='outer').fillna(0)
merged = merged[merged['item'] != '']

if plant_choice:
    merged = merged[merged['plant'].isin(plant_choice)]
if cat_choice:
    merged = merged[merged['category'].isin(cat_choice)]
if cust_choice:
    merged = merged[merged['customer'].isin(cust_choice)]
if item_search.strip():
    merged = merged[merged['item'].str.contains(item_search.strip(), case=False, na=False)]

metric_cols = [c for c in merged.columns if any(k in c for k in loaded_labels) and ('dem' in c or 'sup' in c)]
merged['row_total'] = merged[metric_cols].abs().sum(axis=1)
merged = merged[merged['row_total'] > 0].drop(columns=['row_total'])

st.caption(f"{len(merged)} items — {', '.join(segs_to_use)}" + (f" · Plant: {', '.join(plant_choice)}" if plant_choice else "") + (f" · Category: {', '.join(cat_choice)}" if cat_choice else "") + (f" · Customer: {', '.join(cust_choice)}" if cust_choice else "") + (f" · Search: '{item_search}'" if item_search.strip() else ""))

show_months_toggle = st.toggle("Display Monthly Breakdown", value=False)

h = '<table class="item-table"><thead><tr>'
h += '<th class="left" rowspan="2">Item No</th><th class="left" rowspan="2">Customer</th><th class="left" rowspan="2">Category</th><th class="left" rowspan="2">Plant</th>'
for k in loaded_labels:
    months = file_data[k]["months"]
    col_span = 7 if show_months_toggle else 4
    h += f'<th colspan="{col_span}">{k} &nbsp;({months[0]}–{months[2]})</th>'
if "M1" in loaded_labels and "M2" in loaded_labels:
    h += '<th colspan="2">Δ M2 vs M1</th>'
if "M2" in loaded_labels and "M3" in loaded_labels:
    h += '<th colspan="2">Δ M3 vs M2</th>'
h += '</tr><tr>'
for k in loaded_labels:
    months = file_data[k]["months"]
    if show_months_toggle:
        h += f'<th>Dem (Q)</th><th>Sup (Q)</th><th>Sup (₹)</th><th>Var (Q)</th><th>{months[0]}</th><th>{months[1]}</th><th>{months[2]}</th>'
    else:
        h += f'<th>Dem (Q)</th><th>Sup (Q)</th><th>Sup (₹)</th><th>Var (Q)</th>'
if "M1" in loaded_labels and "M2" in loaded_labels:
    h += '<th>Sup Δ</th><th>Dem Δ</th>'
if "M2" in loaded_labels and "M3" in loaded_labels:
    h += '<th>Sup Δ</th><th>Dem Δ</th>'
h += '</tr></thead><tbody>'

rows = ''
for _, r in merged.iterrows():
    cells = f'<td class="left">{r["item"]}</td><td class="left">{r["customer"]}</td><td class="left">{r["category"]}</td><td class="left">{r["plant"]}</td>'
    for k in loaded_labels:
        dem = r.get(f"{k}_dem_q", 0)
        sup = r.get(f"{k}_sup_q", 0)
        var = sup - dem
        cells += (f'<td>{fmt_q(dem)}</td>'
                  f'<td>{fmt_q(sup)}</td>'
                  f'<td>{fmt_v(r.get(f"{k}_sup_v", 0))}</td>'
                  f'<td>{delta_html(var)}</td>')
        if show_months_toggle:
            cells += (f'<td>{fmt_q(r.get(f"{k}_sup_m1",0))}</td>'
                      f'<td>{fmt_q(r.get(f"{k}_sup_m2",0))}</td>'
                      f'<td>{fmt_q(r.get(f"{k}_sup_m3",0))}</td>')
    if "M1" in loaded_labels and "M2" in loaded_labels:
        cells += (f'<td>{delta_html(r.get("M2_sup_q",0)-r.get("M1_sup_q",0))}</td>'
                  f'<td>{delta_html(r.get("M2_dem_q",0)-r.get("M1_dem_q",0))}</td>')
    if "M2" in loaded_labels and "M3" in loaded_labels:
        cells += (f'<td>{delta_html(r.get("M3_sup_q",0)-r.get("M2_sup_q",0))}</td>'
                  f'<td>{delta_html(r.get("M3_dem_q",0)-r.get("M2_dem_q",0))}</td>')
    rows += f'<tr>{cells}</tr>'

st.markdown(h + rows + '</tbody></table>', unsafe_allow_html=True)
st.download_button(label="Export Item Drilldown to Excel", data=to_excel(merged), file_name="item_drilldown.xlsx")

# ── Dynamic Charts ────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Visualization & Cross-Quarter Analysis</div>', unsafe_allow_html=True)

if len(loaded_labels) > 1:
    st.markdown("**Cross-Quarter Trend Analysis**")
    
    t_col1, t_col2, t_col3 = st.columns([1.5, 1.5, 1.5])
    with t_col1:
        trend_seg = st.selectbox("Select Segment", active_segments, key="trend_seg")
        
    # Extract available plants and categories for the selected segment across all files
    avail_plants = set()
    avail_cats = set()
    for k in loaded_labels:
        df_k = file_data[k]["df"]
        seg_subset = df_k[df_k['segment'] == trend_seg]
        avail_plants.update(seg_subset['plant'].unique().tolist())
        avail_cats.update(seg_subset['category'].unique().tolist())
        
    avail_plants = sorted([p for p in avail_plants if p])
    avail_cats = sorted([c for c in avail_cats if c])
    
    # Row 2: Secondary Filters
    f_col1, f_col2, f_col3 = st.columns([1.5, 1.5, 1.5])
    with f_col1:
        trend_plants = st.multiselect("Filter by Plant", options=avail_plants, default=avail_plants, key="t_plant_filt")
    with f_col2:
        trend_cats = st.multiselect("Filter by Category", options=avail_cats, default=avail_cats, key="t_cat_filt")
    with f_col3:
        st.markdown("<br>", unsafe_allow_html=True)
        breakdown_dim = st.radio("Breakdown Dimension", ["None", "Category", "Plant"], horizontal=True, key="trend_breakdown")
        
    trend_data = []
    for k in loaded_labels:
        df_k = file_data[k]["df"]
        seg_df = df_k[df_k['segment'] == trend_seg]
        
        # Apply filters before aggregation
        if trend_plants:
            seg_df = seg_df[seg_df['plant'].isin(trend_plants)]
        if trend_cats:
            seg_df = seg_df[seg_df['category'].isin(trend_cats)]
        
        if breakdown_dim == "Category":
            grp = seg_df.groupby('category').agg(Demand=('dem_q', 'sum'), Supply=('sup_q', 'sum')).reset_index()
            for _, row in grp.iterrows():
                if row['Demand'] > 0 or row['Supply'] > 0:
                    trend_data.append({'Quarter': k, 'Category': row['category'], 'Metric': 'Demand', 'Quantity': row['Demand']})
                    trend_data.append({'Quarter': k, 'Category': row['category'], 'Metric': 'Supply', 'Quantity': row['Supply']})
        elif breakdown_dim == "Plant":
            grp = seg_df.groupby('plant').agg(Demand=('dem_q', 'sum'), Supply=('sup_q', 'sum')).reset_index()
            for _, row in grp.iterrows():
                if row['Demand'] > 0 or row['Supply'] > 0:
                    trend_data.append({'Quarter': k, 'Category': row['plant'], 'Metric': 'Demand', 'Quantity': row['Demand']})
                    trend_data.append({'Quarter': k, 'Category': row['plant'], 'Metric': 'Supply', 'Quantity': row['Supply']})
        else:
            total_dem = seg_df['dem_q'].sum()
            total_sup = seg_df['sup_q'].sum()
            trend_data.append({'Quarter': k, 'Category': 'Total', 'Metric': 'Demand', 'Quantity': total_dem})
            trend_data.append({'Quarter': k, 'Category': 'Total', 'Metric': 'Supply', 'Quantity': total_sup})
            
    trend_df = pd.DataFrame(trend_data)
    
    if not trend_df.empty:
        trend_df['Quarter'] = pd.Categorical(trend_df['Quarter'], categories=loaded_labels, ordered=True)
        trend_df = trend_df.sort_values(['Quarter', 'Metric'])
        
        fig_trend = go.Figure()
        
        if breakdown_dim in ["Category", "Plant"]:
            categories = trend_df['Category'].unique()
            colors = px.colors.qualitative.Plotly
            
            for i, cat in enumerate(categories):
                cat_df = trend_df[trend_df['Category'] == cat]
                
                dem_df = cat_df[cat_df['Metric'] == 'Demand']
                fig_trend.add_trace(go.Bar(
                    x=dem_df['Quarter'], y=dem_df['Quantity'],
                    name=f"{cat} (Demand)", offsetgroup=0, legendgroup='Demand',
                    marker_color=colors[i % len(colors)], opacity=0.6,
                    hoverinfo='name+y'
                ))
                
                sup_df = cat_df[cat_df['Metric'] == 'Supply']
                fig_trend.add_trace(go.Bar(
                    x=sup_df['Quarter'], y=sup_df['Quantity'],
                    name=f"{cat} (Supply)", offsetgroup=1, legendgroup='Supply',
                    marker_color=colors[i % len(colors)], opacity=1.0,
                    hoverinfo='name+y'
                ))
                
            total_dem_tr = trend_df[trend_df['Metric'] == 'Demand'].groupby('Quarter')['Quantity'].sum().reset_index()
            total_sup_tr = trend_df[trend_df['Metric'] == 'Supply'].groupby('Quarter')['Quantity'].sum().reset_index()
            
            fig_trend.add_trace(go.Scatter(x=total_dem_tr['Quarter'], y=total_dem_tr['Quantity'], mode='lines+markers', name='Total Demand Trend', line=dict(color='black', dash='dot', width=2)))
            fig_trend.add_trace(go.Scatter(x=total_sup_tr['Quarter'], y=total_sup_tr['Quantity'], mode='lines+markers', name='Total Supply Trend', line=dict(color='red', width=2)))
            
            fig_trend.update_layout(barmode='stack', title=f"Supply vs Demand Trajectory: {trend_seg} (by {breakdown_dim})", xaxis_title="Quarter Marker", yaxis_title="Quantity")
            
        else:
            dem_df = trend_df[trend_df['Metric'] == 'Demand']
            sup_df = trend_df[trend_df['Metric'] == 'Supply']
            
            fig_trend.add_trace(go.Bar(x=dem_df['Quarter'], y=dem_df['Quantity'], name='Demand', marker_color='#1f77b4'))
            fig_trend.add_trace(go.Bar(x=sup_df['Quarter'], y=sup_df['Quantity'], name='Supply', marker_color='#2ca02c'))
            
            fig_trend.add_trace(go.Scatter(x=dem_df['Quarter'], y=dem_df['Quantity'], mode='lines+markers', name='Demand Trend Line', line=dict(color='#17becf', dash='dot', width=3)))
            fig_trend.add_trace(go.Scatter(x=sup_df['Quarter'], y=sup_df['Quantity'], mode='lines+markers', name='Supply Trend Line', line=dict(color='#d62728', width=3)))
            
            fig_trend.update_layout(barmode='group', title=f"Supply vs Demand Trajectory: {trend_seg}", xaxis_title="Quarter Marker", yaxis_title="Quantity")
            
        st.plotly_chart(fig_trend, use_container_width=True)
else:
    st.info("Upload M2 and/or M3 files to generate cross-quarter trend analysis.")

st.markdown("---")

target_k = st.selectbox("Select Target Quarter for Static Analysis", loaded_labels, key="chart_quarter")

if target_k:
    chart_df = file_data[target_k]["df"]
    
    col_pie, col_bar = st.columns(2)
    
    with col_pie:
        st.markdown("**Demand Distribution by Segment**")
        pie_seg = st.selectbox("Filter Segment", active_segments, key="pie_seg")
        
        pie_data = chart_df[chart_df['segment'] == pie_seg]
        if not pie_data.empty:
            pie_data_grouped = pie_data.groupby('category')['dem_q'].sum().reset_index()
            fig_pie = px.pie(pie_data_grouped, names='category', values='dem_q', 
                             title=f"Category Breakdown ({pie_seg})", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("No data available for selected segment.")
            
    with col_bar:
        st.markdown("**Item-Level Supply vs Demand**")
        bar_seg = st.selectbox("Filter Segment for Items", active_segments, key="bar_seg")
        
        items_in_seg = chart_df[chart_df['segment'] == bar_seg]['item'].unique().tolist()
        bar_items = st.multiselect("Select Specific Items", items_in_seg, 
                                   default=items_in_seg[:5] if items_in_seg else [], key="bar_items")
        
        if bar_items:
            bar_data = chart_df[(chart_df['segment'] == bar_seg) & (chart_df['item'].isin(bar_items))]
            bar_data_grouped = bar_data.groupby('item').agg(Demand=('dem_q', 'sum'), Supply=('sup_q', 'sum')).reset_index()
            
            bar_melted = bar_data_grouped.melt(id_vars='item', value_vars=['Demand', 'Supply'], 
                                               var_name='Metric', value_name='Quantity')
            
            fig_bar = px.bar(bar_melted, x='item', y='Quantity', color='Metric', barmode='group', 
                             title=f"Variance Comparison ({target_k})",
                             color_discrete_map={'Demand': '#1f77b4', 'Supply': '#2ca02c'})
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("Select at least one item to display the bar chart.")