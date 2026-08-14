"""
Agovor Fleet Sizing Calculator (Leg-Row only / All-Rows)
=========================================================
Pipeline:
  1. Footprint hectares AND lineal metres by region -- read automatically from
     'Footprint_Rationalisation.xlsx' in the same folder as this app
     (the 'Footprint rationalisation' tab), aggregated straight to
     region-level Travel Rows / Leg Rows hectares AND lineal metres. A
     sidebar filter lets you pick ONE region -- everything below (hectares
     shown, lineal metres, and all fleet sizing) is scoped to that region only.
     -> Lineal metres are read DIRECTLY from the workbook's own "LINEAL
     METERS - Travel Rows / Leg Rows / All Rows" block (rows 55-68), which
     already applies each site's own lineal-metres-per-hectare rate (e.g.
     2-row vs 3-row tunnels use different rates in that workbook). This app
     no longer re-derives lineal metres itself with a single flat constant,
     so it automatically stays correct if the workbook's per-tunnel-type
     rates are ever revised -- no changes needed here.
  2. Mode: Leg-Rows only (Agovor LR)  OR  All Rows (Agovor mowing leg+travel rows)
  3. Agovor operating rate (from Corindi trip-time study), PLUS:
       - Mowing rotation (days) -- how often the same ground must be re-mown
         (grass grows faster in summer / slower in winter, so rotation replaces
         a fixed "summer/winter" toggle -- just set the rotation length directly)
       - Spraying operating rate; spraying rotation is derived from the mowing
         rotation via the Mow:Spray ratio (not entered separately)
       - Mow:Spray ratio -- single field like "1:1" or "1:2". 1:1 = mow and spray
         at the same time (same rotation). 1:2 = mow once, spray every 2 days
         (spraying rotation = mowing rotation x 2/1). Mowing and spraying each
         get their own dedicated weekly capacity -- this ratio does NOT split a
         shared pool of operating days between them.
       - Shift type (checkbox) -- Single Shift = 7 hrs/device/day (mornings
         only). Double Shift = 14 hrs/device/day (day + night operation).
         ASSUMPTION: this hour count is applied INDEPENDENTLY to both mowing
         and spraying (i.e. Double Shift gives mowing its own 14 hrs AND
         spraying its own 14 hrs on the same tractor, on their respective
         scheduled days) -- NOT split as 7+7 across the two tasks in one day.
         If that's not the intended meaning, this needs revisiting.

  UNITS -- STANDARDIZED ON LINEAL METRES/HR THROUGHOUT (both modes):
     Previously, Leg-Rows-only mode ran on ha/hr (matching how the Corindi
     trip-time study reported its throughput) while All-Rows mode ran on
     lineal m/hr (since All-Rows demand can only be expressed as a distance
     -- travel aisles get double-counted for the leg-rows off them). That
     meant the "Spray rate" input's unit flipped every time you changed
     Mode, and a stale ha/hr number left over from LR mode would silently
     distort the sprayer/tractor count if not re-entered after switching to
     All-Rows.
     Now: the LR-mode ha/hr entry is converted to lineal m/hr the instant
     it's entered (using the region's own blended m/ha, back-calculated from
     the workbook's actual Leg-Rows lineal metres and hectares -- see
     `_regional_lm_per_ha` below), and EVERY downstream calculation --
     demand, mowing rate, spray rate, capacities -- works in lineal m/hr
     regardless of which mode is selected. The Spray rate field always
     says "lineal m/hr" and never needs to be re-entered on a mode switch.
  4. Fleet sizing -> Tractors / Mowers / Sprayers required for the SELECTED
     REGION, PER YEAR, where annual demand = footprint lineal metres x
     (mowing or spraying cycles/year from each rotation), because the area
     gets revisited many times a year, not once. Tractors = the larger of
     mowers/sprayers required, since one tractor can't run both attachments
     at the same time.
  5. Fleet growth plan -> devices bought stay in service even if demand later
     dips (no disposals), so each year's owned fleet = the running peak of
     the region's requirement so far. Shows what's newly needed each year on
     top of what's already owned from prior years.

Run with:  streamlit run app.py
Requires: Footprint_Rationalisation.xlsx inside a 'Data' subfolder next to this script,
          i.e. <this script's folder>/Data/Footprint_Rationalisation.xlsx.
"""

import os
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Agovor Fleet Sizing Calculator", layout="wide")

YEARS = ["CY26", "CY27", "CY28", "CY29", "CY30"]

# Resolve relative to this script's own location (not the current working
# directory), so the app finds the file the same way whether it's launched
# with `streamlit run app.py` from this folder or from somewhere else.
APP_DIR = os.path.dirname(os.path.abspath(__file__))
FOOTPRINT_FILE = os.path.join(APP_DIR, "Data", "Footprint_Rationalisation_New.xlsx")
FOOTPRINT_SHEET = "Footprint rationalisation"

# Fixed layout of the source blocks within the sheet (Excel row numbers).
# These match the standard 'Footprint rationalisation' tab layout --
# adjust here directly if that layout ever changes.

# Hectares block ("Hectares - Travel Rows (Grande)" / "Hectares - Leg Rows")
HEADER_ROW = 27
FIRST_DATA_ROW = 29
LAST_DATA_ROW = 42

# Lineal Metres block ("LINEAL METERS - Travel Rows (Grande)" / "- Leg Rows"
# / "- All Rows (Agovor)") -- these are the workbook's OWN computed values,
# already reflecting each site's tunnel-specific lineal-metres/ha rate, so
# this app reads them as-is instead of re-deriving them from hectares.
LM_HEADER_ROW = 53
LM_FIRST_DATA_ROW = 55
LM_LAST_DATA_ROW = 68


def get_region(label: str) -> str:
    """Extract the region/site name from a 'Site / Tunnel type' label,
    e.g. 'Corindi - Blues (3x rows/tunnel)' -> 'Corindi'."""
    return str(label).split(" - ")[0].strip()


st.title("Agovor Fleet Sizing Calculator")
st.caption(
    "Leg-Rows only (LR) or All Rows (leg + travel combined) — from budgeted "
    "hectares by region through to how many Tractors, Mowers and Sprayers "
    "you'd need for the selected region, per year. All rates and demand now "
    "run in lineal metres/hr throughout, regardless of mode."
)
st.markdown(
    "`01 Footprint hectares by region` → `02 Leg/Travel split` "
    "→ `03 Mode, rotation, mow:spray ratio & shift` → `04 Fleet sizing (per year)` "
    "→ `05 Fleet growth plan (year-over-year carryover)`"
)

# ===========================================================================
# 01 — FOOTPRINT HECTARES + LINEAL METRES BY REGION (auto-read from local file)
# ===========================================================================
st.header("01 · Budgeted hectares by region")

if not os.path.exists(FOOTPRINT_FILE):
    st.error(
        f"Could not find **{FOOTPRINT_FILE}**. Place the workbook (containing a "
        f"'{FOOTPRINT_SHEET}' tab) in a **Data** subfolder next to this script and rerun."
    )
    st.stop()

try:
    raw = pd.read_excel(FOOTPRINT_FILE, sheet_name=FOOTPRINT_SHEET, header=None)
except Exception as e:
    st.error(f"Could not read '{FOOTPRINT_SHEET}' from {FOOTPRINT_FILE}: {e}")
    st.stop()

# ---------------------------------------------------------------------------
# Read HECTARES block (used only for the display table + the LR ha/hr ->
# lineal m/hr conversion below; fleet-sizing demand itself uses the lineal
# metres block read further down, not these hectares).
# ---------------------------------------------------------------------------
try:
    year_row_idx = HEADER_ROW  # year labels sit one row below the block header
    year_labels_travel = raw.iloc[year_row_idx, 0:5].tolist()   # cols A-E
    year_labels_leg = raw.iloc[year_row_idx, 11:16].tolist()    # cols L-P
    site_block = raw.iloc[FIRST_DATA_ROW - 1 : LAST_DATA_ROW]

    rows_out = []
    for i in range(len(site_block)):
        label = site_block.iloc[i, 5]  # col F
        if pd.isna(label):
            continue
        row = {"Site / Tunnel type": str(label)}
        for j, yr in enumerate(year_labels_travel):
            yr_str = str(yr).strip()
            if yr_str in YEARS:
                row[f"Travel_{yr_str}"] = pd.to_numeric(site_block.iloc[i, j], errors="coerce")
        for j, yr in enumerate(year_labels_leg):
            yr_str = str(yr).strip()
            if yr_str in YEARS:
                row[f"Leg_{yr_str}"] = pd.to_numeric(site_block.iloc[i, 11 + j], errors="coerce")
        rows_out.append(row)

    site_df = pd.DataFrame(rows_out)
    if site_df.empty:
        st.error(
            f"Parsed 0 site/tunnel rows from the Hectares block in '{FOOTPRINT_SHEET}' using "
            f"the fixed layout (header row {HEADER_ROW}, data rows {FIRST_DATA_ROW}-{LAST_DATA_ROW}). "
            "Check the file still matches this layout."
        )
        st.stop()
except Exception as e:
    st.error(
        f"Could not parse the Hectares block in '{FOOTPRINT_SHEET}' using the fixed layout "
        f"(header row {HEADER_ROW}, data rows {FIRST_DATA_ROW}-{LAST_DATA_ROW}): {e}"
    )
    st.stop()

# ---------------------------------------------------------------------------
# Read LINEAL METRES block DIRECTLY -- these are the workbook's own computed
# figures (already reflecting each site's tunnel-specific lineal-metres/ha
# rate), so this app does NOT recompute them from hectares with a flat
# constant. This is the fix for lineal metres appearing "stuck" on an old
# value even after the source workbook's per-tunnel-type rates changed --
# previously this app ignored this block entirely and derived its own
# (uniform) figure from hectares instead.
# ---------------------------------------------------------------------------
try:
    lm_year_labels_travel = raw.iloc[LM_HEADER_ROW, 0:5].tolist()    # cols A-E
    lm_year_labels_leg = raw.iloc[LM_HEADER_ROW, 11:16].tolist()     # cols L-P
    lm_year_labels_all = raw.iloc[LM_HEADER_ROW, 22:27].tolist()     # cols W-AA
    lm_site_block = raw.iloc[LM_FIRST_DATA_ROW - 1 : LM_LAST_DATA_ROW]

    lm_rows_out = []
    for i in range(len(lm_site_block)):
        label = lm_site_block.iloc[i, 5]  # col F
        if pd.isna(label):
            continue
        row = {"Site / Tunnel type": str(label)}
        for j, yr in enumerate(lm_year_labels_travel):
            yr_str = str(yr).strip()
            if yr_str in YEARS:
                row[f"TravelLm_{yr_str}"] = pd.to_numeric(lm_site_block.iloc[i, j], errors="coerce")
        for j, yr in enumerate(lm_year_labels_leg):
            yr_str = str(yr).strip()
            if yr_str in YEARS:
                row[f"LegLm_{yr_str}"] = pd.to_numeric(lm_site_block.iloc[i, 11 + j], errors="coerce")
        for j, yr in enumerate(lm_year_labels_all):
            yr_str = str(yr).strip()
            if yr_str in YEARS:
                row[f"AllLm_{yr_str}"] = pd.to_numeric(lm_site_block.iloc[i, 22 + j], errors="coerce")
        lm_rows_out.append(row)

    lm_site_df = pd.DataFrame(lm_rows_out)
    if lm_site_df.empty:
        st.error(
            f"Parsed 0 site/tunnel rows from the Lineal Metres block in '{FOOTPRINT_SHEET}' "
            f"using the fixed layout (header row {LM_HEADER_ROW}, data rows "
            f"{LM_FIRST_DATA_ROW}-{LM_LAST_DATA_ROW}). Check the file still matches this layout."
        )
        st.stop()
except Exception as e:
    st.error(
        f"Could not parse the Lineal Metres block in '{FOOTPRINT_SHEET}' using the fixed layout "
        f"(header row {LM_HEADER_ROW}, data rows {LM_FIRST_DATA_ROW}-{LM_LAST_DATA_ROW}): {e}"
    )
    st.stop()

# ---------------------------------------------------------------------------
# Aggregate to REGION level -- every region found in the file, before any
# selection is applied.
# ---------------------------------------------------------------------------
region_year_travel_all = {}
region_year_leg_all = {}

for _, row in site_df.iterrows():
    region = get_region(row["Site / Tunnel type"])
    region_year_travel_all.setdefault(region, {yr: 0.0 for yr in YEARS})
    region_year_leg_all.setdefault(region, {yr: 0.0 for yr in YEARS})
    for yr in YEARS:
        t = row.get(f"Travel_{yr}", 0)
        l = row.get(f"Leg_{yr}", 0)
        region_year_travel_all[region][yr] += t if pd.notna(t) else 0
        region_year_leg_all[region][yr] += l if pd.notna(l) else 0

region_year_travellm_all = {}
region_year_leglm_all = {}
region_year_alllm_all = {}

for _, row in lm_site_df.iterrows():
    region = get_region(row["Site / Tunnel type"])
    region_year_travellm_all.setdefault(region, {yr: 0.0 for yr in YEARS})
    region_year_leglm_all.setdefault(region, {yr: 0.0 for yr in YEARS})
    region_year_alllm_all.setdefault(region, {yr: 0.0 for yr in YEARS})
    for yr in YEARS:
        t = row.get(f"TravelLm_{yr}", 0)
        l = row.get(f"LegLm_{yr}", 0)
        a = row.get(f"AllLm_{yr}", 0)
        region_year_travellm_all[region][yr] += t if pd.notna(t) else 0
        region_year_leglm_all[region][yr] += l if pd.notna(l) else 0
        region_year_alllm_all[region][yr] += a if pd.notna(a) else 0

all_regions = sorted(region_year_travel_all.keys())

if not all_regions:
    st.error(f"No regions could be identified from '{FOOTPRINT_SHEET}'.")
    st.stop()

# ===========================================================================
# SIDEBAR — REGION FILTER (single selection, drives everything below)
# ===========================================================================
st.sidebar.header("Region filter")
selected_region = st.sidebar.selectbox(
    "Region",
    options=all_regions,
    help="Everything on this page -- hectares, lineal metres, and all fleet "
         "sizing calculations -- is scoped to whichever region is selected here.",
)

# From this point on, "regions" is just the one selected region, so every
# downstream calculation (Sections 02-05) is automatically scoped to it
# without needing separate filtering logic in each section.
regions = [selected_region]
region_year_travel = {selected_region: region_year_travel_all[selected_region]}
region_year_leg = {selected_region: region_year_leg_all[selected_region]}
region_year_travellm = {selected_region: region_year_travellm_all[selected_region]}
region_year_leglm = {selected_region: region_year_leglm_all[selected_region]}
region_year_alllm = {selected_region: region_year_alllm_all[selected_region]}

# Both single-pass footprint bases, ALWAYS in lineal metres, for either mode
# -- read directly from the workbook's own Lineal Metres block, NOT
# recomputed here. All Rows already accounts for travel aisles being
# traversed twice (in/out) to service the leg-rows off them, because that's
# how the workbook's own '=(A*2)+L' formula built the All Rows figure.
region_year_lm_legonly = {
    region: {yr: region_year_leglm[region][yr] for yr in YEARS} for region in regions
}
region_year_lm_allrows = {
    region: {yr: region_year_alllm[region][yr] for yr in YEARS} for region in regions
}


def _regional_lm_per_ha(region: str) -> float:
    """Blended lineal-metres-per-hectare for the region's Leg Rows, back-
    calculated from the workbook's own figures (leg Lm / leg Ha) rather than
    a flat constant -- used only to convert the LR-mode ha/hr operating-rate
    entry into lineal m/hr. Falls back to 0 if the region has no leg hectares
    (e.g. Soil-only sites), in which case the ha/hr entry can't be converted
    and mowing rate should be checked."""
    total_leg_ha = sum(region_year_leg[region][yr] for yr in YEARS)
    total_leg_lm = sum(region_year_leglm[region][yr] for yr in YEARS)
    if total_leg_ha <= 0:
        return 0.0
    return total_leg_lm / total_leg_ha


st.caption(
    f"Read from **Data/{os.path.basename(FOOTPRINT_FILE)}** ('{FOOTPRINT_SHEET}' tab) — hectares from rows "
    f"{FIRST_DATA_ROW}-{LAST_DATA_ROW}, lineal metres read directly from the workbook's own "
    f"computed Lineal Metres block (rows {LM_FIRST_DATA_ROW}-{LM_LAST_DATA_ROW}), so any "
    "tunnel-specific lineal-metres/ha rate already baked into that workbook is picked up "
    f"automatically. Showing region: **{selected_region}**."
)

# Region-level Travel Rows / Leg Rows summary for the selected region only
region_summary_rows = []
for region in regions:
    row_data = {"Region": region}
    for yr in YEARS:
        row_data[f"{yr} Ha-Travel"] = region_year_travel[region][yr]
        row_data[f"{yr} Ha-Leg"] = region_year_leg[region][yr]
    region_summary_rows.append(row_data)

region_summary_df = pd.DataFrame(region_summary_rows).set_index("Region")
st.dataframe(
    region_summary_df.style.format("{:,.1f}"),
    use_container_width=True,
)

# ===========================================================================
# 02 — LEG / TRAVEL SPLIT -> LINEAL METRES (selected region only)
# ===========================================================================
st.header(f"02 · Leg-row / Travel-row split -> Lineal metres — {selected_region}")

totals = {"Ha-Travel": {}, "Ha-Leg": {}, "Lm-LegOnly": {}, "Lm-AllRows": {}}
for yr in YEARS:
    totals["Ha-Travel"][yr] = region_year_travel[selected_region][yr]
    totals["Ha-Leg"][yr] = region_year_leg[selected_region][yr]
    totals["Lm-LegOnly"][yr] = region_year_leglm[selected_region][yr]
    totals["Lm-AllRows"][yr] = region_year_alllm[selected_region][yr]

totals_df = pd.DataFrame(
    {
        "Total Ha - Travel Rows": totals["Ha-Travel"],
        "Total Ha - Leg Rows": totals["Ha-Leg"],
        "Total Lineal Metres - Leg Rows only": totals["Lm-LegOnly"],
        "Total Lineal Metres - All Rows": totals["Lm-AllRows"],
    }
).T

st.dataframe(totals_df.style.format("{:,.1f}"), use_container_width=True)
st.caption(
    "Lineal metres above are read directly from the workbook's own 'LINEAL METERS' section "
    "(rows 53-75) — this app no longer re-derives them from hectares using a flat constant, so "
    "any per-tunnel-type lineal-metres/ha rate the workbook applies (e.g. different rates for "
    "2-row vs 3-row tunnels) is picked up automatically without needing changes here. "
    "These are single-pass (one mow) totals — Section 04 scales them up by mowing frequency."
)

# ===========================================================================
# 03 — MODE, ROTATION, MOW:SPRAY RATIO, SHIFT TYPE & DEVICE OPERATING RATE
# ===========================================================================
st.header("03 · Mode, mowing cadence & Agovor operating rate")

mode = st.radio(
    "Mowing mode",
    ["Leg Rows only (LR)", "All Rows (Leg + Travel combined)"],
    horizontal=True,
)

st.subheader("Mowing cadence")
cad1, cad2, cad3 = st.columns(3)
rotation_days = cad1.number_input(
    "Mowing rotation (days)",
    min_value=1, value=7, step=1,
    help="How often the same ground must be re-mown, driven by grass growth. "
         "Use a shorter rotation for fast summer growth (e.g. 7 days) and a longer "
         "one for slow winter growth (e.g. 14-21 days) — set directly instead of "
         "picking a season.",
)
mow_spray_ratio_text = cad2.text_input(
    "Mow:Spray ratio",
    value="1:1",
    help="How often spraying happens relative to mowing. 1:1 = mow and spray at the same "
         "time (spraying keeps pace with the mowing rotation). 1:2 = mow once, spray every "
         "2 days (spraying happens half as often as mowing).",
)

double_shift = cad3.checkbox(
    "Double Shift (Day + Night operation)",
    value=False,
    help="Unchecked = Single Shift (mornings only) = 7 hrs/device/day. "
         "Checked = Double Shift (day + night) = 14 hrs/device/day. "
         "This directly sets Hrs/device/day for both mowing and spraying below "
         "— it's no longer manually editable, since it's fully determined by "
         "the shift setting.",
)

device_hours_per_day = 14.0 if double_shift else 7.0

st.caption(
    f"**{'Double Shift' if double_shift else 'Single Shift'}** selected → "
    f"**{device_hours_per_day:g} hrs/device/day**, applied to BOTH mowing and spraying "
    "independently below (i.e. mowing gets its own full "
    f"{device_hours_per_day:g} hours and spraying gets its own full "
    f"{device_hours_per_day:g} hours, on their respective scheduled days — this is NOT "
    "split as half-and-half across the two tasks in a single day)."
)

try:
    _mow_part, _spray_part = mow_spray_ratio_text.split(":")
    mow_ratio_n = float(_mow_part.strip())
    spray_ratio_n = float(_spray_part.strip())
    if mow_ratio_n <= 0:
        raise ValueError
    _ratio_valid = True
except (ValueError, AttributeError):
    mow_ratio_n, spray_ratio_n = 1.0, 1.0
    _ratio_valid = False
    st.warning(f"⚠️ Couldn't parse '{mow_spray_ratio_text}' as a ratio like '1:2' — defaulting to 1:1.")

spray_rotation_days = max(1, round(rotation_days * (spray_ratio_n / mow_ratio_n)))

st.caption(
    f"Mow:Spray ratio {mow_ratio_n:g}:{spray_ratio_n:g} → spraying rotation = mowing rotation "
    f"({rotation_days:.0f}d) × {spray_ratio_n:g}/{mow_ratio_n:g} = **{spray_rotation_days} days**. "
    f"Mowing and spraying each get their own dedicated capacity below (they're not splitting the "
    f"same operating days — 1:1 means both happen on the same schedule, not that each only gets "
    f"half the week)."
)

# ---------------------------------------------------------------------------
# MOWING RATE — always entered in whatever unit the source study reported it
# in for that mode, but converted to lineal m/hr IMMEDIATELY so every
# downstream number (capacity, demand, spray rate) is in one consistent unit.
# For LR mode, the ha/hr -> lineal m/hr conversion uses the SELECTED REGION'S
# OWN blended Leg-Rows lineal-metres/ha (back-calculated from the workbook's
# figures via _regional_lm_per_ha), not a flat constant, so it stays
# consistent with whatever tunnel-specific rates that region actually uses.
# ---------------------------------------------------------------------------
col1, col2, col3, col4 = st.columns(4)

regional_lm_per_ha = _regional_lm_per_ha(selected_region)

if mode == "Leg Rows only (LR)":
    ha_per_hr = col1.number_input(
        "Hectares / hr (from Corindi trip-time study)",
        value=0.62, step=0.01,
        help="Entered in ha/hr since that's how the Corindi trip-time study reported it. "
             "Converted to lineal m/hr immediately below, using the selected region's own "
             "blended Leg-Rows lineal-metres/ha (from the workbook), so it matches All-Rows mode.",
    )
    lm_per_hr = ha_per_hr * regional_lm_per_ha
    hrs_per_day = device_hours_per_day
    col2.metric("Hrs / device / day (mowing)", f"{hrs_per_day:g}")
    days_per_week = col3.number_input("Operating days / week", value=5.0, step=0.5)
    weeks_per_year = col4.number_input("Weeks worked / year", value=48.0, step=1.0)
    if regional_lm_per_ha <= 0:
        st.warning(
            f"⚠️ **{selected_region}** has 0 leg hectares in the workbook, so a lineal-metres/ha "
            "rate can't be derived for it and the ha/hr rate above can't be converted. Mowing "
            "rate below will show as 0 lineal m/hr until this is resolved."
        )
    st.caption(
        f"Converted: {ha_per_hr:.3f} ha/hr × {regional_lm_per_ha:,.1f} m/ha "
        f"({selected_region}'s own blended Leg-Rows rate) = **{lm_per_hr:,.0f} lineal m/hr**."
    )
else:
    km_h = col1.number_input("Speed (km/h)", value=3.0, step=0.1)
    pct_cutting = col2.number_input("% time actually cutting", value=75.0, step=1.0) / 100
    hrs_per_day = device_hours_per_day
    col3.metric("Hrs / device / day (mowing)", f"{hrs_per_day:g}")
    days_per_week = col4.number_input("Operating days / week", value=5.0, step=0.5)
    weeks_per_year = st.number_input("Weeks worked / year", value=48.0, step=1.0)
    lm_per_hr = km_h * 1000 * pct_cutting  # lineal metres/hr actually cut
    st.caption(f"Derived: {lm_per_hr:,.0f} lineal m/hr actual cutting rate.")

# Everything from here on is in lineal metres, regardless of mode
demand_unit = "lineal m"
weekly_capacity = lm_per_hr * hrs_per_day * days_per_week  # lineal metres/week/device
annual_capacity_per_device = weekly_capacity * weeks_per_year

st.caption(
    f"Mowing capacity uses the full {days_per_week:g} operating days/week — this capacity "
    "assumption (weeks worked/year, hrs/day, days/week) applies to the selected region below. "
    "If sites have materially different mowing seasons or rotations, revisit this per region "
    "as you switch the sidebar selection."
)

if annual_capacity_per_device <= 0:
    st.warning(
        f"⚠️ Mowing capacity/device works out to **0 {demand_unit}/yr**, so mowers required will "
        f"show as 0. Check: mowing rate = {lm_per_hr:,.0f} lm/hr, "
        f"Hrs/device/day = {hrs_per_day:g}, Operating days/week = {days_per_week:g}, "
        f"Weeks/year = {weeks_per_year:g}. Any one of these at 0 zeroes out the whole fleet."
    )

st.subheader("Spraying operating rate")
st.caption(
    f"Assumes spraying covers the same footprint as mowing (same {demand_unit} basis from Section "
    "02) — only the rate and rotation differ. **The spray rate is now always in lineal m/hr, "
    "the same unit as the mowing rate above, no matter which Mode is selected** — this removes "
    "the #1 place fleet sizing used to go wrong: a stale rate left over from switching Mode used "
    "to silently distort the Sprayers (and therefore Tractors) count."
)
sp1, sp2 = st.columns(2)
_spray_rate_default = lm_per_hr
spray_rate = sp1.number_input(
    f"Spray rate ({demand_unit}/hr)",
    value=float(round(_spray_rate_default, 2)),
    step=10.0,
    key="spray_rate_lm",  # single fixed key -- no longer keyed by mode, since unit no longer changes
    help=f"Same order of magnitude as your mowing rate above ({_spray_rate_default:,.0f} {demand_unit}/hr) "
         "unless spraying is genuinely faster/slower per pass. Always lineal m/hr now, so this value "
         "carries over cleanly if you switch Mode.",
)
spray_hrs_per_day = device_hours_per_day
sp2.metric("Hrs / device / day (spraying)", f"{spray_hrs_per_day:g}")

st.caption(
    f"Spraying rotation is **{spray_rotation_days} days**, derived from the Mow:Spray ratio above "
    f"— spraying uses the same {days_per_week:g} operating days/week as mowing (its own dedicated "
    "capacity, not a shared/reduced share of the week)."
)

spray_weekly_capacity = spray_rate * spray_hrs_per_day * days_per_week
spray_annual_capacity_per_device = spray_weekly_capacity * weeks_per_year

if spray_rate > 0 and (spray_rate < _spray_rate_default / 20 or spray_rate > _spray_rate_default * 20):
    st.warning(
        f"⚠️ Spray rate ({spray_rate:g} {demand_unit}/hr) is a very different order of magnitude "
        f"from the mowing rate ({_spray_rate_default:,.0f} {demand_unit}/hr) — both are now in the "
        "same unit, so this gap is a genuine rate difference, not a units mix-up. Double check it's "
        "intentional; it will materially change the Sprayers (and therefore Tractors) count below."
    )

if spray_annual_capacity_per_device <= 0:
    st.warning(
        f"⚠️ Spraying capacity/device works out to **0 {demand_unit}/yr**, so sprayers required "
        f"will show as 0. Check: Spray rate = {spray_rate:g} {demand_unit}/hr, "
        f"Hrs/device/day (spraying) = {spray_hrs_per_day:g}, Operating days/week = {days_per_week:g}."
    )

# ===========================================================================
# 04 — FLEET SIZING (per year) — Tractors / Mowers / Sprayers, selected region
# ===========================================================================
st.header(f"04 · Fleet sizing — {selected_region}, per year")

season_days_per_year = weeks_per_year * 7  # calendar days spanned by the operating season
mowing_cycles_per_year = season_days_per_year / rotation_days if rotation_days > 0 else 1.0
spraying_cycles_per_year = season_days_per_year / spray_rotation_days if spray_rotation_days > 0 else 1.0

st.caption(
    f"Mowing rotation {rotation_days:.0f}d → **{mowing_cycles_per_year:.1f} mowing cycles/yr**. "
    f"Spraying rotation {spray_rotation_days:.0f}d → **{spraying_cycles_per_year:.1f} spraying "
    "cycles/yr**. The region's single-pass footprint (Section 02, read directly from the "
    "workbook's Lineal Metres block) is multiplied by the relevant cycle count, then divided by "
    "that task's annual capacity/device, rounded up. **Tractors** are the shared carrier — since "
    "one tractor can't mow and spray at the same time, tractor count per year is the larger of "
    "mowers-required or sprayers-required (the smaller task shares the same tractors on its "
    "scheduled days). Mowers and sprayers are the attachment counts needed to actually cover "
    "each task's demand."
)

# Single-pass footprint, always in lineal metres, matching the selected mode
demand_map_single_pass = region_year_lm_legonly if mode == "Leg Rows only (LR)" else region_year_lm_allrows

total_mow_demand_check = demand_map_single_pass[selected_region][YEARS[0]]
if total_mow_demand_check <= 0:
    basis = "Leg Rows only" if mode == "Leg Rows only (LR)" else "All Rows (Travel + Leg)"
    st.warning(
        f"⚠️ {YEARS[0]} mowing demand for **{selected_region}** is **0**, using the "
        f"'{basis}' basis. Check the Section 01 summary table above, or switch Mode above."
    )

mow_demand_map = {
    region: {yr: demand_map_single_pass[region][yr] * mowing_cycles_per_year for yr in YEARS}
    for region in regions
}
spray_demand_map = {
    region: {yr: demand_map_single_pass[region][yr] * spraying_cycles_per_year for yr in YEARS}
    for region in regions
}


def _size_fleet(demand_map, capacity_per_device):
    out = {}
    for region in regions:
        out[region] = {}
        for yr in YEARS:
            demand = demand_map[region][yr]
            if capacity_per_device <= 0 or demand <= 0:
                out[region][yr] = 0
            else:
                out[region][yr] = int(-(-demand // capacity_per_device))  # ceil
    return out


mowers_by_region_year = _size_fleet(mow_demand_map, annual_capacity_per_device)
sprayers_by_region_year = _size_fleet(spray_demand_map, spray_annual_capacity_per_device)
tractors_by_region_year = {
    region: {yr: max(mowers_by_region_year[region][yr], sprayers_by_region_year[region][yr]) for yr in YEARS}
    for region in regions
}


def _to_table(by_region_year):
    t = pd.DataFrame(by_region_year).T
    t = t[YEARS]
    return t


mowers_table = _to_table(mowers_by_region_year)
sprayers_table = _to_table(sprayers_by_region_year)
tractors_table = _to_table(tractors_by_region_year)

st.subheader("Units required")
units_summary = pd.DataFrame(
    {
        "Tractors": tractors_table.loc[selected_region],
        "Mowers": mowers_table.loc[selected_region],
        "Sprayers": sprayers_table.loc[selected_region],
    }
).T
st.dataframe(units_summary, use_container_width=True)
st.caption(
    "This is the raw per-year requirement per asset type for the selected region (can go up "
    "or down year to year); Section 05 below turns this into the actual fleet you'd own, since "
    "devices aren't disposed of if demand later dips."
)

with st.expander("Show Tractors by year"):
    st.dataframe(tractors_table, use_container_width=True)
    st.bar_chart(tractors_table.T)
with st.expander("Show Mowers by year"):
    st.dataframe(mowers_table, use_container_width=True)
    st.caption(f"Mowing capacity/device: {annual_capacity_per_device:,.1f} {demand_unit}/yr.")
with st.expander("Show Sprayers by year"):
    st.dataframe(sprayers_table, use_container_width=True)
    st.caption(f"Spraying capacity/device: {spray_annual_capacity_per_device:,.1f} {demand_unit}/yr.")

# ===========================================================================
# 05 — FLEET GROWTH PLAN (year-over-year carryover)
# ===========================================================================
st.header(f"05 · Fleet growth plan — {selected_region}, how CY27 carries over from CY26")
st.caption(
    "**How year progression works:** Section 04 gives the region's raw requirement for "
    "each year *independently* (as if starting from zero every year). That's not realistic — "
    "you don't sell devices off and rebuy them. So here, **owned fleet in a given year = the "
    "running maximum of the requirement up to and including that year**. If CY26 needs 5 "
    "tractors and CY27 only needs 4, you still own 5 in CY27 (no disposals) — the extra one "
    "just carries over unused. If CY27 needs 8, you only need to buy 3 *new* ones (8 − 5), "
    "because the 5 from CY26 are already in the fleet. That's the 'New devices that year' row "
    "below — it's what actually needs to be purchased/leased that year, not the running total."
)


def _growth(by_region_year):
    owned = {}
    for region in regions:
        running_max = 0
        owned[region] = {}
        for yr in YEARS:
            running_max = max(running_max, by_region_year[region][yr])
            owned[region][yr] = running_max
    owned_total = {yr: owned[selected_region][yr] for yr in YEARS}
    new_devices = {}
    prev = 0
    for yr in YEARS:
        new_devices[yr] = owned_total[yr] - prev
        prev = owned_total[yr]
    return owned, owned_total, new_devices


owned_by_region_year, owned_total_by_year, new_devices_by_year = _growth(tractors_by_region_year)
mowers_owned_by_region_year, mowers_owned_total_by_year, mowers_new_by_year = _growth(mowers_by_region_year)
sprayers_owned_by_region_year, sprayers_owned_total_by_year, sprayers_new_by_year = _growth(sprayers_by_region_year)

growth_summary = pd.DataFrame(
    {
        "Tractors owned (cumulative)": owned_total_by_year,
        "New tractors that year": new_devices_by_year,
        "Mowers owned (cumulative)": mowers_owned_total_by_year,
        "New mowers that year": mowers_new_by_year,
        "Sprayers owned (cumulative)": sprayers_owned_total_by_year,
        "New sprayers that year": sprayers_new_by_year,
    }
).T
st.dataframe(growth_summary, use_container_width=True)

fm1, fm2, fm3 = st.columns(3)
fm1.metric(f"Tractors owned by {YEARS[-1]}", f"{int(owned_total_by_year[YEARS[-1]])}")
fm2.metric(f"Mowers owned by {YEARS[-1]}", f"{int(mowers_owned_total_by_year[YEARS[-1]])}")
fm3.metric(f"Sprayers owned by {YEARS[-1]}", f"{int(sprayers_owned_total_by_year[YEARS[-1]])}")

with st.expander("Show carryover — Tractors (raw requirement vs. owned-with-carryover)"):
    raw_req_df = pd.DataFrame(tractors_by_region_year).T[YEARS]
    owned_df = pd.DataFrame(owned_by_region_year).T[YEARS]
    st.write("Raw requirement per year (Section 04, no carryover):")
    st.dataframe(raw_req_df, use_container_width=True)
    st.write("Owned fleet per year (running peak, no disposals — what you'd actually hold):")
    st.dataframe(owned_df, use_container_width=True)
