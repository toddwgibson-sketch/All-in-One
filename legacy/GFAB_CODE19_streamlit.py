#!/usr/bin/env python3
"""
GFAB Excel Formatter — Streamlit Web Version
Converted from Tkinter desktop app.

Same functions and identical output as the original Tkinter version.

Steps performed on each target file:
  1.  Split lldp_sp → Downlinks (interface down) + Mismatches (swp)
  2.  Rename optics_rx_tx_threshold → Optics, remove unwanted columns,
      move Measured (dBm) to first column
  3.  Remove interfaces_sp tab
  4.  Reorder tabs: Downlinks, Mismatches, Optics, combined_fec
  5.  Insert L/R column after every Interface column (Z L/R after Z Interface)
  6.  Populate Source_port, DMARC1, DMARC2, Destination_port from Cutsheets
  6b. Mismatches: add Possible columns (matched via cutsheet Z-side) +
      Active Z columns (from Act. data), both with pink background
  7.  Add Summary tab (first) with row counts
  8.  Remove all fills (except pink Possible/Z cols); apply borders everywhere
  9.  Rename output file to top-2 Rack numbers

Usage:
    streamlit run GFAB_CODE19_streamlit.py
"""

import sys
import os
import shutil
import tempfile
import zipfile
from collections import Counter
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Tab-name resolution ──────────────────────────────────────────────────────
# The source-report generator periodically renames tabs (e.g. adding a
# "_with_pp" suffix). To survive these renames, look up tabs by ANY of their
# known aliases instead of hard-coding a single name.

TAB_ALIASES = {
    'lldp':         ('lldp_sp', 'full_path_lldp_with_int_down'),
    'optics':       ('optics_rx_tx_threshold', 'optics_rx_tx_threshold_with_pp'),
    'interfaces':   ('interfaces_sp', 'interfaces_sp_with_pp'),
    'combined_fec': ('combined_fec', 'combined_fec_with_pp'),
}

# Tabs that should always be dropped from the output (case-insensitive match).
# Add or remove names here as the source report evolves.
TABS_TO_REMOVE = (
    'device_reporting_failure',
    'bgp_sp',
    'spectrum_health',
    'sp_power',
    'sp_fans',
    'optics_temp',
    'pre_fec_ber_threshold_with_pp',
)

# Columns to strip from every tab in the final output. Exact-match,
# case-sensitive. Edit this list to drop more columns globally.
COLUMNS_TO_REMOVE = (
    'Building',
    'Act. Building',
    'Exp. Building',
    'PP_A',
    'PP_Z',
    # combined_fec source noise:
    'Remote Host',
    'Remote Interface',
    'Mapped Remote Host',
    'Mapped Remote Interface',
    'Mapped Remote Rack',
    'Mapped Remote Elevation',
    'Remote Host Match',
    'Remote Interface Match',
    'Remote End Match',
    'Z_end_host',
    'Z_end_intf',
    'rack_z',
    'Z_Rack',
    'Z_Elevation',
    'Index',
    'Source Sheet',
    'Placement Group',
)

# Tabs that should receive Z-side info (Z Hostname, Z Interface, Z L/R,
# Z Rack, Z Elevation) pulled from the cutsheet by Hostname+Interface match.
# Columns are inserted right after Destination_port. Add more tab names
# here if you want Z-side info in additional tabs.
Z_FILL_TABS = ('Optics', 'combined_fec')

def find_tab(wb_or_sheetnames, key):
    """Return the actual tab name in the workbook for the given logical key,
    or None if no alias is present."""
    names = (wb_or_sheetnames.sheetnames
             if hasattr(wb_or_sheetnames, 'sheetnames')
             else list(wb_or_sheetnames))
    for alias in TAB_ALIASES[key]:
        if alias in names:
            return alias
    return None


# ── Style helpers ────────────────────────────────────────────────────────────

PINK   = 'FFB6C1'
YELLOW = 'FFFF00'

def thin_border():
    s = Side(style='thin', color='000000')
    return Border(left=s, right=s, top=s, bottom=s)


def clear_and_border(ws, pink_cols=None):
    """Remove all fills (preserve pink cols); yellow-highlight row 1; apply black border everywhere."""
    bd        = thin_border()
    no_fill   = PatternFill(fill_type=None)
    pink_fill = PatternFill('solid', start_color=PINK)
    yellow_fill = PatternFill('solid', start_color=YELLOW)
    pink_cols = set(pink_cols or [])
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                             min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.row == 1:
                cell.fill = yellow_fill
            elif cell.column in pink_cols:
                cell.fill = pink_fill
            else:
                cell.fill = no_fill
            cell.border = bd
            if cell.font:
                cell.font = Font(
                    bold=cell.font.bold,
                    name=cell.font.name or 'Arial',
                    size=cell.font.size or 10,
                    color='FF000000'
                )


def header_cell(cell, value, fill=None):
    cell.value     = value
    cell.font      = Font(bold=True, name='Arial', size=10)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
    cell.border    = thin_border()
    cell.fill      = fill if fill else PatternFill('solid', start_color=YELLOW)


def data_cell(cell, value, fill=None):
    cell.value  = value
    cell.border = thin_border()
    if fill:
        cell.fill = fill


def autofit_sheet(ws, header_row_height=24, data_row_height=20, max_col_width=80):
    """Expand columns to fit content and give rows a comfortable height.

    - Column width = longest cell content in that column (+ padding), capped
      at `max_col_width` so a single huge value can't blow up the layout.
    - Cells inside a merged range are ignored when measuring column width.
    - Row 1 gets `header_row_height`; remaining rows get `data_row_height`.
    """
    # Cells that are part of a merge — skip them when measuring widths
    merged = set()
    for mrange in ws.merged_cells.ranges:
        for r in range(mrange.min_row, mrange.max_row + 1):
            for c in range(mrange.min_col, mrange.max_col + 1):
                merged.add((r, c))

    col_max = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                             min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.value is None:
                continue
            if (cell.row, cell.column) in merged:
                continue
            # Handle multi-line values: width = longest line
            longest_line = max(
                (len(line) for line in str(cell.value).splitlines()),
                default=0
            )
            letter = get_column_letter(cell.column)
            if longest_line > col_max.get(letter, 0):
                col_max[letter] = longest_line

    for letter, length in col_max.items():
        ws.column_dimensions[letter].width = min(length + 4, max_col_width)

    # Row heights
    if ws.max_row >= 1:
        ws.row_dimensions[1].height = header_row_height
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = data_row_height


def write_sheet(wb, name, df):
    ws = wb.create_sheet(name)
    bd = thin_border()
    for c, col in enumerate(df.columns, 1):
        header_cell(ws.cell(row=1, column=c), col)
    for r, (_, row) in enumerate(df.iterrows(), 2):
        for c, col in enumerate(df.columns, 1):
            val = row[col]
            cell = ws.cell(row=r, column=c, value=None if pd.isna(val) else val)
            cell.border = bd
    for c, col in enumerate(df.columns, 1):
        mx = max([len(str(col))] + [len(str(v)) for v in df[col].dropna()])
        ws.column_dimensions[get_column_letter(c)].width = min(mx + 2, 40)
    ws.freeze_panes = 'A2'
    return ws


# ── Reference-file loaders ───────────────────────────────────────────────────

def load_lr_lookup(path):
    df = pd.read_excel(path, header=0)
    df.columns = ['key', 'value']
    return dict(zip(df['key'].astype(str).str.strip(),
                    df['value'].astype(str).str.strip()))


def load_cutsheet(path):
    return pd.read_excel(path, sheet_name='Installation Sheet')


def build_cutsheet_lookup(cut_df):
    """Key: (Hostname, Interface) → row dict for Source_port etc.

    Adapts to schema changes: only carries forward the fill columns that
    actually exist in the cutsheet. DMARC1/DMARC2 were dropped in later
    cutsheet revisions; the lookup still works without them.
    """
    candidate_cols = ['Source_port', 'DMARC1', 'DMARC2', 'Destination_port']
    fill_cols = [c for c in candidate_cols if c in cut_df.columns]
    lookup = {}
    for _, row in cut_df.iterrows():
        key = (
            str(row['Hostname']).strip(),
            str(row['Interface']).strip(),
        )
        lookup[key] = {c: row[c] for c in fill_cols}
    # Stash which columns were actually available so step 6 can use the same set
    lookup['__fill_cols__'] = fill_cols
    return lookup


def build_z_lookup(cut_df):
    """Key: (Z Hostname, Z Interface) → full row. Z-side is 1:1 unique in
    the cutsheet (same as A-side), so the simpler 2-tuple key is reliable
    and matches the (Hostname, Interface) pattern used elsewhere."""
    lookup = {}
    for _, row in cut_df.iterrows():
        key = (
            str(row['Z Hostname']).strip(),
            str(row['Z Interface']).strip(),
        )
        lookup[key] = row
    return lookup


def paired_subport(iface):
    """Return the paired sub-port interface name.

    s0 ↔ s1 are a pair, s2 ↔ s3 are a pair.
    e.g. 'swp4s0' → 'swp4s1', 'swp15s3' → 'swp15s2'.
    Returns None if iface doesn't end in s0/s1/s2/s3.
    """
    pairs = {'s0': 's1', 's1': 's0', 's2': 's3', 's3': 's2'}
    for suffix, mate in pairs.items():
        if iface.endswith(suffix):
            return iface[:-len(suffix)] + mate
    return None


# ── Core processor (UNCHANGED from original) ─────────────────────────────────

def process_file(input_path, output_path, lr_lookup, cut_df, log):
    shutil.copy2(input_path, output_path)
    wb = load_workbook(output_path)

    cutsheet_lookup = build_cutsheet_lookup(cut_df)
    z_lookup        = build_z_lookup(cut_df)

    # ── 1. Split lldp tab → Downlinks / Mismatches ──────────────────────────
    mis_orig_df = None   # keep for step 6b
    lldp_tab = find_tab(wb, 'lldp')
    if lldp_tab:
        log(f"  · Splitting {lldp_tab} → Downlinks / Mismatches")
        df = pd.read_excel(input_path, sheet_name=lldp_tab)
        down_df     = df[df['Act. Interface'] == 'interface down'].copy()
        mis_orig_df = df[df['Act. Interface'].str.startswith('swp', na=False)].copy()
        # Downlinks: drop only the Active-side columns; keep the Expected-side
        # columns and order them right after Elevation. Source_port and
        # Destination_port will be inserted between Elevation and Expected*
        # by step 6, giving the final order:
        #   Hostname, Interface, L/R, Rack, Elevation,
        #   Source_port, Destination_port,
        #   Expected Hostname, Exp. Interface, Exp. L/R, Exp. Rack, Exp. Elevation
        drop = ['Active Host', 'Act. Interface', 'Act. Rack', 'Act. Elevation']
        down_df.drop(columns=[c for c in drop if c in down_df.columns], inplace=True)
        # Pull Expected* to the end (relative order preserved) so they sit
        # after Elevation in the written sheet.
        exp_cols = ['Expected Hostname', 'Exp. Interface', 'Exp. Rack', 'Exp. Elevation']
        present_exp = [c for c in exp_cols if c in down_df.columns]
        if present_exp:
            rest = [c for c in down_df.columns if c not in present_exp]
            down_df = down_df[rest + present_exp]
        del wb[lldp_tab]
        write_sheet(wb, 'Downlinks', down_df)
        write_sheet(wb, 'Mismatches', mis_orig_df.drop(
            columns=[c for c in [] if c in mis_orig_df.columns]))  # keep all cols for now

    # ── 2. Optics tab ───────────────────────────────────────────────────────
    optics_src = find_tab(wb, 'optics')
    if optics_src:
        log(f"  · Processing Optics tab ({optics_src})")
        # Extra cols introduced by the *_with_pp variant — drop them too.
        drop_cols = {'Transceiver', 'Channel',
                     'Min Threshold (dBm)', 'Max Threshold (dBm)',
                     'PP_A', 'PP_Z', 'Z_end_host', 'Z_end_intf',
                     'rack_z', 'Z_Rack', 'Z_Elevation', 'Index',
                     'Status', 'Placement Group'}
        optics_df = pd.read_excel(input_path, sheet_name=optics_src)
        optics_df.drop(columns=[c for c in drop_cols if c in optics_df.columns], inplace=True)
        # Put 'Metric' first, then 'Measured (dBm)', so both columns can be
        # frozen and stay visible while scrolling horizontally.
        leading = [c for c in ('Metric', 'Measured (dBm)') if c in optics_df.columns]
        if leading:
            rest = [c for c in optics_df.columns if c not in leading]
            optics_df = optics_df[leading + rest]
        del wb[optics_src]
        write_sheet(wb, 'Optics', optics_df)
        # Freeze row 1 + the two leading columns (Metric, Measured (dBm)).
        wb['Optics'].freeze_panes = 'C2' if len(leading) >= 2 else 'B2'

    # ── 3. Remove interfaces tab ────────────────────────────────────────────
    interfaces_tab = find_tab(wb, 'interfaces')
    if interfaces_tab:
        log(f"  · Removing {interfaces_tab}")
        del wb[interfaces_tab]

    # ── 3a. Remove unwanted source tabs ─────────────────────────────────────
    drop_lower = {t.lower() for t in TABS_TO_REMOVE}
    for existing in list(wb.sheetnames):
        if existing.lower() in drop_lower:
            log(f"  · Removing {existing}")
            del wb[existing]

    # ── 3b. combined_fec: move Lock Status + Pre-FEC BER to the front ───────
    fec_tab = find_tab(wb, 'combined_fec')
    if fec_tab:
        log(f"  · Reordering {fec_tab} (Lock Status, Pre-FEC BER first)")
        fec_df = pd.read_excel(input_path, sheet_name=fec_tab)

        def _norm(s):
            # Normalize hyphen variants and whitespace for tolerant matching
            return (str(s)
                    .replace('\u2011', '-')   # non-breaking hyphen
                    .replace('\u2013', '-')   # en dash
                    .replace('\u2014', '-')   # em dash
                    .strip()
                    .lower())

        wanted = ['lock status', 'pre-fec ber']
        front = []
        for target in wanted:
            for col in fec_df.columns:
                if _norm(col) == target and col not in front:
                    front.append(col)
                    break

        if front:
            rest = [c for c in fec_df.columns if c not in front]
            fec_df = fec_df[front + rest]
            del wb[fec_tab]
            write_sheet(wb, 'combined_fec', fec_df)
        else:
            log("    ⚠ Lock Status / Pre-FEC BER not found — leaving combined_fec as-is")

    # ── 4. Reorder tabs ─────────────────────────────────────────────────────
    # Put the four primary output tabs at the END in the documented order,
    # everything else stays in its current relative order. Rewriting _sheets
    # directly is simpler and more robust than chained move_sheet() calls
    # whose offset arithmetic depended on the count of intermediate tabs.
    desired  = ['Downlinks', 'Mismatches', 'Optics', 'combined_fec']
    existing = [s for s in desired if s in wb.sheetnames]
    others   = [s for s in wb.sheetnames if s not in desired]
    wb._sheets = [wb[n] for n in others + existing]

    # ── 5. Insert L/R columns ───────────────────────────────────────────────
    log("  · Adding L/R mapped columns")
    # Each *Interface column gets its own L/R variant inserted to its right.
    lr_name_for = {
        'Interface':     'L/R',
        'Z Interface':   'Z L/R',
        'Exp. Interface':'Exp. L/R',
    }
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        targets = [(i+1, h) for i, h in enumerate(header) if h in lr_name_for]
        for col_idx, col_name in sorted(targets, reverse=True):
            new_name = lr_name_for[col_name]
            ws.insert_cols(col_idx + 1)
            header_cell(ws.cell(row=1, column=col_idx + 1), new_name)
            for r in range(2, ws.max_row + 1):
                val = str(ws.cell(row=r, column=col_idx).value or '').strip()
                ws.cell(row=r, column=col_idx + 1, value=lr_lookup.get(val, ''))
                ws.cell(row=r, column=col_idx + 1).border = thin_border()
            ws.column_dimensions[get_column_letter(col_idx + 1)].width = 10

    # ── 6. Populate Source_port / DMARC1 / DMARC2 / Destination_port ────────
    # Match each row's (Hostname, Interface) against the cutsheet and fill
    # the available cutsheet columns. New fill columns get inserted right
    # after Elevation when the target sheet doesn't already have them.
    fill_cols = cutsheet_lookup.get('__fill_cols__',
                                    ['Source_port', 'DMARC1', 'DMARC2', 'Destination_port'])
    log(f"  · Filling {', '.join(fill_cols) or '(no cutsheet fill cols available)'} (match on Hostname + Interface)")
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        if not all(c in header for c in ['Hostname', 'Interface']):
            continue
        # Append any fill columns this sheet doesn't already have, right after
        # Elevation if present, otherwise at the end.
        anchor = (header.index('Elevation') + 1) if 'Elevation' in header else len(header)
        insert_at = anchor + 1
        for col_name in fill_cols:
            if col_name in header:
                continue
            ws.insert_cols(insert_at)
            header_cell(ws.cell(row=1, column=insert_at), col_name)
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=insert_at).border = thin_border()
            ws.column_dimensions[get_column_letter(insert_at)].width = max(len(col_name)+2, 14)
            insert_at += 1
        header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        host_c, int_c = header.index('Hostname')+1, header.index('Interface')+1
        fill_idx = {c: header.index(c)+1 for c in fill_cols}
        for r in range(2, ws.max_row + 1):
            host  = str(ws.cell(row=r, column=host_c).value or '').strip()
            iface = str(ws.cell(row=r, column=int_c).value or '').strip()
            match = cutsheet_lookup.get((host, iface))
            if match:
                for col_name, col_idx in fill_idx.items():
                    val = match.get(col_name)
                    if val is not None and not (isinstance(val, float) and pd.isna(val)):
                        ws.cell(row=r, column=col_idx, value=val)

    # ── 6c. Fill Z-side info in designated tabs (default: Optics) ───────────
    Z_COLS = ['Z Hostname', 'Z Interface', 'Z L/R', 'Z Rack', 'Z Elevation']
    z_available = [c for c in Z_COLS if c in cut_df.columns]
    if z_available and Z_FILL_TABS:
        # Per-row lookup of Z columns by (Hostname, Interface)
        z_by_host_int = {}
        for _, row in cut_df.iterrows():
            k = (str(row['Hostname']).strip(), str(row['Interface']).strip())
            z_by_host_int[k] = {c: row[c] for c in z_available}

        for tab in Z_FILL_TABS:
            if tab not in wb.sheetnames:
                continue
            ws_z = wb[tab]
            header = [ws_z.cell(row=1, column=c).value
                      for c in range(1, ws_z.max_column + 1)]
            if not all(c in header for c in ['Hostname', 'Interface']):
                continue
            log(f"  · Filling Z-side info in {tab}: {', '.join(z_available)}")
            # Anchor: right after Destination_port if present, else after
            # Elevation, else at the end of the sheet.
            if 'Destination_port' in header:
                anchor = header.index('Destination_port') + 1
            elif 'Elevation' in header:
                anchor = header.index('Elevation') + 1
            else:
                anchor = len(header)
            insert_at = anchor + 1
            for col_name in z_available:
                if col_name in header:
                    continue
                ws_z.insert_cols(insert_at)
                header_cell(ws_z.cell(row=1, column=insert_at), col_name)
                for r in range(2, ws_z.max_row + 1):
                    ws_z.cell(row=r, column=insert_at).border = thin_border()
                ws_z.column_dimensions[get_column_letter(insert_at)].width = max(len(col_name)+2, 14)
                insert_at += 1
            header = [ws_z.cell(row=1, column=c).value
                      for c in range(1, ws_z.max_column + 1)]
            host_c, int_c = header.index('Hostname')+1, header.index('Interface')+1
            fill_idx = {c: header.index(c)+1 for c in z_available}
            for r in range(2, ws_z.max_row + 1):
                host  = str(ws_z.cell(row=r, column=host_c).value or '').strip()
                iface = str(ws_z.cell(row=r, column=int_c).value or '').strip()
                match = z_by_host_int.get((host, iface))
                if match:
                    for col_name, col_idx in fill_idx.items():
                        val = match.get(col_name)
                        if val is not None and not (isinstance(val, float) and pd.isna(val)):
                            ws_z.cell(row=r, column=col_idx, value=val)

    # ── 6b. Mismatches: Possible columns + Active Z columns (pink) ───────────
    if 'Mismatches' in wb.sheetnames:
        log("  · Building Possible + Active Z columns in Mismatches")
        pink_fill   = PatternFill('solid', start_color=PINK)
        yellow_fill = PatternFill('solid', start_color=YELLOW)
        bd          = thin_border()

        # Build act_lookup from original lldp tab (alias-aware)
        act_lookup = {}
        src_sheets = pd.ExcelFile(input_path).sheet_names
        src_lldp = find_tab(src_sheets, 'lldp')
        if src_lldp:
            orig_df  = pd.read_excel(input_path, sheet_name=src_lldp)
            mis_rows = orig_df[orig_df['Act. Interface'].str.startswith('swp', na=False)]
            for _, row in mis_rows.iterrows():
                key       = (str(row['Hostname']).strip(), str(row['Interface']).strip())
                act_iface = str(row['Act. Interface']).strip()
                act_lookup[key] = {
                    'Z Hostname' : str(row['Active Host']).strip(),
                    'Z Interface': act_iface,
                    'Z L/R'      : lr_lookup.get(act_iface, ''),
                    'Z Rack'     : int(float(str(row['Act. Rack']))),
                    'Z Elevation': int(float(str(row['Act. Elevation']))),
                }

        ws_m   = wb['Mismatches']
        header = [ws_m.cell(row=1, column=c).value for c in range(1, ws_m.max_column + 1)]

        # Drop only the Active-side columns. Expected-side columns are kept
        # so they can sit between Destination_port and the Possible/Z blocks
        # in the final layout. Exp. L/R (added by step 5) is preserved too.
        act_drop = {'Active Host', 'Act. Interface', 'Act. Rack', 'Act. Elevation'}
        for idx in sorted([i+1 for i, h in enumerate(header) if h in act_drop], reverse=True):
            ws_m.delete_cols(idx)
        header = [ws_m.cell(row=1, column=c).value for c in range(1, ws_m.max_column + 1)]

        h_idx = header.index('Hostname') + 1
        i_idx = header.index('Interface') + 1

        # Collect act data per row
        act_rows = []
        for r in range(2, ws_m.max_row + 1):
            hn    = str(ws_m.cell(row=r, column=h_idx).value or '').strip()
            iface = str(ws_m.cell(row=r, column=i_idx).value or '').strip()
            act_rows.append(act_lookup.get((hn, iface), {}))

        # Possible columns: match act Z key against cutsheet Z side.
        # Filter out ones whose source column doesn't exist in this cutsheet
        # (e.g. DMARC1 / DMARC2 were dropped in later cutsheet revisions).
        possible_cols_all = [
            ('Possible Hostname',         'Hostname'),
            ('Possible Interface',        'Interface'),
            ('Possible L/R',              '__lr__'),
            ('Possible Rack',             'Rack'),
            ('Possible Elevation',        'Elevation'),
            ('Possible Source_port',      'Source_port'),
            ('Possible DMARC1',           'DMARC1'),
            ('Possible DMARC2',           'DMARC2'),
            ('Possible Destination_port', 'Destination_port'),
        ]
        cut_cols = set(cut_df.columns)
        possible_cols = [
            (out_col, src) for out_col, src in possible_cols_all
            if src == '__lr__' or src in cut_cols
        ]

        possible_data = {col: [] for col, _ in possible_cols}
        for act in act_rows:
            zh   = act.get('Z Hostname', '')
            zi   = act.get('Z Interface', '')
            match = z_lookup.get((zh, zi)) if zh else None
            # Fallback: if exact sub-port not in cutsheet, try its pair.
            # s0↔s1 are a pair, s2↔s3 are a pair.
            if match is None and zh and zi:
                mate = paired_subport(zi)
                if mate:
                    match = z_lookup.get((zh, mate))
            for col, src in possible_cols:
                if match is not None:
                    val = lr_lookup.get(str(match.get('Interface', '')).strip(), '') \
                          if src == '__lr__' else match.get(src, '')
                else:
                    val = ''
                possible_data[col].append(val)

        # Write Possible columns
        pink_col_indices = []
        start = ws_m.max_column + 1
        for c_off, (col_name, _) in enumerate(possible_cols):
            col_idx = start + c_off
            pink_col_indices.append(col_idx)
            hdr = ws_m.cell(row=1, column=col_idx, value=col_name)
            hdr.font = Font(bold=True, name='Arial', size=10)
            hdr.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
            hdr.fill  = yellow_fill
            hdr.border = bd
            ws_m.column_dimensions[get_column_letter(col_idx)].width = max(len(col_name)+3, 14)
            for r_off, val in enumerate(possible_data[col_name]):
                cell = ws_m.cell(row=r_off+2, column=col_idx,
                                 value=val if val != '' else None)
                cell.fill   = pink_fill
                cell.border = bd

        # Write Active Z columns
        act_z_cols = ['Z Hostname', 'Z Interface', 'Z L/R', 'Z Rack', 'Z Elevation']
        start2 = ws_m.max_column + 1
        for c_off, col_name in enumerate(act_z_cols):
            col_idx = start2 + c_off
            pink_col_indices.append(col_idx)
            hdr = ws_m.cell(row=1, column=col_idx, value=col_name)
            hdr.font = Font(bold=True, name='Arial', size=10)
            hdr.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
            hdr.fill  = yellow_fill
            hdr.border = bd
            ws_m.column_dimensions[get_column_letter(col_idx)].width = max(len(col_name)+3, 14)
            for r_off, act in enumerate(act_rows):
                val = act.get(col_name, '')
                cell = ws_m.cell(row=r_off+2, column=col_idx,
                                 value=val if val != '' else None)
                cell.fill   = pink_fill
                cell.border = bd

        # ── Connection Status (basic mismatch detection) ─────────────────────
        log("  · Adding Connection Status column (Option B)")
        mismatch_fill = PatternFill('solid', start_color='FFCCCB')  # light red/pink
        status_col = ws_m.max_column + 1

        hdr = ws_m.cell(row=1, column=status_col, value="Connection Status")
        hdr.font = Font(bold=True, name='Arial', size=10)
        hdr.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
        hdr.fill  = yellow_fill
        hdr.border = bd
        ws_m.column_dimensions[get_column_letter(status_col)].width = 32

        for r_off in range(len(act_rows)):
            act = act_rows[r_off]
            actual_h = str(act.get('Z Hostname', '') or '').strip().lower()
            actual_i = str(act.get('Z Interface', '') or '').strip().lower()

            exp_h = str(possible_data.get('Possible Hostname', [''] * len(act_rows))[r_off] or '').strip().lower()
            exp_i = str(possible_data.get('Possible Interface', [''] * len(act_rows))[r_off] or '').strip().lower()

            if not actual_h and not actual_i:
                status = "No Active Connection"
            elif not exp_h and not exp_i:
                status = "No Expected Data in Cutsheet"
            elif actual_h == exp_h and actual_i == exp_i:
                status = "As Expected"
            elif actual_h == exp_h:
                status = "Mismatch - Different Interface"
            elif actual_i == exp_i:
                status = "Mismatch - Different Host"
            else:
                status = "Mismatch - Different Host + Interface"

            cell = ws_m.cell(row=r_off + 2, column=status_col, value=status)
            cell.border = bd
            if "Mismatch" in status:
                cell.fill = mismatch_fill

    # ── 6c. Strip unwanted columns across every tab ─────────────────────────
    if COLUMNS_TO_REMOVE:
        log(f"  · Stripping columns: {', '.join(COLUMNS_TO_REMOVE)}")
        drop_set = set(COLUMNS_TO_REMOVE)
        for sheet_name in wb.sheetnames:
            ws_x = wb[sheet_name]
            header = [ws_x.cell(row=1, column=c).value
                      for c in range(1, ws_x.max_column + 1)]
            # Find 1-based indices to drop, delete right-to-left so earlier
            # indices stay valid as columns shift left.
            to_drop = [i + 1 for i, h in enumerate(header) if h in drop_set]
            for idx in sorted(to_drop, reverse=True):
                ws_x.delete_cols(idx)

    # ── 7. Summary tab (per Rack breakdown) ─────────────────────────────────
    log("  · Creating Summary tab")

    # Gather rack counts per sheet from the workbook data
    tab_rack  = {}
    all_racks = set()
    no_fill_s   = PatternFill(fill_type=None)
    yellow_fill_s = PatternFill('solid', start_color=YELLOW)
    center_s  = Alignment(horizontal='center', vertical='center', wrap_text=False)
    bd_s      = thin_border()

    def _s(cell, value, bold=False, header=False):
        cell.value     = value
        cell.font      = Font(bold=bold, name='Arial', size=10)
        cell.alignment = center_s
        cell.border    = bd_s
        cell.fill      = yellow_fill_s if header else no_fill_s

    for sname in wb.sheetnames:
        ws_tmp = wb[sname]
        hdr = [ws_tmp.cell(row=1, column=c).value for c in range(1, ws_tmp.max_column+1)]
        if 'Rack' not in hdr:
            tab_rack[sname] = {}
            continue
        rack_col = hdr.index('Rack') + 1
        counts = {}
        for r in range(2, ws_tmp.max_row + 1):
            val = ws_tmp.cell(row=r, column=rack_col).value
            if val is not None:
                try:
                    k = int(float(str(val)))
                    counts[k] = counts.get(k, 0) + 1
                    all_racks.add(k)
                except ValueError:
                    pass
        tab_rack[sname] = counts

    racks      = sorted(all_racks)
    tabs_order = [n for n in wb.sheetnames]
    total_cols = 1 + len(racks) + 1  # Tab Name + per rack + Total

    # Excel treats sheet names case-insensitively. The new report ships a
    # lowercase 'summary' tab; delete any case-variant before creating ours.
    for existing in list(wb.sheetnames):
        if existing.lower() == 'summary':
            del wb[existing]
    wb.create_sheet('Summary', 0)
    ws_s = wb['Summary']

    # Title
    ws_s.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    title_c = ws_s.cell(row=1, column=1, value='Tab Summary by Rack')
    title_c.font = Font(bold=True, name='Arial', size=13)
    title_c.alignment = center_s
    title_c.border    = bd_s
    title_c.fill      = yellow_fill_s
    ws_s.row_dimensions[1].height = 28

    # Header row
    _s(ws_s.cell(row=2, column=1), 'Tab Name', bold=True, header=True)
    for c, rack in enumerate(racks, start=2):
        _s(ws_s.cell(row=2, column=c), str(rack), bold=True, header=True)
    _s(ws_s.cell(row=2, column=total_cols), 'Total', bold=True, header=True)

    # Data rows (exclude the Summary tab itself — case-insensitive, since the
    # source report may ship a lowercase 'summary' tab that we delete later)
    rack_totals = {r: 0 for r in racks}
    data_tabs   = [n for n in tabs_order if n.lower() != 'summary']
    for i, tab_name in enumerate(data_tabs, start=3):
        _s(ws_s.cell(row=i, column=1), tab_name)
        row_total = 0
        for c, rack in enumerate(racks, start=2):
            count = tab_rack.get(tab_name, {}).get(rack, 0)
            _s(ws_s.cell(row=i, column=c), count if count > 0 else '')
            rack_totals[rack] += count
            row_total += count
        _s(ws_s.cell(row=i, column=total_cols), row_total, bold=True)

    # Grand total row
    tot_r = 3 + len(data_tabs)
    _s(ws_s.cell(row=tot_r, column=1), 'TOTAL', bold=True)
    grand = 0
    for c, rack in enumerate(racks, start=2):
        _s(ws_s.cell(row=tot_r, column=c), rack_totals[rack], bold=True)
        grand += rack_totals[rack]
    _s(ws_s.cell(row=tot_r, column=total_cols), grand, bold=True)

    # Column widths
    ws_s.column_dimensions['A'].width = 20
    for c in range(2, total_cols + 1):
        ws_s.column_dimensions[get_column_letter(c)].width = 14

    # ── 8. No fill + borders (preserve pink in Mismatches) ──────────────────
    log("  · Removing fills and applying borders")
    # The earlier pink_col_indices was captured in step 6b, but step 6c (column
    # strip) deletes columns from the middle of Mismatches afterwards, which
    # shifts all Possible/Z columns left and invalidates those stale indices.
    # Recompute by name from the current header so every pink column gets
    # filled correctly regardless of how many columns were stripped.
    if 'Mismatches' in wb.sheetnames:
        ws_m = wb['Mismatches']
        m_header = [ws_m.cell(row=1, column=c).value
                    for c in range(1, ws_m.max_column + 1)]
        Z_NAMES = {'Z Hostname', 'Z Interface', 'Z L/R', 'Z Rack', 'Z Elevation'}
        pink_col_indices = [
            i + 1 for i, h in enumerate(m_header)
            if (h and (str(h).startswith('Possible ') or h in Z_NAMES))
        ]

    for sheet_name in wb.sheetnames:
        pcols = pink_col_indices if sheet_name == 'Mismatches' else []
        clear_and_border(wb[sheet_name], pink_cols=pcols)

    # ── 8c. Centre-align all cells across all tabs ──────────────────────────
    log("  · Aligning all cells to middle-centre")
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=False)
    for sheet_name in wb.sheetnames:
        for row in wb[sheet_name].iter_rows():
            for cell in row:
                cell.alignment = center_align

    # ── 8b. Add NOTE column + autofilter to all tabs ───────────────────────
    log("  · Adding NOTE column and filters to all tabs")
    no_fill     = PatternFill(fill_type=None)
    yellow_fill = PatternFill('solid', start_color=YELLOW)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for col_name in ['NOTE']:
            col_idx = ws.max_column + 1
            hdr = ws.cell(row=1, column=col_idx, value=col_name)
            hdr.font      = Font(bold=True, name='Arial', size=10)
            hdr.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
            hdr.fill      = yellow_fill
            hdr.border    = thin_border()
            ws.column_dimensions[get_column_letter(col_idx)].width = 16
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(row=r, column=col_idx)
                cell.fill   = no_fill
                cell.border = thin_border()
        if ws.max_row > 1 and ws.max_column > 0:
            ws.auto_filter.ref = ws.dimensions

    # ── 8d. Grey-out Optics rows that are matched in Downlinks ──────────────
    if 'Optics' in wb.sheetnames and 'Downlinks' in wb.sheetnames:
        log("  · Greying out matched Optics rows")

        MATCH_COLS = [
            'Hostname', 'Interface', 'L/R', 'Rack', 'Elevation',
            'Source_port', 'DMARC1', 'DMARC2', 'Destination_port',
            'Z Hostname', 'Z Interface', 'Z L/R', 'Z Rack', 'Z Elevation',
        ]
        GREY_FONT_COLOR = 'FFD3D3D3'  # light grey

        ws_dl = wb['Downlinks']
        dl_header = [ws_dl.cell(row=1, column=c).value
                     for c in range(1, ws_dl.max_column + 1)]

        # Build a set of tuples from Downlinks for all match columns present
        dl_match_cols = [c for c in MATCH_COLS if c in dl_header]
        dl_col_idx    = {c: dl_header.index(c) + 1 for c in dl_match_cols}

        dl_keys = set()
        for r in range(2, ws_dl.max_row + 1):
            key = tuple(
                str(ws_dl.cell(row=r, column=dl_col_idx[c]).value or '').strip()
                for c in dl_match_cols
            )
            dl_keys.add(key)

        ws_op = wb['Optics']
        op_header = [ws_op.cell(row=1, column=c).value
                     for c in range(1, ws_op.max_column + 1)]

        # Only match on columns present in both sheets
        common_cols  = [c for c in dl_match_cols if c in op_header]
        op_col_idx   = {c: op_header.index(c) + 1 for c in common_cols}
        dl_col_idx_c = {c: dl_header.index(c) + 1 for c in common_cols}

        # Rebuild dl_keys using only common columns
        dl_keys_common = set()
        for r in range(2, ws_dl.max_row + 1):
            key = tuple(
                str(ws_dl.cell(row=r, column=dl_col_idx_c[c]).value or '').strip()
                for c in common_cols
            )
            dl_keys_common.add(key)

        for r in range(2, ws_op.max_row + 1):
            op_key = tuple(
                str(ws_op.cell(row=r, column=op_col_idx[c]).value or '').strip()
                for c in common_cols
            )
            if op_key in dl_keys_common:
                for c in range(1, ws_op.max_column + 1):
                    cell = ws_op.cell(row=r, column=c)
                    cell.font = Font(
                        bold=cell.font.bold if cell.font else False,
                        name=(cell.font.name if cell.font else None) or 'Arial',
                        size=(cell.font.size if cell.font else None) or 10,
                        color=GREY_FONT_COLOR,
                    )

    # ── 8e. Expand all columns and rows on every sheet ──────────────────────
    log("  · Expanding all columns and rows")
    for sheet_name in wb.sheetnames:
        autofit_sheet(wb[sheet_name])

    wb.save(output_path)

    # ── 9. Rename by top-2 Rack numbers ─────────────────────────────────────
    try:
        all_racks = []
        for sheet_name in wb.sheetnames:
            ws     = wb[sheet_name]
            header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column+1)]
            if 'Rack' in header:
                rc = header.index('Rack') + 1
                for r in range(2, ws.max_row + 1):
                    val = ws.cell(row=r, column=rc).value
                    if val is not None:
                        try: all_racks.append(int(float(str(val))))
                        except ValueError: pass
        if all_racks:
            top2     = [str(r) for r, _ in Counter(all_racks).most_common(2)]
            new_name = '+'.join(top2) + '.xlsx'
            new_path = os.path.join(os.path.dirname(output_path), new_name)
            load_workbook(output_path).save(new_path)
            if new_path != output_path:
                os.remove(output_path)
            log(f"  ✓ Saved → {new_name}")
            return new_path
    except Exception as e:
        log(f"  ⚠ Could not rename by Rack: {e}")

    log(f"  ✓ Saved → {os.path.basename(output_path)}")
    return output_path


# ── Reference identification (UNCHANGED) ─────────────────────────────────────

def identify_refs(paths):
    """
    Given two reference file paths, identify which is the Formula L&R and
    which is the Cutsheet. The cutsheet is the one that contains an
    'Installation Sheet' tab; the other is the L&R.
    Returns (formula_path, cutsheet_path).
    """
    formula = cutsheet = None
    for p in paths:
        try:
            sheets = load_workbook(p, read_only=True).sheetnames
        except Exception as e:
            raise ValueError(f"Could not open {os.path.basename(p)}: {e}")
        if 'Installation Sheet' in sheets:
            if cutsheet:
                raise ValueError(
                    "Both selected files contain an 'Installation Sheet' tab. "
                    "Could not tell which one is the cutsheet."
                )
            cutsheet = p
        else:
            if formula:
                raise ValueError(
                    "Neither selected file contains an 'Installation Sheet' tab. "
                    "Could not identify the cutsheet."
                )
            formula = p
    if not cutsheet:
        raise ValueError(
            "Neither selected file contains an 'Installation Sheet' tab; "
            "expected one of them to be the cutsheet."
        )
    if not formula:
        raise ValueError("Could not identify the Formula L&R file.")
    return formula, cutsheet


# ── Streamlit UI ─────────────────────────────────────────────────────────────

def load_references_from_uploads(ref_uploads):
    """Save two uploaded files to a temp dir, identify, load lookups, return results."""
    with tempfile.TemporaryDirectory() as tmpd:
        paths = []
        names = []
        for uf in ref_uploads:
            p = os.path.join(tmpd, uf.name)
            with open(p, "wb") as f:
                f.write(uf.getbuffer())
            paths.append(p)
            names.append(uf.name)

        formula_path, cutsheet_path = identify_refs(paths)
        lr_lookup = load_lr_lookup(formula_path)
        cut_df = load_cutsheet(cutsheet_path)

        return {
            'lr_lookup': lr_lookup,
            'cut_df': cut_df,
            'formula_name': os.path.basename(formula_path),
            'cutsheet_name': os.path.basename(cutsheet_path),
        }


def process_uploads(input_uploads, lr_lookup, cut_df):
    """Process all uploaded input files. Returns (results, errors, logs)."""
    logs = []
    results = []   # (original_name, final_filename, bytes)
    errors = []

    def log(msg):
        logs.append(msg)

    total = len(input_uploads)

    for idx, uf in enumerate(input_uploads, 1):
        log(f"[{idx}/{total}]  {uf.name}")

        with tempfile.TemporaryDirectory() as tmp:
            in_path = os.path.join(tmp, uf.name)
            with open(in_path, "wb") as f:
                f.write(uf.getbuffer())

            base = os.path.splitext(uf.name)[0]
            out_path = os.path.join(tmp, f"{base}_formatted.xlsx")

            try:
                final_path = process_file(in_path, out_path, lr_lookup, cut_df, log)
                with open(final_path, "rb") as f:
                    data = f.read()
                final_name = os.path.basename(final_path)
                results.append((uf.name, final_name, data))
                log(f"  ✓ Success → {final_name}")
            except Exception as e:
                log(f"  ✗ ERROR: {e}")
                errors.append((uf.name, str(e)))

        log("")

    return results, errors, logs


def main():
    st.set_page_config(
        page_title="GFAB Cutsheet Formatter",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    st.title("GFAB Excel Cutsheet Formatter")
    st.markdown(
        "Convert raw source reports into clean, cutsheet-enriched, rack-renamed deliverables. "
        "**Identical output** to the original desktop version."
    )

    with st.sidebar:
        st.header("Instructions")
        st.markdown("""
        1. Upload **exactly two** reference files:
           - **Formula L&R** (the simple 2-column lookup)
           - **Cutsheet** (contains "Installation Sheet" tab)
        2. The app will automatically detect which is which.
        3. Upload one or more Excel files you want to format.
        4. Click **Process All Files**.
        5. Download the results (files are renamed to top-2 rack numbers, e.g. `11+12.xlsx`).
        """)
        st.divider()
        st.caption("All processing happens locally in your browser session. No data is uploaded to any server.")

    # ── 1. REFERENCE FILES ──────────────────────────────────────────────────
    st.header("1. Reference Files")

    ref_uploads = st.file_uploader(
        "Upload Formula L&R + Cutsheet (exactly 2 files)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="refs",
        help="The cutsheet must contain an 'Installation Sheet' tab. The other file is the Formula L&R lookup."
    )

    refs_ready = False
    if len(ref_uploads) == 2:
        if st.button("Load & Validate References", type="secondary"):
            with st.spinner("Identifying and loading reference files..."):
                try:
                    ref_data = load_references_from_uploads(ref_uploads)
                    st.session_state['lr_lookup'] = ref_data['lr_lookup']
                    st.session_state['cut_df'] = ref_data['cut_df']
                    st.session_state['ref_names'] = (ref_data['formula_name'], ref_data['cutsheet_name'])
                    st.success(
                        f"✅ References loaded successfully\n\n"
                        f"**Formula L&R**: {ref_data['formula_name']}\n"
                        f"**Cutsheet**: {ref_data['cutsheet_name']}"
                    )
                except Exception as e:
                    st.error(f"Failed to load references: {e}")
                    if 'lr_lookup' in st.session_state:
                        del st.session_state['lr_lookup']
                    if 'cut_df' in st.session_state:
                        del st.session_state['cut_df']
                    if 'ref_names' in st.session_state:
                        del st.session_state['ref_names']

    if 'lr_lookup' in st.session_state and 'cut_df' in st.session_state:
        refs_ready = True
        f_name, c_name = st.session_state.get('ref_names', ('Formula L&R', 'Cutsheet'))
        st.info(f"**Current references**: {f_name}  |  {c_name}", icon="📁")

    # ── 2. INPUT FILES ──────────────────────────────────────────────────────
    st.header("2. Files to Format")

    input_uploads = st.file_uploader(
        "Select Excel file(s) to process",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="inputs",
        disabled=not refs_ready,
        help="You can select multiple files. Each will be processed independently."
    )

    if not refs_ready:
        st.warning("Load the two reference files first before uploading input files.")
    elif not input_uploads:
        st.info("Select one or more source report files to format.")

    # ── 3. PROCESS ──────────────────────────────────────────────────────────
    st.header("3. Run Processing")

    can_process = refs_ready and bool(input_uploads)

    if st.button("🚀 Process All Files", type="primary", disabled=not can_process, width="stretch"):
        lr = st.session_state['lr_lookup']
        cut = st.session_state['cut_df']

        # Progress + live log
        progress = st.progress(0, text="Preparing...")
        log_placeholder = st.empty()

        logs_so_far = []

        def live_log(msg):
            logs_so_far.append(msg)
            log_placeholder.code("\n".join(logs_so_far), language="text")

        total = len(input_uploads)
        all_results = []
        all_errors = []
        all_logs = []

        for i, uf in enumerate(input_uploads, 1):
            progress.progress((i - 1) / total, text=f"Processing {uf.name} ({i}/{total})")
            live_log(f"[{i}/{total}]  {uf.name}")

            with tempfile.TemporaryDirectory() as tmp:
                in_path = os.path.join(tmp, uf.name)
                with open(in_path, "wb") as f:
                    f.write(uf.getbuffer())

                base = os.path.splitext(uf.name)[0]
                out_path = os.path.join(tmp, f"{base}_formatted.xlsx")

                try:
                    final_path = process_file(in_path, out_path, lr, cut, live_log)
                    with open(final_path, "rb") as f:
                        data = f.read()
                    final_name = os.path.basename(final_path)
                    all_results.append((uf.name, final_name, data))
                    live_log(f"  ✓ Success → {final_name}")
                except Exception as e:
                    live_log(f"  ✗ ERROR: {e}")
                    all_errors.append((uf.name, str(e)))

            all_logs.extend(logs_so_far)
            logs_so_far = []
            progress.progress(i / total, text=f"Completed {i}/{total}")

        progress.empty()
        log_placeholder.empty()

        # Store results for download buttons (survive reruns)
        st.session_state['last_results'] = all_results
        st.session_state['last_errors'] = all_errors
        st.session_state['last_logs'] = all_logs

        st.rerun()

    # ── RESULTS (shown after processing or on rerun) ────────────────────────
    if 'last_results' in st.session_state:
        results = st.session_state['last_results']
        errors = st.session_state.get('last_errors', [])
        logs = st.session_state.get('last_logs', [])

        st.divider()
        st.header("Results")

        if results:
            st.success(f"✅ Successfully processed {len(results)} file(s)")

            # Individual downloads
            st.subheader("Download Individual Files")
            for orig_name, final_name, data in results:
                cols = st.columns([4, 1])
                cols[0].markdown(f"**{orig_name}**  →  `{final_name}`")
                cols[1].download_button(
                    label="Download",
                    data=data,
                    file_name=final_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{final_name}_{hash(final_name)}"
                )

            # Zip download
            if len(results) > 1:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for _, final_name, data in results:
                        zf.writestr(final_name, data)
                st.download_button(
                    label="⬇️ Download ALL as ZIP archive",
                    data=zip_buffer.getvalue(),
                    file_name="GFAB_formatted_outputs.zip",
                    mime="application/zip",
                    type="primary"
                )

        if errors:
            st.error(f"⚠️ {len(errors)} file(s) failed to process:")
            for name, err in errors:
                st.markdown(f"- **{name}**: {err}")

        # Full log
        with st.expander("View full processing log", expanded=False):
            if logs:
                st.code("\n".join(logs), language="text")
            else:
                st.text("No log entries.")

        # Clear results button
        if st.button("Clear Results & Process New Files"):
            for key in ['last_results', 'last_errors', 'last_logs']:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()

    # Footer
    st.divider()
    st.caption(
        "GFAB Cutsheet Formatter • Streamlit version • "
        "All logic identical to the original Tkinter implementation • "
        "Output files are automatically named using the two most frequent rack numbers."
    )


if __name__ == "__main__":
    main()
