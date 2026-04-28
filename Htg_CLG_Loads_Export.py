import os
import math
from pathlib import Path

import iesve
import pythoncom
import win32com.client as win32


# ============================================================
# CONFIG ONLY
# ============================================================
TEMPLATE_PATH = r"C:\Users\juan.losada\OneDrive - Tetra Tech, Inc\Desktop\HLE-CoDE Building Performance\133 - IES input-output file\Drafting\IES Input Sheet WIP.xlsx"
HTG_FILE = "TL - Loads.htg"
CLG_FILE = "TL - Loads.clg"

GROUPING_SCHEME_NAME = "Building"
ROOM_GROUP_NAME = "Main building"

HTG_SHEET_NAME = "IES - heat loss data (.htg)"
CLG_SHEET_NAME = "IES - heat gain data (.clg)"

HTG_MARKER = "IES ZONE HEAT LOSS OUTPUTS"
CLG_MARKER = "IES ZONE HEAT GAIN OUTPUTS"
SOLAR_MARKER = "Peak time table - Solar gain"

# Primary driver (fallback logic still applied if missing)
PEAK_DRIVER = "Cooling + dehum plant load (kW)"
WRITE_CLG_COMBINED_SUMMARY = True
DEBUG_PRINT_Z_VARS = False


# ============================================================
# HEADERS
# ============================================================
HTG_HEADERS = [
    "Room Name",
    "Room Area (m²)",
    "Air temperature (°C)",
    "Dry resultant temperature (°C)",
    "External conduction gain (kW)",
    "Internal conduction gain (kW)",
    "Infiltration gain (kW)",
    "Steady state heating plant load (kW)",
    "Running total heating load (kW)",
]

CLG_HEADERS = [
    "Room Name",
    "Room Area (m²)",
    "Peak date",
    "Peak time",
    "Air temperature (°C)",
    "Dry resultant temperature (°C)",
    "Internal gain (kW)",
    "Solar gain (kW)",
    "Conduction gain (kW)",
    "Infiltration gain (kW)",
    "Cooling + dehum plant load (kW)",
    "Space conditioning sensible (kW)",
]


# ============================================================
# HELPERS
# ============================================================
def log(msg):
    print(msg, flush=True)


def to_excel_value(v):
    if hasattr(v, "item"):
        try:
            v = v.item()
        except Exception:
            pass
    if isinstance(v, (list, tuple)):
        v = v[0] if v else ""
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return ""
    if isinstance(v, (int, float, str, bool)) or v is None:
        return v
    return str(v)


def scalar(x):
    if x is None:
        return None
    if hasattr(x, "item"):
        try:
            return x.item()
        except Exception:
            pass
    if isinstance(x, (list, tuple)):
        return x[0] if x else None
    try:
        return float(x)
    except Exception:
        return x


def get_room_results_safe(rr, room_id, aps_var, vista_var, var_level='z', start_day=-1, end_day=-1):
    try:
        return rr.get_room_results(room_id, aps_var, vista_var, var_level, start_day, end_day)
    except TypeError:
        return rr.get_room_results(room_id, aps_var, vista_var, start_day, end_day)


def month_time_from_hour_index(hour_idx):
    month_names = ["May", "June", "July", "August", "September"]
    month_index = max(0, min(4, hour_idx // 24))
    month = month_names[month_index]
    hour_1_24 = (hour_idx % 24) + 1
    return month, f"{hour_1_24:02d}:00"



def find_marker_cell_fast(ws, marker_text):
    xlValues = -4163
    xlWhole = 1
    xlByRows = 1
    xlNext = 1
    hit = ws.Cells.Find(
        What=marker_text,
        After=ws.Cells(1, 1),
        LookIn=xlValues,
        LookAt=xlWhole,
        SearchOrder=xlByRows,
        SearchDirection=xlNext,
        MatchCase=False
    )
    if hit is None:
        return None
    return hit.Row, hit.Column


def write_2d_block(ws, start_row, start_col, rows_2d):
    if not rows_2d:
        return

    # normalize row widths
    ncols = max(len(r) for r in rows_2d) if rows_2d else 0
    if ncols == 0:
        return

    normalized = []
    for r in rows_2d:
        rr = [to_excel_value(v) for v in r]
        if len(rr) < ncols:
            rr.extend([""] * (ncols - len(rr)))
        elif len(rr) > ncols:
            rr = rr[:ncols]
        normalized.append(tuple(rr))   # tuple row for COM

    data = tuple(normalized)           # tuple of tuples for COM
    nrows = len(data)

    rng = ws.Range(ws.Cells(start_row, start_col),
                   ws.Cells(start_row + nrows - 1, start_col + ncols - 1))
    rng.Value = data


def safe_kw_at(series, idx):
    if series is None:
        return ""
    try:
        v = float(series[idx])
        if math.isnan(v) or math.isinf(v):
            return ""
        return round(v / 1000.0, 3)
    except Exception:
        return ""


def safe_peak(series):
    vals = []
    for i, v in enumerate(series):
        vv = scalar(v)   # <-- key change
        try:
            fv = float(vv)
            if math.isnan(fv) or math.isinf(fv):
                continue
            vals.append((i, fv))
        except Exception:
            continue
    if not vals:
        return None, None
    idx, mx = max(vals, key=lambda t: t[1])
    return idx, mx


def pick_peak_driver_series(series, room_id, room_name, preferred_driver):
    order = [
        preferred_driver,
        "Cooling + dehum plant load (kW)",
        "Space conditioning sensible (kW)",
        "Solar gain (kW)",
        "Internal gain (kW)",
        "Air temperature (°C)",
    ]
    seen = set()
    for k in order:
        if k in seen:
            continue
        seen.add(k)
        s = series.get(k)
        if s is not None:
            if k != preferred_driver:
                log(f"[CLG][INFO] Fallback peak driver for {room_id} ({room_name}) -> {k}")
            return s, k
    return None, None


def print_clg_var_availability_summary(rr):
    wanted = {
        "Air temperature",
        "Dry resultant temperature",
        "Internal gain",
        "Solar gain",
        "Conduction gain",
        "Infiltration gain",
        "Cooling + dehum plant load",
        "Space conditioning sensible",
    }
    try:
        vars_all = rr.get_variables()
        z_vars = [v for v in vars_all if v.get("model_level") == "z"]
        display_names = {v.get("display_name") for v in z_vars}
        present = sorted([x for x in wanted if x in display_names])
        missing = sorted([x for x in wanted if x not in display_names])
        log(f"[CLG][INFO] z-level variables present ({len(present)}): {present}")
        log(f"[CLG][INFO] z-level variables missing ({len(missing)}): {missing}")
    except Exception as e:
        log(f"[CLG][WARN] Could not compute variable availability summary: {e}")


# ============================================================
# IES COLLECTION
# ============================================================
def collect_heating_data(results_reader, htg_file_path, room_ids_to_analyse):
    log(f"[HTG] Opening: {htg_file_path}")
    rr = results_reader.open(str(htg_file_path))
    try:
        rooms = rr.get_room_list()
        hl_data = []
        running_total_kw = 0.0

        for room in rooms:
            name, room_id, room_area, room_volume = room
            if room_id not in room_ids_to_analyse:
                continue

            np_air_temp = get_room_results_safe(rr, room_id, 'Room air temperature', 'Air temperature', 'z')
            np_dry_resultant_temp = get_room_results_safe(rr, room_id, 'Comfort temperature', 'Dry resultant temperature', 'z')
            np_external_conduction_gain = get_room_results_safe(rr, room_id, 'Conduction from ext elements', 'External conduction gain', 'z')
            np_internal_conduction_gain = get_room_results_safe(rr, room_id, 'Conduction from int surfaces', 'Internal conduction gain', 'z')
            np_infiltration_gain = get_room_results_safe(rr, room_id, 'Infiltration gain', 'Infiltration gain', 'z')
            np_steady_state_heating_plant_load = get_room_results_safe(
                rr, room_id, 'Room units steady state htg load', 'Steady state heating plant load', 'z'
            )

            steady_w = scalar(np_steady_state_heating_plant_load)
            steady_kw_raw = float(steady_w) / 1000.0
            running_total_kw += steady_kw_raw

            hl_data.append([
                name,
                round(float(room_area), 2),
                round(float(scalar(np_air_temp)), 2),
                round(float(scalar(np_dry_resultant_temp)), 2),
                round(float(scalar(np_external_conduction_gain)) / 1000.0, 2),
                round(float(scalar(np_internal_conduction_gain)) / 1000.0, 2),
                round(float(scalar(np_infiltration_gain)) / 1000.0, 2),
                round(steady_kw_raw, 2),
                round(running_total_kw, 2),
            ])

        log(f"[HTG] Rows prepared: {len(hl_data)}")
        return hl_data
    finally:
        rr.close()


def collect_cooling_data(results_reader, clg_file_path, room_ids_to_analyse, peak_driver=PEAK_DRIVER):
    log(f"[CLG] Opening: {clg_file_path}")
    rr = results_reader.open(str(clg_file_path))
    try:
        print_clg_var_availability_summary(rr)
        rooms = rr.get_room_list()

        var_map = {
            "Air temperature (°C)": ("Room air temperature", "Air temperature"),
            "Dry resultant temperature (°C)": ("Comfort temperature", "Dry resultant temperature"),
            "Internal gain (kW)": ("Casual gains", "Internal gain"),
            "Solar gain (kW)": ("Window solar gains", "Solar gain"),
            "Conduction gain (kW)": ("Conduction gain", "Conduction gain"),
            "Infiltration gain (kW)": ("Infiltration gain", "Infiltration gain"),
            "Cooling + dehum plant load (kW)": ("Room units cooling + dehum load", "Cooling + dehum plant load"),
            "Space conditioning sensible (kW)": ("System plant etc. gains", "Space conditioning sensible"),
        }


        hg_data = []
        solar_peaks_table = [["Room Name", "Peak date", "Peak time", "Max solar gain (kW)"]]
        all_driver_series = []

        for room in rooms:
            name, room_id, room_area, room_volume = room
            if room_id not in room_ids_to_analyse:
                continue

            series = {}
            failed = False

            for k, (aps_name, vista_name) in var_map.items():
                try:
                    s = get_room_results_safe(rr, room_id, aps_name, vista_name, 'z')
                    if s is None:
                        failed = True
                        log(f"[CLG][WARN] Required field missing: {k} for room {room_id} ({name})")
                        break
                    series[k] = s
                except Exception as e:
                    failed = True
                    log(f"[CLG][WARN] Failed {k} for room {room_id} ({name}) -> {e}")
                    break

            if failed:
                continue

            driver_series, used_driver = pick_peak_driver_series(series, room_id, name, peak_driver)
            if driver_series is None:
                log(f"[CLG][WARN] No usable peak driver for room {room_id} ({name}); skipping")
                continue

            peak_hour, peak_val = safe_peak(driver_series)
            if peak_hour is None:
                log(f"[CLG][WARN] Peak driver has no valid numeric values for room {room_id} ({name}); skipping")
                continue

            peak_month, peak_time = month_time_from_hour_index(peak_hour)

            solar_series = series["Solar gain (kW)"]
            solar_max_hour, solar_max = safe_peak(solar_series)

            if solar_max_hour is None:
                solar_month, solar_time, solar_kw = "", "", ""
                log(f"[CLG][WARN] No valid solar values for {room_id} ({name})")
            else:
                solar_month, solar_time = month_time_from_hour_index(int(solar_max_hour))
                solar_kw = round(float(solar_max) / 1000.0, 3)

            solar_peaks_table.append([str(name), str(solar_month), str(solar_time), solar_kw])

            hg_data.append([
                name,
                round(float(room_area), 2),
                peak_month,
                peak_time,
                round(float(series["Air temperature (°C)"][peak_hour]), 2),
                round(float(series["Dry resultant temperature (°C)"][peak_hour]), 2),
                safe_kw_at(series["Internal gain (kW)"], peak_hour),
                safe_kw_at(series["Solar gain (kW)"], peak_hour),
                safe_kw_at(series["Conduction gain (kW)"], peak_hour),
                safe_kw_at(series["Infiltration gain (kW)"], peak_hour),
                safe_kw_at(series["Cooling + dehum plant load (kW)"], peak_hour),
                safe_kw_at(series["Space conditioning sensible (kW)"], peak_hour),
            ])
            all_driver_series.append(driver_series)

        combined_summary = False
        if all_driver_series:
            n_hours = len(all_driver_series[0])
            combined = []
            for h in range(n_hours):
                s = 0.0
                valid = False
                for arr in all_driver_series:
                    try:
                        v = float(arr[h])
                        if math.isnan(v) or math.isinf(v):
                            continue
                        s += v
                        valid = True
                    except Exception:
                        continue
                combined.append(s if valid else float("nan"))

            valid_combined = [(i, v) for i, v in enumerate(combined) if not (math.isnan(v) or math.isinf(v))]
            if valid_combined:
                c_hour, c_peak = max(valid_combined, key=lambda t: t[1])
                m, t = month_time_from_hour_index(c_hour)
                combined_summary = [round(float(c_peak / 1000), 3), m, t, peak_driver]

        log(f"[CLG] Rows prepared: {len(hg_data)}")
        return hg_data, combined_summary, solar_peaks_table
    finally:
        rr.close()


# ============================================================
# EXCEL COM WRITE
# ============================================================
def write_results_to_template_com(
    template_path,
    hl_data,
    hg_data,
    solar_peaks_table,
    clg_combined_summary=None,
    htg_sheet_name=HTG_SHEET_NAME,
    clg_sheet_name=CLG_SHEET_NAME,
    htg_marker=HTG_MARKER,
    clg_marker=CLG_MARKER,
    solar_marker=SOLAR_MARKER,   # <-- new
):
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = None
    try:
        log("[XLSX] Opening workbook...")
        wb = excel.Workbooks.Open(str(template_path))

        ws_htg = wb.Worksheets(htg_sheet_name)
        ws_clg = wb.Worksheets(clg_sheet_name)

        htg_anchor = find_marker_cell_fast(ws_htg, htg_marker)
        clg_anchor = find_marker_cell_fast(ws_clg, clg_marker)
        solar_anchor = find_marker_cell_fast(ws_clg, solar_marker)

        if htg_anchor is None:
            raise ValueError(f'Marker "{htg_marker}" not found in sheet "{htg_sheet_name}".')
        if clg_anchor is None:
            raise ValueError(f'Marker "{clg_marker}" not found in sheet "{clg_sheet_name}".')
        if solar_anchor is None:
            raise ValueError(f'Marker "{solar_marker}" not found in sheet "{clg_sheet_name}".')
        
        log(f"[XLSX] HTG marker at: {htg_anchor}")
        log(f"[XLSX] CLG marker at: {clg_anchor}")
        log(f"[XLSX] SOLAR marker at: {solar_anchor}")

        htg_header_row, htg_col = htg_anchor[0] + 1, htg_anchor[1]
        htg_data_row = htg_header_row + 1
        clg_header_row, clg_col = clg_anchor[0] + 1, clg_anchor[1]
        clg_data_row = clg_header_row + 1

        write_2d_block(ws_htg, htg_header_row, htg_col, [HTG_HEADERS])
        write_2d_block(ws_clg, clg_header_row, clg_col, [CLG_HEADERS])

        write_2d_block(ws_htg, htg_data_row, htg_col, hl_data)
        write_2d_block(ws_clg, clg_data_row, clg_col, hg_data)


        # choose first (top-left) match; change to solar_hits[-1] if you prefer last
        solar_start_row, solar_start_col = solar_anchor
        solar_block = [["Peak time table - Solar gain maximums"]] + solar_peaks_table

        log(f"[XLSX] Solar table rows to write: {len(solar_block)}")
        log(f"[XLSX] Writing solar table at row {solar_start_row}, col {solar_start_col}")

        write_2d_block(ws_clg, solar_start_row, solar_start_col, solar_block)

        if WRITE_CLG_COMBINED_SUMMARY and clg_combined_summary:
            summary_row = clg_data_row + len(hg_data) + 1
            summary_block = [[
                f"Combined peak ({clg_combined_summary[3]}) (kW)", clg_combined_summary[0],
                "Month", clg_combined_summary[1], "Time", clg_combined_summary[2]
            ]]
            write_2d_block(ws_clg, summary_row, clg_col, summary_block)

        log("[XLSX] Saving workbook...")
        wb.Save()
        log("[XLSX] Save complete.")
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=True)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


# ============================================================
# VALIDATION / RESOLUTION
# ============================================================
def resolve_rooms(room_groups):
    schemes = room_groups.get_grouping_schemes()
    scheme_handle = None
    for s in schemes:
        if s["name"] == GROUPING_SCHEME_NAME:
            scheme_handle = s["handle"]
            break
    if scheme_handle is None:
        raise RuntimeError(f'Grouping scheme "{GROUPING_SCHEME_NAME}" not found.')

    groups = room_groups.get_room_groups(scheme_handle)
    rooms = []
    for g in groups:
        if g["name"] == ROOM_GROUP_NAME:
            rooms = g["rooms"]
            break
    if not rooms:
        raise RuntimeError(f'No rooms in group "{ROOM_GROUP_NAME}".')
    return rooms


def validate_inputs(project_folder):
    vista_path = project_folder / "Vista"
    if not vista_path.exists():
        raise RuntimeError(f"Vista folder not found: {vista_path}")

    files = os.listdir(vista_path)
    lower_files = [f.lower() for f in files]
    if HTG_FILE.lower() not in lower_files:
        raise RuntimeError(f'HTG_FILE "{HTG_FILE}" not found in {vista_path}')
    if CLG_FILE.lower() not in lower_files:
        raise RuntimeError(f'CLG_FILE "{CLG_FILE}" not found in {vista_path}')

    template = Path(TEMPLATE_PATH)
    if not template.exists():
        raise RuntimeError(f"TEMPLATE_PATH not found: {template}")
    if template.suffix.lower() != ".xlsx":
        raise RuntimeError("TEMPLATE_PATH must point to .xlsx")

    return template, vista_path


# ============================================================
# MAIN
# ============================================================
def main():
    log("=== HGHL Export v3.3 (clean CLG + robust solar peak) ===")

    project = iesve.VEProject.get_current_project()
    results_reader = iesve.ResultsReader
    room_groups = iesve.RoomGroups()

    project_folder = Path(project.path)
    log(f"[INFO] Project folder: {project_folder}")
    log(f"[CFG] HTG_FILE = {HTG_FILE}")
    log(f"[CFG] CLG_FILE = {CLG_FILE}")
    log(f"[CFG] TEMPLATE_PATH = {TEMPLATE_PATH}")
    log(f"[CFG] PEAK_DRIVER = {PEAK_DRIVER}")

    template_path, vista_path = validate_inputs(project_folder)
    htg_path = vista_path / HTG_FILE
    clg_path = vista_path / CLG_FILE

    rooms = resolve_rooms(room_groups)
    log(f"[RUN] Rooms in group: {len(rooms)}")

    if DEBUG_PRINT_Z_VARS:
        rr_dbg = results_reader.open(str(clg_path))
        try:
            z_vars = [v for v in rr_dbg.get_variables() if v.get("model_level") == "z"]
            log(f"[DEBUG] Room-level vars count: {len(z_vars)}")
            for v in z_vars[:120]:
                log(f"[DEBUG] {v.get('display_name')} | {v.get('aps_varname')}")
        finally:
            rr_dbg.close()

    hl_data = collect_heating_data(results_reader, htg_path, rooms)
    hg_data, clg_combined_summary, solar_peaks_table = collect_cooling_data(
        results_reader, clg_path, rooms, peak_driver=PEAK_DRIVER
    )

    if not hg_data:
        log("[CLG][WARN] No cooling rows prepared. Workbook will still be written (headers + tables).")

    write_results_to_template_com(
        template_path=template_path,
        hl_data=hl_data,
        hg_data=hg_data,
        solar_peaks_table=solar_peaks_table,
        clg_combined_summary=clg_combined_summary,
    )

    log(f"[SUCCESS] Results written to: {template_path}")
    os.startfile(str(template_path))


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"[ERROR] {e}")