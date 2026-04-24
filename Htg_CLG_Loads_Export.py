import os
import math
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox as messagebox

import iesve
import pythoncom
import win32com.client as win32


# -----------------------------
# Utilities
# -----------------------------
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


def month_time_from_hour_index(hour_idx):
    month_names = ["May", "June", "July", "August", "September"]
    month_index = hour_idx // 24
    month_index = max(0, min(4, month_index))
    month = month_names[month_index]
    hour_1_24 = (hour_idx % 24) + 1
    return month, f"{hour_1_24:02d}:00"


def find_marker_cell(ws, marker_text, max_rows=20000, max_cols=200):
    for r in range(1, max_rows + 1):
        for c in range(1, max_cols + 1):
            v = ws.Cells(r, c).Value
            if isinstance(v, str) and v.strip() == marker_text:
                return r, c
    return None


def pick_template_file_terminal():
    """
    Fast file picker with no persistent Tk window.
    Prints selected file to terminal.
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        parent=root,
        title="Select template workbook (.xlsx)",
        filetypes=[("Excel workbook", "*.xlsx")]
    )
    root.destroy()

    if path:
        log(f"[TEMPLATE] Selected: {path}")
        return Path(path)

    log("[TEMPLATE] No file selected.")
    return None


# -----------------------------
# IES data collection
# -----------------------------
def collect_heating_data(results_reader, htg_file_name, room_ids_to_analyse):
    log("[HTG] Opening results...")
    rr = results_reader.open(htg_file_name)
    rooms = rr.get_room_list()

    hl_data = []
    running_total_kw = 0.0

    for room in rooms:
        name, room_id, a, b = room
        name, volume, room_area, room_volume = room

        if room_id not in room_ids_to_analyse:
            continue

        np_air_temp = rr.get_room_results(room_id, 'Room air temperature', 'Air temperature', 'z')
        np_dry_resultant_temp = rr.get_room_results(room_id, 'Comfort temperature', 'Dry resultant temperature', 'z')
        np_external_conduction_gain = rr.get_room_results(room_id, 'Conduction from ext elements', 'External conduction gain', 'z')
        np_internal_conduction_gain = rr.get_room_results(room_id, 'Conduction from int surfaces', 'Internal conduction gain', 'z')
        np_infiltration_gain = rr.get_room_results(room_id, 'Infiltration gain', 'Infiltration gain', 'z')
        np_steady_state_heating_plant_load = rr.get_room_results(
            room_id, 'Room units steady state htg load', 'Steady state heating plant load', 'z'
        )

        steady_kw_raw = float(np_steady_state_heating_plant_load / 1000)
        running_total_kw += steady_kw_raw

        row = [
            name,
            round(room_area, 2),
            round(float(np_air_temp), 2),
            round(float(np_dry_resultant_temp), 2),
            round(float(np_external_conduction_gain / 1000), 2),
            round(float(np_internal_conduction_gain / 1000), 2),
            round(float(np_infiltration_gain / 1000), 2),
            round(steady_kw_raw, 2),
            round(running_total_kw, 2),
        ]
        hl_data.append(row)

    log(f"[HTG] Rows prepared: {len(hl_data)}")
    return hl_data


def collect_cooling_data(results_reader, clg_file_name, room_ids_to_analyse):
    log("[CLG] Opening results...")
    rr = results_reader.open(clg_file_name)
    rooms = rr.get_room_list()

    hg_data = []
    all_space_con_series = []

    for room in rooms:
        name, room_id, a, b = room
        name, volume, room_area, room_volume = room

        if room_id not in room_ids_to_analyse:
            continue

        np_air_temp = rr.get_room_results(room_id, 'Room air temperature', 'Air temperature', 'z')
        np_dry_resultant_temp = rr.get_room_results(room_id, 'Comfort temperature', 'Dry resultant temperature', 'z')
        np_internal_gain = rr.get_room_results(room_id, 'Casual gains', 'Internal gain', 'z')
        np_solar_gain = rr.get_room_results(room_id, 'Window solar gains', 'Solar gain', 'z')
        np_conduction_gain = rr.get_room_results(room_id, 'Conduction gain', 'Conduction gain', 'z')
        np_infiltration_gain = rr.get_room_results(room_id, 'Infiltration gain', 'Infiltration gain', 'z')
        np_space_conditioning_sensible = rr.get_room_results(
            room_id, 'System plant etc. gains', 'Space conditioning sensible', 'z'
        )

        all_space_con_series.append(np_space_conditioning_sensible)

        peak_value = min(np_space_conditioning_sensible)
        peak_hour = list(np_space_conditioning_sensible).index(peak_value)
        peak_month, peak_time = month_time_from_hour_index(peak_hour)

        row = [
            name,
            round(room_area, 2),
            peak_month,
            peak_time,
            round(float(np_air_temp[peak_hour]), 2),
            round(float(np_dry_resultant_temp[peak_hour]), 2),
            round(float(np_internal_gain[peak_hour] / 1000), 2),
            round(float(np_solar_gain[peak_hour] / 1000), 2),
            round(float(np_conduction_gain[peak_hour] / 1000), 2),
            round(float(np_infiltration_gain[peak_hour] / 1000), 2),
            round(float(peak_value / 1000), 2),
        ]
        hg_data.append(row)

    if all_space_con_series:
        n_hours = len(all_space_con_series[0])
        combined = [sum(series[h] for series in all_space_con_series) for h in range(n_hours)]
        combined_peak = min(combined)
        combined_hour = combined.index(combined_peak)
        m, t = month_time_from_hour_index(combined_hour)
        hg_data.append([combined_peak, m, t])

    log(f"[CLG] Rows prepared: {len(hg_data)}")
    return hg_data


# -----------------------------
# Excel COM writing
# -----------------------------
def write_results_to_template_com(template_path, hl_data, hg_data):
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = None
    try:
        log("[XLSX] Opening workbook...")
        wb = excel.Workbooks.Open(str(template_path))

        ws_htg = wb.Worksheets("IES - heat loss data (.htg)")
        ws_clg = wb.Worksheets("IES - heat gain data (.clg)")

        htg_marker = find_marker_cell(ws_htg, "IES ZONE HEAT LOSS OUTPUTS")
        clg_marker = find_marker_cell(ws_clg, "IES ZONE HEAT GAIN OUTPUTS")

        if htg_marker is None:
            raise ValueError('Marker "IES ZONE HEAT LOSS OUTPUTS" not found in sheet "IES - heat loss data (.htg)".')
        if clg_marker is None:
            raise ValueError('Marker "IES ZONE HEAT GAIN OUTPUTS" not found in sheet "IES - heat gain data (.clg)".')

        htg_start_row, htg_start_col = htg_marker[0] + 1, htg_marker[1]
        clg_start_row, clg_start_col = clg_marker[0] + 1, clg_marker[1]

        log(f"[XLSX] HTG anchor marker at R{htg_marker[0]}C{htg_marker[1]}, writing starts R{htg_start_row}C{htg_start_col}")
        log(f"[XLSX] CLG anchor marker at R{clg_marker[0]}C{clg_marker[1]}, writing starts R{clg_start_row}C{clg_start_col}")

        # HTG
        r = htg_start_row
        for row in hl_data:
            c = htg_start_col
            for val in row:
                ws_htg.Cells(r, c).Value = to_excel_value(val)
                c += 1
            r += 1

        # CLG (exclude combined summary row if shape=[value, month, time])
        cooling_rows = hg_data[:-1] if hg_data and len(hg_data[-1]) == 3 else hg_data
        r = clg_start_row
        for row in cooling_rows:
            c = clg_start_col
            for val in row:
                ws_clg.Cells(r, c).Value = to_excel_value(val)
                c += 1
            r += 1

        log("[XLSX] Saving workbook...")
        wb.Save()
        log("[XLSX] Save complete.")
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()
        pythoncom.CoUninitialize()


# -----------------------------
# Optional GUI wrapper for file selection + run
# -----------------------------
class Window(tk.Frame):
    def __init__(self, master, project, results_reader, room_groups):
        super().__init__(master)
        self.master = master
        self.project = project
        self.project_folder = Path(project.path)
        self.results_reader = results_reader
        self.room_groups = room_groups

        self.template_path_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")

        self.run_btn = None
        self._build_ui()

    def _build_ui(self):
        self.master.title("HGHL export")
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)

        tk.Label(self, text="Select one .htg and one .clg, then template .xlsx").grid(row=0, column=0, columnspan=3, sticky="w")
        tk.Button(self, text="Create Grouping Scheme", command=self.create_grouping).grid(row=1, column=0, sticky="w", pady=(8, 0))

        tk.Label(self, text="Vista files:").grid(row=2, column=0, sticky="w", pady=(8, 2))
        self.listbox = tk.Listbox(self, selectmode=tk.MULTIPLE, width=80, height=10)
        self.listbox.grid(row=3, column=0, columnspan=3, sticky="nsew")

        vista_path = self.project_folder / "Vista"
        files = os.listdir(vista_path) if vista_path.exists() else []
        for f in [x for x in files if x.lower().endswith(".htg")] + [x for x in files if x.lower().endswith(".clg")]:
            self.listbox.insert(tk.END, f)

        tk.Label(self, text="Template (.xlsx):").grid(row=4, column=0, sticky="w", pady=(8, 2))
        tk.Entry(self, textvariable=self.template_path_var, width=80).grid(row=5, column=0, columnspan=2, sticky="ew")
        tk.Button(self, text="Browse", command=self.browse_template).grid(row=5, column=2, sticky="w")

        self.run_btn = tk.Button(self, text="Run", command=self.run_calc)
        self.run_btn.grid(row=6, column=0, sticky="w", pady=(10, 0))
        tk.Button(self, text="Cancel", command=self.master.destroy).grid(row=6, column=1, sticky="w", pady=(10, 0))

        tk.Label(self, textvariable=self.status_var, fg="blue").grid(row=7, column=0, columnspan=3, sticky="w", pady=(6, 0))

        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)
        self.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

    def browse_template(self):
        path = filedialog.askopenfilename(
            parent=self.master,
            title="Select template workbook (.xlsx)",
            filetypes=[("Excel workbook", "*.xlsx")]
        )
        if path:
            self.template_path_var.set(path)
            log(f"[GUI] Template selected: {path}")

    def create_grouping(self):
        schemes = self.room_groups.get_grouping_schemes()
        if any(s["name"] == "HGHL Analysis" for s in schemes):
            messagebox.showinfo("Grouping scheme", "Grouping scheme already exists.")
            return

        scheme_index = self.room_groups.create_grouping_scheme("HGHL Analysis")
        self.room_groups.create_room_group(scheme_index, "Analyse HGHL Results")
        self.room_groups.create_room_group(scheme_index, "Do Not Analyse")
        messagebox.showinfo("Grouping scheme", "Created. Assign rooms in VE.")
        log("[GUI] Grouping scheme created.")

    def get_selected_files(self):
        selected = self.listbox.curselection()
        if not selected:
            messagebox.showerror("Selection error", "Select one .htg and one .clg file.")
            return None, None

        selected_files = [self.listbox.get(i) for i in selected]
        htg = [f for f in selected_files if f.lower().endswith(".htg")]
        clg = [f for f in selected_files if f.lower().endswith(".clg")]

        if len(htg) != 1 or len(clg) != 1:
            messagebox.showerror("Selection error", f"Need exactly 1 .htg + 1 .clg (got {len(htg)} + {len(clg)}).")
            return None, None

        return htg[0], clg[0]

    def get_rooms_to_analyse(self):
        schemes = self.room_groups.get_grouping_schemes()
        scheme_handle = None
        for s in schemes:
            if s["name"] == "HGHL Analysis":
                scheme_handle = s["handle"]
                break

        if scheme_handle is None:
            messagebox.showerror("Room group error", "Create grouping scheme first.")
            return None

        groups = self.room_groups.get_room_groups(scheme_handle)
        rooms = []
        for g in groups:
            if g["name"] == "Analyse HGHL Results":
                rooms = g["rooms"]
                break

        if not rooms:
            messagebox.showerror("Room group error", "No rooms in 'Analyse HGHL Results'.")
            return None

        return rooms

    def _set_busy(self, busy):
        self.run_btn.config(state="disabled" if busy else "normal")
        self.status_var.set("Running..." if busy else "Ready")
        self.master.update_idletasks()

    def run_calc(self):
        rooms = self.get_rooms_to_analyse()
        if not rooms:
            return

        htg_file, clg_file = self.get_selected_files()
        if not htg_file or not clg_file:
            return

        template = self.template_path_var.get().strip()
        if not template:
            messagebox.showerror("Template error", "Select template workbook.")
            return

        template_path = Path(template)
        if template_path.suffix.lower() != ".xlsx":
            messagebox.showerror("Template error", "Template must be .xlsx")
            return

        self._set_busy(True)

        def worker():
            try:
                log("[RUN] Collecting heating data...")
                hl_data = collect_heating_data(self.results_reader, htg_file, rooms)

                log("[RUN] Collecting cooling data...")
                hg_data = collect_cooling_data(self.results_reader, clg_file, rooms)

                log("[RUN] Writing workbook...")
                write_results_to_template_com(template_path, hl_data, hg_data)

                def ok():
                    self._set_busy(False)
                    log(f"[RUN] Done: {template_path}")
                    self.master.destroy()

                self.master.after(0, ok)
            except Exception as e:
                def fail():
                    self._set_busy(False)
                    log(f"[ERROR] {e}")
                    messagebox.showerror("Run error", str(e))
                self.master.after(0, fail)

        threading.Thread(target=worker, daemon=True).start()


# -----------------------------
# Entry points
# -----------------------------
def run_terminal_mode():
    """
    No persistent window.
    - Picks template via dialog
    - Uses first .htg and first .clg found in Vista (can adjust as needed)
    - Prints all progress to terminal
    """
    project = iesve.VEProject.get_current_project()
    results_reader = iesve.ResultsReader
    room_groups = iesve.RoomGroups()

    project_folder = Path(project.path)
    vista_path = project_folder / "Vista"
    if not vista_path.exists():
        log(f"[ERROR] Vista folder not found: {vista_path}")
        return

    files = os.listdir(vista_path)
    htg_files = [f for f in files if f.lower().endswith(".htg")]
    clg_files = [f for f in files if f.lower().endswith(".clg")]
    if not htg_files or not clg_files:
        log("[ERROR] Need at least one .htg and one .clg in Vista folder.")
        return

    htg_file = htg_files[0]
    clg_file = clg_files[0]
    log(f"[FILES] Using HTG: {htg_file}")
    log(f"[FILES] Using CLG: {clg_file}")

    # get rooms from grouping
    schemes = room_groups.get_grouping_schemes()
    scheme_handle = None
    for s in schemes:
        if s["name"] == "HGHL Analysis":
            scheme_handle = s["handle"]
            break
    if scheme_handle is None:
        log("[ERROR] Grouping scheme 'HGHL Analysis' not found.")
        return

    groups = room_groups.get_room_groups(scheme_handle)
    rooms = []
    for g in groups:
        if g["name"] == "Analyse HGHL Results":
            rooms = g["rooms"]
            break
    if not rooms:
        log("[ERROR] No rooms in 'Analyse HGHL Results'.")
        return

    template_path = pick_template_file_terminal()
    if template_path is None:
        return

    try:
        log("[RUN] Collecting heating data...")
        hl_data = collect_heating_data(results_reader, htg_file, rooms)

        log("[RUN] Collecting cooling data...")
        hg_data = collect_cooling_data(results_reader, clg_file, rooms)

        log("[RUN] Writing workbook...")
        write_results_to_template_com(template_path, hl_data, hg_data)

        log(f"[SUCCESS] Written to: {template_path}")
        os.startfile(str(template_path))
    except Exception as e:
        log(f"[ERROR] {e}")


def run_gui_mode():
    project = iesve.VEProject.get_current_project()
    results_reader = iesve.ResultsReader
    room_groups = iesve.RoomGroups()

    root = tk.Tk()
    Window(root, project, results_reader, room_groups)
    root.mainloop()


if __name__ == "__main__":
    # Set to True for "no persistent window" mode requested
    TERMINAL_MODE = True

    if TERMINAL_MODE:
        run_terminal_mode()
    else:
        run_gui_mode()