import csv
from pathlib import Path
import iesve


# Set your files here
RESULT_FILES = [
    r"C:\_VE_Projects\IES-Input-output\Loads-20260121-JL-IES_Origin-TLY 22-23\vista\TL - Loads.clg",
    r"C:\_VE_Projects\IES-Input-output\Loads-20260121-JL-IES_Origin-TLY 22-23\vista\TL - Loads.htg",
    r"C:\_VE_Projects\IES-Input-output\Loads-20260121-JL-IES_Origin-TLY 22-23\vista\TL - Apache.aps",
]

OUT_DIR = Path(r"C:\_VE_Projects\IES-Input-output")
OUT_DIR.mkdir(parents=True, exist_ok=True)


def export_variables_csv(results_path):
    results_path = Path(results_path)
    rr = iesve.ResultsReader.open(str(results_path))
    try:
        vars_list = rr.get_variables() or []
        out_csv = OUT_DIR / f"{results_path.stem}_variables.csv"

        # Union of all keys so we don't miss columns
        all_keys = set()
        for v in vars_list:
            all_keys.update(v.keys())
        fieldnames = sorted(all_keys)

        with out_csv.open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            for row in vars_list:
                w.writerow({k: row.get(k, "") for k in fieldnames})

        print(f"[OK] {results_path.name} -> {out_csv}")
    finally:
        rr.close()


for fp in RESULT_FILES:
    p = Path(fp)
    if not p.exists():
        print(f"[WARN] Missing file: {p}")
        continue
    export_variables_csv(p)