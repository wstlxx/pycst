"""Read Tables\\1D Results\\Mix 1DC_1 and write to CSV."""

import csv
import sys
from pathlib import Path

import numpy as np

from cst_interface import CSTInterface

PROJECT = Path(r"D:/workspace/cstwork/original/20260609.cst")
TREE = r"Tables\1D Results\Mix 1DC_1"
OUT_DIR = Path(r"D:/workspace/cstwork/pycst")


def main() -> int:
    if not PROJECT.exists():
        print(f"Project not found: {PROJECT}")
        return 1

    cst = CSTInterface()
    try:
        cst.connect_to_cst_or_start_it()
        cst.silent_mode = True
        cst.open_project(str(PROJECT))

        # First: discover RunIDs for that tree item.
        ids = cst.get_result_ids_from_tree_item(TREE)
        print(f"RunIDs at {TREE!r}: {ids}")

        # Then: read 1D result (real or complex).
        x, y, info = cst.get_1d_result_from_tree_item(
            TREE, run_id_filter=None, query_x=None)
        print(f"x.shape = {x.shape}, y.shape = {y.shape}, y.dtype = {y.dtype}")
        print(f"info: res_type={info.get('res_type')!r}, "
              f"x_label={info.get('x_label')!r}, "
              f"y_label={info.get('y_label')!r}, "
              f"title={info.get('title')!r}")

        # Build CSV columns.
        if y.dtype.kind == "c":
            col_data = [np.asarray(x, dtype=float).ravel()]
            col_names = ["x"]
            for j, run_id in enumerate(info["run_ids"]):
                col_data.append(np.asarray(y[:, j].real, dtype=float).ravel())
                col_names.append(f"y_re[{run_id}]")
                col_data.append(np.asarray(y[:, j].imag, dtype=float).ravel())
                col_names.append(f"y_im[{run_id}]")
        else:
            col_data = [np.asarray(x, dtype=float).ravel()]
            col_data += [np.asarray(y[:, j], dtype=float).ravel()
                         for j in range(y.shape[1])]
            col_names = ["x"] + [f"y[{rid}]" for rid in info["run_ids"]]

        n_rows = len(col_data[0])
        print(f"\nCSV: {n_rows} rows, {len(col_data)} columns")
        for i, c in enumerate(col_data):
            print(f"  col {i} {col_names[i]!r}: "
                  f"shape={getattr(c, 'shape', '?')}, "
                  f"len={len(c) if hasattr(c, '__len__') else 'N/A'}")


        out = OUT_DIR / "mix_1dc_1.csv"
        with out.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(col_names)
            for i in range(n_rows):
                row = [c[i] for c in col_data]
                w.writerow([f"{v:.10g}" for v in row])
        print(f"\nWrote {out}  ({n_rows} rows, {len(col_names)} columns)")
        print("First 3 rows:")
        with out.open("r", encoding="utf-8") as f:
            for i, line in enumerate(f):
                if i >= 4:
                    break
                print("  " + line.rstrip())

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"FAILED: {e}", file=sys.stderr)
        return 2
    finally:
        try:
            cst.close_project(save_project=False)
        except Exception:
            pass

    return 0


if __name__ == "__main__":
    sys.exit(main())
