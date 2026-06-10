"""
pycst usage demo.

Importable module — no top-level side effects, no `__main__` block.
Each function demonstrates one logical group of `CSTInterface` methods,
with comments explaining what each call does and what it returns.

Typical usage:

    from cst_interface import CSTInterface
    import demo_api

    cst = demo_api.setup()                  # connect + open project
    n  = demo_api.demo_param_read(cst)      # read parameters
    demo_api.demo_param_write(cst, "a1", 5.1)
    demo_api.demo_solve(cst)
    x, y, info = demo_api.demo_1d_result(cst, r"Tables\\1D Results\\Mix 1DC_1")
    demo_api.demo_close(cst)                # close project
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import numpy as np

from cst_interface import CSTInterface


# ── Constants ───────────────────────────────────────────────────────
# Path to the test project used by the demos. Override by passing your
# own path to `setup(...)`.
DEFAULT_PROJECT = Path(r"D:/workspace/cstwork/original/20260609.cst")

# Tree path used in result-reading demos. Adjust to whatever your
# project actually contains.
MIX_TREE_ITEM = r"Tables\1D Results\Mix 1DC_1"
S_TREE_PATH   = r"1D Results\S-Parameters"
Z_TREE_PATH   = r"1D Results\Z Matrix"


# ── Setup / teardown ───────────────────────────────────────────────
def setup(project: str | Path | None = None,
          silent: bool = True) -> CSTInterface:
    """
    Connect to (or start) CST and open a project. Returns the live
    `CSTInterface`. Does not save anything on close unless you ask.

    - `project` may be:
        * a path to a `.cst` file  → opens that project
        * None                    → attaches to whichever project is
                                   already active in CST
    - `silent=True` suppresses progress dots from `open_project`.
    """
    cst = CSTInterface()
    # Lazily binds the COM dispatch handle for `CSTStudio.Application`.
    # Starts CST if no instance is running.
    cst.connect_to_cst_or_start_it()
    cst.silent_mode = silent
    if project is None:
        # No file → attach to the currently active project, if any.
        cst.open_project()
    else:
        # Resolves to an absolute path internally and polls
        # `GetProjectPath` until the project is fully loaded.
        cst.open_project(Path(project))
    return cst


def demo_close(cst: CSTInterface, save: bool = False) -> None:
    """
    Close the current project. `save=True` saves first; default is
    to discard any unsaved changes. Falls back to quitting CST if
    the project handle is gone.
    """
    cst.close_project(save_project=save)


# ── Parameter reading ──────────────────────────────────────────────
def demo_param_read(cst: CSTInterface) -> dict[str, float]:
    """
    Read every parameter in the project. Returns a `name → value` dict.

    Demonstrates:
      * `get_parameter_list`
      * `get_number_of_parameters`
      * `get_parameter_value`        (index-based, fastest)
      * `get_parameter_expression`   (the raw design expression string)
      * `get_current_value_of_all_parameters`
      * `get_current_expression_of_all_parameters`
      * `parameter_exists`           (case-insensitive existence check)
    """
    # Number of parameters declared in the project.
    n = cst.get_number_of_parameters()
    # Names in declaration order. Internal loop uses 0-based indices.
    names = cst.get_parameter_list()
    assert len(names) == n

    # Read values one at a time by *index* (1-based in the public API).
    # Index access avoids a name lookup, so it is faster than name access.
    values_by_index = [cst.get_parameter_value(i + 1) for i in range(n)]

    # Read the design expression (e.g. "1.5*sqrt(2)") — useful when the
    # parameter is a formula rather than a literal number.
    expressions = [cst.get_parameter_expression(i + 1) for i in range(n)]

    # Bulk helpers — equivalent to the loops above, but as a single tuple.
    all_names, all_values = cst.get_current_value_of_all_parameters()
    all_names2, all_exprs  = cst.get_current_expression_of_all_parameters()

    # Check whether a parameter exists (case-insensitive). The second
    # argument throws ValueError instead of returning False.
    has_a1 = cst.parameter_exists("a1", throw_error_if_not=False)
    # Resolve a name to its 1-based index (also case-insensitive).
    a1_idx = cst.get_parameter_index_by_name("A1")
    # And the reverse: index → name.
    a1_name = cst.get_parameter_name_by_index(a1_idx)

    return dict(zip(all_names, all_values))


# ── Parameter writing ──────────────────────────────────────────────
def demo_param_write(cst: CSTInterface, name: str, value: float) -> None:
    """
    Set a single parameter, rebuilding the model once.

    Demonstrates:
      * `change_parameter`           (existence check + store + rebuild)
      * `store_parameter`            (no existence check — creates if absent)
    """
    # `change_parameter` checks that `name` exists, then either calls
    # StoreParameter (string value) or StoreDoubleParameter (numeric),
    # and finally RebuildOnParametricChange(False, False) to refresh the
    # geometry. `update_structure=False` skips the rebuild.
    cst.change_parameter(name, value, update_structure=True)

    # `store_parameter` is the same minus the existence check, so use
    # it when you want to create a brand-new parameter:
    # cst.store_parameter("new_param", 1.23, update_structure=True)


def demo_param_batch(cst: CSTInterface) -> None:
    """
    Update several parameters in one go, rebuilding only at the end.
    This matches the MATLAB `ChangeParameters` pattern.

    Demonstrates:
      * `change_parameters`          (loop with single final rebuild)
    """
    # Single rebuild no matter how many parameters are changed —
    # much faster than calling `change_parameter` in a Python loop.
    cst.change_parameters(
        ["a1", "a2", "b1"],     # names (str or iterable of str)
        [5.1, 5.2, 13.5],        # matching values (scalar or iterable)
        update_structure=True,   # only the last change triggers rebuild
    )


# ── Solving ────────────────────────────────────────────────────────
def demo_solve(cst: CSTInterface, tries: int = 1) -> str:
    """
    Run the active solver. Returns the solver type that was used.

    Demonstrates:
      * `solve`
    """
    # Inspect the solver type ("HF Time Domain" / "HF Frequency Domain"
    # / "HF Eigenmode") that the project is currently configured for.
    solver_type = cst._proj.GetSolverType()
    # Start the solver; on transient failure it retries up to `tries`
    # times, sleeping 2 s between attempts. For real runs this call
    # blocks for as long as the simulation needs.
    cst.solve(number_of_tries=tries)
    return solver_type


# ── Result tree navigation ─────────────────────────────────────────
def demo_tree(cst: CSTInterface) -> list[str]:
    """
    Walk the result tree under `1D Results`. Returns all immediate
    children of the parent path.

    Demonstrates:
      * `tree_item_exists`
      * `get_result_tree_item_children`
      * `get_result_ids_from_tree_item`
    """
    # Boolean check; second arg raises ValueError when False.
    if not cst.tree_item_exists("1D Results"):
        return []

    # Immediate children — for this project, "S-Parameters",
    # "Port Information", "Power", "Reference Impedance", etc.
    top_children = cst.get_result_tree_item_children("1D Results")

    # Children of a specific sub-tree.
    s_kids: list[str] = []
    if cst.tree_item_exists(S_TREE_PATH):
        s_kids = cst.get_result_tree_item_children(S_TREE_PATH)
        # Each kid has one or more RunIDs; with multiple parametric-sweep
        # runs they show up as "3D:RunID:0", "3D:RunID:1", ...
        first = s_kids[0]
        run_ids = cst.get_result_ids_from_tree_item(first)

    return top_children


# ── 1D result reading ──────────────────────────────────────────────
def demo_1d_result(cst: CSTInterface,
                   tree_item: str = MIX_TREE_ITEM
                   ) -> tuple[np.ndarray, np.ndarray, dict[str, Any]]:
    """
    Fetch a single 1D result from the result tree. Returns
    `(x, y, info)` where:

      * `x`    : 1-D `np.ndarray` of the X axis (e.g. frequency)
      * `y`    : 2-D `np.ndarray`, shape `(len(x), n_runs)`, real or
                 complex depending on the source result
      * `info` : dict with `tree_item`, `run_ids`, `res_type` ("1D" or
                 "1DC"), `x_label`, `y_label`, `title`, and (when
                 available) `impedance`.

    Demonstrates:
      * `get_1d_result_from_tree_item`
      * `cst.nearest_or_interpolated` property (Exact / Nearest / Interpolated)
      * `run_id_filter` (positive = include, negative / `~` = exclude)
      * `query_x` (sample at specific X values)
    """
    # Choose how off-grid query points are handled:
    #   "Exact"        → must hit a simulated point
    #   "Nearest"      → snap to the closest point (default)
    #   "Interpolated" → linear-interpolate between bracketing points
    cst.nearest_or_interpolated = "Nearest"

    # Read all runs, full X axis:
    x, y, info = cst.get_1d_result_from_tree_item(
        tree_item,
        run_id_filter=None,   # accept every available RunID
        query_x=None,         # full X axis; do not sub-sample
    )

    # Example: keep only RunID 2 (positive → include).
    _x2, y2, _info2 = cst.get_1d_result_from_tree_item(
        tree_item, run_id_filter=[2])

    # Example: keep every run *except* RunID 1 (negative → exclude).
    _x3, y3, _info3 = cst.get_1d_result_from_tree_item(
        tree_item, run_id_filter=[-1])

    # Example: query a few specific frequencies, interpolating if needed.
    cst.nearest_or_interpolated = "Interpolated"
    xq, yq, _ = cst.get_1d_result_from_tree_item(
        tree_item, run_id_filter=None, query_x=[1.0, 2.5, 5.0])

    return x, y, info


# ── S / Z parameters ───────────────────────────────────────────────
def demo_s_params(cst: CSTInterface
                  ) -> tuple[np.ndarray, np.ndarray, dict[str, Any]]:
    """
    Read the S-parameter matrix. Returns
    `(s, freqs, info)` with shapes:

      * `s`     : `(n_rows, n_cols, n_freqs, n_runs)`, complex
      * `freqs` : `(n_freqs,)`, real
      * `info`  : dict with `tree_path`, `rows`, `cols`, `elements`

    Demonstrates:
      * `get_s_params`
    """
    # All matrix elements under the default S-parameter tree.
    s, freqs, info = cst.get_s_params()

    # Sub-select specific ports (or specific RunIDs, or specific freqs).
    s11, f11, _ = cst.get_s_params(
        i_rows=[1], i_cols=[1],            # only S11
        run_id_filter=[-1],                # all but the most recent run
        freqs_to_get=None,                 # full frequency axis
    )
    return s, freqs, info


def demo_z_params(cst: CSTInterface
                  ) -> tuple[np.ndarray, np.ndarray, dict[str, Any]]:
    """
    Read the Z-parameter matrix. Same shape contract as S-parameters.
    Not all projects contain a Z Matrix; in that case this raises.

    Demonstrates:
      * `get_z_params`
    """
    z, freqs, info = cst.get_z_params()
    return z, freqs, info


# ── End-to-end convenience: a typical "change → solve → fetch" loop ─
def demo_resim(cst: CSTInterface,
               param_name: str = "a1",
               new_value: float = 5.1,
               tree_item: str = MIX_TREE_ITEM
               ) -> tuple[np.ndarray, np.ndarray, dict[str, Any]]:
    """
    Re-run a simulation for one parameter change and pull the result.
    Convenience that exercises the full change → solve → read chain.

    Returns `(x, y, info)` from `get_1d_result_from_tree_item`.
    """
    # 1. Update geometry parameter. Single rebuild (faster than
    #    rebuilding after every individual change).
    cst.change_parameter(param_name, new_value, update_structure=True)

    # 2. Run the configured solver. Blocks until done.
    cst.solve(number_of_tries=1)

    # 3. Read the result we care about. Default = latest run, full X.
    cst.nearest_or_interpolated = "Nearest"
    return cst.get_1d_result_from_tree_item(tree_item)
