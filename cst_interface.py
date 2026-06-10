"""
CSTInterface - Python interface to CST Microwave Studio via COM automation.
Requires: Windows, pywin32, numpy

【首次使用】先生成早绑定缓存（只需运行一次）：
    python -m win32com.client.makepy "D:\\Program Files (x86)\\CST Studio Suite 2024\\Patches\\BASE_2023-09-01\\AMD64\\CST DESIGN ENVIRONMENT_AMD64.exe"
"""

import os
import re
import tempfile
import time
from pathlib import Path

import numpy as np

try:
    import win32com.client
    from win32com.client import gencache
    import pythoncom
except ImportError:
    raise ImportError("pywin32 is required: pip install pywin32")


_CST_APP_PROGID = "CSTStudio.Application"

_DISPATCH_METHOD  = pythoncom.DISPATCH_METHOD
_DISPATCH_PROPGET = pythoncom.DISPATCH_PROPERTYGET


class _Proj:
    """
    对 PyIDispatch 的轻量长寿命包装。
    【终极修复】彻底移除 DISPID 缓存机制。根据 CST 底层特性，
    每次调用方法前必须实时通过 GetIDsOfNames 强刷映射，以此确保连续调用绝不失效。
    """

    def __init__(self, raw: "PyIDispatch"):  # noqa: F821
        object.__setattr__(self, "_raw", raw)

    def _invoke(self, name: str, flags: int, *args):
        raw = object.__getattribute__(self, "_raw")
        # 核心修复：绝不缓存 DISPID，每次 Invoke 前都强行重新获取一次，刷新通道上下文
        dispid = raw.GetIDsOfNames(0, name)
        return raw.Invoke(dispid, 0, flags, True, *args)

    # ── 公开方法（由动态派发器接管） ─────────────────────────

    def GetProjectPath(self, path_type: str = "Project") -> str:
        return self._invoke("GetProjectPath", _DISPATCH_METHOD, path_type)

    def Save(self):
        self._invoke("Save", _DISPATCH_METHOD)

    def CloseProject(self, filename: str):
        # [修复] 顺应底层映射，直接在项目级安全调用 CloseProject (DISPID 153)
        self._invoke("CloseProject", _DISPATCH_METHOD, filename)

    def GetSolverType(self) -> str:
        return self._invoke("GetSolverType", _DISPATCH_METHOD)

    def GetNumberOfParameters(self) -> int:
        return self._invoke("GetNumberOfParameters", _DISPATCH_METHOD)

    def GetParameterName(self, index: int) -> str:
        return self._invoke("GetParameterName", _DISPATCH_METHOD, index)

    def GetParameterNValue(self, index: int) -> float:
        return self._invoke("GetParameterNValue", _DISPATCH_METHOD, index)

    def GetParameterSValue(self, index: int) -> str:
        return self._invoke("GetParameterSValue", _DISPATCH_METHOD, index)

    def DoesParameterExist(self, name: str) -> bool:
        return bool(self._invoke("DoesParameterExist", _DISPATCH_METHOD, name))

    def StoreGlobalDataValue(self, key: str, value: str):
        self._invoke("StoreGlobalDataValue", _DISPATCH_METHOD, key, value)

    def RestoreGlobalDataValue(self, key: str) -> str:
        return self._invoke("RestoreGlobalDataValue", _DISPATCH_METHOD, key)

    def StoreParameter(self, name: str, value: str):
        self._invoke("StoreParameter", _DISPATCH_METHOD, name, value)

    def StoreDoubleParameter(self, name: str, value: float):
        self._invoke("StoreDoubleParameter", _DISPATCH_METHOD, name, value)

    def RebuildOnParametricChange(self, rebuild: bool = False,
                                  update: bool = False) -> bool:
        return bool(self._invoke("RebuildOnParametricChange", _DISPATCH_METHOD,
                                 rebuild, update))

    # ── 子对象衍生 ───────────────────────────────────────────

    def Resulttree(self) -> "_Proj":
        raw = self._invoke("Resulttree", _DISPATCH_PROPGET)
        return _Proj(raw)

    def Solver(self) -> "_Proj":
        raw = self._invoke("Solver", _DISPATCH_PROPGET)
        return _Proj(raw)

    def FDSolver(self) -> "_Proj":
        raw = self._invoke("FDSolver", _DISPATCH_PROPGET)
        return _Proj(raw)

    def EigenmodeSolver(self) -> "_Proj":
        raw = self._invoke("EigenmodeSolver", _DISPATCH_PROPGET)
        return _Proj(raw)

    # ── Resulttree 子方法 ─────────────────────────────────────

    def GetResultFromTreeItem(self, item: str, run_id: str):
        raw = self._invoke("GetResultFromTreeItem", _DISPATCH_METHOD,
                           item, run_id)
        return _Result(raw)

    def GetImpedanceResultFromTreeItem(self, item: str, run_id: str):
        raw = self._invoke("GetImpedanceResultFromTreeItem", _DISPATCH_METHOD,
                           item, run_id)
        return _Result(raw)

    def TreeItemHasImpedance(self, item: str, run_id: str) -> bool:
        return bool(self._invoke("TreeItemHasImpedance", _DISPATCH_METHOD,
                                 item, run_id))

    def DoesTreeItemExist(self, item: str) -> bool:
        return bool(self._invoke("DoesTreeItemExist", _DISPATCH_METHOD, item))

    def GetResultIDsFromTreeItem(self, item: str):
        return self._invoke("GetResultIDsFromTreeItem", _DISPATCH_METHOD, item)

    def GetFirstChildName(self, item: str) -> str:
        return self._invoke("GetFirstChildName", _DISPATCH_METHOD, item)

    def GetNextItemName(self, item: str) -> str:
        return self._invoke("GetNextItemName", _DISPATCH_METHOD, item)

    # ── Solver 子方法 ─────────────────────────────────────────

    def Start(self) -> bool:
        return bool(self._invoke("Start", _DISPATCH_METHOD))


class _Result:
    """对 CST result COM 对象的轻量包装。"""

    def __init__(self, raw: "PyIDispatch"):  # noqa: F821
        object.__setattr__(self, "_raw", raw)

    def _invoke(self, name: str, flags: int, *args):
        raw = object.__getattribute__(self, "_raw")
        dispid = raw.GetIDsOfNames(0, name)
        return raw.Invoke(dispid, 0, flags, True, *args)

    def GetResultObjectType(self) -> str:
        return self._invoke("GetResultObjectType", _DISPATCH_METHOD)

    def GetArray(self, name: str):
        return self._invoke("GetArray", _DISPATCH_METHOD, name)

    def GetXLabel(self) -> str:
        return self._invoke("GetXLabel", _DISPATCH_METHOD)

    def GetYLabel(self) -> str:
        return self._invoke("GetYLabel", _DISPATCH_METHOD)

    def GetTitle(self) -> str:
        return self._invoke("GetTitle", _DISPATCH_METHOD)


def _get_app():
    """连接或启动 CST，返回早绑定 IApplication。"""
    try:
        return gencache.EnsureDispatch(_CST_APP_PROGID)
    except Exception:
        try:
            return win32com.client.GetActiveObject(_CST_APP_PROGID)
        except Exception:
            return win32com.client.Dispatch(_CST_APP_PROGID)


def _get_proj(app) -> "_Proj | None":
    """获取常驻项目指针"""
    try:
        raw = app._oleobj_.Invoke(1, 0, _DISPATCH_METHOD, True)
        if raw is None:
            return None
        return _Proj(raw)
    except Exception:
        return None


class CSTInterface:
    """Python interface to CST Microwave Studio via COM automation."""

    VERSION = "1.5.0_pure_nocache_invoke"
    CST_VERSION_TESTED = "2024"

    def __init__(self, project_file: str | None = None):
        self._app    = None
        self._proj   = None   # _Proj 实例
        self._solver = None
        self._silent_mode = False
        self._nearest_or_interpolated = "Nearest"
        self._temp_dir = os.path.join(tempfile.gettempdir(), "PyCSTTempFiles")

        if project_file is not None:
            self.open_project(project_file)

    # =========================================================================
    # Properties
    # =========================================================================

    @property
    def application(self):
        return self._app

    @property
    def project(self):
        return self._proj

    @property
    def silent_mode(self) -> bool:
        return self._silent_mode

    @silent_mode.setter
    def silent_mode(self, value: bool):
        self._silent_mode = bool(value)

    @property
    def nearest_or_interpolated(self) -> str:
        return self._nearest_or_interpolated

    @nearest_or_interpolated.setter
    def nearest_or_interpolated(self, value: str):
        valid = ("Exact", "Nearest", "Interpolated")
        if value not in valid:
            raise ValueError(f'Must be one of {valid}, got "{value}"')
        self._nearest_or_interpolated = value

    # =========================================================================
    # Connection & Project Management
    # =========================================================================

    def connect_to_cst_or_start_it(self):
        self._app = _get_app()

    def open_project(self, full_file_name: str | None = None):
        if self._app is None:
            self.connect_to_cst_or_start_it()

        if full_file_name is None:
            self._proj = _get_proj(self._app)
            if self._proj is None:
                raise RuntimeError("No active project in CST.")
        else:
            full_file_name = str(Path(full_file_name).resolve())
            self._app.OpenFile(full_file_name)

            self._print("Waiting for project to load")
            deadline = time.time() + 60
            while time.time() < deadline:
                time.sleep(1)
                self._print(".")
                try:
                    proj = _get_proj(self._app)
                    if proj is None:
                        continue
                    # 动态探测路径以验证加载状态
                    p_path = proj.GetProjectPath("Project")
                    if isinstance(p_path, str) and p_path.strip() != "":
                        self._proj = proj
                        break
                except Exception:
                    continue
            else:
                raise RuntimeError(
                    f'Timeout: project did not finish loading: "{full_file_name}"')
            self._print("\n")

        try:
            proj_path = self._proj.GetProjectPath("Project")
        except Exception:
            proj_path = full_file_name if full_file_name else "Unknown"
        self._print(f'Successfully connected to project: "{proj_path}"\n')

    def close_project(self, full_file_name: str | None = None,
                      save_project: bool = False):
        if self._app is None:
            self.connect_to_cst_or_start_it()
        if full_file_name is None:
            self._check_project_is_open(True)
            try:
                full_file_name = self._proj.GetProjectPath("Project") + ".cst"
            except Exception:
                self._app.Quit()
                return
        else:
            full_file_name = str(Path(full_file_name).resolve())
        
        if save_project and self._proj:
            try:
                self._proj.Save()
            except Exception:
                pass

        # [修复] 优先通过无缓存的项目级代理安全调用 CloseProject (DISPID 153)
        if self._proj:
            try:
                self._proj.CloseProject(full_file_name)
                return
            except Exception:
                pass
        
        # 兜底清理
        try:
            self._app.CloseProject(full_file_name)
        except Exception:
            try:
                self._app.Quit()
            except Exception:
                pass

    # =========================================================================
    # Parameter Management
    # =========================================================================

    def get_number_of_parameters(self) -> int:
        self._check_project_is_open(True)
        n = self._proj.GetNumberOfParameters()
        if not isinstance(n, int):
            raise RuntimeError(f"GetNumberOfParameters() returned: {repr(n)}")
        return n

    def get_parameter_list(self) -> list[str]:
        self._check_project_is_open(True)
        n = self.get_number_of_parameters()
        # [修复] 随着每一轮迭代都在底层强刷 GetIDsOfNames，22个变量名将畅通无阻地全部读出！
        return [str(self._proj.GetParameterName(i)) for i in range(n)]

    def parameter_exists(self, param_name: str,
                         throw_error_if_not: bool = False) -> bool:
        self._check_project_is_open(True)
        res = self._proj.DoesParameterExist(param_name)
        if throw_error_if_not and not res:
            raise ValueError(f'Parameter "{param_name}" does not exist.')
        return res

    def get_parameter_name_by_index(self, param_index: int) -> str:
        self._check_project_is_open(True)
        if param_index < 1:
            raise ValueError("param_index must be >= 1")
        return str(self._proj.GetParameterName(param_index - 1))

    def get_parameter_index_by_name(self, param_name: str) -> int:
        self._check_project_is_open(True)
        if not param_name:
            raise ValueError("param_name cannot be None/empty")
        self.parameter_exists(param_name, True)
        for i, n in enumerate(self.get_parameter_list()):
            if n.lower() == param_name.lower():
                return i + 1
        raise ValueError(f'Parameter "{param_name}" not found.')

    def get_parameter_value(self, param_name_or_index: str | int) -> float:
        self._check_project_is_open(True)
        idx = (self.get_parameter_index_by_name(param_name_or_index)
               if isinstance(param_name_or_index, str)
               else int(param_name_or_index))
        if idx < 1:
            raise ValueError("Parameter index must be >= 1")
        return float(self._proj.GetParameterNValue(idx - 1))

    def get_parameter_expression(self, param_name_or_index: str | int) -> str:
        self._check_project_is_open(True)
        idx = (self.get_parameter_index_by_name(param_name_or_index)
               if isinstance(param_name_or_index, str)
               else int(param_name_or_index))
        return str(self._proj.GetParameterSValue(idx - 1))

    def get_current_value_of_all_parameters(
            self) -> tuple[list[str], np.ndarray]:
        names = self.get_parameter_list()
        values = np.array([self.get_parameter_value(i + 1)
                           for i in range(len(names))])
        return names, values

    def get_current_expression_of_all_parameters(
            self) -> tuple[list[str], list[str]]:
        names = self.get_parameter_list()
        return names, [self.get_parameter_expression(i + 1)
                       for i in range(len(names))]

    # =========================================================================
    # Solving
    # =========================================================================

    def solve(self, number_of_tries: int = 1):
        self._check_project_is_open(True)
        solver_type = self._proj.GetSolverType()
        solver_map = {
            "HF Time Domain":      "Solver",
            "HF Frequency Domain": "FDSolver",
            "HF Eigenmode":        "EigenmodeSolver",
        }
        if solver_type not in solver_map:
            raise RuntimeError(f'Unknown solver type "{solver_type}".')
        self._solver = getattr(self._proj, solver_map[solver_type])()
        self._run_solver(number_of_tries)

    # =========================================================================
    # Result Tree
    # =========================================================================

    def tree_item_exists(self, tree_item: str,
                         throw_error_if_not: bool = False) -> bool:
        self._check_project_is_open(True)
        rt = self._proj.Resulttree()
        res = rt.DoesTreeItemExist(tree_item)
        if throw_error_if_not and not res:
            raise ValueError(f'Tree item "{tree_item}" does not exist.')
        return res

    def get_result_ids_from_tree_item(self, tree_item: str) -> list[str]:
        self._check_project_is_open(True)
        rt = self._proj.Resulttree()
        if not rt.DoesTreeItemExist(tree_item):
            raise ValueError(f'Tree item "{tree_item}" does not exist.')
        ids = rt.GetResultIDsFromTreeItem(tree_item)
        return self._ensure_list(ids)

    def get_result_tree_item_children(self,
                                      parent_tree_item: str) -> list[str]:
        self._check_project_is_open(True)
        rt = self._proj.Resulttree()
        if not rt.DoesTreeItemExist(parent_tree_item):
            raise ValueError(
                f'Parent tree item "{parent_tree_item}" does not exist.')
        children, child = [], rt.GetFirstChildName(parent_tree_item)
        while child:
            children.append(child)
            child = rt.GetNextItemName(child)
        return children

    def get_1d_result_from_tree_item(self, tree_item: str,
                                     run_id_filter=None,
                                     query_x=None) -> tuple[np.ndarray,
                                                            np.ndarray,
                                                            dict]:
        self._check_project_is_open(True)
        rt = self._proj.Resulttree()
        if not rt.DoesTreeItemExist(tree_item):
            raise ValueError(f'Tree item "{tree_item}" does not exist.')

        run_ids = self._apply_run_id_filter(
            self._ensure_list(rt.GetResultIDsFromTreeItem(tree_item)),
            run_id_filter)
        if not run_ids:
            info = {"run_ids": [], "tree_item": tree_item}
            return np.array([], dtype=float), np.empty((0, 0)), info

        first = rt.GetResultFromTreeItem(tree_item, run_ids[0])
        res_type = str(first.GetResultObjectType())
        if res_type not in ("1D", "1DC"):
            raise RuntimeError(
                f'Tree item "{tree_item}" contains unsupported result '
                f'type "{res_type}".')

        x = self._to_1d_array(first.GetArray("x"), dtype=float)
        x_out, ix = self._find_queried_x(x, query_x)

        first_y = self._read_result_y(first, res_type)
        dtype = complex if np.iscomplexobj(first_y) else float
        y_out = np.empty((len(x_out), len(run_ids)), dtype=dtype)
        y_out[:, 0] = self._get_indexed_y(x, first_y, ix, query_x)

        info = {
            "tree_item": tree_item,
            "run_ids": run_ids,
            "res_type": res_type,
            "x_label": self._safe_result_attr(first, "GetXLabel"),
            "y_label": self._safe_result_attr(first, "GetYLabel"),
            "title": self._safe_result_attr(first, "GetTitle"),
        }

        has_impedance = False
        try:
            has_impedance = rt.TreeItemHasImpedance(tree_item, run_ids[0])
        except Exception:
            has_impedance = False

        for col, run_id in enumerate(run_ids[1:], start=1):
            obj = rt.GetResultFromTreeItem(tree_item, run_id)
            typ = str(obj.GetResultObjectType())
            if typ != res_type:
                raise RuntimeError(
                    f'Result "{tree_item}" run ID "{run_id}" has type '
                    f'"{typ}", expected "{res_type}".')
            xr = self._to_1d_array(obj.GetArray("x"), dtype=float)
            yr = self._read_result_y(obj, typ)
            self._check_x_is_same(x, xr, run_id, query_x)
            y_out[:, col] = self._get_indexed_y(xr, yr, ix, query_x)

        if has_impedance:
            z_out = np.empty((len(x_out), len(run_ids)), dtype=complex)
            for col, run_id in enumerate(run_ids):
                try:
                    z_obj = rt.GetImpedanceResultFromTreeItem(tree_item, run_id)
                    z_type = str(z_obj.GetResultObjectType())
                    xr = self._to_1d_array(z_obj.GetArray("x"), dtype=float)
                    zr = self._read_result_y(z_obj, z_type)
                    self._check_x_is_same(x, xr, run_id, query_x)
                    z_out[:, col] = self._get_indexed_y(xr, zr, ix, query_x)
                except Exception:
                    z_out[:, col] = np.nan
            if z_out.size and np.allclose(z_out, z_out.flat[0], equal_nan=True):
                info["impedance"] = z_out.flat[0]
            else:
                info["impedance"] = z_out

        return x_out, y_out, info

    def get_s_params(self, run_id_filter=None, i_rows=None, i_cols=None,
                     port_mode=None, freqs_to_get=None):
        return self._get_syz_params(r"1D Results\S-Parameters", run_id_filter,
                                    i_rows, i_cols, port_mode, freqs_to_get)

    def get_z_params(self, run_id_filter=None, i_rows=None, i_cols=None,
                     port_mode=None, freqs_to_get=None):
        return self._get_syz_params(r"1D Results\Z Matrix", run_id_filter,
                                    i_rows, i_cols, port_mode, freqs_to_get)

    # =========================================================================
    # Parameter Mutation
    # =========================================================================

    def store_parameter(self, param_name: str, param_value,
                        update_structure: bool = True):
        self._check_project_is_open(True)
        if not param_name:
            raise ValueError("param_name cannot be None/empty")
        value = self._unwrap_scalar_cell(param_value)
        if isinstance(value, str):
            self._proj.StoreParameter(param_name, value)
        elif isinstance(value, (int, float, np.number)):
            self._proj.StoreDoubleParameter(param_name, float(value))
        else:
            raise TypeError("param_value must be a string or numeric scalar")
        if update_structure:
            res = self._proj.RebuildOnParametricChange(False, False)
            if not res:
                raise RuntimeError(
                    "Error during structure update on parametric change.")
        return self

    def change_parameter(self, param_name: str, param_value,
                         update_structure: bool = True):
        self.parameter_exists(param_name, True)
        return self.store_parameter(param_name, param_value, update_structure)

    def change_parameters(self, param_names, param_values,
                          update_structure: bool = True):
        self._check_project_is_open(True)
        if isinstance(param_names, str):
            param_names = [param_names]
        if np.isscalar(param_values) or isinstance(param_values, str):
            param_values = [param_values]
        names = list(param_names)
        values = list(param_values)
        if len(names) != len(values):
            raise ValueError("param_names and param_values must have same length")
        for i, (name, value) in enumerate(zip(names, values)):
            self.change_parameter(name, value,
                                  update_structure and i == len(names) - 1)
        return self

    # =========================================================================
    # Private Helpers
    # =========================================================================

    def _get_syz_params(self, tree_path: str, run_id_filter=None,
                        i_rows=None, i_cols=None, port_mode=None,
                        freqs_to_get=None):
        self._check_project_is_open(True)
        children = self.get_result_tree_item_children(tree_path)
        elems = []
        for child in children:
            parsed = self._parse_matrix_tree_item(child)
            if parsed is None:
                continue
            row, col, row_mode, col_mode = parsed
            if i_rows is not None and row not in self._as_int_set(i_rows):
                continue
            if i_cols is not None and col not in self._as_int_set(i_cols):
                continue
            if port_mode is not None and (row_mode not in (None, port_mode)
                                          or col_mode not in (None, port_mode)):
                continue
            elems.append((child, row, col))
        if not elems:
            raise ValueError(f'No matrix elements found under "{tree_path}".')

        max_row = max(row for _, row, _ in elems)
        max_col = max(col for _, _, col in elems)
        if i_rows is not None:
            rows = sorted(self._as_int_set(i_rows))
        else:
            rows = list(range(1, max_row + 1))
        if i_cols is not None:
            cols = sorted(self._as_int_set(i_cols))
        else:
            cols = list(range(1, max_col + 1))
        row_pos = {row: i for i, row in enumerate(rows)}
        col_pos = {col: i for i, col in enumerate(cols)}

        data = None
        freqs = None
        infos = {}
        for child, row, col in elems:
            if row not in row_pos or col not in col_pos:
                continue
            x, y, info = self.get_1d_result_from_tree_item(
                child, run_id_filter, freqs_to_get)
            if data is None:
                freqs = x
                data = np.full((len(rows), len(cols), len(x), y.shape[1]),
                               np.nan + 0j, dtype=complex)
            elif len(x) != len(freqs) or not np.allclose(x, freqs):
                raise RuntimeError(
                    f'Frequency axis for "{child}" does not match previous '
                    "matrix elements.")
            data[row_pos[row], col_pos[col], :, :] = y
            infos[(row, col)] = info

        info = {
            "tree_path": tree_path,
            "rows": rows,
            "cols": cols,
            "elements": infos,
        }
        return data, freqs, info

    @staticmethod
    def _parse_matrix_tree_item(tree_item: str):
        name = tree_item.split("\\")[-1]
        m = re.search(r"\D+(\d+)\((\d+)\)\D+(\d+)\((\d+)\)", name)
        if m:
            row, row_mode, col, col_mode = map(int, m.groups())
            return row, col, row_mode, col_mode
        m = re.search(r"\D+(\d+)\D+(\d+)", name)
        if m:
            row, col = map(int, m.groups())
            return row, col, None, None
        return None

    @staticmethod
    def _as_int_set(values) -> set[int]:
        if np.isscalar(values):
            return {int(values)}
        return {int(v) for v in values}

    @staticmethod
    def _unwrap_scalar_cell(value):
        if isinstance(value, (list, tuple)):
            if len(value) != 1:
                raise ValueError("list/tuple parameter values must be scalar")
            return value[0]
        return value

    @staticmethod
    def _run_id_string(run_id) -> str:
        if isinstance(run_id, str):
            s = run_id[1:] if run_id[:1] in ("~", "-") else run_id
            if s.startswith("3D:RunID:"):
                return s
            return f"3D:RunID:{int(s)}"
        if isinstance(run_id, float) and np.isnan(run_id):
            return "3D:RunID:0"
        return f"3D:RunID:{abs(int(run_id))}"

    @classmethod
    def _apply_run_id_filter(cls, run_ids, run_id_filter) -> list[str]:
        ids = [str(r) for r in run_ids if str(r)]
        if run_id_filter is None:
            return ids
        if isinstance(run_id_filter, (list, tuple, np.ndarray)) and len(run_id_filter) == 0:
            return ids
        if isinstance(run_id_filter, str):
            filters = [run_id_filter]
        elif np.isscalar(run_id_filter):
            filters = [run_id_filter]
        else:
            filters = list(run_id_filter)

        include, exclude = [], []
        for flt in filters:
            is_exclude = False
            if isinstance(flt, str):
                is_exclude = flt[:1] in ("~", "-")
            else:
                try:
                    is_exclude = float(flt) < 0 or np.isnan(float(flt))
                except Exception:
                    is_exclude = False
            target = cls._run_id_string(flt)
            (exclude if is_exclude else include).append(target)

        if include:
            ids = [r for r in ids if r in include]
        if exclude:
            ids = [r for r in ids if r not in exclude]
        return ids

    def _find_queried_x(self, x: np.ndarray, query_x):
        if query_x is None or (hasattr(query_x, "__len__") and len(query_x) == 0):
            return x.copy(), np.arange(len(x), dtype=float)
        q = np.atleast_1d(np.array(query_x, dtype=float)).flatten()
        x_out = np.empty(len(q), dtype=float)
        ix = np.empty(len(q), dtype=float)
        for i, qx in enumerate(q):
            tol = 1e-6 * max(abs(qx), 1.0)
            exact = np.where(np.abs(x - qx) < tol)[0]
            if exact.size:
                ix[i] = exact[0]
                x_out[i] = x[int(ix[i])]
                continue
            if qx < np.min(x) or qx > np.max(x):
                raise ValueError(
                    f"Query X={qx:g} is out of bounds [{np.min(x):g}, "
                    f"{np.max(x):g}].")
            if self._nearest_or_interpolated == "Exact":
                raise ValueError(f"X={qx:g} was not found in simulation data.")
            if self._nearest_or_interpolated == "Nearest":
                ix[i] = int(np.argmin(np.abs(x - qx)))
                x_out[i] = x[int(ix[i])]
            else:
                ix[i] = np.nan
                x_out[i] = qx
        return x_out, ix

    @staticmethod
    def _get_indexed_y(x: np.ndarray, y: np.ndarray, ix: np.ndarray,
                       query_x) -> np.ndarray:
        out = np.empty(len(ix), dtype=y.dtype)
        direct = ~np.isnan(ix)
        if np.any(direct):
            out[direct] = y[ix[direct].astype(int)]
        if np.any(~direct):
            q = np.atleast_1d(np.array(query_x, dtype=float)).flatten()
            interp_x = q[~direct]
            if np.iscomplexobj(y):
                out[~direct] = (np.interp(interp_x, x, y.real)
                                + 1j * np.interp(interp_x, x, y.imag))
            else:
                out[~direct] = np.interp(interp_x, x, y)
        return out

    @staticmethod
    def _to_1d_array(obj, dtype=None) -> np.ndarray:
        if obj is None:
            return np.array([], dtype=dtype or float)
        return np.array(list(obj), dtype=dtype).flatten()

    @classmethod
    def _read_result_y(cls, obj: _Result, res_type: str) -> np.ndarray:
        if res_type == "1DC":
            re_y = cls._to_1d_array(obj.GetArray("yre"), dtype=float)
            im_y = cls._to_1d_array(obj.GetArray("yim"), dtype=float)
            return re_y + 1j * im_y
        return cls._to_1d_array(obj.GetArray("y"), dtype=float)

    @staticmethod
    def _safe_result_attr(obj: _Result, method_name: str):
        try:
            return getattr(obj, method_name)()
        except Exception:
            return ""

    @staticmethod
    def _check_x_is_same(expected: np.ndarray, actual: np.ndarray,
                         run_id: str, query_x):
        if query_x is not None and not (hasattr(query_x, "__len__")
                                        and len(query_x) == 0):
            return
        if expected.shape != actual.shape or not np.allclose(expected, actual):
            raise RuntimeError(
                f'X axis for RunID "{run_id}" is not the same as previous runs.')

    def _check_project_is_open(self, throw_error: bool = False) -> bool:
        if self._proj is None:
            if throw_error:
                raise RuntimeError("No project is open.")
            return False
        try:
            # 使用高频调用安全的 GetProjectPath 作为心跳检测
            p = self._proj.GetProjectPath("Project")
            if not isinstance(p, str):
                raise Exception("heartbeat failed")
        except Exception:
            self._proj = _get_proj(self._app)
            if self._proj is None:
                if throw_error:
                    raise RuntimeError("Project connection lost.")
                return False
        return True

    def _print(self, msg: str):
        if not self._silent_mode:
            print(msg, end="", flush=True)

    def _run_solver(self, n_tries: int):
        for _ in range(n_tries):
            if self._solver.Start():
                return
            time.sleep(2)
        raise RuntimeError(f"Solver failed after {n_tries} tries.")

    @staticmethod
    def _ensure_list(obj) -> list:
        if obj is None:
            return []
        if isinstance(obj, (list, tuple)):
            return list(obj)
        return [obj]

    @staticmethod
    def _get_array(obj, name: str) -> np.ndarray:
        arr = obj.GetArray(name)
        return (np.array([], dtype=float) if arr is None
                else np.array(list(arr), dtype=float).flatten())

    def about(self):
        print("*" * 60)
        print("* PyCST Interface — Pure Stateful No-Cache Invoke")
        print(f"* Version : {self.VERSION}")
        print(f"* Tested  : CST Studio Suite {self.CST_VERSION_TESTED}")
        print("*" * 60)
