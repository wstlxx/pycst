"""
CSTInterface - Python interface to CST Microwave Studio via COM automation.
Requires: Windows, pywin32, numpy

【首次使用】先生成早绑定缓存（只需运行一次）：
    python -m win32com.client.makepy "D:\\Program Files (x86)\\CST Studio Suite 2024\\Patches\\BASE_2023-09-01\\AMD64\\CST DESIGN ENVIRONMENT_AMD64.exe"
"""

import os
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

    # =========================================================================
    # Private Helpers
    # =========================================================================

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
        return list(obj) if isinstance(obj, (list, tuple)) else [obj]

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
