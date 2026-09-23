# excel_mcp/document_tools.py
"""Document management tools for Excel — MCP 工具（显式 schema 版）

包含 bundle 生命周期、探查、合并、拆分、导出等文件级操作。

使用方式：
    tools = DocumentTools(registry)
    for tool_def in tools.get_tool_definitions():
        server.register(tool_def, lambda args, n=tool_def["name"]: tools.execute(n, args))
"""
from __future__ import annotations

import json
from typing import Any, Dict, List, Optional

import pandas as pd

from excel_mcp_server.bundle_registry import ExcelBundleRegistry
from excel_mcp_server.excel_bundle import DEFAULT_EXTENSIONS


def _json(data: Any) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2, default=str)


def _fmt_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.1f}{unit}"
        n /= 1024
    return f"{n:.1f}TB"


class DocumentTools:
    """Excel 文件级操作 MCP 工具集。"""

    def __init__(self, registry: ExcelBundleRegistry) -> None:
        self.registry = registry





    @staticmethod
    def _normalize_paths(paths) -> List[str]:
        """把 Agent 传来的各种形式的 paths 统一成 list[str]。

        能处理的输入：
            - list / tuple                → 直接规整
            - 'C:/path'                   → ['C:/path']
            - '["C:/a", "C:/b"]'          → ['C:/a', 'C:/b']   （标准 JSON）
            - '["C:/a", "C:/b"'           → ['C:/a', 'C:/b']   （缺右括号）
            - '["C:/a"'                   → ['C:/a']           （缺右括号，单元素）
            - '[C:/a, C:/b]'              → ['C:/a', 'C:/b']   （无引号）
        """
        import json
        import re

        if paths is None:
            return []

        # 1. 已经是 list / tuple
        if isinstance(paths, (list, tuple)):
            return [str(p).strip() for p in paths if str(p).strip()]

        # 2. 非字符串，无法处理
        if not isinstance(paths, str):
            return []

        s = paths.strip()
        if not s:
            return []

        # 3. 看起来像数组 → 尝试解析
        if s.startswith("[") or s.startswith("("):
            # 3.1 先试标准 JSON
            try:
                parsed = json.loads(s)
                if isinstance(parsed, list):
                    return [str(p).strip() for p in parsed if str(p).strip()]
                if isinstance(parsed, str):
                    return [parsed.strip()] if parsed.strip() else []
            except json.JSONDecodeError:
                pass

            # 3.2 JSON 失败（缺 ] 等）→ 手动修复
            fixed = s
            if not fixed.endswith("]") and not fixed.endswith(")"):
                fixed += "]"          # 补上缺失的右括号
            try:
                parsed = json.loads(fixed)
                if isinstance(parsed, list):
                    return [str(p).strip() for p in parsed if str(p).strip()]
            except json.JSONDecodeError:
                pass

            # 3.3 还是不行 → 暴力剥离方括号和引号，按逗号拆
            inner = s.strip("[]()").strip()
            parts = re.split(r'\s*,\s*', inner)
            cleaned = [p.strip().strip('"').strip("'").strip() for p in parts]
            return [p for p in cleaned if p]

        # 4. 普通字符串 → 单个路径
        return [s]



    # ======================================================
    #                    工具定义（schema）
    # ======================================================
    def get_tool_definitions(self) -> List[Dict[str, Any]]:
        return [
            {
                "name": "excel_load_bundle",
                "description": (
                    "加载一个或多个 Excel 文件/目录到内存 bundle（惰性加载）。"
                    "支持指定扩展名（.xlsx / .xls）、文件名通配、是否递归。"
                    "加载后返回 bundle_id，后续所有操作都用这个 bundle_id 引用。"
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "paths": {
                            "oneOf": [
                                {"type": "string"},
                                {"type": "array", "items": {"type": "string"}},
                            ],
                            "description": (
                                "文件或目录路径。"
                                "可以传单个字符串（如 'C:/data'），"
                                "也可以传数组（如 ['C:/a', 'C:/b']）。"
                                "推荐传字符串以避免解析问题。"
                            ),
                        },
                        "bundle_id": {
                            "type": "string",
                            "description": "自定义 bundle 名（可选，不给则自动生成）",
                        },
                        "extensions": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": (
                                f"目录扫描时的扩展名过滤。不传则默认收 {list(DEFAULT_EXTENSIONS)}。"
                                "例如：['.xlsx'] 只收 xlsx；['.xlsx', '.xls'] 两种都要。"
                            ),
                        },
                        "pattern": {
                            "type": "string",
                            "description": "额外的文件名通配，如 '2024*.xlsx'。可选。",
                        },
                        "recursive": {
                            "type": "boolean",
                            "description": "目录是否递归扫描子目录（默认 false）",
                            "default": False,
                        },
                        "cache_limit": {
                            "type": "integer",
                            "description": "最多缓存多少个 DataFrame（LRU 淘汰，默认 100）",
                            "default": 100,
                        },
                        "make_default": {
                            "type": "boolean",
                            "description": "是否设为默认 bundle（默认 true）",
                            "default": True,
                        },
                    },
                    "required": ["paths"],
                },
            },
            {
                "name": "excel_close_bundle",
                "description": "释放一个 bundle 占用的内存。",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "bundle_id": {"type": "string", "description": "要关闭的 bundle_id"},
                    },
                    "required": ["bundle_id"],
                },
            },
            {
                "name": "excel_close_all_bundles",
                "description": "释放所有 bundle。",
                "inputSchema": {"type": "object", "properties": {}},
            },
            {
                "name": "excel_list_bundles",
                "description": "列出当前进程内所有已加载的 bundle。",
                "inputSchema": {"type": "object", "properties": {}},
            },
            {
                "name": "excel_list_files",
                "description": "列出某个 bundle 内已注册的所有文件。",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "bundle_id": {
                            "type": "string",
                            "description": "不传则用默认 bundle",
                        },
                    },
                },
            },
            {
                "name": "excel_list_sheets",
                "description": (
                    "列出 sheet。给定 file_key 只列该文件的；不给定则列所有文件的。"
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "bundle_id": {"type": "string"},
                        "file_key": {
                            "type": "string",
                            "description": "可选。不传则列出所有文件的 sheet",
                        },
                    },
                },
            },
            {
                "name": "excel_get_bundle_info",
                "description": (
                    "查看 bundle 整体信息。"
                    "with_shape=false（默认）只返回文件与 sheet 名，不读数据；"
                    "with_shape=true 附带每个 sheet 的行列数（会触发读盘）。"
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "bundle_id": {"type": "string"},
                        "with_shape": {"type": "boolean", "default": False},
                    },
                },
            },
            {
                "name": "excel_get_columns",
                "description": "获取指定 sheet 的列名与行数（走缓存）。",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "file_key": {"type": "string"},
                        "sheet_key": {"type": "string"},
                        "bundle_id": {"type": "string"},
                    },
                    "required": ["file_key", "sheet_key"],
                },
            },
            {
                "name": "excel_preview_sheet",
                "description": "预览指定 sheet 的前 n 行。",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "file_key": {"type": "string"},
                        "sheet_key": {"type": "string"},
                        "n": {"type": "integer", "default": 5},
                        "bundle_id": {"type": "string"},
                    },
                    "required": ["file_key", "sheet_key"],
                },
            },
            {
                "name": "excel_merge_files",
                "description": (
                    "合并 bundle 内多个文件的 df 为一个新 df，写回 bundle。"
                    "结果默认写到 '__merged__' 这个虚拟 file_key。"
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "file_keys": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "要合并的 file_key 列表；不传=全部",
                        },
                        "sheet_key": {
                            "type": "string",
                            "description": "每个文件取哪个 sheet；不传=各自第一个",
                        },
                        "target_file_key": {"type": "string", "default": "__merged__"},
                        "remove_sources": {
                            "type": "boolean",
                            "description": (
                                "合并后是否移除源文件的 df（默认 false）。"
                                "若为 true，合并后 bundle 里只剩下 target_file_key 一个结果。"
                            ),
                            "default": False,
                        },
                        "target_sheet_key": {"type": "string", "default": "Sheet1"},
                        "add_source_column": {"type": "boolean", "default": True},
                        "source_column_name": {"type": "string", "default": "源文件"},
                        "bundle_id": {"type": "string"},
                    },
                },
            },
            {
                "name": "excel_merge_from_paths",
                "description": (
                    "【一步到位】从路径加载所有 Excel，合并为一个 sheet，"
                    "导出到指定 xlsx。不依赖跨调用的 bundle 状态，"
                    "推荐在 stdio 传输模式下使用。"
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "paths": {
                            "oneOf": [
                                {"type": "string"},
                                {"type": "array",
                                 "items": {"type": "string"}},
                            ],
                            "description": (
                                "文件或目录路径。推荐传字符串。"
                            ),
                        },
                        "out_path": {
                            "type": "string",
                            "description": "输出 xlsx 的完整路径",
                        },
                        "sheet_key": {
                            "type": "string",
                            "description": (
                                "每个文件取哪个 sheet；不传=各自第一个"
                            ),
                        },
                        "add_source_column": {
                            "type": "boolean", "default": True,
                        },
                        "source_column_name": {
                            "type": "string", "default": "源文件",
                        },
                        "extensions": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "目录扫描扩展名过滤，如 ['.xlsx']",
                        },
                        "pattern": {
                            "type": "string",
                            "description": "文件名通配，如 '2024*.xlsx'",
                        },
                        "recursive": {
                            "type": "boolean", "default": False,
                        },
                    },
                    "required": ["paths", "out_path"],
                },
            },
            {
                "name": "excel_split_by_column",
                "description": (
                    "按指定列的唯一值把一个 sheet 拆成多个 df 写回 bundle。"
                    "结果写到 '<target_prefix 或 file_key>__<value>'，"
                    "sheet_key 保持不变。"
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "file_key": {"type": "string"},
                        "sheet_key": {"type": "string"},
                        "column": {"type": "string"},
                        "target_prefix": {"type": "string"},
                        "bundle_id": {"type": "string"},
                    },
                    "required": ["file_key", "sheet_key", "column"],
                },
            },
            {
                "name": "excel_export_bundle",
                "description": (
                    "把 bundle 内所有 df 导出到一个多 sheet 的 Excel。"
                    "sheet 命名自动处理重名/超长。"
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "out_path": {"type": "string"},
                        "file_keys": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "只导出指定文件；不传=全部",
                        },
                        "sheet_key": {
                            "type": "string",
                            "description": "只导出该 sheet 名；不传=全部",
                        },
                        "bundle_id": {"type": "string"},
                    },
                    "required": ["out_path"],
                },
            },
            {
                "name": "excel_export_split",
                "description": "每个 (file_key, sheet_key) 单独导出一个 .xlsx 文件。",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "out_dir": {"type": "string"},
                        "file_keys": {
                            "type": "array",
                            "items": {"type": "string"},
                        },
                        "name_sep": {"type": "string", "default": "__"},
                        "bundle_id": {"type": "string"},
                    },
                    "required": ["out_dir"],
                },
            },
        ]

    # ======================================================
    #                    执行分发
    # ======================================================
    def execute(self, tool_name: str, arguments: Dict[str, Any]) -> str:
        try:
            match tool_name:
                case "excel_load_bundle":
                    return self._load_bundle(**arguments)
                case "excel_close_bundle":
                    return self._close_bundle(**arguments)
                case "excel_close_all_bundles":
                    return self._close_all_bundles()
                case "excel_list_bundles":
                    return self._list_bundles()
                case "excel_list_files":
                    return self._list_files(**arguments)
                case "excel_list_sheets":
                    return self._list_sheets(**arguments)
                case "excel_get_bundle_info":
                    return self._get_bundle_info(**arguments)
                case "excel_get_columns":
                    return self._get_columns(**arguments)
                case "excel_preview_sheet":
                    return self._preview_sheet(**arguments)
                case "excel_merge_files":
                    return self._merge_files(**arguments)
                case "excel_split_by_column":
                    return self._split_by_column(**arguments)
                case "excel_export_bundle":
                    return self._export_bundle(**arguments)
                case "excel_export_split":
                    return self._export_split(**arguments)
                case "excel_merge_from_paths":
                    return self._merge_from_paths(**arguments)
                case _:
                    return _json({"success": False,
                                  "error": f"未知工具: {tool_name}"})
        except TypeError as e:
            # 参数不匹配时的友好提示
            return _json({"success": False,
                          "error": f"参数错误: {e}",
                          "hint": "请检查 inputSchema 中定义的参数名和类型"})
        except Exception as e:
            return _json({"success": False, "error": str(e)})

    # ======================================================
    #                    生命周期
    # ======================================================
    def _load_bundle(
        self,
        paths: Any,
        bundle_id: Optional[str] = None,
        extensions: Optional[List[str]] = None,
        pattern: Optional[str] = None,
        recursive: bool = False,
        cache_limit: int = 100,
        make_default: bool = True,
    ) -> str:

        # 统一路径格式
        normalized = self._normalize_paths(paths)
        if not normalized:
            return _json({
                "success": False,
                "error": f"paths 为空或格式无法识别: {paths!r}",
                "hint": "可传字符串 'C:/data' 或数组 ['C:/a', 'C:/b']",
            })


        bid = self.registry.create(name=bundle_id, make_default=make_default,
                                   cache_limit=cache_limit)
        bundle = self.registry.get(bid)
        result = bundle.load(normalized, extensions=extensions,
                             pattern=pattern, recursive=recursive)
        return _json({
            "success": True,
            "bundle_id": bid,
            "file_count": len(bundle.files()),
            "loaded_files": result["loaded"],
            "errors": result["errors"],
        })

    def _close_bundle(self, bundle_id: str) -> str:
        ok = self.registry.close(bundle_id)
        return _json({
            "success": ok,
            "bundle_id": bundle_id,
            "remaining": [b["bundle_id"] for b in self.registry.list_all()],
        })

    def _close_all_bundles(self) -> str:
        n = self.registry.close_all()
        return _json({"success": True, "closed_count": n})

    # ======================================================
    #                    探查
    # ======================================================
    def _list_bundles(self) -> str:
        return _json({
            "success": True,
            "bundles": self.registry.list_all(),
            "default": self.registry.default_id(),
        })

    def _list_files(self, bundle_id: Optional[str] = None) -> str:
        bundle = self.registry.get(bundle_id)
        files = []
        for fkey in bundle.files():
            info = bundle.file_info(fkey)
            if info is None:
                continue
            info["size_human"] = _fmt_size(info["size"])
            files.append(info)
        return _json({"success": True, "count": len(files), "files": files})

    def _list_sheets(
        self,
        bundle_id: Optional[str] = None,
        file_key: Optional[str] = None,
    ) -> str:
        bundle = self.registry.get(bundle_id)
        if file_key:
            if not bundle.has(file_key):
                return _json({"success": False,
                              "error": f"file_key 不存在: {file_key}",
                              "available": bundle.files()})
            return _json({"success": True, "file_key": file_key,
                          "sheets": bundle.sheets(file_key)})
        return _json({
            "success": True,
            "files": [{"file_key": fk, "sheets": bundle.sheets(fk)}
                      for fk in bundle.files()],
        })

    def _get_bundle_info(
        self,
        bundle_id: Optional[str] = None,
        with_shape: bool = False,
    ) -> str:
        bundle = self.registry.get(bundle_id)
        return _json({"success": True, **bundle.summary(with_shape=with_shape)})

    def _get_columns(
        self,
        file_key: str,
        sheet_key: str,
        bundle_id: Optional[str] = None,
    ) -> str:
        bundle = self.registry.get(bundle_id)
        df = bundle.get(file_key, sheet_key)
        return _json({"success": True, "file_key": file_key,
                      "sheet_key": sheet_key,
                      "columns": list(df.columns),
                      "row_count": int(len(df))})

    def _preview_sheet(
        self,
        file_key: str,
        sheet_key: str,
        n: int = 5,
        bundle_id: Optional[str] = None,
    ) -> str:
        bundle = self.registry.get(bundle_id)
        df = bundle.get(file_key, sheet_key)
        return _json({
            "success": True,
            "file_key": file_key,
            "sheet_key": sheet_key,
            "columns": list(df.columns),
            "rows_total": int(len(df)),
            "preview": df.head(n).to_dict(orient="records"),
        })

    # ======================================================
    #                    合并 / 拆分
    # ======================================================
    def _merge_files(
        self,
        file_keys: Optional[List[str]] = None,
        sheet_key: Optional[str] = None,
        target_file_key: str = "__merged__",
        remove_sources: bool = False,
        target_sheet_key: str = "Sheet1",
        add_source_column: bool = True,
        source_column_name: str = "源文件",
        bundle_id: Optional[str] = None,
    ) -> str:
        bundle = self.registry.get(bundle_id)
        keys = bundle.files() if file_keys is None else list(file_keys)
        if not keys:
            return _json({"success": False, "error": "没有要合并的文件"})

        dfs: List[pd.DataFrame] = []
        errors: List[Dict] = []
        skipped_empty = 0

        for fkey in keys:
            if not bundle.has(fkey):
                errors.append({"file_key": fkey, "error": "文件不存在"})
                continue
            sk = sheet_key or (bundle.sheets(fkey)[0] if bundle.sheets(fkey) else None)
            if not sk:
                errors.append({"file_key": fkey, "error": "文件无 sheet"})
                continue
            try:
                df = bundle.get(fkey, sk)
                if df.empty:
                    skipped_empty += 1
                    continue
                if add_source_column:
                    col = source_column_name
                    if col in df.columns:
                        col = f"_src_{col}"
                    df = df.assign(**{col: fkey})
                dfs.append(df)
            except Exception as e:
                errors.append({"file_key": fkey, "sheet_key": sk, "error": str(e)})

        if not dfs:
            return _json({"success": False, "error": "没有可合并的数据",
                          "errors": errors})

        merged = pd.concat(dfs, ignore_index=True, copy=False)
        bundle.set(target_file_key, target_sheet_key, merged)

        removed = []
        if remove_sources:
            for fkey in keys:
                if fkey != target_file_key and bundle.remove(fkey):
                    removed.append(fkey)

        return _json({
            "success": True,
            "target_file_key": target_file_key,
            "target_sheet_key": target_sheet_key,
            "source_count": len(dfs),
            "removed_sources": removed,
            "skipped_empty": skipped_empty,
            "total_rows": int(len(merged)),
            "columns": list(merged.columns),
            "errors": errors,
        })

    def _merge_from_paths(
        self,
        paths: Any,
        out_path: str,
        sheet_key: Optional[str] = None,
        add_source_column: bool = True,
        source_column_name: str = "源文件",
        extensions: Optional[List[str]] = None,
        pattern: Optional[str] = None,
        recursive: bool = False,
    ) -> str:
        """一步到位：加载路径 → 合并所有 df → 导出到一个 xlsx。

        不依赖跨调用的 bundle 状态，适合 stdio 传输模式。
        """
        normalized = self._normalize_paths(paths)
        if not normalized:
            return _json({
                "success": False,
                "error": f"paths 为空或格式无法识别: {paths!r}",
            })

        # 用临时 bundle，函数结束时销毁
        temp_bid = self.registry.create(make_default=False)
        bundle = self.registry.get(temp_bid)

        try:
            # 1. 加载
            load_result = bundle.load(
                normalized,
                extensions=extensions,
                pattern=pattern,
                recursive=recursive,
            )
            if not bundle.files():
                return _json({
                    "success": False,
                    "error": "未加载到任何文件",
                    "loaded": load_result,
                })

            # 2. 合并
            dfs: List[pd.DataFrame] = []
            errors: List[Dict] = []
            skipped_empty = 0

            for fkey in bundle.files():
                sk = sheet_key or (bundle.sheets(fkey)[0]
                                   if bundle.sheets(fkey) else None)
                if not sk:
                    errors.append({"file_key": fkey, "error": "文件无 sheet"})
                    continue
                try:
                    df = bundle.get(fkey, sk)
                    if df.empty:
                        skipped_empty += 1
                        continue
                    if add_source_column:
                        col = source_column_name
                        if col in df.columns:
                            col = f"_src_{col}"
                        df = df.assign(**{col: fkey})
                    dfs.append(df)
                except Exception as e:
                    errors.append({"file_key": fkey, "sheet_key": sk,
                                   "error": str(e)})

            if not dfs:
                return _json({
                    "success": False,
                    "error": "没有可合并的数据",
                    "errors": errors,
                })

            merged = pd.concat(dfs, ignore_index=True, copy=False)
            bundle.set("__merged__", "Sheet1", merged)

            # 3. 导出（只导合并结果）
            export_result = bundle.export(out_path, file_keys=["__merged__"])

            return _json({
                "success": True,
                "out_path": out_path,
                "loaded_files": load_result["loaded"],
                "source_count": len(dfs),
                "skipped_empty": skipped_empty,
                "total_rows": int(len(merged)),
                "columns": list(merged.columns),
                "errors": errors,
                "export": export_result,
            })
        finally:
            # 无论成功失败都清理临时 bundle
            self.registry.close(temp_bid)




    def _split_by_column(
        self,
        file_key: str,
        sheet_key: str,
        column: str,
        target_prefix: Optional[str] = None,
        bundle_id: Optional[str] = None,
    ) -> str:
        bundle = self.registry.get(bundle_id)
        df = bundle.get(file_key, sheet_key)
        if column not in df.columns:
            return _json({"success": False,
                          "error": f"列不存在: {column}",
                          "available": list(df.columns)})

        prefix = target_prefix or file_key
        created: List[str] = []
        for value, group in df.groupby(column, sort=False):
            sub_key = f"{prefix}__{value}"
            bundle.set(sub_key, sheet_key, group.reset_index(drop=True))
            created.append(sub_key)

        return _json({
            "success": True,
            "source": f"{file_key}.{sheet_key}",
            "column": column,
            "created_count": len(created),
            "created_keys": created,
        })

    # ======================================================
    #                    导出
    # ======================================================
    def _export_bundle(
        self,
        out_path: str,
        file_keys: Optional[List[str]] = None,
        sheet_key: Optional[str] = None,
        bundle_id: Optional[str] = None,
    ) -> str:
        bundle = self.registry.get(bundle_id)
        return _json(bundle.export(out_path, file_keys=file_keys,
                                   sheet_key=sheet_key))

    def _export_split(
        self,
        out_dir: str,
        file_keys: Optional[List[str]] = None,
        name_sep: str = "__",
        bundle_id: Optional[str] = None,
    ) -> str:
        bundle = self.registry.get(bundle_id)
        return _json(bundle.export_split(out_dir, file_keys=file_keys,
                                         name_sep=name_sep))