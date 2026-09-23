"""ExcelBundle —— 多 Excel / 多 Sheet 的内存容器

设计要点：
- 惰性加载：add() 只读 sheet 索引，不读数据
- LRU 缓存：常用 sheet 常驻内存
- 目录扫描灵活：可选递归、可选扩展名、可选通配 pattern
"""
from __future__ import annotations

import os
import re
import fnmatch
import hashlib
import logging
from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Tuple

import pandas as pd

logger = logging.getLogger(__name__)

# ---------- 常量 ----------
DEFAULT_EXTENSIONS = (".xlsx", ".xls")     # 用户不传 extensions 时的默认
_LEGACY_XLS_EXTENSIONS = (".xls",)         # 走 xlrd 的旧格式
_MODERN_XLSX_EXTENSIONS = (".xlsx", ".xlsm")  # 走 openpyxl 的现代格式
_EXCEL_SHEET_NAME_LIMIT = 31
_INVALID_SHEET_CHARS = re.compile(r'[\[\]:*?/\\]')



# ============================================================
#                    读取引擎选择
# ============================================================
def _pick_engine(path: str) -> Optional[str]:
    """根据扩展名选 pandas 读取引擎；未识别的返回 None 让 pandas 自行判断。"""
    low = path.lower()
    if low.endswith(_LEGACY_XLS_EXTENSIONS):
        return "xlrd"
    if low.endswith(_MODERN_XLSX_EXTENSIONS):
        return "openpyxl"
    return None

# ============================================================
#                    目录扫描
# ============================================================
def scan_excel_files(
    directory: str,
    extensions: Optional[List[str]] = None,
    pattern: Optional[str] = None,
    recursive: bool = False,
) -> List[str]:
    """扫描目录下匹配的 Excel 文件。

    Args:
        directory:  要扫描的目录
        extensions: 扩展名过滤，如 [".xlsx"] 或 [".xlsx", ".xls"]。
                    None = 默认 [".xlsx", ".xls"]。
        pattern:    额外的文件名通配，如 "2024*.xlsx"。
                    None = 仅按 extensions 过滤。
        recursive:  是否递归扫描子目录。

    Returns:
        匹配文件的绝对路径列表（已排序）

    实现说明：
        用 os.scandir 而非 Path.glob —— DirEntry 会缓存 stat 结果，
        在目录条目多时比 glob 快 2~3 倍。
    """
    root = Path(directory)
    if not root.is_dir():
        return []

    exts = tuple(e.lower() for e in (extensions or DEFAULT_EXTENSIONS))

    def _match(name: str) -> bool:
        low = name.lower()
        if not low.endswith(exts):
            return False
        if pattern is not None and not fnmatch.fnmatch(low, pattern.lower()):
            return False
        return True

    results: List[str] = []

    def _scan(dir_path: Path) -> None:
        try:
            with os.scandir(dir_path) as it:
                for entry in it:
                    try:
                        if entry.is_file(follow_symlinks=False):
                            if _match(entry.name):
                                results.append(entry.path)
                        elif recursive and entry.is_dir(follow_symlinks=False):
                            _scan(Path(entry.path))
                    except OSError:
                        continue
        except OSError as e:
            logger.warning("扫描目录失败 %s: %s", dir_path, e)

    _scan(root)
    results.sort()
    return results


# ============================================================
#                    命名辅助
# ============================================================
def sanitize_sheet_name(name: str) -> str:
    s = str(name) if name is not None else ""
    if not s:
        return "_"
    return _INVALID_SHEET_CHARS.sub("_", s) if _INVALID_SHEET_CHARS.search(s) else s


def _truncate_with_hash(name: str, limit: int) -> str:
    h = hashlib.md5(name.encode("utf-8")).hexdigest()[:6]
    suffix = f"_{h}"
    keep = limit - len(suffix)
    return f"{name[:keep]}{suffix}" if keep > 0 else h[:limit]


# ============================================================
#                    ExcelBundle
# ============================================================
@dataclass
class _FileMeta:
    path: str            # 磁盘路径；虚拟 df 时为 ""
    sheets: List[str]
    size: int
    mtime: float


class ExcelBundle:
    """多 Excel / 多 Sheet 容器（惰性加载 + LRU 缓存）"""

    def __init__(self, cache_limit: int = 20) -> None:
        self._meta: "OrderedDict[str, _FileMeta]" = OrderedDict()
        self._cache: "OrderedDict[Tuple[str, str], pd.DataFrame]" = OrderedDict()
        self._cache_limit = max(1, cache_limit)

    # ---------------- 加载 ----------------
    def add(self, file_path: str,
            file_key: Optional[str] = None,
            sheet_name=None) -> str:
        """注册一个 Excel 文件。

        Args:
            file_path: 文件路径
            file_key:  自定义 key（默认用文件名主干）
            sheet_name: None=注册所有 sheet；list=只注册指定 sheet

        Returns:
            实际使用的 file_key
        """
        p = Path(file_path)
        if not p.is_file():
            raise FileNotFoundError(f"文件不存在: {file_path}")
        resolved = str(p.resolve())

        with pd.ExcelFile(resolved) as xl:
            all_sheets = list(xl.sheet_names)

        if sheet_name is None:
            sheets = all_sheets
        elif isinstance(sheet_name, str):
            sheets = [sheet_name] if sheet_name in all_sheets else []
            if not sheets:
                raise ValueError(f"sheet 不存在: {sheet_name}，可用: {all_sheets}")
        else:
            sheets = [s for s in sheet_name if s in all_sheets]

        fkey = file_key or p.stem
        st = p.stat()
        self._meta[fkey] = _FileMeta(
            path=resolved, sheets=sheets, size=st.st_size, mtime=st.st_mtime,
        )
        return fkey

    def add_dir(
        self,
        dir_path: str,
        extensions: Optional[List[str]] = None,
        pattern: Optional[str] = None,
        recursive: bool = False,
    ) -> List[str]:
        """扫描目录并注册所有匹配文件。

        Args:
            extensions: 扩展名过滤，如 [".xlsx"]（只 xlsx）或
                        [".xlsx", ".xls"]（两者都要）。None=默认两者。
            pattern:    文件名通配，如 "2024*.xlsx"。None=不额外过滤。
            recursive:  是否递归子目录。

        Returns:
            注册成功的 file_key 列表
        """
        files = scan_excel_files(dir_path, extensions, pattern, recursive)
        keys: List[str] = []
        for fp in files:
            try:
                keys.append(self.add(fp))
            except Exception as e:
                logger.warning("注册失败 %s: %s", fp, e)
        return keys

    def load(
        self,
        paths: List[str],
        extensions: Optional[List[str]] = None,
        pattern: Optional[str] = None,
        recursive: bool = False,
    ) -> Dict[str, List[str]]:
        """便捷方法：接受文件/目录混合列表，统一加载。

        Returns:
            {"loaded": [...], "errors": [{path, error}, ...]}
        """
        loaded: List[str] = []
        errors: List[Dict] = []
        for p in paths:
            pth = Path(p)
            try:
                if pth.is_dir():
                    loaded.extend(self.add_dir(str(pth), extensions,
                                               pattern, recursive))
                elif pth.is_file():
                    loaded.append(self.add(str(pth)))
                else:
                    errors.append({"path": str(p), "error": "路径不存在"})
            except Exception as e:
                errors.append({"path": str(p), "error": str(e)})
        return {"loaded": loaded, "errors": errors}

    # ---------------- 查询 ----------------
    def files(self) -> List[str]:
        return list(self._meta.keys())

    def sheets(self, file_key: str) -> List[str]:
        meta = self._meta.get(file_key)
        return list(meta.sheets) if meta else []

    def has(self, file_key: str, sheet_key: Optional[str] = None) -> bool:
        if file_key not in self._meta:
            return False
        return True if sheet_key is None else sheet_key in self._meta[file_key].sheets

    def get(self, file_key: str, sheet_key: str) -> pd.DataFrame:
        """取 df（命中缓存直接返回，否则读盘并缓存）。"""
        key = (file_key, sheet_key)
        cached = self._cache.get(key)
        if cached is not None:
            self._cache.move_to_end(key)
            return cached

        meta = self._meta.get(file_key)
        if meta is None:
            raise KeyError(f"file_key 不存在: {file_key}")
        if sheet_key not in meta.sheets:
            raise KeyError(f"sheet 不存在: {sheet_key}，可用: {meta.sheets}")

        engine = _pick_engine(meta.path)
        df = pd.read_excel(meta.path, sheet_name=sheet_key, engine=engine)

        self._cache[key] = df
        self._cache.move_to_end(key)
        self._evict_if_needed()
        return df

    def set(self, file_key: str, sheet_key: str, df: pd.DataFrame) -> None:
        """写入/替换 df（用于合并结果等）。未注册的 file_key 自动创建为虚拟文件。"""
        meta = self._meta.get(file_key)
        if meta is None:
            self._meta[file_key] = _FileMeta(
                path="", sheets=[sheet_key], size=0, mtime=0.0,
            )
        elif sheet_key not in meta.sheets:
            meta.sheets.append(sheet_key)

        key = (file_key, sheet_key)
        self._cache[key] = df
        self._cache.move_to_end(key)
        self._evict_if_needed()

    def _evict_if_needed(self) -> None:
        while len(self._cache) > self._cache_limit:
            self._cache.popitem(last=False)

    # ---------------- 释放 ----------------
    def drop_cache(self, file_key: Optional[str] = None) -> int:
        """丢弃缓存（保留文件注册信息）。返回释放数量。"""
        if file_key is None:
            n = len(self._cache)
            self._cache.clear()
            return n
        keys = [k for k in self._cache if k[0] == file_key]
        for k in keys:
            self._cache.pop(k, None)
        return len(keys)

    def remove(self, file_key: str) -> bool:
        """彻底移除文件（元数据 + 缓存）。"""
        if file_key not in self._meta:
            return False
        self.drop_cache(file_key)
        del self._meta[file_key]
        return True

    # ---------------- 概览 ----------------
    def items(self) -> Iterator[Tuple[str, str, pd.DataFrame]]:
        """惰性遍历 (file_key, sheet_key, df)。"""
        for fkey in list(self._meta.keys()):
            for skey in list(self._meta[fkey].sheets):
                try:
                    yield fkey, skey, self.get(fkey, skey)
                except Exception as e:
                    logger.warning("跳过 %s.%s: %s", fkey, skey, e)

    def summary(self, with_shape: bool = False) -> Dict:
        """结构化概览。

        Args:
            with_shape: True 时附带每个 sheet 的行列数（会触发读盘）
        """
        files_info = []
        for fkey, meta in self._meta.items():
            sheets_info = []
            for skey in meta.sheets:
                item = {"sheet_key": skey}
                if with_shape:
                    try:
                        df = self.get(fkey, skey)
                        item["rows"] = int(len(df))
                        item["columns"] = list(df.columns)
                    except Exception as e:
                        item["error"] = str(e)
                sheets_info.append(item)
            files_info.append({
                "file_key": fkey,
                "path": meta.path,
                "size": meta.size,
                "mtime": meta.mtime,
                "sheets": sheets_info,
            })
        return {
            "file_count": len(files_info),
            "cache_count": len(self._cache),
            "cache_limit": self._cache_limit,
            "files": files_info,
        }

    def file_info(self, file_key: str) -> Optional[Dict]:
        """返回单个文件的元信息（供外部工具/接口使用）。"""
        meta = self._meta.get(file_key)
        if meta is None:
            return None
        return {
            "file_key": file_key,
            "path": meta.path,
            "size": meta.size,
            "mtime": meta.mtime,
            "sheet_count": len(meta.sheets),
            "sheets": list(meta.sheets),
        }


    # ---------------- 导出 ----------------
    def export(
        self,
        out_path: str,
        file_keys: Optional[List[str]] = None,
        sheet_key: Optional[str] = None,
    ) -> Dict:
        """把 bundle 内所有 df 导出到一个多 sheet 的 Excel。

        sheet 命名策略：
            skey → fkey_skey（跨文件同名）→ 截断+哈希（超 31）→ 数字后缀
        所有改名都会 WARNING log 并汇总。
        """
        keys = self.files() if file_keys is None else list(file_keys)

        entries: List[Tuple[str, str, pd.DataFrame]] = []
        skey_freq: Dict[str, int] = {}
        for fkey in keys:
            sheets = ([sheet_key] if sheet_key else self.sheets(fkey))
            for skey in sheets:
                if not self.has(fkey, skey):
                    continue
                try:
                    df = self.get(fkey, skey)
                except Exception as e:
                    logger.warning("读取失败 %s.%s: %s", fkey, skey, e)
                    continue
                if df.empty:
                    continue
                entries.append((fkey, skey, df))
                skey_freq[skey] = skey_freq.get(skey, 0) + 1

        if not entries:
            return {"success": False, "error": "没有可导出的数据"}

        used: set = set()
        counter: Dict[str, int] = {}
        renames: List[Tuple[str, str, str]] = []
        written: List[str] = []

        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            for fkey, skey, df in entries:
                name = sanitize_sheet_name(skey)

                if skey_freq[skey] > 1:
                    fb = sanitize_sheet_name(f"{fkey}_{skey}")
                    if fb != name:
                        renames.append((name, fb, "跨文件同名"))
                    name = fb

                if len(name) > _EXCEL_SHEET_NAME_LIMIT:
                    short = _truncate_with_hash(name, _EXCEL_SHEET_NAME_LIMIT)
                    renames.append((name, short, "超长"))
                    name = short

                if name in used:
                    base = name
                    i = counter.get(base, 0)
                    while True:
                        sfx = f"_{i}"
                        cand = base[: _EXCEL_SHEET_NAME_LIMIT - len(sfx)] + sfx
                        i += 1
                        if cand not in used:
                            break
                    counter[base] = i
                    renames.append((base, cand, "重名"))
                    name = cand
                else:
                    counter.setdefault(name, 1)

                used.add(name)
                df.to_excel(writer, sheet_name=name, index=False)
                written.append(name)

        _log_renames(renames)
        return {
            "success": True,
            "output_path": str(Path(out_path).resolve()),
            "sheet_count": len(written),
            "sheets": written,
        }

    def export_split(
        self,
        out_dir: str,
        file_keys: Optional[List[str]] = None,
        name_sep: str = "__",
    ) -> Dict:
        """每个 (file_key, sheet_key) 单独导出一个 .xlsx。"""
        out = Path(out_dir)
        out.mkdir(parents=True, exist_ok=True)

        keys = self.files() if file_keys is None else list(file_keys)
        written: List[str] = []
        failed: List[Dict] = []
        seen: set = set()

        for fkey in keys:
            for skey in self.sheets(fkey):
                try:
                    df = self.get(fkey, skey)
                except Exception as e:
                    failed.append({"file": f"{fkey}.{skey}", "error": str(e)})
                    continue
                if df.empty:
                    continue

                raw = f"{fkey}{name_sep}{skey}"
                safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", raw).strip().rstrip(". ") or "_"
                safe = safe[:145]

                final = safe
                i = 1
                while final.lower() in seen:
                    sfx = f"_{i}"
                    final = safe[: 150 - len(sfx)] + sfx
                    i += 1
                seen.add(final.lower())

                fp = out / f"{final}.xlsx"
                try:
                    df.to_excel(fp, index=False)
                    written.append(str(fp))
                except Exception as e:
                    failed.append({"file": str(fp), "error": str(e)})

        return {"success": True, "written": written, "failed": failed}


# ============================================================
#                    内部辅助
# ============================================================
def _log_renames(renames: List[Tuple[str, str, str]], preview: int = 10) -> None:
    if not renames:
        return
    logger.warning("导出时对 %d 个 sheet 名做了调整：", len(renames))
    for intent, final, reason in renames[:preview]:
        logger.warning("    %s → %s（%s）", intent, final, reason)
    if len(renames) > preview:
        logger.warning("    ... 其余 %d 条已省略", len(renames) - preview)


