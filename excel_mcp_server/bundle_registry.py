"""ExcelBundleRegistry —— 进程级 bundle 注册表

- 单例：跨 MCP 工具调用共享
- 支持多个 bundle 并存（多任务/多用户）
- 支持 default bundle（单任务场景可省略 bundle_id）
"""
from __future__ import annotations

import threading
import uuid
from typing import Dict, List, Optional

from excel_bundle import ExcelBundle


class ExcelBundleRegistry:
    _instance: Optional["ExcelBundleRegistry"] = None
    _lock = threading.RLock()

    def __new__(cls) -> "ExcelBundleRegistry":
        with cls._lock:
            if cls._instance is None:
                inst = super().__new__(cls)
                inst._bundles: Dict[str, ExcelBundle] = {}
                inst._default: Optional[str] = None
                cls._instance = inst
            return cls._instance

    # ---------------- 创建 ----------------
    def create(
        self,
        name: Optional[str] = None,
        make_default: bool = True,
        cache_limit: int = 20,
    ) -> str:
        with self._lock:
            bid = name or f"b_{uuid.uuid4().hex[:8]}"
            if bid in self._bundles:
                raise ValueError(f"bundle 已存在: {bid}")
            self._bundles[bid] = ExcelBundle(cache_limit=cache_limit)
            if make_default or self._default is None:
                self._default = bid
            return bid

    # ---------------- 获取 ----------------
    def get(self, bundle_id: Optional[str] = None) -> ExcelBundle:
        with self._lock:
            bid = bundle_id or self._default
            if bid is None:
                raise RuntimeError("没有默认 bundle，请先 load_bundle")
            if bid not in self._bundles:
                raise KeyError(f"bundle 不存在: {bid}，可用: {list(self._bundles)}")
            return self._bundles[bid]

    def default_id(self) -> Optional[str]:
        return self._default

    def exists(self, bundle_id: str) -> bool:
        return bundle_id in self._bundles

    # ---------------- 枚举 ----------------
    def list_all(self) -> List[Dict]:
        with self._lock:
            return [
                {
                    "bundle_id": bid,
                    "is_default": bid == self._default,
                    "file_count": len(b.files()),
                    "cache_count": len(b._cache),
                }
                for bid, b in self._bundles.items()
            ]

    # ---------------- 关闭 ----------------
    def close(self, bundle_id: str) -> bool:
        with self._lock:
            if bundle_id not in self._bundles:
                return False
            del self._bundles[bundle_id]
            if self._default == bundle_id:
                self._default = next(iter(self._bundles), None)
            return True

    def close_all(self) -> int:
        with self._lock:
            n = len(self._bundles)
            self._bundles.clear()
            self._default = None
            return n


# 全局单例
REGISTRY = ExcelBundleRegistry()