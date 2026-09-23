"""共享工具：Excel 读写、JSON 序列化"""
import json
from pathlib import Path
from typing import Optional
import pandas as pd


def json_out(data) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2, default=str)


def read_excel(fp: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    p = Path(fp)
    if not p.exists():
        raise FileNotFoundError(f"文件不存在: {fp}")
    engine = "xlrd" if p.suffix.lower() == ".xls" else "openpyxl"
    return pd.read_excel(fp, sheet_name=sheet_name, engine=engine)


def write_excel(df: pd.DataFrame, fp: str, sheet_name: str = "Sheet1") -> None:
    p = Path(fp)
    p.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(fp, sheet_name=sheet_name, index=False, engine="openpyxl")