"""Universal data loading helpers (files/URLs -> DataFrame)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from ._deps import pd, require_deps

_CSV_ALIASES = {".csv", ".tsv", ".txt"}
_EXCEL_ALIASES = {".xlsx", ".xls", ".xlsm"}


def read_csv(path: str | Path, **kwargs: Any) -> Any:
    """Read a CSV/TSV file into a DataFrame.

    Extra keyword arguments are forwarded to :func:`pandas.read_csv`.
    """
    require_deps("pandas")
    return pd.read_csv(path, **kwargs)


def read_json(path: str | Path, **kwargs: Any) -> Any:
    """Read a JSON file into a DataFrame.

    Extra keyword arguments are forwarded to :func:`pandas.read_json`.
    """
    require_deps("pandas")
    return pd.read_json(path, **kwargs)


def read_excel(
    path: str | Path,
    sheet_name: str | int | list[str] | list[int] | None = 0,
    **kwargs: Any,
) -> Any:
    """Read an Excel workbook sheet into a DataFrame.

    Extra keyword arguments are forwarded to :func:`pandas.read_excel`.
    """
    require_deps("pandas")
    try:
        import openpyxl  # noqa: F401
    except ImportError:  # pragma: no cover
        raise ImportError(
            "读取 Excel 需要 openpyxl, 请安装可选依赖: pip install wei-data-shu[analysis] "
            "(或 wei-data-shu[excel])"
        ) from None
    return pd.read_excel(path, sheet_name=sheet_name, **kwargs)


def read_any(path: str | Path, **kwargs: Any) -> Any:
    """Read a data file into a DataFrame by dispatching on its extension.

    Supported: CSV/TSV/TXT, JSON, Excel (xlsx/xls/xlsm).
    """
    suffix = Path(path).suffix.lower()
    if suffix in _CSV_ALIASES:
        sep = kwargs.pop("sep", "\t" if suffix == ".tsv" else ",")
        return read_csv(path, sep=sep, **kwargs)
    if suffix == ".json":
        return read_json(path, **kwargs)
    if suffix in _EXCEL_ALIASES:
        return read_excel(path, **kwargs)
    raise ValueError(
        f"不支持的文件类型: {suffix!r}. 支持的扩展名: {sorted(_CSV_ALIASES | _EXCEL_ALIASES | {'.json'})}"
    )


__all__ = ["read_csv", "read_json", "read_excel", "read_any"]
