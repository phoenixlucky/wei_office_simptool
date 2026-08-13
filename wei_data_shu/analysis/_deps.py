"""Optional dependency helpers for the analysis domain."""

from typing import Any, cast

try:
    import numpy as np
except ImportError:  # pragma: no cover
    np = cast(Any, None)

try:
    import pandas as pd
except ImportError:  # pragma: no cover
    pd = cast(Any, None)

try:
    from matplotlib import pyplot as plt
except ImportError:  # pragma: no cover
    plt = cast(Any, None)


def require_deps(*dep_names: str) -> None:
    """Raise a friendly ImportError if any required optional dependency is missing."""
    available = {
        "numpy": np,
        "pandas": pd,
        "matplotlib": plt,
    }
    missing = [name for name in dep_names if available.get(name) is None]
    if missing:
        missing_list = ", ".join(missing)
        raise ImportError(
            f"当前功能缺少依赖: {missing_list}. 请安装可选依赖: pip install wei-data-shu[analysis]"
        )


__all__ = [
    "np",
    "pd",
    "plt",
    "require_deps",
]
