"""Quick plotting helpers built on matplotlib + pandas.

Every function returns a ``matplotlib.figure.Figure``. Pass ``save_path`` to
write the figure to disk, or ``show=True`` to display it interactively.

.. note:: 中文标签需要先配置中文字体, 例如::

    import matplotlib
    matplotlib.rcParams["font.sans-serif"] = ["SimHei"]
    matplotlib.rcParams["axes.unicode_minus"] = False
"""

from __future__ import annotations

from typing import Any, Iterable

from ._deps import np, pd, plt, require_deps

_DEFAULT_FIGSIZE = (10, 5)


def _require_and_finalize(fig: Any, save_path: str | None, show: bool) -> Any:
    if save_path is not None:
        fig.savefig(save_path, bbox_inches="tight")
    if show:
        plt.show()
    return fig


def _numeric_cols(df: Any, cols: Iterable[str] | None = None) -> list[str]:
    if cols is not None:
        return [cols] if isinstance(cols, str) else list(cols)
    return list(df.select_dtypes(include=["number"]).columns)


def plot_line(
    df: Any,
    x: str | None = None,
    y: str | Iterable[str] | None = None,
    title: str | None = None,
    figsize: tuple[int, int] = _DEFAULT_FIGSIZE,
    save_path: str | None = None,
    show: bool = False,
    **kwargs: Any,
) -> Any:
    """Line chart: ``y`` (one or more columns) against ``x`` (default: index)."""
    require_deps("pandas", "matplotlib")
    fig, ax = plt.subplots(figsize=figsize)
    if x is not None and y is not None:
        y_cols = [y] if isinstance(y, str) else list(y)
        data = df.set_index(x)[y_cols]
    else:
        data = df
    data.plot(kind="line", ax=ax, **kwargs)
    ax.set_title(title or "")
    ax.grid(alpha=0.3)
    return _require_and_finalize(fig, save_path, show)


def plot_bar(
    df: Any,
    x: str,
    y: str | Iterable[str] | None = None,
    title: str | None = None,
    stacked: bool = False,
    horizontal: bool = False,
    figsize: tuple[int, int] = _DEFAULT_FIGSIZE,
    save_path: str | None = None,
    show: bool = False,
    **kwargs: Any,
) -> Any:
    """Bar chart of ``y`` (default: all numeric columns) grouped by ``x``."""
    require_deps("pandas", "matplotlib")
    fig, ax = plt.subplots(figsize=figsize)
    y_cols = [y] if isinstance(y, str) else list(y) if y is not None else _numeric_cols(df)
    data = df.groupby(x, observed=True)[y_cols].mean(numeric_only=True)
    kind = "barh" if horizontal else "bar"
    data.plot(kind=kind, stacked=stacked, ax=ax, **kwargs)
    ax.set_title(title or "")
    ax.grid(alpha=0.3, axis="x" if horizontal else "y")
    return _require_and_finalize(fig, save_path, show)


def plot_hist(
    df: Any,
    col: str | Iterable[str] | None = None,
    bins: int = 20,
    title: str | None = None,
    figsize: tuple[int, int] = _DEFAULT_FIGSIZE,
    save_path: str | None = None,
    show: bool = False,
    **kwargs: Any,
) -> Any:
    """Histogram of one or more numeric columns (default: all numeric columns)."""
    require_deps("pandas", "matplotlib")
    cols = [col] if isinstance(col, str) else list(col) if col is not None else _numeric_cols(df)
    if not cols:
        raise ValueError("没有可绘制的数值列")
    n = len(cols)
    fig, axes = plt.subplots(1, n, figsize=(figsize[0] * n, figsize[1]), squeeze=False)
    for ax, c in zip(axes.flat, cols):
        df[c].dropna().plot(kind="hist", bins=bins, ax=ax, **kwargs)
        ax.set_title(f"{title or ''} {c}".strip())
        ax.grid(alpha=0.3)
    return _require_and_finalize(fig, save_path, show)


def plot_box(
    df: Any,
    cols: Iterable[str] | None = None,
    title: str | None = None,
    figsize: tuple[int, int] = _DEFAULT_FIGSIZE,
    save_path: str | None = None,
    show: bool = False,
    **kwargs: Any,
) -> Any:
    """Box plot of numeric columns (default: all numeric columns)."""
    require_deps("pandas", "matplotlib")
    target = _numeric_cols(df, cols)
    if not target:
        raise ValueError("没有可绘制的数值列")
    fig, ax = plt.subplots(figsize=figsize)
    df[target].boxplot(ax=ax, rot=45, **kwargs)
    ax.set_title(title or "")
    ax.grid(alpha=0.3, axis="y")
    return _require_and_finalize(fig, save_path, show)


def plot_scatter(
    df: Any,
    x: str,
    y: str,
    color: str | None = None,
    title: str | None = None,
    figsize: tuple[int, int] = _DEFAULT_FIGSIZE,
    save_path: str | None = None,
    show: bool = False,
    **kwargs: Any,
) -> Any:
    """Scatter plot of ``x`` vs ``y`` (``color`` optionally colors by a column)."""
    require_deps("pandas", "matplotlib")
    fig, ax = plt.subplots(figsize=figsize)
    ax.scatter(df[x], df[y], c=df[color] if color else None, alpha=0.6, **kwargs)
    ax.set_xlabel(x)
    ax.set_ylabel(y)
    ax.set_title(title or "")
    ax.grid(alpha=0.3)
    return _require_and_finalize(fig, save_path, show)


def plot_pie(
    df: Any,
    col: str,
    title: str | None = None,
    figsize: tuple[int, int] = (6, 6),
    save_path: str | None = None,
    show: bool = False,
    **kwargs: Any,
) -> Any:
    """Pie chart of the value counts of a categorical column."""
    require_deps("pandas", "matplotlib")
    counts = df[col].value_counts()
    fig, ax = plt.subplots(figsize=figsize)
    ax.pie(counts, labels=counts.index, autopct="%1.1f%%", **kwargs)
    ax.set_title(title or f"{col} 分布")
    return _require_and_finalize(fig, save_path, show)


def plot_corr_heatmap(
    df: Any,
    cols: Iterable[str] | None = None,
    method: str = "pearson",
    annot: bool = True,
    cmap: str = "coolwarm",
    title: str | None = None,
    figsize: tuple[int, int] = (8, 6),
    save_path: str | None = None,
    show: bool = False,
    **kwargs: Any,
) -> Any:
    """Correlation heatmap of numeric columns.

    ``method`` is passed to ``df.corr`` (pearson / kendall / spearman).
    """
    require_deps("pandas", "matplotlib")
    target = _numeric_cols(df, cols)
    corr = df[target].corr(method=method)
    fig, ax = plt.subplots(figsize=figsize)
    im = ax.imshow(corr, cmap=cmap, vmin=-1, vmax=1, **kwargs)
    ax.set_xticks(range(len(corr)), corr.columns, rotation=45, ha="right")
    ax.set_yticks(range(len(corr)), corr.columns)
    if annot:
        for i in range(len(corr)):
            for j in range(len(corr)):
                ax.text(j, i, f"{corr.iloc[i, j]:.2f}", ha="center", va="center", fontsize=8)
    fig.colorbar(im, ax=ax)
    ax.set_title(title or f"相关性热力图 ({method})")
    return _require_and_finalize(fig, save_path, show)


__all__ = [
    "plot_line",
    "plot_bar",
    "plot_hist",
    "plot_box",
    "plot_scatter",
    "plot_pie",
    "plot_corr_heatmap",
]
