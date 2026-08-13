"""Analysis domain exports (data loading, cleaning, visualization)."""

from importlib import import_module

__all__ = [
    "read_csv",
    "read_json",
    "read_excel",
    "read_any",
    "DataCleaner",
    "plot_line",
    "plot_bar",
    "plot_hist",
    "plot_box",
    "plot_scatter",
    "plot_pie",
    "plot_corr_heatmap",
    "setup_chinese_font",
]

_EXPORTS = {
    "read_csv": ("wei_data_shu.analysis.io", "read_csv"),
    "read_json": ("wei_data_shu.analysis.io", "read_json"),
    "read_excel": ("wei_data_shu.analysis.io", "read_excel"),
    "read_any": ("wei_data_shu.analysis.io", "read_any"),
    "DataCleaner": ("wei_data_shu.analysis.cleaning", "DataCleaner"),
    "plot_line": ("wei_data_shu.analysis.charts", "plot_line"),
    "plot_bar": ("wei_data_shu.analysis.charts", "plot_bar"),
    "plot_hist": ("wei_data_shu.analysis.charts", "plot_hist"),
    "plot_box": ("wei_data_shu.analysis.charts", "plot_box"),
    "plot_scatter": ("wei_data_shu.analysis.charts", "plot_scatter"),
    "plot_pie": ("wei_data_shu.analysis.charts", "plot_pie"),
    "plot_corr_heatmap": ("wei_data_shu.analysis.charts", "plot_corr_heatmap"),
    "setup_chinese_font": ("wei_data_shu.analysis.charts", "setup_chinese_font"),
}


def __getattr__(name: str):
    target = _EXPORTS.get(name)
    if target is None:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    module_name, attr_name = target
    module = import_module(module_name)
    return getattr(module, attr_name)
