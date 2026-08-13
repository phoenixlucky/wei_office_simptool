"""Mail domain exports."""

from importlib import import_module

__all__ = ["DailyEmailReport", "MailError"]

_EXPORTS = {
    "DailyEmailReport": "DailyEmailReport",
    "MailError": "MailError",
}


def __getattr__(name: str):
    target = _EXPORTS.get(name)
    if target is None:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    module = import_module("wei_data_shu.mail.report")
    return getattr(module, target)
