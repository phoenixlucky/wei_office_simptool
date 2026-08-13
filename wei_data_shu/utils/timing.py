"""Timing helpers."""

from __future__ import annotations

import time
from functools import wraps
from typing import Any, Callable, TypeVar, cast

F = TypeVar("F", bound=Callable[..., Any])


def fn_timer(func: F) -> F:
    """Wrap ``func`` to print its wall-clock runtime and return ``(result, elapsed)``."""

    @wraps(func)
    def function_timer(*args: Any, **kwargs: Any) -> tuple[Any, float]:
        t0 = time.perf_counter()
        result = func(*args, **kwargs)
        t1 = time.perf_counter()
        elapsed_time = t1 - t0
        print(f"Total time running {func.__name__}: {elapsed_time:.2f} seconds")
        return result, elapsed_time

    return cast(F, function_timer)


__all__ = ["fn_timer"]
