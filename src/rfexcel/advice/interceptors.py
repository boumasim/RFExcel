from __future__ import annotations

import inspect
from collections.abc import Callable
from functools import wraps
from typing import TYPE_CHECKING, Concatenate, ParamSpec, TypeVar, cast

from rfexcel.backend.interfaces.i_library import ISetExcel

if TYPE_CHECKING:
    from rfexcel.backend.writer.xls_writer import XlsWriter
    
P = ParamSpec("P")
R = TypeVar("R")

MANAGED_COMPONENTS = {"resource", "style", "metadata", "reader", "writer"}

def auto_convert_xls_to_xlsx[**P, R](
	method: Callable[Concatenate[XlsWriter, P], R],
) -> Callable[Concatenate[XlsWriter, P], R]:
    @wraps(method)
    def wrapper(self: XlsWriter, *args: P.args, **kwargs: P.kwargs) -> R:
        ref: ISetExcel = self.resolve_weak_ref()
        ref.xls_to_xlsx()

        sig = inspect.signature(method)
        bound = sig.bind(self, *args, **kwargs)

        for param_name in MANAGED_COMPONENTS:
            if param_name in bound.arguments:
                bound.arguments[param_name] = getattr(ref, param_name)
        
        new_method = getattr(ref.writer, method.__name__)
        return cast(R, new_method(*bound.args[1:], **bound.kwargs))
    
    return wrapper