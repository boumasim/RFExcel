from __future__ import annotations

import inspect
from functools import wraps
from typing import (TYPE_CHECKING, Callable, Concatenate, ParamSpec, TypeVar,
                    cast)

if TYPE_CHECKING:
    from rfexcel.backend.writer.xls_writer import XlsWriter
    from rfexcel.RFExcel import RFExcel

P = ParamSpec("P")
R = TypeVar("R")

MANAGED_COMPONENTS = {"resource", "style", "metadata", "reader", "writer"}

def auto_convert_xls_to_xlsx(method: Callable[Concatenate[XlsWriter, P], R]
                             ) -> Callable[Concatenate[XlsWriter, P], R]:
    @wraps(method)
    def wrapper(self: XlsWriter, *args: P.args, **kwargs: P.kwargs) -> R:
        ref: RFExcel = self.resolve_weak_ref()
        ref.xls_to_xlsx()

        sig = inspect.signature(method)
        bound = sig.bind(self, *args, **kwargs)

        for param_name in MANAGED_COMPONENTS:
            if param_name in bound.arguments:
                bound.arguments[param_name] = getattr(ref, param_name)
        
        new_method = getattr(ref.writer, method.__name__)
        return cast(R, new_method(*bound.args[1:], **bound.kwargs))
    
    return wrapper