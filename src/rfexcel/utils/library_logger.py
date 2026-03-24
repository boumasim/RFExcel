import logging
from typing import Protocol, runtime_checkable


@runtime_checkable
class LoggerProtocol(Protocol):
    def info(self, msg: str) -> None: ...
    def warn(self, msg: str) -> None: ...


class _StdlibAdapter:

    def __init__(self) -> None:
        self._log = logging.getLogger("rfexcel")

    def info(self, msg: str) -> None:
        self._log.info(msg)

    def warn(self, msg: str) -> None:
        self._log.warning(msg)


class LibraryLogger:

    def __init__(self) -> None:
        self._delegate: LoggerProtocol = _StdlibAdapter()

    def configure(self, delegate: LoggerProtocol) -> None:
        self._delegate = delegate

    def info(self, msg: str) -> None:
        self._delegate.info(msg)

    def warn(self, msg: str) -> None:
        self._delegate.warn(msg)


logger = LibraryLogger()
