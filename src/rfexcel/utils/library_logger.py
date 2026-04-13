from __future__ import annotations

import logging
import types
from abc import ABC, abstractmethod
from typing import override

from robot.api import logger as robot_logger


class ILogger(ABC):
    @abstractmethod
    def info(self, msg: str) -> None:
        pass

    @abstractmethod
    def warn(self, msg: str) -> None:
        pass

    @abstractmethod
    def error(self, msg: str) -> None:
        pass


class DefaultLogger(ILogger):
    _instance: DefaultLogger | None = None
    _log: logging.Logger

    def __new__(cls) -> DefaultLogger:
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._log = logging.getLogger("rfexcel")
        return cls._instance

    @override
    def info(self, msg: str) -> None:
        self._log.info(msg)

    @override
    def warn(self, msg: str) -> None:
        self._log.warning(msg)

    @override
    def error(self, msg: str) -> None:
        self._log.error(msg)


class RobotLogger(ILogger):
    _instance: RobotLogger | None = None
    _log: types.ModuleType

    def __new__(cls) -> RobotLogger:
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._log = robot_logger
        return cls._instance

    @override
    def info(self, msg: str) -> None:
        self._log.info(msg)

    @override
    def warn(self, msg: str) -> None:
        self._log.warn(msg)

    @override
    def error(self, msg: str) -> None:
        self._log.error(msg)


class LibraryLogger:
    def __init__(self) -> None:
        self._delegate: ILogger = DefaultLogger()

    def configure(self, delegate: ILogger) -> None:
        self._delegate = delegate

    def info(self, msg: str) -> None:
        self._delegate.info(msg)

    def warn(self, msg: str) -> None:
        self._delegate.warn(msg)

    def error(self, msg: str) -> None:
        self._delegate.error(msg)


logger = LibraryLogger()
