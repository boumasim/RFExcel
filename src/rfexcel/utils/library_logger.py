from abc import ABC, abstractmethod
import logging
from typing import override
from robot.api import logger as robot_logger


class LoggerProtocol(ABC):
    @abstractmethod
    def info(self, msg: str) -> None:
        pass

    @abstractmethod
    def warn(self, msg: str) -> None:
        pass

    @abstractmethod
    def error(self, msg: str) -> None:
        pass


class DefaultLogger(LoggerProtocol):

    def __init__(self) -> None:
        self._log = logging.getLogger("rfexcel")

    @override
    def info(self, msg: str) -> None:
        self._log.info(msg)

    @override
    def warn(self, msg: str) -> None:
        self._log.warning(msg)

    @override
    def error(self, msg: str) -> None:
        self._log.error(msg)

class RobotLogger(LoggerProtocol):

    def __init__(self) -> None:
        self._log = robot_logger

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
        self._delegate: LoggerProtocol = DefaultLogger()

    def configure(self, delegate: LoggerProtocol) -> None:
        self._delegate = delegate

    def info(self, msg: str) -> None:
        self._delegate.info(msg)

    def warn(self, msg: str) -> None:
        self._delegate.warn(msg)

    def error(self, msg: str) -> None:
        self._delegate.error(msg)
logger = LibraryLogger()
