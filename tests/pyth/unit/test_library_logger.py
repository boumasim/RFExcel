import pytest

from rfexcel.utils.library_logger import DefaultLogger, LoggerProtocol, RobotLogger

SINGLETON_LOGGERS = [DefaultLogger, RobotLogger]


@pytest.mark.parametrize("logger_class", SINGLETON_LOGGERS)
def test_logger_creates_instance_on_first_call(logger_class: LoggerProtocol) -> None:
    instance = logger_class()
    assert instance is not None


@pytest.mark.parametrize("logger_class", SINGLETON_LOGGERS)
def test_logger_reuses_instance_on_subsequent_calls(
    logger_class: LoggerProtocol,
) -> None:
    first = logger_class()
    second = logger_class()
    assert first is second
