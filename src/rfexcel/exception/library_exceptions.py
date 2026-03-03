class RFExcelException(Exception):
    """Base exception for RFExcelLibrary related errors"""

    def __init__(self, msg: str):
        self._message: str = msg
        super().__init__(self._message)

class FileFormatNotSupportedException(RFExcelException):
    def __init__(self, msg: str = "File format not supported in this library"):
        super().__init__(msg)

class FileAlreadyExistsException(RFExcelException):
    def __init__(self, msg: str = "File with same name already exits"):
        super().__init__(msg)

class FileDoesNotExistException(RFExcelException):
    def __init__(self, path: str):
        super().__init__(f"File {path} does not exist")

class LibraryException(RFExcelException):
    """Exception for invalid operations on library objects"""
    def __init__(self, msg: str):
        super().__init__(msg)

class RowIndexOutOfBoundsException(RFExcelException):
    """Exception when row index is out of valid range"""
    def __init__(self, row_index: int, msg: str = ""):
        if msg:
            super().__init__(msg)
        else:
            super().__init__(f"Row index {row_index} is out of bounds")

class StreamingViolationException(RFExcelException):
    """Exception when trying to read backwards in streaming mode"""
    def __init__(self, row_index: int, last_read: int):
        super().__init__(
            f"Cannot read row {row_index} in streaming mode. "
            f"Already read up to row {last_read}. "
            f"Streaming only supports forward-only sequential access."
        )

class OperationNotSupportedForFormat(RFExcelException):
    """Exception raised when an operation is not supported for a specific file format"""
    def __init__(self, msg: str = "This operation is not supported for the current file format"):
        super().__init__(msg)
        
class NotSupportedInReadOnlyMode(RFExcelException):
    """Exception raised when trying to perform a write operation in read-only mode"""
    def __init__(self, msg: str = "This operation is not supported in read-only mode"):
        super().__init__(msg)