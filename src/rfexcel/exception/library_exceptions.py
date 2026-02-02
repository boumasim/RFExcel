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