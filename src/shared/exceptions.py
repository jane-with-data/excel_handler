"""
Custom exception classes for the application.

All application-specific exceptions inherit from BaseAppError
for consistent error handling throughout the application.
"""

__all__ = [
    "BaseAppError",
    "ValidationError",
    "ExcelFileNotFoundError",
    "ZaloAPIError",
    "ConfigError",
]


class BaseAppError(Exception):
    """
    Base exception for all application-specific errors.
    
    This class provides a consistent error interface with
    optional error codes for categorization and tracking.
    """
    
    def __init__(self, message: str, error_code: str = "") -> None:
        """
        Initialize exception.
        
        Args:
            message: Human-readable error message
            error_code: Optional error code for categorization
        
        Example:
            >>> raise BaseAppError("Something went wrong", error_code="GENERAL_ERROR")
        """
        self.message = message
        self.error_code = error_code
        super().__init__(self.message)
    
    def __str__(self) -> str:
        """Return formatted error message."""
        if self.error_code:
            return f"[{self.error_code}] {self.message}"
        return self.message


class ValidationError(BaseAppError):
    """
    Raised when data validation fails.
    
    This exception is raised when input data doesn't meet
    required constraints or format specifications.
    """
    
    def __init__(self, message: str, field: str = "") -> None:
        """
        Initialize validation error.
        
        Args:
            message: Description of validation failure
            field: Optional field name that failed validation
        """
        super().__init__(message, error_code="VALIDATION_ERROR")
        self.field = field


class ExcelFileNotFoundError(BaseAppError):
    """
    Raised when Excel file cannot be found or read.
    
    This exception indicates file access issues such as:
    - File doesn't exist
    - File is corrupted
    - Permission denied
    """
    
    def __init__(self, message: str) -> None:
        """
        Initialize file not found error.
        
        Args:
            message: Description of file issue
        """
        super().__init__(message, error_code="FILE_NOT_FOUND")

class ConfigError(BaseAppError):
    """
    Raised when configuration is invalid or incomplete.
    
    This exception indicates issues with application
    configuration such as missing required settings.
    """
    
    def __init__(self, message: str) -> None:
        """
        Initialize configuration error.
        
        Args:
            message: Description of config issue
        
        Example:
            >>> raise ConfigError("Missing ZALO_API_KEY in .env")
        """
        super().__init__(message, error_code="CONFIG_ERROR")