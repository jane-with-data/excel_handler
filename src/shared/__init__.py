"""Shared utilities package.

Contains shared utilities used across the application:
- configs: Configuration management
- constants: Application constants
- exceptions: Custom exception definitions
- settings: Legacy settings (deprecated)
"""

from shared.exceptions import BaseAppError
from shared.configs import get_config

__all__ = ["BaseAppError", "get_config"]