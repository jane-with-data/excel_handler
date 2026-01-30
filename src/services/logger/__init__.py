"""Logger service package.

Provides centralized logging configuration using the Singleton pattern.

Classes:
    Logger: Singleton logger with file and console handlers

Usage:
    >>> from src.services.logger_service.logger import get_logger
    >>> logger = get_logger()
    >>> logger.info("Application started")
"""

from services.logger.logger import get_logger

__all__ = ["get_logger"]
