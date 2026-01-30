"""
Configuration management module.

This module provides application-wide configuration using Pydantic settings
with support for environment variables and .env files.

Example:
    >>> from src.shared.configs import get_config
    >>> config = get_config()
    >>> print(config.zalo_api_key)
"""

from pathlib import Path
from typing import Optional

from pydantic_settings import BaseSettings

__all__ = ["AppConfig", "get_config"]


class AppConfig(BaseSettings):
    """
    Application configuration.
    
    Configuration is loaded from environment variables and .env file.
    All settings can be overridden via environment variables.
    
    Attributes:
        project_root: Root directory of project
        data_dir: Data directory path
        input_dir: Input data directory
        output_dir: Output results directory
        logs_dir: Logs directory
        zalo_api_key: Zalo API key
        zalo_url: Zalo service URL
        max_retries: Maximum retry attempts
        retry_delay_seconds: Delay between retries
    
    Example:
        Configure via .env file:
            PROJECT_ROOT=/path/to/project
            ZALO_API_KEY=your_key_here
            ZALO_URL=https://chat.zalo.me
    """
    
    # Paths
    project_root: Path = Path(__file__).parent.parent.parent
    data_dir: Path = project_root / "data"
    input_dir: Path = data_dir / "input"
    output_dir: Path = data_dir / "output"
    logs_dir: Path = data_dir / "logs"
    temp_dir: Path = data_dir / "temp"
    
    # Retry settings
    max_retries: int = 3
    retry_delay_seconds: float = 2.0
    network_check_timeout_seconds: int = 5
    network_check_max_wait_seconds: int = 300
    
    # Excel I/O
    input_excel_file: str = "phone_check.xlsx"
    output_excel_file: str = "phone_check_result.xlsx"
    backup_progress_file: str = "backup_progress.json"
    auto_save_interval: int = 10
    
    # Logging
    log_level: str = "INFO"
    
    class Config:
        """Pydantic configuration."""
        env_file = ".env"
        case_sensitive = False
        extra = "ignore"  # Ignore unknown env vars
    
    def create_directories(self) -> None:
        """
        Create required directories if they don't exist.
        
        Example:
            >>> config = AppConfig()
            >>> config.create_directories()
        """
        for path in [self.data_dir, self.input_dir, self.output_dir, self.logs_dir, self.temp_dir]:
            path.mkdir(parents=True, exist_ok=True)


# Singleton pattern for config
_config_instance: Optional[AppConfig] = None


def get_config() -> AppConfig:
    """
    Get singleton configuration instance.
    
    Returns:
        AppConfig: Application configuration
    
    Example:
        >>> config = get_config()
        >>> print(config.zalo_url)
    """
    global _config_instance
    if _config_instance is None:
        _config_instance = AppConfig()
        _config_instance.create_directories()
    return _config_instance