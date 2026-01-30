"""
Excel reader module for handling input data.

This module provides utilities to read and validate Excel files
with required column checking and error handling.
"""

import logging
from pathlib import Path
from typing import List, Optional
import pandas as pd

from shared.exceptions import (
    ExcelFileNotFoundError,
    ValidationError,
)

logger = logging.getLogger(__name__)

__all__ = ["ExcelReader"]


class ExcelReader:
    """
    Reads and validates Excel files.
    
    This class handles loading Excel files and validating that
    all required columns are present before returning data.
    
    Attributes:
        file_path (Path): Path to the Excel file
        required_columns (List[str]): List of required column names
    
    Raises:
        ExcelFileNotFoundError: If file doesn't exist or can't be read
        ValidationError: If required columns are missing
    
    Example:
        >>> reader = ExcelReader(
        ...     "data/input/phones.xlsx",
        ...     required_columns=["phone", "status"]
        ... )
        >>> df = reader.read()
        >>> print(df.shape)  # (1000, 2)
    """
    
    def __init__(
        self,
        file_path: str,
        required_columns: Optional[List[str]] = None,
    ) -> None:
        """
        Initialize ExcelReader.
        
        Args:
            file_path: Path to Excel file (str or Path-like)
            required_columns: List of required column names.
                              Defaults to empty list (no validation).
        """
        self.file_path = Path(file_path)
        self.required_columns = required_columns or []
    
    def read(self) -> pd.DataFrame:
        """
        Read and validate Excel file.
        
        Returns:
            DataFrame with all data from the Excel file
        
        Raises:
            ExcelFileNotFoundError: If file doesn't exist or is unreadable
            ValidationError: If required columns are missing
        
        Example:
            >>> reader = ExcelReader("data.xlsx", required_columns=["id"])
            >>> df = reader.read()  # Raises if "id" column missing
        """
        # Check if file exists
        self._check_file_exists()
        
        # Read Excel file
        df = self._read_excel_file()
        
        # Validate required columns
        self._validate_columns(df)
        
        logger.info(
            f"Successfully read Excel file: {self.file_path} "
            f"({df.shape[0]} rows, {df.shape[1]} columns)"
        )
        return df
    
    def _check_file_exists(self) -> None:
        """
        Check if Excel file exists.
        
        Raises:
            ExcelFileNotFoundError: If file not found
        """
        if not self.file_path.exists():
            raise ExcelFileNotFoundError(
                f"Excel file not found: {self.file_path}"
            )
    
    def _read_excel_file(self) -> pd.DataFrame:
        """
        Read Excel file into DataFrame.
        
        Returns:
            DataFrame with Excel data
        
        Raises:
            ExcelFileNotFoundError: If file cannot be read
        """
        try:
            logger.debug(f"Reading Excel file: {self.file_path}")
            df = pd.read_excel(self.file_path)
            return df
        except pd.errors.ParserError as e:
            logger.error(f"Invalid Excel format: {e}", exc_info=True)
            raise ExcelFileNotFoundError(
                f"Cannot parse Excel file: {self.file_path}"
            ) from e
        except Exception as e:
            logger.error(f"Error reading Excel file: {e}", exc_info=True)
            raise ExcelFileNotFoundError(
                f"Cannot read Excel file: {self.file_path}"
            ) from e
    
    def _validate_columns(self, df: pd.DataFrame) -> None:
        """
        Validate that required columns exist in DataFrame.
        
        Args:
            df: DataFrame to validate
        
        Raises:
            ValidationError: If required columns are missing
        """
        if not self.required_columns:
            # No validation needed
            return
        
        # Check for missing columns
        available_cols = set(df.columns)
        required_set = set(self.required_columns)
        missing_cols = required_set - available_cols
        
        if missing_cols:
            raise ValidationError(
                message=(
                    f"Excel file missing required columns: {sorted(missing_cols)}. "
                    f"Available columns: {sorted(available_cols)}"
                ),
                field="columns"
            )