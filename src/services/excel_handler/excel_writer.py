"""
Excel writer module for exporting data to Excel files.

This module provides utilities to write DataFrames to Excel files
with formatting support.
"""

import logging
from pathlib import Path
from typing import Optional, List

import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.dataframe import dataframe_to_rows

from shared.exceptions import (
    ExcelFileNotFoundError,
    ConfigError,
)
# from shared.configs import get_config
from shared.constants import PATH
from services.logger.logger import get_logger
from services.excel_handler.set_sheet_formatter import SetSheetFormatter
logger = logging.getLogger(__name__)

__all__ = ["ExcelWriter"]


class ExcelWriter:
    """
    Write DataFrames to Excel files with formatting.
    
    This class handles creating Excel workbooks, adding sheets,
    writing data, and applying formatting.
    
    Attributes:
        workbook (Workbook): The openpyxl Workbook object
        formatter: Sheet formatter (optional)
    
    Example:
        >>> writer = ExcelWriter()
        >>> writer.write_data(df, "Sheet1", "output.xlsx")
    """
    
    def __init__(self) -> None:
        """
        Initialize ExcelWriter.
        
        Args:
            formatter: Optional formatter object for styling sheets.
                      If not provided, no formatting is applied.
        """
        self.workbook = self._create_workbook()
        self.formatter = SetSheetFormatter(self.workbook)
    
    def _create_workbook(self) -> Workbook:
        """
        Create new empty Workbook.
        
        Returns:
            Empty Workbook object with default sheet removed
        """
        wb = Workbook()
        # Remove default empty sheet
        if wb.active:
            wb.remove(wb.active)
        return wb
    
    def add_sheet(
        self,
        df: pd.DataFrame,
        sheet_name: str = "Sheet1"
    ) -> Worksheet:
        """
        Add DataFrame as a new sheet to workbook.
        
        Args:
            df: DataFrame to add
            sheet_name: Name for the new sheet. Default: "Sheet1"
        
        Returns:
            The created Worksheet object
        
        Raises:
            ValueError: If sheet_name is invalid
        """
        if not sheet_name or not isinstance(sheet_name, str):
            raise ValueError(f"sheet_name must be non-empty string, got: {sheet_name}")
        
        logger.debug(f"Creating sheet '{sheet_name}' with {len(df)} rows")
        
        # Create new sheet
        ws = self.workbook.create_sheet(title=sheet_name)
        
        # Write data to sheet
        for row_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True)):
            ws.append(row)
            if row_idx % 100 == 0:
                logger.debug(f"Wrote {row_idx} rows to sheet")
        
        logger.info(f"Successfully added sheet '{sheet_name}' with {len(df)} rows")
        return ws
    
    def format_sheet(
        self,
        ws: Worksheet,
        df: pd.DataFrame
    ) -> None:
        """
        Apply formatting to a worksheet.
        
        Args:
            ws: Worksheet to format
            df: DataFrame (for reference during formatting)
        """
        if self.formatter is None:
            logger.debug("No formatter available, skipping sheet formatting")
            return
        
        logger.debug(f"Applying formatting to sheet '{ws.title}'")
        self.formatter.run_pipeline(ws, df)
    
    def save(self, file_path: str) -> None:
        """
        Save workbook to file.
        
        Args:
            file_path: Path where to save the file
        
        Raises:
            ExcelFileNotFoundError: If file cannot be written
        
        Example:
            >>> writer.save("data/output/results.xlsx")
        """
        file_path = Path(file_path)
        
        try:
            logger.info(f"Saving workbook to: {file_path}")
            file_path.parent.mkdir(parents=True, exist_ok=True)
            self.workbook.save(str(file_path))
            logger.info(f"Successfully saved: {file_path}")
        
        except PermissionError as e:
            logger.error(f"Permission denied when saving: {file_path}", exc_info=True)
            raise ExcelFileNotFoundError(
                f"Cannot write to file (permission denied): {file_path}"
            ) from e
        
        except Exception as e:
            logger.error(f"Error saving workbook to {file_path}: {e}", exc_info=True)
            raise ExcelFileNotFoundError(
                f"Cannot save Excel file: {file_path}"
            ) from e
    

    def export_pipeline(
        self,
        file_name,
        dfs: List[pd.DataFrame],
        sheet_name:  Optional[list[str]] = None
    ) -> None:
        """
        Write DataFrame to Excel file in one step.
        
        Convenience method that combines add_sheet, format_sheet, and save.
        
        Args:
            df: DataFrame to write
            sheet_name: Name for the sheet
            output_path: Output file path. If not provided, uses config default.
        """
        if not file_name:
            raise ConfigError("file_name must be provided for export")
        file_path = PATH["OUTPUT_DIR"] / file_name
        
        logger.info(f"Starting write pipeline for {file_path}")
        
        
        # Add sheets
        for i, df in enumerate(dfs):
            ws = self.add_sheet(df, sheet_name[i] if sheet_name else f"Sheet{i+1}")
            
            # Format sheets
            self.format_sheet(ws, df)
            
        # Save
        self.save(file_path)
        
        logger.info(f"Write pipeline complete: {file_path}")
        
if __name__ == "__main__":
    # Initialize list to hold dataframes
    df = []
    
    # Create first demo dataframe
    df1 = pd.DataFrame({
        'Name': ['Alice', 'Bob', 'Charlie'],
        'Age': [25, 30, 35],
        'City': ['New York', 'London', 'Paris']
    })
    df.append(df1)
    # Create second demo dataframe
    df2 = pd.DataFrame({
        'Product': ['Laptop', 'Phone', 'Tablet'],
        'number_decimal': [9909, 0.699, 1.399],
        'number_no_decimal': [15430, 25, 15]
    })
    df.append(df2)
    print("DataFrame 1:")
    print(df1)
    print("\nDataFrame 2:")
    print(df2)

    ew = ExcelWriter()
    ew.export_pipeline(
        file_name="demo_output.xlsx",
        dfs=df,
        sheet_name=["People", "Products"]
    )