"""Application constants and configuration values.

This module defines all constants used throughout the application,
including data type configurations, styling, and Excel format settings.

Constants are defined with type hints and Final annotations to prevent
accidental modifications.
"""

from typing import Final, Dict, List, Any, Set

# =============================================================================
# SYSTEM & PROJECT INFORMATION
DEFAULT_ENCODING: Final[str] = "utf-8"
MAX_WIDTH: Final[int] = 60

# LOGGING =============================================================================
LOG: Final[Dict[str, Any]] = {
    "SET_LEVEL_FILE": "DEBUG",
    "SET_LEVEL_CONSOLE": "DEBUG",
    "NAME": "app.log",
    "LEVEL": "INFO",
    "WHEN": "midnight",
    "INTERVAL": 1,
    "BACKUP_COUNT": 30
}

# =============================================================================
# THÔNG SỐ WORKBOOK & WORKSHEET
# column_name -> data_type -> style_name

# Declare `data type` for each `data field`
DECLARE_DATA_TYPE_CONFIG: Final[Dict[str, str]] = {
    "dated_created": "date_time",
    "date_modified": "date",
    "date_checked": "date_time",
    "date": "date_time",
    "number_no_decimal": "int",
    "number_decimal": "float",
    "currency": "currency",
    "percentage": "percentage"
}

# Map `data_type` vs `data_style`
MAP_DATA_TYPE_STYLE_NAME_CONFIG: Final[Dict[str, str]] = {
    "date": "date_style",
    "date_time": "date_time_style",
    "int": "int_style",
    "float": "float_style",
    "currency": "currency_style",
    "percentage": "percentage_style",
    "default_data_type": "default_style"
}

# Declare format for each data type style
DATA_TYPE_STYLE_NAME_CONFIG: Final[Dict[str, Any]] = {
    # Định dạng Ngày tháng
    "date_style": {
        "format": "dd/mm/yyyy"
    },

    "date_time_style": {
        "format": "dd/mm/yyyy hh:mm:ss"
    },
    
    # Định dạng Số
    "int_style": {
        "format": "#,##0"
    },
    
    "float_style": {
        "format": "#,##0.00"
    },

    # Định dạng Tiền tệ
    "currency_style": {
        "format": '"$"#,##0;[Red]-#,##0',
    },

    # Định dạng Phần trăm
    "percentage_style": {
        "format": "0.00%",
        # "alignment_horizontal": "center",
    },
    
    "default_style": {
        "format": "General",
        # "alignment_horizontal": "center",
    }  
}

# Cấu hình format style cho Workbook/Worksheet
VISUAL_STYLE_NAME_CONFIG: Final[Dict[str, Any]] = {
    # Phong cách tiêu đề (Header)
    "header_style": {
        "font_name": "Calibri",
        "font_size": 12,
        "font_bold": True,
        "font_italic": False,   
        "font_vertAlign": None,
        "font_underline": 'none',
        "font_strike": False,
        "font_color": "FFFFFF",
        "fill_type": "solid",
        "fill_start_color": "366092",
        "fill_end_color": "366092",
        "alignment_horizontal": "center",
        "alignment_vertical": "center",
        "alignment_wrap_text": False
    },
    
    # Phong cách mặc định cho phần thân (Body)
    "body_style": {
        "font_name": "Calibri",
        "font_size": 11,
        "font_bold": False,
        "font_italic": False,   
        "font_vertAlign": None,
        "font_underline": 'none',
        "font_strike": False,
        "font_color": "000000",
        "fill_type": None,
        "fill_start_color": "FFFFFF",
        "fill_end_color": "FFFFFF",
        "alignment_horizontal": "left",
        "alignment_vertical": "top",
        "alignment_wrap_text": False
    },

    # Tự động căn chỉnh kích thước (Auto adjust)
    "auto_adjust_style": {
        "auto_adjust_width": True,
        "max_column_width": 60,
        "auto_adjust_height": False,
        "max_column_height": 15
    },
    
    # Hiệu ứng dòng kẻ sọc (Alternating rows)
    "alternating_style": {
        "alternating_rows": True,
        "fill_start_color": "F8F9FA",
        "fill_end_color": "FFFFFF",
    },
    
    # Freeze Panes config
    "freeze_panes_style": {
        "mode_on": True,
        "freeze_cell": "B2"
    },
    
    # Filter config
    "filter_style": {
        "mode_on": True
    }
}