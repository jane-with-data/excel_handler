"""
Excel named style builder module.

This module create all named style needed.
"""

from openpyxl.styles import NamedStyle, Font, PatternFill, Alignment
from typing import List
from shared.constants import DATA_TYPE_STYLE_NAME_CONFIG, VISUAL_STYLE_NAME_CONFIG
from services.logger.logger import get_logger
logger = get_logger()

def _build_named_style(style_name: str, style_type: str) -> NamedStyle:
    """
    Convert raw config to actual NamedStyle object.
    
    Args:
        style_name: Name of the style
        style_type: Type of style ('VISUAL_STYLE_NAME_CONFIG' or 'DATA_TYPE_STYLE_NAME_CONFIG')
    
    Returns:
        NamedStyle: Configured NamedStyle object or None if config not found
    """
    logger.debug(f"style_name: {style_name}, style_type: {style_type}")
    
    # Init NamedStyle
    named_style_obj = NamedStyle(name=style_name)
    
    # Process: STYLE_NAME_CONFIG
    if style_type == 'VISUAL_STYLE_NAME_CONFIG':
        style_name_cfg = VISUAL_STYLE_NAME_CONFIG.get(style_name)
        if not style_name_cfg:
            logger.warning(f"Config style '{style_name}' not found in VISUAL_STYLE_NAME_CONFIG")
            return None
        
        logger.debug("Initializing visual format style")
        
        # Init `font`
        if "font_name" in style_name_cfg:
            named_style_obj.font = Font(
                name=style_name_cfg.get("font_name", 'Calibri'),
                size=style_name_cfg.get("font_size", 11),
                bold=style_name_cfg.get("font_bold", False),
                italic=style_name_cfg.get("font_italic", False),
                vertAlign=style_name_cfg.get("font_vertAlign", None),
                underline=style_name_cfg.get("font_underline", 'none'),
                strike=style_name_cfg.get("font_strike", False),
                color=style_name_cfg.get("font_color", '000000')
            )
        
        # Init `pattern_fill`
        if "fill_type" in style_name_cfg:
            named_style_obj.fill = PatternFill(
                fill_type=style_name_cfg.get("fill_type", None),
                start_color=style_name_cfg.get("fill_start_color", "FFFFFF"),
                end_color=style_name_cfg.get("fill_end_color", "FFFFFF")
            )
        
        # Init `alignment`
        if "alignment_horizontal" in style_name_cfg:
            named_style_obj.alignment = Alignment(
                horizontal=style_name_cfg.get("alignment_horizontal", 'left'),
                vertical=style_name_cfg.get("alignment_vertical", "top"),
                wrap_text=style_name_cfg.get("alignment_wrap_text", False)
            )
    
    # Process: DATA_TYPE_STYLE_NAME_CONFIG
    elif style_type == 'DATA_TYPE_STYLE_NAME_CONFIG':
        logger.debug("Initializing data type format style")
        style_name_cfg = DATA_TYPE_STYLE_NAME_CONFIG.get(style_name)
        if not style_name_cfg:
            logger.warning(f"Config style '{style_name}' not found in DATA_TYPE_STYLE_NAME_CONFIG")
            return None
        
        # Init `format`
        if "format" in style_name_cfg:
            named_style_obj.number_format = style_name_cfg.get("format", 'General')
    
        # Init `alternating_style`
            
        # Init `auto_adjust_style`
        
        # Init `freeze_panes_style`
        
        # Init `filter_style`
        
    return named_style_obj

def build_bulk_named_style() -> List[NamedStyle]:
    # Init full `VISUAL_STYLE_NAME_CONFIG`
    res_lst = []
    for snc in VISUAL_STYLE_NAME_CONFIG.keys():
        snc_ = _build_named_style(snc, "VISUAL_STYLE_NAME_CONFIG")
        res_lst.append(snc_)
        
    # Init full `DATA_TYPE_STYLE_NAME_CONFIG`
    for snc in DATA_TYPE_STYLE_NAME_CONFIG.keys():
        snc_ = _build_named_style(snc, "DATA_TYPE_STYLE_NAME_CONFIG")
        res_lst.append(snc_)
    return res_lst
