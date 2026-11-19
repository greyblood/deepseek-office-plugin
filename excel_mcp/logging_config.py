"""
Logging configuration for DeepSeek Excel Studio MCP Server.
Ensures all logs write strictly to stderr to preserve stdout for MCP JSON-RPC protocol transport.
Never exposes sensitive cell contents or workbook records in log outputs.
"""

import logging
import sys
from typing import Optional


def setup_logging(level: str = "INFO", log_file: Optional[str] = None) -> logging.Logger:
    """
    Configures application-wide logging.
    
    CRITICAL: In MCP stdio mode, stdout is reserved for JSON-RPC frames.
    All logging MUST be directed to sys.stderr or an external file.
    """
    numeric_level = getattr(logging, level.upper(), logging.INFO)
    
    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] [excel-mcp] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    
    root_logger = logging.getLogger("excel_mcp")
    root_logger.setLevel(numeric_level)
    root_logger.handlers.clear()
    
    # Stderr stream handler
    stderr_handler = logging.StreamHandler(sys.stderr)
    stderr_handler.setLevel(numeric_level)
    stderr_handler.setFormatter(formatter)
    root_logger.addHandler(stderr_handler)
    
    # Optional file handler
    if log_file:
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setLevel(numeric_level)
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)
        
    return root_logger


# Default logger
logger = logging.getLogger("excel_mcp")
