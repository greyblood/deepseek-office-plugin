"""
Configuration module for DeepSeek Excel Studio MCP Server.
"""

import os
from pathlib import Path
from typing import List, Optional
from pydantic import BaseModel, Field


class ServerConfig(BaseModel):
    """Server runtime configuration."""
    
    # Allowed directories for file operations. Empty means current working directory.
    allowed_directories: List[Path] = Field(default_factory=list)
    
    # Whether to automatically create snapshots before destructive mutations
    auto_backup: bool = True
    
    # Explicit custom backup directory. If None, uses `.dsh_backup` inside workbook dir
    custom_backup_dir: Optional[Path] = None
    
    # Safety limits
    max_file_size_bytes: int = 100 * 1024 * 1024  # 100 MB
    max_rows_read: int = 10_000
    max_cols_read: int = 200
    
    # Supported file extensions
    allowed_extensions: tuple[str, ...] = (".xlsx", ".xlsm")
    
    # Require preview and confirmation for destructive operations
    require_confirmation: bool = True
    
    # Logging level
    log_level: str = "INFO"

    @classmethod
    def from_env(cls) -> "ServerConfig":
        """Load configuration from environment variables."""
        allowed_dirs_env = os.getenv("EXCEL_MCP_ALLOWED_DIRS", "").strip()
        allowed_dirs: List[Path] = []
        if allowed_dirs_env:
            # Support semicolon on Windows (to preserve drive letters like C:\...) and colon on POSIX
            separator = ";" if (os.name == "nt" or ";" in allowed_dirs_env) else ":"
            for raw_path in allowed_dirs_env.split(separator):
                p = raw_path.strip()
                if p:
                    allowed_dirs.append(Path(p).resolve())
        else:
            # Default to current working directory
            allowed_dirs.append(Path.cwd().resolve())

        auto_backup = os.getenv("EXCEL_MCP_AUTO_BACKUP", "true").lower() in ("1", "true", "yes")
        
        custom_backup_env = os.getenv("EXCEL_MCP_BACKUP_DIR", "").strip()
        custom_backup_dir = Path(custom_backup_env).resolve() if custom_backup_env else None
        
        max_file_size_mb = int(os.getenv("EXCEL_MCP_MAX_FILE_SIZE_MB", "100"))
        max_rows_read = int(os.getenv("EXCEL_MCP_MAX_ROWS_READ", "10000"))
        max_cols_read = int(os.getenv("EXCEL_MCP_MAX_COLS_READ", "200"))
        require_confirmation = os.getenv("EXCEL_MCP_REQUIRE_CONFIRMATION", "true").lower() in ("1", "true", "yes")
        log_level = os.getenv("EXCEL_MCP_LOG_LEVEL", "INFO").upper()

        return cls(
            allowed_directories=allowed_dirs,
            auto_backup=auto_backup,
            custom_backup_dir=custom_backup_dir,
            max_file_size_bytes=max_file_size_mb * 1024 * 1024,
            max_rows_read=max_rows_read,
            max_cols_read=max_cols_read,
            require_confirmation=require_confirmation,
            log_level=log_level,
        )


# Global default configuration instance
config = ServerConfig.from_env()
