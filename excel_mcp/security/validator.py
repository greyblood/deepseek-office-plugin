"""
Path validation and directory containment security module.
Enforces strict sandboxing against path traversal and unauthorized filesystem access.
"""

from pathlib import Path
from typing import List, Optional, Tuple
from excel_mcp.config import config
from excel_mcp.logging_config import logger


class SecurityError(Exception):
    """Raised when an operation violates security policy or sandboxing rules."""
    pass


def validate_file_path(
    path_str: str | Path,
    must_exist: bool = True,
    allowed_dirs: Optional[List[Path]] = None,
    allowed_extensions: Optional[Tuple[str, ...]] = None,
) -> Path:
    """
    Validates and canonicalizes a file path.
    
    1. Resolves symlinks and relative segments (preventing '../' directory traversal).
    2. Validates that the resolved target lies within one of the allowed directories.
    3. Checks file extension against the whitelist (.xlsx, .xlsm).
    4. Validates existence and file size if required.
    
    Returns the resolved Path object.
    Raises SecurityError or FileNotFoundError if invalid.
    """
    if not path_str or not str(path_str).strip():
        raise ValueError("File path cannot be empty.")

    target = Path(path_str).expanduser().resolve()
    
    # 1. Extension check
    valid_exts = allowed_extensions or config.allowed_extensions
    if target.suffix.lower() not in [ext.lower() for ext in valid_exts]:
        raise SecurityError(
            f"Unsupported file format '{target.suffix}'. "
            f"Allowed extensions are: {', '.join(valid_exts)}"
        )

    # 2. Directory confinement check
    dirs_to_check = allowed_dirs if allowed_dirs is not None else config.allowed_directories
    if dirs_to_check:
        is_allowed = False
        for allowed_dir in dirs_to_check:
            try:
                # Target must be relative to the allowed root (i.e. inside it)
                target.relative_to(allowed_dir)
                is_allowed = True
                break
            except ValueError:
                continue

        if not is_allowed:
            allowed_roots = [str(d) for d in dirs_to_check]
            logger.warning("Access denied: %s outside allowed roots", target)
            raise SecurityError(
                f"Access denied: '{target.name}' is outside permitted directories. "
                f"Configured allowed roots: {allowed_roots}"
            )

    # 3. Existence and size check
    if must_exist:
        if not target.exists():
            raise FileNotFoundError(f"Workbook not found at path: '{target}'")
        if not target.is_file():
            raise ValueError(f"Path exists but is not a file: '{target}'")
        
        file_size = target.stat().st_size
        if file_size > config.max_file_size_bytes:
            max_mb = config.max_file_size_bytes // (1024 * 1024)
            actual_mb = file_size / (1024 * 1024)
            raise SecurityError(
                f"File size ({actual_mb:.1f} MB) exceeds maximum permitted limit ({max_mb} MB)."
            )

    return target


def validate_directory_path(
    dir_str: str | Path,
    must_exist: bool = True,
    allowed_dirs: Optional[List[Path]] = None,
) -> Path:
    """
    Validates that a directory path is legitimate and contained within allowed roots.
    """
    if not dir_str or not str(dir_str).strip():
        raise ValueError("Directory path cannot be empty.")

    target = Path(dir_str).expanduser().resolve()
    dirs_to_check = allowed_dirs if allowed_dirs is not None else config.allowed_directories
    
    if dirs_to_check:
        is_allowed = False
        for allowed_dir in dirs_to_check:
            try:
                target.relative_to(allowed_dir)
                is_allowed = True
                break
            except ValueError:
                # Also allow if the target is an exact parent or equal to allowed_dir
                if target == allowed_dir:
                    is_allowed = True
                    break

        if not is_allowed:
            raise SecurityError(
                f"Access denied: Directory '{target}' is outside permitted roots."
            )

    if must_exist:
        if not target.exists():
            raise FileNotFoundError(f"Directory not found: '{target}'")
        if not target.is_dir():
            raise ValueError(f"Path is not a directory: '{target}'")

    return target
