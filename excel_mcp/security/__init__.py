"""
Security and validation package for DeepSeek Excel Studio.
"""

from .validator import validate_file_path, validate_directory_path, SecurityError
from .backup import BackupManager
from .guard import ConfirmationGuard

__all__ = [
    "validate_file_path",
    "validate_directory_path",
    "SecurityError",
    "BackupManager",
    "ConfirmationGuard",
]
