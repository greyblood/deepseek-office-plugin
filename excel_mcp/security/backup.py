"""
Automated snapshot and backup management module.
Creates defensive pre-modification copies before any destructive operation.
"""

from datetime import datetime
import shutil
import uuid
from pathlib import Path
from typing import List, Optional
from excel_mcp.config import config
from excel_mcp.logging_config import logger


class BackupManager:
    """Manages file backups and snapshots before writes/modifications."""

    BACKUP_FOLDER_NAME = ".dsh_backup"

    @classmethod
    def get_backup_directory(cls, file_path: Path) -> Path:
        """Determines the backup folder for a given target file."""
        if config.custom_backup_dir:
            backup_dir = config.custom_backup_dir
        else:
            backup_dir = file_path.parent / cls.BACKUP_FOLDER_NAME
        backup_dir.mkdir(parents=True, exist_ok=True)
        return backup_dir

    @classmethod
    def create_backup(cls, file_path: Path) -> Optional[Path]:
        """
        Creates a timestamped snapshot of the target file if it exists.
        Returns the path to the backup file, or None if backups are disabled or file doesn't exist.
        """
        if not config.auto_backup or not file_path.exists():
            return None

        backup_dir = cls.get_backup_directory(file_path)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        unique_token = uuid.uuid4().hex[:6]
        backup_name = f"{file_path.stem}_{timestamp}_{unique_token}.bak{file_path.suffix}"
        backup_path = backup_dir / backup_name

        try:
            shutil.copy2(file_path, backup_path)
            logger.info("Created backup of %s -> %s", file_path.name, backup_path.name)
            return backup_path
        except Exception as e:
            logger.error("Failed to create backup for %s: %s", file_path, e)
            raise IOError(f"Defensive backup creation failed for '{file_path}': {e}") from e

    @classmethod
    def list_backups(cls, file_path: Path) -> List[Path]:
        """Lists all existing backup files for a given workbook."""
        backup_dir = cls.get_backup_directory(file_path)
        if not backup_dir.exists():
            return []

        pattern = f"{file_path.stem}_*.bak{file_path.suffix}"
        backups = sorted(
            backup_dir.glob(pattern),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        return backups

    @classmethod
    def restore_backup(cls, backup_path: Path, target_path: Path) -> None:
        """Restores a backup to the target workbook path."""
        if not backup_path.exists():
            raise FileNotFoundError(f"Backup file not found: {backup_path}")

        # Create a backup of the current state before restoring
        if target_path.exists():
            cls.create_backup(target_path)

        shutil.copy2(backup_path, target_path)
        logger.info("Restored %s from backup %s", target_path.name, backup_path.name)
