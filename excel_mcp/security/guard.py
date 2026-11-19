"""
Confirmation and preview guard for destructive operations.
Ensures destructive actions (e.g. deletions, overwriting, bulk clearing)
require explicit confirmation or return an actionable impact preview.
"""

from typing import Any, Dict, Optional
from excel_mcp.config import config
from excel_mcp.logging_config import logger


class ConfirmationRequiredError(Exception):
    """Raised when a destructive operation requires confirmation."""
    def __init__(self, action: str, impact: Dict[str, Any], message: Optional[str] = None):
        self.action = action
        self.impact = impact
        self.message = message or (
            f"Action '{action}' is destructive. Review the impact preview and set confirm=True to proceed."
        )
        super().__init__(self.message)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "status": "confirmation_required",
            "action": self.action,
            "message": self.message,
            "impact_preview": self.impact,
        }


class ConfirmationGuard:
    """Enforces preview-and-confirm workflows for dangerous operations."""

    @staticmethod
    def check_confirmation(
        action_name: str,
        confirmed: bool,
        impact_details: Dict[str, Any],
        custom_message: Optional[str] = None,
    ) -> None:
        """
        Validates if confirmation is granted.
        If config.require_confirmation is enabled and confirmed is False, raises ConfirmationRequiredError.
        """
        if config.require_confirmation and not confirmed:
            logger.info("Confirmation required for action: %s", action_name)
            raise ConfirmationRequiredError(
                action=action_name,
                impact=impact_details,
                message=custom_message,
            )
