"""Minimal shared module - No caching, direct Outlook operations only."""

from .logging_config import get_logger

logger = get_logger(__name__)

# Note: This module previously contained email caching functionality.
# All caching has been removed. The system now works directly with Outlook.