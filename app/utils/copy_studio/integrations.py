"""Copy Studio provider capabilities. No connector sends outreach."""
from __future__ import annotations

import os


class IntegrationUnavailable(RuntimeError):
    """A source has not been configured or is unavailable (safe for the UI)."""


class MeetAlfredProvider:
    """Reserved boundary, deliberately disconnected until a real API exists.

    Do not reuse Pulse's historic Meet Alfred importer here: this release's
    LinkedIn references come exclusively from reviewed Copy Bank content.
    """

    @staticmethod
    def capabilities() -> dict:
        return {
            "id": "meet_alfred", "name": "Meet Alfred", "channel": "linkedin",
            "status": "disabled", "configured": False, "read_copy": False,
            "read_performance": False, "send": False,
            "reason": "Future integration. No API connection is implemented or enabled.",
        }

    def sync(self, *args, **kwargs):
        raise IntegrationUnavailable("Meet Alfred is disabled. Use LinkedIn Copy Bank examples.")


def capabilities() -> list[dict]:
    database = bool(os.environ.get("SUPABASE_URL") and os.environ.get("SUPABASE_SERVICE_ROLE_KEY"))
    sources = [
        ("smartlead", "Smartlead", bool(os.environ.get("SMARTLEAD_API_KEY")),
         "Server API credential required; mapped campaigns only.", "email"),
        ("manual", "Manual Google Sheets", bool(os.environ.get("GOOGLE_SHEETS_SA_JSON")) and database,
         "Share campaign sheets with the configured Google service account as Viewer.", "email"),
        ("copy_bank", "Copy Bank", database,
         "Imports are reviewed before reuse. LinkedIn examples have no performance claim.", "both"),
        ("read_ai", "Read AI onboarding", bool(os.environ.get("SUPABASE_URL") and os.environ.get("SUPABASE_SERVICE_ROLE_KEY")),
         "Reads complete stored webhook transcripts. Calls missing from the dashboard need the existing Read AI ingestion.", "both"),
    ]
    result = [{"id": key, "name": name, "channel": channel,
               "status": "configured_not_verified" if configured else "not_configured",
               "configured": configured, "send": False, "reason": reason}
              for key, name, configured, reason, channel in sources]
    return result + [MeetAlfredProvider.capabilities()]
