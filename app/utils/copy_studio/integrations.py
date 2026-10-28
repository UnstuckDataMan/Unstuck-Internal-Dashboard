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
    agency = bool(os.environ.get("COPY_STUDIO_AGENCY_ID") or os.environ.get("PULSE_AGENCY_ID"))
    scoped_database = database and agency
    sources = [
        ("smartlead", "Smartlead", bool(os.environ.get("SMARTLEAD_API_KEY")) and scoped_database,
         "Requires the server API credential and scoped database mapping. Only this client's mapped campaigns are read; exact sent versions need review before reuse.", "email"),
        ("manual", "Manual Google Sheets", bool(os.environ.get("GOOGLE_SHEETS_SA_JSON")) and scoped_database,
         "Share campaign sheets with the configured Google service account as Viewer.", "email"),
        ("copy_bank", "Copy Bank", scoped_database,
         "This client's published email and LinkedIn copy is imported for review. LinkedIn examples have no performance claim.", "both"),
        ("read_ai", "Read AI client calls", scoped_database,
         "Reads all mapped client transcripts and uniquely domain-matched unassigned calls from stored webhooks. Full account history is not assumed.", "both"),
    ]
    result = [{"id": key, "name": name, "channel": channel,
               "status": "configured_not_verified" if configured else "not_configured",
               "configured": configured, "live_verified": False, "send": False,
               "read_copy": key in {"copy_bank", "smartlead", "manual"},
               "read_performance": key in {"smartlead", "manual"},
               "reference_review_required": key in {"copy_bank", "smartlead", "manual"},
               "scope": "selected_client", "reason": reason}
              for key, name, configured, reason, channel in sources]
    return result + [MeetAlfredProvider.capabilities()]
