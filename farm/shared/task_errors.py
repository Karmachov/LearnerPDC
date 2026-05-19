"""
Celery task errors with user-safe messages (no tracebacks in the result backend).
"""


class ReportTaskError(Exception):
    """Raised when report generation fails; message is safe to show in the UI."""

    def __init__(self, user_message: str):
        self.user_message = user_message
        super().__init__(user_message)


def user_facing_task_error(exc: BaseException) -> str:
    """Map internal exceptions to short, safe messages for clients."""
    if isinstance(exc, ReportTaskError):
        return exc.user_message

    msg = str(exc).strip()
    if not msg:
        return "Report generation failed. Please check your Excel file and try again."

    lowered = msg.lower()
    if "faculty not found" in lowered:
        return "Your account could not be found. Please sign in again."
    if "decryption failed" in lowered or "master_key" in lowered:
        return "Signing credentials could not be decrypted. Re-upload your key in Profile Settings."
    if "soft time limit" in lowered or "time limit" in lowered:
        return "Report generation timed out. Try a smaller file or fewer students."
    if "excel" in lowered or "spreadsheet" in lowered or "sheet" in lowered:
        return f"Excel processing error: {msg[:200]}"
    if "no output path" in lowered:
        return "No students matched your filters. Check semester and learner type settings."
    if "digital signing" in lowered or "signing failed" in lowered:
        return msg[:300]
    if "pdf conversion" in lowered or "libreoffice" in lowered:
        return msg[:300]

    return msg[:300]
