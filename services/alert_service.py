from __future__ import annotations

from config.settings import ALERT_RULES


def build_alerts(action_counts: dict[str, int] | None) -> list[dict[str, str]]:
    action_counts = action_counts or {}

    overdue = int(action_counts.get("OVERDUE / FOLLOW UP", 0))
    returned = int(action_counts.get("RETURNED BY NV5 / NEED RESUBMIT", 0))
    need_update = int(action_counts.get("UPDATE TRACKING TO CLOSED", 0))

    alerts: list[dict[str, str]] = []

    if overdue > int(ALERT_RULES["overdue_threshold"]):
        alerts.append(
            {
                "level": "critical",
                "title": "Overdue documents detected",
                "message": f"{overdue} overdue item(s) require urgent follow-up.",
            }
        )

    if returned > int(ALERT_RULES["returned_threshold"]):
        alerts.append(
            {
                "level": "warning",
                "title": "Returned documents exceed threshold",
                "message": f"{returned} returned item(s) may require resubmission planning.",
            }
        )

    if need_update > int(ALERT_RULES["need_update_threshold"]):
        alerts.append(
            {
                "level": "info",
                "title": "Tracking update required",
                "message": f"{need_update} item(s) may need status update to closed.",
            }
        )

    if not alerts:
        alerts.append(
            {
                "level": "success",
                "title": "No major alert",
                "message": "No urgent overdue or returned document risk detected.",
            }
        )

    return alerts
