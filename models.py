"""Shared constants used by the database, Excel sync, and UI layers."""

STATUS_NOT_READY = "Not Ready"
STATUS_READY_PREP = "Ready Prep"
STATUS_IN_PREP = "In Prep"
STATUS_READY_QA = "Ready QA"
STATUS_IN_QA = "In QA"
STATUS_COMPLETED = "Completed"

SECTION_DEFAULT = "Unsectioned"

LANE_READY_PREP = "Ready to Claim"
LANE_IN_PREP = "Currently Being Prepped"
LANE_READY_QA = "Ready for QA"
LANE_IN_QA = "Currently Being QA"
LANE_COMPLETED = "Completed"

LANE_STATUSES = {
    LANE_READY_PREP: STATUS_READY_PREP,
    LANE_IN_PREP: STATUS_IN_PREP,
    LANE_READY_QA: STATUS_READY_QA,
    LANE_IN_QA: STATUS_IN_QA,
    LANE_COMPLETED: STATUS_COMPLETED,
}
