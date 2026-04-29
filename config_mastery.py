from __future__ import annotations

from dataclasses import dataclass

STATUS_OK = "OK"
STATUS_PARTIAL = "PARTIAL"
STATUS_FLAG = "FLAG"

# C013 and C014 are manual on-call controls. They are always OK in scoring
# and never appear in the findings list. Their importance weights are retained
# in IMPORTANCE for reference but are excluded from compute_score() via
# SCORING_EXCLUDED.
SCORING_EXCLUDED = {'C013', 'C014'}

# Maximum number of findings rows written to the output sheet.
MAX_FINDINGS = 24

PRIORITY_POINTS = {10: -18, 9: -15, 8: -13, 7: -11, 6: -9, 5: -7, 4: -5, 3: -3, 2: -2, 1: 0}
IMPACT_LABEL = {10: 'Critical', 9: 'High', 8: 'High', 7: 'Medium', 6: 'Medium', 5: 'Medium', 4: 'Low', 3: 'Low', 2: 'Visibility', 1: 'Visibility'}

IMPORTANCE = {
    'C001': 10, 'C002': 8,  'C003': 9,  'C004': 6,  'C005': 8,  'C006': 8,
    'C007': 9,  'C008': 6,  'C009': 5,  'C010': 9,  'C011': 10, 'C012': 6,
    'C013': 7,  'C014': 8,
}

CONTROL_NAMES = {
    'C001': 'Objective Clearly Defined',
    'C002': 'Objective vs Near-Term Alignment',
    'C003': 'Account Challenges Documented',
    'C004': 'Seasonality Awareness',
    'C005': 'Operational Constraints Awareness',
    'C006': 'Client Journey Map',
    'C007': 'Narrative Consistency',
    'C008': 'Sales Concentration Matches Account Story',
    'C009': 'Client Contact Cadence (last 6 months)',
    'C010': 'Customizations Documented & Justified',
    'C011': 'Target Spend / KPI Targets Documented',
    'C012': 'Tagging / Segmentation Logic Clear',
    'C013': 'On-Call Interaction Quality',
    'C014': 'On-Call Explanation Quality',
}

WHY = {
    'C001': 'A clear objective is the starting point for every strategy decision. Without it, the team cannot prioritize correctly or explain trade-offs to the client.',
    'C002': 'The objective context needs KPI targets, efficiency limits, and a timeframe so the CoE can review the account with the right focus. All five elements are required.',
    'C003': 'Knowing the active challenges helps the team avoid repeating mistakes and explains why certain metrics are moving. Generic descriptions do not help the reviewer.',
    'C004': 'Seasonal accounts need a documented plan. If seasonality is not captured, the team may invest too much or too little at the wrong time.',
    'C005': 'When the account has operational constraints, they must be written down so any reviewer understands the limits before making changes.',
    'C006': 'The Client Journey Map shows where the client is in their lifecycle. Without it, strategy decisions may not match the client stage.',
    'C007': 'The efficiency targets set in the project must respect the limits agreed with the client. A looser project target means the account may be run outside the agreed boundaries.',
    'C008': 'A concentrated account needs different attention than a diversified one. The documented story must match what the data actually shows.',
    'C009': 'Regular client contact keeps the account story current. A long gap between calls means the documented goals may no longer reflect what the client actually wants.',
    'C010': 'Framework exceptions must be documented so the CoE can tell if they are intentional. Undocumented customizations look like errors during a review.',
    'C011': 'When a spend target is set, actual daily spend should stay close to it. A large gap means the account is either under-delivering or at risk of overspending.',
    'C012': 'Clear product tagging shows the team understands the portfolio and is managing products at the right priority level.',
    'C013': 'Manual review required during the QR call — interaction quality cannot be checked from system data.',
    'C014': 'Manual review required during the QR call — explanation quality cannot be checked from system data.',
}

SOURCES = {
    'C001': '38_Client_Success_Insights_Repo!AM7',
    'C002': '38_Client_Success_Insights_Repo!AY7',
    'C003': '38_Client_Success_Insights_Repo!BN7',
    'C004': 'Health seasonality logic on AM7 + narrative cross-check in AY7 / BN7 / T7',
    'C005': '38_Client_Success_Insights_Repo!AL7 + narrative cross-check',
    'C006': '39_Client_Journey_Insights_Data!H7',
    'C007': '38_Client_Success_Insights_Repo!O7/AX7 vs 54_Project_Dataset_on_SF!J7/K7',
    'C008': '38_Client_Success_Insights_Repo!AU7 + 07_KPIs_by_Parent_ASIN_by_Month',
    'C009': '37_Gong_Call_Insights_for_Sales!P:P',
    'C010': 'Framework personalization tabs + 54_Project_Dataset_on_SF!T7',
    'C011': '54_Project_Dataset_on_SF + 02_Date_Range_KPIs__Date_Range_',
    'C012': '14_Campaign_Performance_by_Adve!X:AB / Tag1-Tag5',
    'C013': 'Manual review required — QR presentation call',
    'C014': 'Manual review required — QR presentation call',
}


@dataclass(frozen=True)
class ControlResult:
    status: str
    what: str = ''
    why: str = ''
    source: str = ''
