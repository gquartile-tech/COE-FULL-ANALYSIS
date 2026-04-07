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
    'C001': 'Clear objective framing is the anchor for judging account mastery and whether the account is being managed against the right business outcome.',
    'C002': 'The primary objective should include KPI logic, context, and constraints so the intended strategy is analyzable.',
    'C003': 'Documented challenges show the CSM understands the real blockers affecting account performance.',
    'C004': 'Seasonality should be documented only when it exists, and omitted when it does not.',
    'C005': 'When the YES/NO toggle is used, the related context should also be visible in the narrative fields for the reviewer.',
    'C006': 'A client journey map is a basic signal that the account setup and customer stage are being tracked.',
    'C007': 'Project-level KPI targets should not be looser than the declared account constraints.',
    'C008': 'Understanding concentration is important because highly concentrated accounts need different prioritization and risk management.',
    'C009': 'Recent client contact cadence is evidence that the account narrative is current and client-informed.',
    'C010': 'Framework exceptions should be visible in CS notes so the CoE can tell whether they are intentional and justified.',
    'C011': 'When targets are documented, they should be reasonably aligned with current KPI reality.',
    'C012': 'Product tagging and segmentation reveal whether the account structure supports the intended strategy.',
    'C013': 'Manual review required during QR presentation call — interaction quality cannot be assessed from system data.',
    'C014': 'Manual review required during QR presentation call — explanation quality cannot be assessed from system data.',
}

SOURCES = {
    'C001': '38_Client_Success_Insights_Repo!AY7',
    'C002': '38_Client_Success_Insights_Repo!AM7',
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
