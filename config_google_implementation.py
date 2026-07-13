"""
config_google_implementation.py
12 controls: I001-I012
Two blocks: Access & Connectivity (I001-I008), Feed & Data Sources (I009-I012)

Auto-evaluated from Databricks:
  I001 — CLIENT_SUCCESS proxy (SF record presence + field completeness)
  I002 — ADVERTISER_DETAILS (Tab 34) + ACCOUNT_LINKS (Tab 40)
  I003 — CAMPAIGN_GOLD (Tab 13) conversions vs spend proxy
  I006 — STRIPE_INFO (Tab 15) billing status
  I008 — FEED_PRODUCTS (Tab 30) availability as disapproval proxy
  I010 — FEED_PRODUCTS (Tab 30) LastUpdatedAt freshness proxy
  I011 — FEED_PRODUCTS (Tab 30) same freshness proxy (hours granularity)
  I012 — STRIPE_INFO (Tab 15) IsConnect field

Manual / UI-only (SCORING_EXCLUDED, return PARTIAL with reviewer note):
  I004 — Purchase on Primary Conversion Tag (Google Ads UI)
  I005 — Confirm Conversion Tag on Campaigns (Google Ads UI)
  I007 — Policy Violations = Zero (Google Ads Policy Manager)
  I009 — GMC Feed Duplication Check (GMC UI)
"""
from __future__ import annotations
from config import STATUS_OK, STATUS_FLAG, STATUS_PARTIAL, ControlResult

PILLAR = "google_implementation"

# Manual UI-only controls — excluded from score, return PARTIAL with reviewer note
SCORING_EXCLUDED = {"I004", "I005", "I007", "I009"}

MAX_FINDINGS = 12

# New scoring scale from updated spreadsheet
# Importance → Priority (points lost on FLAG; PARTIAL = 50%)
PRIORITY_POINTS = {
    10: -30,
    9:  -27,
    8:  -24,
    7:  -21,
    6:  -18,
    5:  -15,
    4:  -12,
    3:   -9,
    2:   -6,
    1:    0,
}

IMPACT_LABEL = {
    10: "Critical", 9: "High", 8: "High", 7: "Medium", 6: "Medium",
    5:  "Medium",   4: "Low",  3: "Low",  2: "Visibility", 1: "Visibility",
}

# Importance values pulled from the (U) tab
IMPORTANCE = {
    "I001": 5,   # Medium / -15
    "I002": 8,   # High / -24
    "I003": 10,  # Critical / -30
    "I004": 5,   # Medium / -15 (manual, scoring excluded)
    "I005": 5,   # Medium / -15 (manual, scoring excluded)
    "I006": 5,   # Medium / -15  (NOTE: spreadsheet shows -15 not -30 — using col 13 value)
    "I007": 10,  # Critical / -30 (manual, scoring excluded)
    "I008": 10,  # Critical / -30
    "I009": 7,   # Medium / -21 (manual, scoring excluded)
    "I010": 5,   # Medium / -15
    "I011": 5,   # Medium / -15 (no priority in sheet — treating as 5)
    "I012": 5,   # Visibility / 0 in sheet but we treat as 5 for scoring
}

CONTROL_NAMES = {
    "I001": "Salesforce Access Confirmed",
    "I002": "GMC + GA4 Linked to Google Ads",
    "I003": "Conversion Tag Active",
    "I004": "Purchase on Primary Conversion Tag",
    "I005": "Confirm Conversion Tag on Campaigns",
    "I006": "Billing Status Active",
    "I007": "Policy Violations = Zero",
    "I008": "Product Disapproval Rate < 10%",
    "I009": "GMC Feed Duplication Check",
    "I010": "Quartile Portal Feed Active",
    "I011": "Quartile Portal Last Update",
    "I012": "Shopify / E-commerce Platform Connected",
}

# WHY = exact text from the (U) tab "Why It Matters" column
WHY = {
    "I001": "Without Salesforce access, the agent cannot read targets, constraints, or CS context — all pillar evaluations degrade.",
    "I002": "Disconnected GMC pauses all Shopping/PMAX product ads. Missing GA4 removes audience targeting and total-sales visibility.",
    "I003": "A broken conversion tag disables Smart Bidding signal entirely. This is the most critical technical control in the account.",
    "I004": "Dual primary tags cause Smart Bidding to optimize toward low-intent actions, inflating conversion volume while ROAS degrades silently.",
    "I005": "Smart Bidding at campaign level optimizes for the specified goal. Wrong goal = wrong optimization = wasted spend.",
    "I006": "A billing failure pauses the entire account immediately — zero impressions, zero revenue until resolved.",
    "I007": "Unresolved policy violations restrict ad serving and, in severe cases, result in full account suspension.",
    "I008": "High product disapproval rates directly reduce the addressable inventory for Shopping and PMAX campaigns.",
    "I009": "Feed duplication causes GMC attribute conflicts that can lead to mass disapprovals and inconsistent product data.",
    "I010": "A stale feed means product data (titles, labels, availability) in Google is outdated. Performance degradation is silent but compounding.",
    "I011": "Daily feed refresh is required for DPL labels and product segmentation to reflect current account state.",
    "I012": "A disconnected platform means total sales cannot be tracked — TACoS, organic sales, and full-funnel reporting are all blind.",
}

# HOW = verification steps from the (U) tab "How" column (used in What We Saw when manual)
HOW = {
    "I001": "Salesforce > Account > Project. Confirm access and that Profile ID, Budget, and ROAS Goal fields are populated.",
    "I002": "Google Ads > Tools > Linked Accounts. Verify both GMC and GA4 connections show as Active.",
    "I003": "Google Ads > Tools > Conversions. Confirm status = Active, last conversion date is recent, no 'Inactive' or 'Unverified' tags.",
    "I004": "Google Ads > Tools > Conversions. Check 'Primary action' column — only Purchase type should be Primary. All others must be set to Secondary.",
    "I005": "Google Ads > Campaigns > Settings > Conversion goals. Verify each PMAX and Search campaign is set to Purchase.",
    "I006": "Tab 15_Stripe_and_Account_Info: status field (col 13). Also verify in Google Ads > Billing > Summary.",
    "I007": "Google Ads > Policy Manager. Confirm zero active violations. Flag any 'Limited' or 'Disapproved' account-level policies.",
    "I008": "GMC > Overview > Products Dashboard. Count approved vs disapproved. Also check product_v2 Availability field.",
    "I009": "GMC > Data Sources. Confirm only one primary feed source is active. Flag duplicate or conflicting supplemental feeds.",
    "I010": "QT Portal > Google Channel > Feed Export > Last Completed. FLAG if last successful export is more than 48 hours ago.",
    "I011": "QT Portal > Google Channel > Feed Export > Last Completed. Check timestamp vs current date.",
    "I012": "Tab 15_Stripe_and_Account_Info: IsConnect col 9. Also check QT Portal > Settings > Connected Channels.",
}
