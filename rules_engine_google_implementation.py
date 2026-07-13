"""
rules_engine_google_implementation.py
12 controls: I001–I012
Two blocks: Access & Connectivity (I001–I009), Feed & Data Sources (I010–I012)

Auto-evaluated from Databricks:
  I001 — CLIENT_SUCCESS proxy (SF record presence + key field completeness)
  I002 — ADVERTISER_DETAILS (Tab 33/34) + ACCOUNT_LINKS (Tab 39/40)
  I003 — CAMPAIGN_GOLD (Tab 12/13): conversions vs spend proxy
  I006 — STRIPE_INFO (Tab 14/15): billing status
  I008 — FEED_PRODUCTS (Tab 29/30): availability as disapproval proxy
  I010 — FEED_PRODUCTS (Tab 29/30): LastUpdatedAt freshness proxy
  I011 — FEED_PRODUCTS (Tab 29/30): same freshness, hours granularity
  I012 — STRIPE_INFO (Tab 14/15): IsConnect field

Manual / UI-only (SCORING_EXCLUDED, return PARTIAL with reviewer note):
  I004 — Purchase on Primary Conversion Tag (Google Ads UI)
  I005 — Confirm Conversion Tag on Campaigns (Google Ads UI)
  I007 — Policy Violations = Zero (Google Ads Policy Manager)
  I009 — GMC Feed Duplication Check (GMC UI)
"""
from __future__ import annotations

import re
from datetime import date, timedelta
from typing import Dict, Optional

import pandas as pd

from config import ControlResult, STATUS_OK, STATUS_FLAG, STATUS_PARTIAL
from config_google_implementation import WHY, HOW, SCORING_EXCLUDED
from reader_databricks_google import (
    GoogleContext, get_sheet, find_col, get_active_campaigns,
    to_float, to_str, money_str, _parse_date,
)

QT_PREFIX = re.compile(r'^QT[_\-]', re.IGNORECASE)


def _manual_partial(cid: str) -> ControlResult:
    """Manual controls: always PARTIAL with the HOW verification steps as What We Saw."""
    return ControlResult(
        STATUS_PARTIAL,
        f"Manual review required. {HOW.get(cid, 'Verify in Google Ads UI.')}",
        WHY[cid],
    )


# ── Controls ──────────────────────────────────────────────────────────────────

def _i001(ctx: GoogleContext) -> ControlResult:
    """Salesforce Access Confirmed — proxy: CS record present + key fields populated"""
    df = get_sheet(ctx, "CLIENT_SUCCESS")
    if df.empty:
        return ControlResult(
            STATUS_FLAG,
            "Tab 22 (Client Success) returned no data. Salesforce record may be missing or inaccessible for this account.",
            WHY["I001"],
        )

    row = df.iloc[0]
    advertiser_id = to_str(row.iloc[80]) if len(row) > 80 else ""
    objective     = to_str(row.iloc[39]) if len(row) > 39 else ""
    budget        = to_float(row.iloc[35]) if len(row) > 35 else None

    missing = []
    if not advertiser_id:
        missing.append("AdvertiserId (col 80) missing")
    if not objective:
        missing.append("Primary Objective (col 39) not set")
    if not budget:
        missing.append("Monthly Budget (col 35) missing or zero")

    if not missing:
        return ControlResult(
            STATUS_OK,
            f"Salesforce record found. AdvertiserId = {advertiser_id}. All key fields populated.",
            WHY["I001"],
        )
    elif len(missing) < 3:
        return ControlResult(
            STATUS_PARTIAL,
            f"Salesforce record found but incomplete. Issues: {'; '.join(missing)}.",
            WHY["I001"],
        )
    return ControlResult(
        STATUS_FLAG,
        f"Salesforce record found but critical fields missing: {'; '.join(missing)}.",
        WHY["I001"],
    )


def _i002(ctx: GoogleContext) -> ControlResult:
    """GMC + GA4 Linked to Google Ads — Tab 33/34 + Tab 39/40"""
    df34 = get_sheet(ctx, "ADVERTISER_DETAILS")
    df40 = get_sheet(ctx, "ACCOUNT_LINKS")

    gmc_linked  = None
    ga4_linked  = None
    merchant_id = ""

    if not df34.empty:
        gmc_col = find_col(df34, ["GMC_Linked", "gmc_linked"])
        mid_col = find_col(df34, ["MerchantID", "merchant_id"])
        if gmc_col:
            gmc_linked = str(df34.iloc[0][gmc_col]).lower() in ("true", "1", "yes")
        if mid_col:
            merchant_id = to_str(df34.iloc[0][mid_col])

    if not df40.empty:
        merch_col  = find_col(df40, ["MerchantStatus", "merchant_status"])
        analyt_col = find_col(df40, ["AnalyticsStatus", "analytics_status"])
        analyt_id  = find_col(df40, ["AnalyticsId", "analytics_id"])
        row40 = df40.iloc[0]

        if gmc_linked is None and merch_col:
            merch_val  = to_str(row40[merch_col])
            gmc_linked = merch_val.lower() not in ("false", "0", "no", "nan", "")

        if analyt_col:
            analyt_val = to_str(row40[analyt_col])
            ga4_linked = analyt_val.lower() not in ("false", "0", "no", "nan", "")
        if analyt_id:
            aid = to_str(row40[analyt_id])
            if aid and aid.lower() not in ("nan", "false", "0", ""):
                ga4_linked = True

    if gmc_linked is None and ga4_linked is None:
        return ControlResult(
            STATUS_PARTIAL,
            "Account link data not available in export. Manual check required: Google Ads > Tools > Linked Accounts.",
            WHY["I002"],
        )

    gmc_s = "linked" if gmc_linked else "NOT linked"
    ga4_s = "linked" if ga4_linked else "NOT linked"
    mid_s = f" (Merchant ID: {merchant_id})" if merchant_id else ""

    if gmc_linked and ga4_linked:
        return ControlResult(STATUS_OK, f"GMC = {gmc_s}{mid_s}. GA4 = {ga4_s}.", WHY["I002"])
    elif gmc_linked or ga4_linked:
        return ControlResult(STATUS_PARTIAL, f"GMC = {gmc_s}{mid_s}. GA4 = {ga4_s}. One link missing.", WHY["I002"])
    return ControlResult(STATUS_FLAG, f"GMC = {gmc_s}. GA4 = {ga4_s}. Both disconnected.", WHY["I002"])


def _i003(ctx: GoogleContext) -> ControlResult:
    """Conversion Tag Active — proxy: campaigns with spend have conversion data"""
    df = get_sheet(ctx, "CAMPAIGN_GOLD")
    if df.empty:
        return ControlResult(STATUS_FLAG, "Campaign Gold Metrics (Tab 12/13) not found. Cannot assess conversion tracking.", WHY["I003"])

    cost_col = find_col(df, ["Cost", "cost"])
    conv_col = find_col(df, ["Conversions", "conversions"])
    if not cost_col or not conv_col:
        return ControlResult(STATUS_FLAG, "Cost or Conversions column not found in Tab 13.", WHY["I003"])

    total_spend = df[cost_col].apply(to_float).fillna(0).sum()
    total_conv  = df[conv_col].apply(to_float).fillna(0).sum()

    if total_spend == 0:
        return ControlResult(STATUS_FLAG, "No spend found in campaign data — cannot validate conversion tag.", WHY["I003"])

    if total_conv > 0:
        return ControlResult(
            STATUS_OK,
            f"Conversion data present. {total_conv:.0f} total conversions across {money_str(total_spend)} spend. Proxy CVR = {total_conv/total_spend:.4f}. Manual tag status verification in Google Ads still recommended.",
            WHY["I003"],
        )
    return ControlResult(
        STATUS_FLAG,
        f"Zero conversions recorded across {money_str(total_spend)} spend. Conversion tag is likely broken or not configured.",
        WHY["I003"],
    )


def _i004(ctx): return _manual_partial("I004")
def _i005(ctx): return _manual_partial("I005")


def _i006(ctx: GoogleContext) -> ControlResult:
    """Billing Status Active — Tab 14/15 Stripe & Account Info"""
    df = get_sheet(ctx, "STRIPE_INFO")
    if df.empty:
        return ControlResult(STATUS_FLAG, "Stripe & Account Info tab (Tab 14/15) not found.", WHY["I006"])

    row = df.iloc[0]
    status_val     = to_str(row.iloc[13]) if len(row) > 13 else ""
    invoice_status = to_str(row.iloc[22]) if len(row) > 22 else ""

    billing_active = status_val.lower() == "active"
    invoice_ok     = invoice_status.lower() in ("paid", "active", "")

    if billing_active and invoice_ok:
        return ControlResult(
            STATUS_OK,
            f"Billing status = active. Invoice status = {invoice_status or 'not flagged'}.",
            WHY["I006"],
        )
    elif billing_active:
        return ControlResult(
            STATUS_PARTIAL,
            f"Billing = active. Invoice status = {invoice_status}. Investigate unpaid invoices before they trigger suspension.",
            WHY["I006"],
        )
    return ControlResult(
        STATUS_FLAG,
        f"Billing status = {status_val or 'unknown'}. Invoice status = {invoice_status or 'unknown'}. Account is at risk of immediate campaign suspension.",
        WHY["I006"],
    )


def _i007(ctx): return _manual_partial("I007")


def _i008(ctx: GoogleContext) -> ControlResult:
    """Product Disapproval Rate < 10% — proxy via feed availability"""
    df = get_sheet(ctx, "FEED_PRODUCTS")
    if df.empty:
        return ControlResult(
            STATUS_PARTIAL,
            "Feed Products tab (Tab 29/30) not found. Direct disapproval count requires manual check in GMC > Overview > Products Dashboard.",
            WHY["I008"],
        )

    avail_col = find_col(df, ["Availability", "availability"])
    if not avail_col:
        return ControlResult(
            STATUS_PARTIAL,
            "Availability column not found in feed. Manual GMC Products Dashboard check required.",
            WHY["I008"],
        )

    total = len(df)
    if total == 0:
        return ControlResult(STATUS_FLAG, "No products found in feed.", WHY["I008"])

    # Use out-of-stock + preorder as a proxy for non-available/disapproved inventory
    unavailable_mask = df[avail_col].astype(str).str.lower().str.contains(
        "out of stock|preorder|pre.order", na=False
    )
    unavailable_count = unavailable_mask.sum()
    pct = unavailable_count / total

    if pct < 0.10:
        return ControlResult(
            STATUS_OK,
            f"{unavailable_count} of {total} feed products ({pct*100:.1f}%) are unavailable/out-of-stock. Below 10% threshold. Note: this is a feed proxy — actual GMC disapproval rate requires manual verification.",
            WHY["I008"],
        )
    return ControlResult(
        STATUS_FLAG,
        f"{unavailable_count} of {total} feed products ({pct*100:.1f}%) are unavailable/out-of-stock — above 10% threshold. This is a feed proxy; confirm actual disapproval rate in GMC > Overview > Products Dashboard.",
        WHY["I008"],
    )


def _i009(ctx): return _manual_partial("I009")


def _i010(ctx: GoogleContext) -> ControlResult:
    """Quartile Portal Feed Active — proxy via feed freshness"""
    df = get_sheet(ctx, "FEED_PRODUCTS")
    if df.empty:
        return ControlResult(
            STATUS_PARTIAL,
            "Feed Products tab not found. Manual check required: QT Portal > Google Channel > Feed Export > Last Completed.",
            WHY["I010"],
        )

    # Try to find a last-updated date column
    date_col = find_col(df, [
        "LastUpdatedAt", "last_updated_at", "UpdatedAt", "updated_at",
        "LastModifiedDate", "ModifiedDate", "FeedDate", "feed_date",
    ])

    total = len(df)
    if total == 0:
        return ControlResult(STATUS_FLAG, "No products in feed. Portal feed may be inactive or broken.", WHY["I010"])

    if date_col is None:
        # No date column — use row count as a presence signal only
        return ControlResult(
            STATUS_OK,
            f"Feed Products tab present with {total} products. No LastUpdatedAt column found — exact feed freshness requires manual check in QT Portal > Google Channel > Feed Export.",
            WHY["I010"],
        )

    ref = ctx.window_end or date.today()
    dates = [_parse_date(v) for v in df[date_col] if not pd.isna(v)]
    dates = [d for d in dates if d is not None]

    if not dates:
        return ControlResult(
            STATUS_PARTIAL,
            "Feed present but no valid dates found in LastUpdatedAt column. Manual freshness check required.",
            WHY["I010"],
        )

    latest = max(dates)
    days_stale = (ref - latest).days

    if days_stale <= 2:
        return ControlResult(
            STATUS_OK,
            f"Feed last updated {latest} ({days_stale} day(s) ago). {total} products. Feed is active and fresh.",
            WHY["I010"],
        )
    elif days_stale <= 7:
        return ControlResult(
            STATUS_PARTIAL,
            f"Feed last updated {latest} ({days_stale} days ago). {total} products. Feed may be stale — check QT Portal export schedule.",
            WHY["I010"],
        )
    return ControlResult(
        STATUS_FLAG,
        f"Feed last updated {latest} ({days_stale} days ago). {total} products. Feed is stale — DPL labels and product segmentation are out of date.",
        WHY["I010"],
    )


def _i011(ctx: GoogleContext) -> ControlResult:
    """Quartile Portal Last Update — same freshness check, hours granularity"""
    df = get_sheet(ctx, "FEED_PRODUCTS")
    if df.empty:
        return ControlResult(
            STATUS_PARTIAL,
            "Feed Products tab not found. Manual check required: QT Portal > Google Channel > Feed Export > Last Completed timestamp.",
            WHY["I011"],
        )

    # Try datetime columns for hour-level granularity
    date_col = find_col(df, [
        "LastUpdatedAt", "last_updated_at", "UpdatedAt", "updated_at",
        "LastModifiedDate", "ModifiedDate",
    ])

    total = len(df)

    if date_col is None:
        return ControlResult(
            STATUS_OK,
            f"Feed Products tab present with {total} products. Hour-level timestamp not available in export — verify daily refresh in QT Portal.",
            WHY["I011"],
        )

    ref = ctx.window_end or date.today()
    dates = [_parse_date(v) for v in df[date_col] if not pd.isna(v)]
    dates = [d for d in dates if d is not None]

    if not dates:
        return ControlResult(
            STATUS_PARTIAL,
            "No valid dates in LastUpdatedAt. Manual Portal feed timestamp check required.",
            WHY["I011"],
        )

    latest = max(dates)
    days_stale = (ref - latest).days

    # Daily refresh is required per the HOW note
    if days_stale == 0:
        return ControlResult(STATUS_OK, f"Feed last updated today ({latest}). Daily refresh confirmed.", WHY["I011"])
    elif days_stale == 1:
        return ControlResult(STATUS_OK, f"Feed last updated {latest} (1 day ago). Within daily refresh tolerance.", WHY["I011"])
    return ControlResult(
        STATUS_FLAG,
        f"Feed last updated {latest} ({days_stale} days ago). Daily refresh requirement not met — DPL labels and segmentation are degrading.",
        WHY["I011"],
    )


def _i012(ctx: GoogleContext) -> ControlResult:
    """Shopify / E-commerce Platform Connected — Tab 14/15 IsConnect"""
    df = get_sheet(ctx, "STRIPE_INFO")
    if df.empty:
        return ControlResult(
            STATUS_PARTIAL,
            "Stripe & Account Info tab not found. Manual platform connection check required: QT Portal > Settings > Connected Channels.",
            WHY["I012"],
        )

    conn_col = find_col(df, ["IsConnect", "is_connect"])
    if not conn_col:
        return ControlResult(
            STATUS_PARTIAL,
            "IsConnect field not found in Tab 15. Manual check required: QT Portal > Settings > Connected Channels.",
            WHY["I012"],
        )

    connected = str(df.iloc[0][conn_col]).lower() in ("true", "1", "yes")
    if connected:
        return ControlResult(
            STATUS_OK,
            "Platform connection (IsConnect) = true. Total sales channel linked — TACoS and full-funnel reporting available.",
            WHY["I012"],
        )
    return ControlResult(
        STATUS_FLAG,
        "Platform connection (IsConnect) = false. Total sales cannot be tracked — TACoS, organic sales, and full-funnel reporting are all blind.",
        WHY["I012"],
    )


# ── Orchestrator ──────────────────────────────────────────────────────────────

_EVALUATORS = {
    "I001": _i001, "I002": _i002, "I003": _i003, "I004": _i004,
    "I005": _i005, "I006": _i006, "I007": _i007, "I008": _i008,
    "I009": _i009, "I010": _i010, "I011": _i011, "I012": _i012,
}


def evaluate_all_implementation(ctx: GoogleContext) -> Dict[str, ControlResult]:
    results = {}
    for cid, fn in _EVALUATORS.items():
        try:
            results[cid] = fn(ctx)
        except Exception as e:
            results[cid] = ControlResult(
                STATUS_FLAG,
                f"Evaluation error: {e}",
                "Internal error — review this control manually.",
            )
    return results
