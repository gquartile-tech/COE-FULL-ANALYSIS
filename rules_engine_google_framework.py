"""
rules_engine_google_framework.py
22 controls: F001–F022
Binary methodology: OK or FLAG only — no PARTIAL statuses.
Manual controls always return OK with a reviewer note.
All auto controls pull from the same tab sources as the old engine,
using the shared reader helpers (get_active_campaigns, get_sheet, find_col).
"""
from __future__ import annotations

import re
from typing import Dict, Optional

import pandas as pd

from config import ControlResult, STATUS_OK, STATUS_FLAG
from config_google_framework import WHY, WYSD, VALID_CAMPAIGN_TYPES, VALID_STRATEGY_TAGS, PROMO_KEYWORDS
from reader_databricks_google import (
    GoogleContext, get_sheet, find_col, get_active_campaigns,
    to_float, to_str, money_str, pct_str,
)

QT_PREFIX = re.compile(r'^QT[_\-]', re.IGNORECASE)


# ── Shared helpers ────────────────────────────────────────────────────────────

def _spend_by_campaign(ctx: GoogleContext) -> dict:
    """Return {CampaignId: total_cost} from Campaign Gold Metrics (Tab 13)."""
    df = get_sheet(ctx, "CAMPAIGN_GOLD")
    if df.empty:
        return {}
    cid  = find_col(df, ["CampaignId", "campaign_id"])
    cost = find_col(df, ["Cost", "cost"])
    if not cid or not cost:
        return {}
    return df.groupby(cid)[cost].sum().to_dict()


def _manual_ok(cid: str) -> ControlResult:
    return ControlResult(
        STATUS_OK,
        f"Manual review required. {HOW_NOTES.get(cid, 'Verify in Google Ads UI.')}",
        WHY[cid],
        WYSD.get(cid, ""),
    )


# Manual reviewer notes — concise instructions for each manual control
HOW_NOTES = {
    "F004": "Open each active Search campaign in Google Ads > Settings > Networks. Confirm 'Display Network' is unchecked.",
    "F006": "Open each active PMAX campaign > Settings. Verify auto-created assets, Final URL expansion, and store goals auto-apply are all disabled.",
    "F008": "Log into QT Portal > Google > Branded Terms. Confirm all trademark terms are uploaded and current.",
    "F013": "Google Ads > Assets > Sitelinks. Confirm minimum 4 approved sitelinks at account level.",
    "F014": "Google Ads > Assets > Structured Snippets. Confirm at least one active structured snippet per relevant category.",
    "F015": "Google Ads > Assets > Callouts. Confirm callout extensions are active and cover key USPs.",
    "F016": "Google Ads > Assets > Business Name. Confirm approved business name asset is present at account level.",
    "F019": "QT Portal > Google Channel > Keyword Expander. Confirm Expander is enabled for all active Search campaigns.",
    "F020": "Google Ads > Asset Groups > Audiences. Confirm minimum 1 custom intent segment and 1 remarketing list per active PMAX asset group.",
    "F021": "Google Ads > Asset Groups > Search Themes. Confirm minimum 10 Search Themes per active PMAX asset group.",
}


# ── Naming helpers ────────────────────────────────────────────────────────────

def _parse_name_tokens(name: str):
    """
    Split a QT campaign name into components.
    Expected format: QT_{CampaignType}_{StrategyTag}_{...descriptor}
    Returns (prefix_ok, type_token, strategy_tag) all lowercased.
    """
    parts = re.split(r'[_\-]', name, maxsplit=3)
    prefix      = parts[0].lower() if len(parts) > 0 else ""
    type_token  = parts[1].lower() if len(parts) > 1 else ""
    strategy    = parts[2].lower() if len(parts) > 2 else ""
    return prefix, type_token, strategy


# ── Controls ──────────────────────────────────────────────────────────────────

def _f001(ctx: GoogleContext) -> ControlResult:
    """Naming Convention — QT Prefix"""
    active = get_active_campaigns(ctx)
    if active.empty:
        return ControlResult(STATUS_FLAG, "No active campaign data found.", WHY["F001"], WYSD["F001"])

    name_col = find_col(active, ["CampaignName", "campaign_name"])
    cid_col  = find_col(active, ["CampaignId", "campaign_id"])
    if not name_col:
        return ControlResult(STATUS_FLAG, "CampaignName column not found.", WHY["F001"], WYSD["F001"])

    spend_map = _spend_by_campaign(ctx)
    non_qt = []
    for _, row in active.iterrows():
        name  = to_str(row[name_col])
        cid   = to_str(row[cid_col]) if cid_col else ""
        spend = spend_map.get(cid, 0) or 0
        if spend > 0 and not QT_PREFIX.match(name):
            non_qt.append(name)

    total = len(active)
    if not non_qt:
        return ControlResult(
            STATUS_OK,
            f"All {total} active campaigns with spend follow QT_ naming.",
            WHY["F001"], WYSD["F001"],
        )

    sample = ", ".join(f"'{n}'" for n in non_qt[:3])
    return ControlResult(
        STATUS_FLAG,
        f"{len(non_qt)} of {total} active campaigns with spend don't follow QT_ naming. Examples: {sample}.",
        WHY["F001"], WYSD["F001"],
    )


def _f002(ctx: GoogleContext) -> ControlResult:
    """Naming Convention — Campaign Type Token"""
    active = get_active_campaigns(ctx)
    if active.empty:
        return ControlResult(STATUS_FLAG, "No active campaign data found.", WHY["F002"], WYSD["F002"])

    name_col = find_col(active, ["CampaignName", "campaign_name"])
    cid_col  = find_col(active, ["CampaignId", "campaign_id"])
    if not name_col:
        return ControlResult(STATUS_FLAG, "CampaignName not found.", WHY["F002"], WYSD["F002"])

    spend_map = _spend_by_campaign(ctx)
    invalid = []
    for _, row in active.iterrows():
        name  = to_str(row[name_col])
        cid   = to_str(row[cid_col]) if cid_col else ""
        spend = spend_map.get(cid, 0) or 0
        if spend > 0 and QT_PREFIX.match(name):
            _, type_token, _ = _parse_name_tokens(name)
            if type_token not in VALID_CAMPAIGN_TYPES:
                invalid.append(f"'{name}' (token: '{type_token}')")

    total_qt = sum(
        1 for _, row in active.iterrows()
        if QT_PREFIX.match(to_str(row[name_col]))
        and (spend_map.get(to_str(row[cid_col]) if cid_col else "", 0) or 0) > 0
    )

    if not invalid:
        return ControlResult(
            STATUS_OK,
            f"All {total_qt} QT_ campaigns with spend have a valid campaign type token.",
            WHY["F002"], WYSD["F002"],
        )
    sample = ", ".join(invalid[:3])
    return ControlResult(
        STATUS_FLAG,
        f"{len(invalid)} QT_ campaign(s) with an invalid or missing type token. Examples: {sample}.",
        WHY["F002"], WYSD["F002"],
    )


def _f003(ctx: GoogleContext) -> ControlResult:
    """Naming Convention — Strategy Tag"""
    active = get_active_campaigns(ctx)
    if active.empty:
        return ControlResult(STATUS_FLAG, "No active campaign data found.", WHY["F003"], WYSD["F003"])

    name_col = find_col(active, ["CampaignName", "campaign_name"])
    cid_col  = find_col(active, ["CampaignId", "campaign_id"])
    if not name_col:
        return ControlResult(STATUS_FLAG, "CampaignName not found.", WHY["F003"], WYSD["F003"])

    spend_map = _spend_by_campaign(ctx)
    invalid = []
    for _, row in active.iterrows():
        name  = to_str(row[name_col])
        cid   = to_str(row[cid_col]) if cid_col else ""
        spend = spend_map.get(cid, 0) or 0
        if spend > 0 and QT_PREFIX.match(name):
            _, _, strategy = _parse_name_tokens(name)
            if strategy not in VALID_STRATEGY_TAGS:
                invalid.append(f"'{name}' (tag: '{strategy}')")

    if not invalid:
        return ControlResult(
            STATUS_OK,
            "All QT_ campaigns with spend have a valid strategy tag in position 3.",
            WHY["F003"], WYSD["F003"],
        )
    sample = ", ".join(invalid[:3])
    return ControlResult(
        STATUS_FLAG,
        f"{len(invalid)} campaign(s) with an invalid or missing strategy tag. Examples: {sample}.",
        WHY["F003"], WYSD["F003"],
    )


def _f004(ctx): return _manual_ok("F004")


def _f005(ctx: GoogleContext) -> ControlResult:
    """Promotion End Dates"""
    active = get_active_campaigns(ctx)
    if active.empty:
        return ControlResult(STATUS_FLAG, "No active campaign data found.", WHY["F005"], WYSD["F005"])

    name_col  = find_col(active, ["CampaignName", "campaign_name"])
    start_col = find_col(active, ["StartDate", "start_date"])
    if not name_col or not start_col:
        return ControlResult(STATUS_OK, "CampaignName or StartDate not available — manual promo check required.", WHY["F005"], WYSD["F005"])

    ref = ctx.window_end or pd.Timestamp.now().date()
    flagged = []
    for _, row in active.iterrows():
        name = to_str(row[name_col])
        if not PROMO_KEYWORDS.search(name):
            continue
        sd = pd.to_datetime(str(row[start_col]), errors="coerce")
        if not pd.isna(sd) and (pd.Timestamp(ref) - sd).days > 60:
            flagged.append(f"'{name}'")

    if not flagged:
        return ControlResult(STATUS_OK, "No active campaigns with promotional naming older than 60 days found.", WHY["F005"], WYSD["F005"])
    return ControlResult(
        STATUS_FLAG,
        f"{len(flagged)} active campaign(s) with promo naming and start date > 60 days ago: {', '.join(flagged[:3])}.",
        WHY["F005"], WYSD["F005"],
    )


def _f006(ctx): return _manual_ok("F006")


def _f007(ctx: GoogleContext) -> ControlResult:
    """Match Type Governance — BROAD Dominance"""
    df = get_sheet(ctx, "KEYWORD_REPORT")
    if df.empty:
        return ControlResult(STATUS_FLAG, "Keyword Report (Tab 09) not found or empty.", WHY["F007"], WYSD["F007"])

    mt_col = find_col(df, ["MatchType", "match_type"])
    if not mt_col:
        return ControlResult(STATUS_FLAG, "MatchType column not found in Tab 09.", WHY["F007"], WYSD["F007"])

    counts = df[mt_col].astype(str).str.upper().value_counts()
    total = counts.sum()
    if total == 0:
        return ControlResult(STATUS_FLAG, "No keywords found in Tab 09.", WHY["F007"], WYSD["F007"])

    broad_pct = counts.get("BROAD", 0) / total
    exact_pct = counts.get("EXACT", 0) / total

    dist = ", ".join(f"{k}: {v/total*100:.1f}%" for k, v in counts.items())

    if broad_pct > 0.80:
        return ControlResult(
            STATUS_FLAG,
            f"BROAD dominance = {broad_pct*100:.1f}% of {total} keywords. {dist}. >80% BROAD is a governance risk.",
            WHY["F007"], WYSD["F007"],
        )
    return ControlResult(
        STATUS_OK,
        f"BROAD share = {broad_pct*100:.1f}%. EXACT share = {exact_pct*100:.1f}%. Match type mix within governance threshold. {dist}.",
        WHY["F007"], WYSD["F007"],
    )


def _f008(ctx): return _manual_ok("F008")


def _f009(ctx: GoogleContext) -> ControlResult:
    """Branded Search Campaign Active"""
    active  = get_active_campaigns(ctx)
    df_gold = get_sheet(ctx, "CAMPAIGN_GOLD")
    df      = active if not active.empty else df_gold
    if df.empty:
        return ControlResult(STATUS_FLAG, "No campaign data found.", WHY["F009"], WYSD["F009"])

    spend_map = _spend_by_campaign(ctx)
    name_col  = find_col(df, ["CampaignName", "campaign_name"])
    ch_col    = find_col(df, ["AdvertisingChannelType"])
    cid_col   = find_col(df, ["CampaignId", "campaign_id"])

    branded = []
    for _, row in df.iterrows():
        name  = to_str(row[name_col]) if name_col else ""
        ch    = to_str(row[ch_col]).upper() if ch_col else ""
        cid   = to_str(row[cid_col]) if cid_col else ""
        tokens = re.split(r'[_\-\s]+', name.upper())
        is_branded = any(t in ("TM", "SKW") or t.startswith("BRAND") for t in tokens)
        if "SEARCH" in ch and is_branded:
            spend = spend_map.get(cid, 0) or 0
            branded.append((name, spend))

    if branded:
        best = max(branded, key=lambda x: x[1])
        return ControlResult(
            STATUS_OK,
            f"Branded Search campaign active: '{best[0]}'. Spend in window: {money_str(best[1])}.",
            WHY["F009"], WYSD["F009"],
        )
    return ControlResult(STATUS_FLAG, "No branded/TM Search campaign found with spend in the window.", WHY["F009"], WYSD["F009"])


def _f010(ctx: GoogleContext) -> ControlResult:
    """Search Term Waste"""
    df11 = get_sheet(ctx, "SEARCH_TERMS")
    df02 = get_sheet(ctx, "DATE_RANGE_KPIS")
    if df11.empty:
        return ControlResult(STATUS_FLAG, "Search Terms Report (Tab 10/11) not found.", WHY["F010"], WYSD["F010"])

    term_col = find_col(df11, ["SearchTerm", "search_term"])
    cost_col = find_col(df11, ["Cost", "cost"])
    conv_col = find_col(df11, ["Conversions", "conversions"])
    if not term_col or not cost_col:
        return ControlResult(STATUS_FLAG, "SearchTerm or Cost column not found.", WHY["F010"], WYSD["F010"])

    avg_cpc = 0.35
    if not df02.empty:
        cpc_col = find_col(df02, ["CPC", "cpc"])
        if cpc_col:
            avg_cpc = to_float(df02.iloc[0][cpc_col]) or 0.35

    threshold = avg_cpc * 3
    waste, total_spend = [], 0.0

    for _, row in df11.iterrows():
        cost = to_float(row[cost_col]) or 0
        total_spend += cost
        conv = to_float(row[conv_col]) if conv_col else 0
        if cost > threshold and (conv is None or conv == 0):
            waste.append((to_str(row[term_col]), cost))

    waste_spend = sum(w[1] for w in waste)
    pct = waste_spend / total_spend if total_spend > 0 else 0

    if pct < 0.05:
        return ControlResult(
            STATUS_OK,
            f"{len(waste)} search term(s) with spend > avg CPC × 3 (${threshold:.2f}) and zero conversions. Waste = {money_str(waste_spend)} ({pct*100:.1f}% of search spend). Within threshold.",
            WHY["F010"], WYSD["F010"],
        )
    return ControlResult(
        STATUS_FLAG,
        f"{len(waste)} waste term(s) with spend > ${threshold:.2f} and zero conversions. Waste spend = {money_str(waste_spend)} ({pct*100:.1f}% of total search spend). Top terms: {', '.join(f'{w[0]} (${w[1]:.2f})' for w in sorted(waste, key=lambda x: -x[1])[:3])}.",
        WHY["F010"], WYSD["F010"],
    )


def _f011(ctx: GoogleContext) -> ControlResult:
    """Negative Keyword Coverage"""
    df = get_sheet(ctx, "NEGATIVE_KEYWORDS")
    if df.empty:
        return ControlResult(STATUS_FLAG, "Negative Keywords tab (Tab 30/31) not found or empty. No keyword negatives detected.", WHY["F011"], WYSD["F011"])

    type_col = find_col(df, ["Type"])
    kw_col   = find_col(df, ["Keyword"])

    if type_col:
        df_kw = df[df[type_col].astype(str).str.upper() == "KEYWORD"].copy()
    else:
        df_kw = df.copy()

    if kw_col:
        actual = df_kw[kw_col].dropna()
        actual = actual[actual.astype(str).str.strip() != ""]
        kw_count = len(actual)
    else:
        kw_count = 0

    listing_rows = (
        len(df[df[type_col].astype(str).str.upper() == "LISTING_GROUP"]) if type_col else 0
    )

    if kw_count > 0:
        return ControlResult(
            STATUS_OK,
            f"{kw_count} keyword negative(s) confirmed ({listing_rows} listing group exclusions also present, not counted as keyword negatives).",
            WHY["F011"], WYSD["F011"],
        )
    return ControlResult(
        STATUS_FLAG,
        f"Zero keyword negatives found in export ({listing_rows} listing group exclusions present, not keyword negatives). Shared exclusion lists also require manual verification.",
        WHY["F011"], WYSD["F011"],
    )


def _f012(ctx: GoogleContext) -> ControlResult:
    """Budget Concentration — PMAX Dominance"""
    df = get_sheet(ctx, "CHANNEL_TYPE")
    if df.empty:
        return ControlResult(STATUS_FLAG, "Channel Type tab (Tab 07) not found.", WHY["F012"], WYSD["F012"])

    ch_col  = find_col(df, ["AdvertisingChannelType"])
    spd_col = find_col(df, ["Spend"])
    pct_col = find_col(df, ["Perc_Spend"])
    if not ch_col:
        return ControlResult(STATUS_FLAG, "AdvertisingChannelType not found in Tab 07.", WHY["F012"], WYSD["F012"])

    total = sum(to_float(r[spd_col]) or 0 for _, r in df.iterrows()) if spd_col else 0
    pcts: dict = {}
    for _, row in df.iterrows():
        ch = to_str(row[ch_col]).upper()
        if pct_col:
            pct = to_float(row[pct_col]) or 0
        elif spd_col and total > 0:
            pct = (to_float(row[spd_col]) or 0) / total
        else:
            pct = 0
        pcts[ch] = pct

    pmax   = pcts.get("PERFORMANCE_MAX", 0)
    search = pcts.get("SEARCH", 0)
    breakdown = ", ".join(f"{k}: {v*100:.1f}%" for k, v in pcts.items())

    if pmax > 0.85 and search < 0.05:
        return ControlResult(
            STATUS_FLAG,
            f"PMAX = {pmax*100:.1f}% of spend, Search = {search*100:.1f}%. No query-level visibility. Channel breakdown: {breakdown}.",
            WHY["F012"], WYSD["F012"],
        )
    return ControlResult(
        STATUS_OK,
        f"PMAX = {pmax*100:.1f}%, Search = {search*100:.1f}%. Channel mix within governance threshold. Breakdown: {breakdown}.",
        WHY["F012"], WYSD["F012"],
    )


def _f013(ctx): return _manual_ok("F013")
def _f014(ctx): return _manual_ok("F014")
def _f015(ctx): return _manual_ok("F015")
def _f016(ctx): return _manual_ok("F016")


def _f017(ctx: GoogleContext) -> ControlResult:
    """Logos Approved — Tab 33 (Assets Extensions)"""
    df = get_sheet(ctx, "ASSETS_EXTENSIONS")
    if df.empty:
        return ControlResult(STATUS_FLAG, "Assets Extensions tab (Tab 32/33) not found. Manual logo check required.", WHY["F017"], WYSD["F017"])

    ft_col     = find_col(df, ["FieldType"])
    approval   = find_col(df, ["PolicyApprovalStatus"])
    status_col = find_col(df, ["Status"])
    if not ft_col:
        return ControlResult(STATUS_FLAG, "FieldType column not found in Tab 33.", WHY["F017"], WYSD["F017"])

    image_types = {"MARKETING_IMAGE", "SQUARE_MARKETING_IMAGE", "PORTRAIT_MARKETING_IMAGE"}
    mask = df[ft_col].astype(str).str.upper().isin(image_types)
    if approval:
        mask &= df[approval].astype(str).str.upper() == "APPROVED"
    if status_col:
        mask &= df[status_col].astype(str).str.upper() == "ENABLED"

    count = mask.sum()
    if count >= 1:
        return ControlResult(STATUS_OK, f"{count} approved image asset(s) found in active PMAX asset groups.", WHY["F017"], WYSD["F017"])
    return ControlResult(STATUS_FLAG, "No approved image assets found. PMAX Display defaults to text-only ads.", WHY["F017"], WYSD["F017"])


def _f018(ctx: GoogleContext) -> ControlResult:
    """Ad Strength — proxy via PerformanceLabel from Tab 33"""
    df = get_sheet(ctx, "ASSETS_EXTENSIONS")
    if df.empty:
        return ControlResult(STATUS_FLAG, "Assets Extensions tab not found. Manual ad strength check required.", WHY["F018"], WYSD["F018"])

    perf_col   = find_col(df, ["PerformanceLabel"])
    status_col = find_col(df, ["Status"])
    if not perf_col:
        return ControlResult(STATUS_FLAG, "PerformanceLabel not found in Tab 33. Manual ad strength check required.", WHY["F018"], WYSD["F018"])

    active = df[df[status_col].astype(str).str.upper() == "ENABLED"] if status_col else df
    counts = active[perf_col].astype(str).str.upper().value_counts().to_dict()
    has_best = counts.get("BEST", 0) > 0
    has_good = counts.get("GOOD", 0) > 0
    all_low  = counts.get("LOW", 0) > 0 and not has_best and not has_good
    parts = ", ".join(f"{k}: {v}" for k, v in sorted(counts.items()) if k not in ("NAN", "UNKNOWN", "NONE"))

    if all_low:
        return ControlResult(
            STATUS_FLAG,
            f"All rated assets show LOW performance label. No BEST or GOOD assets. Distribution: {parts}. Note: this is an asset-level proxy, not account-level ad strength.",
            WHY["F018"], WYSD["F018"],
        )
    return ControlResult(
        STATUS_OK,
        f"Performance label distribution — {parts}. At least one BEST or GOOD asset present.",
        WHY["F018"], WYSD["F018"],
    )


def _f019(ctx): return _manual_ok("F019")
def _f020(ctx): return _manual_ok("F020")
def _f021(ctx): return _manual_ok("F021")


def _f022(ctx: GoogleContext) -> ControlResult:
    """Match Type Governance — EXACT Coverage"""
    df = get_sheet(ctx, "KEYWORD_REPORT")
    if df.empty:
        return ControlResult(STATUS_FLAG, "Keyword Report (Tab 09) not found or empty.", WHY["F022"], WYSD["F022"])

    mt_col = find_col(df, ["MatchType", "match_type"])
    if not mt_col:
        return ControlResult(STATUS_FLAG, "MatchType column not found in Tab 09.", WHY["F022"], WYSD["F022"])

    counts = df[mt_col].astype(str).str.upper().value_counts()
    total = counts.sum()
    if total == 0:
        return ControlResult(STATUS_FLAG, "No keywords found.", WHY["F022"], WYSD["F022"])

    exact_pct = counts.get("EXACT", 0) / total
    broad_pct = counts.get("BROAD", 0) / total
    dist = ", ".join(f"{k}: {v/total*100:.1f}%" for k, v in counts.items())

    if exact_pct < 0.10:
        return ControlResult(
            STATUS_FLAG,
            f"EXACT coverage = {exact_pct*100:.1f}% of {total} keywords — below the 10% minimum threshold. Cannot enforce precise query control on high-value terms. {dist}.",
            WHY["F022"], WYSD["F022"],
        )
    return ControlResult(
        STATUS_OK,
        f"EXACT coverage = {exact_pct*100:.1f}% of {total} keywords. Above 10% threshold. {dist}.",
        WHY["F022"], WYSD["F022"],
    )


# ── Orchestrator ──────────────────────────────────────────────────────────────

_EVALUATORS = {
    "F001": _f001, "F002": _f002, "F003": _f003, "F004": _f004,
    "F005": _f005, "F006": _f006, "F007": _f007, "F008": _f008,
    "F009": _f009, "F010": _f010, "F011": _f011, "F012": _f012,
    "F013": _f013, "F014": _f014, "F015": _f015, "F016": _f016,
    "F017": _f017, "F018": _f018, "F019": _f019, "F020": _f020,
    "F021": _f021, "F022": _f022,
}


def evaluate_all_framework(ctx: GoogleContext) -> Dict[str, ControlResult]:
    results = {}
    for cid, fn in _EVALUATORS.items():
        try:
            results[cid] = fn(ctx)
        except Exception as e:
            results[cid] = ControlResult(
                STATUS_FLAG,
                f"Evaluation error: {e}",
                "Internal error — review this control manually.",
                "",
            )
    return results
