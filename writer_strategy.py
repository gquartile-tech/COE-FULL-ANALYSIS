"""
writer_strategy.py
──────────────────
Writes the Strategy Analysis output workbook.

What it does
────────────
1.  Reads the pre-analysis Databricks workbook via reader_databricks_strategy.
2.  Writes the Questionnaire Survey - AMZ tab (header fields, Salesforce data, Gong notes).
3.  Runs the strategy auto-flag logic and writes OK / PARTIAL / FLAG to column G
    of the STRATEGY OVERVIEW tab for every row that has a data signal.
    Rows with no automatable signal are left as-is (default OK in template).
4.  Reads back every row from STRATEGY OVERVIEW and copies the ones that are
    FLAG or PARTIAL into the Account Strategy _Analysis findings table.
    OK rows that belong to a section with at least one FLAG/PARTIAL are also
    included as context (marked OK).
5.  Writes header, grade, and interpretation to the Analysis tab.
6.  Writes the ChildASIN View tab.
7.  Saves the output .xlsm file.

Auto-flag logic (column G, STRATEGY OVERVIEW)
─────────────────────────────────────────────
Row  3  ACoS target above constraint by >5pp                              → FLAG
Row  3  ACoS target above constraint by 2–5pp                             → PARTIAL
Row  4  ACoS consistently decreasing, changes ≥2 in 30 days              → PARTIAL
Row  6  ACoS being loosened (direction = increasing)                      → FLAG
Row  8  Account has OOB events in period                                  → FLAG
Row 17  Single parent ASIN but bulk multi-ASIN structures active          → FLAG
Row 18  CPC increased YoY by >20%                                         → FLAG
Row 18  CPC increased YoY by 10–20%                                       → PARTIAL
Row 20  Account has OOB events in period (budget constraint note)         → FLAG
Row 29  Framework compliance gap: Imported+NonQT > 40% of spend          → FLAG
Row 29  Framework compliance gap: Imported+NonQT 20–40%                  → PARTIAL
Row 30  SPT present (structure review)                                    → PARTIAL
Row 31  SPT present (narrowing to best sellers)                           → PARTIAL
Row 32  ATM < 3% of spend (severely underweighted)                       → FLAG
Row 32  ATM 3–8% of spend                                                → PARTIAL
Row 37  BA present, review slow movers                                    → PARTIAL
Row 39  BA active but fewer than 2 campaigns (no category segmentation)  → FLAG
Row 45  No SB spend at all                                                → FLAG
Row 47  Imported spend > 0 (import kickoff needed)                       → FLAG
Row 62  No CAT_SP standard naming detected                                → FLAG
Row 62  CAT_ non-standard naming detected                                 → PARTIAL
Row 63  No SBV campaigns                                                  → FLAG
Row 67  No OP / product-target campaigns detected                         → PARTIAL
Row 71  Multiple WATM campaigns (>1) active                               → PARTIAL
Row 82  No SD campaigns and SD impressions = 0                            → FLAG
Row 83  No SD_ATC / ProSuite ATC campaign                                 → PARTIAL
Row 84  SD active but no SD_PRD campaigns                                 → PARTIAL
Row 86  Portfolios present but 0 managed                                  → PARTIAL
Row 86  0 portfolios with budget cap and >3 portfolios                    → FLAG
Row 87  Non-QT + Imported > 50% of spend                                 → FLAG
Row 88  Campaign-level ACoS overrides detected                            → PARTIAL
Row 89  Product-level ACoS overrides detected                             → PARTIAL
Row 92  No RBO configured                                                 → PARTIAL
Row 93  SBV campaigns present but not named with SBV_ convention         → PARTIAL
Row 94  Campaigns not in portfolio > 0                                    → FLAG
Row 95  Campaigns not in portfolio > 0 (rename/unmanaged note)           → FLAG
Row 104 SB active (impressions > 0) but SBV spend = 0                   → FLAG
Row 111 YoY ad sales is negative                                          → FLAG
Row 124 GGS not compliant and SD spend = 0                               → FLAG
Row 125 GGS not compliant and SD spend = 0 (display coverage note)      → FLAG
Row 126 SD active but no SD_FLEX / remarketing campaigns                 → PARTIAL
Row 127 SD active but no ATC retargeting (duplicate of row 83 at GGS)   → PARTIAL
"""

from __future__ import annotations

import math
import os
import re
import sys
from datetime import datetime

import openpyxl
from openpyxl.formatting.rule import (
    ColorScaleRule,
    IconSetRule,
    Rule,
)
from openpyxl.styles import Font
from openpyxl.styles.differential import DifferentialStyle

from reader_databricks_strategy import read_strategy_context, StrategyContext


# ── helpers ───────────────────────────────────────────────────────────────────

def _safe(val, default=''):
    return default if val is None else val


def _read_header(ws):
    account_str = date_range = downloaded = ''
    for row in ws.iter_rows(min_row=1, max_row=4, values_only=True):
        for cell in row:
            if cell and isinstance(cell, str):
                if 'Account:' in cell:
                    account_str = cell
                elif 'Date Range:' in cell:
                    date_range = cell.replace('Date Range: ', '').strip()
                elif 'Downloaded:' in cell:
                    downloaded = cell.replace('Downloaded: ', '').strip()
    return account_str, date_range, downloaded


def _find_header_row(ws, max_scan=10):
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), 1):
        if len([c for c in row if c is not None]) > 3:
            return i
    return None


def _tab_to_dict(ws):
    hr = _find_header_row(ws)
    if hr is None:
        return {}
    rows = list(ws.iter_rows(min_row=hr, max_row=hr + 1, values_only=True))
    if len(rows) < 2:
        return {}
    return {h: rows[1][i] for i, h in enumerate(rows[0]) if h is not None and i < len(rows[1])}


def _tab_to_records_full(ws):
    hr = _find_header_row(ws)
    if hr is None:
        return []
    headers = None
    records = []
    for row in ws.iter_rows(min_row=hr, values_only=True):
        if headers is None:
            headers = list(row)
            continue
        if not any(row):
            continue
        records.append({headers[j]: row[j] for j in range(len(headers)) if headers[j] is not None})
    return records


def _tab_to_records(ws):
    return _tab_to_records_full(ws)


def _latest_record(records):
    if not records:
        return {}
    if len(records) == 1:
        return records[0]
    col = next(
        (k for k in records[0] if re.sub(r'[\s_]', '', k).lower() in ('systemmodstamp', 'modstamp')),
        None
    )
    if col:
        try:
            import pandas as _pd
            return max(records, key=lambda r: _pd.to_datetime(r.get(col), errors='coerce') or _pd.NaT)
        except Exception:
            pass
    return records[0]


def _filter_by_advertiser(records, account_id):
    if not account_id or not records:
        return records
    def _norm(s): return re.sub(r'[\s_]', '', str(s)).lower()
    adv_col = next(
        (k for k in records[0] if 'advertiserid' in _norm(k)),
        None
    )
    if not adv_col:
        return records
    filtered = [r for r in records if str(r.get(adv_col, '')).strip() == str(account_id).strip()]
    return filtered if filtered else records


# ── auto-flag engine ──────────────────────────────────────────────────────────

# ── Control ID → template row lookup (New Strategy Overview tab) ─────────────
_SID_TO_ROW: dict[str, int] = {
    'S001': 2,   'S002': 3,   'S003': 4,   'S004': 5,   'S005': 6,
    'S006': 7,   'S007': 8,   'S008': 9,   'S009': 10,  'S010': 11,
    'S011': 12,  'S012': 13,  'S013': 14,  'S014': 15,  'S015': 16,
    'S016': 17,  'S017': 18,  'S018': 19,  'S019': 20,  'S020': 21,
    'S021': 22,  'S022': 23,  'S023': 24,  'S024': 25,  'S025': 26,
    'S026': 27,  'S027': 28,  'S028': 29,  'S029': 30,  'S030': 31,
    'S031': 32,  'S032': 33,  'S033': 34,  'S034': 35,  'S035': 36,
    'S036': 37,  'S037': 38,  'S038': 39,  'S039': 40,  'S040': 41,
    'S041': 42,  'S042': 43,  'S043': 44,  'S044': 45,  'S045': 46,
    'S046': 47,  'S047': 48,  'S048': 49,  'S049': 50,  'S050': 51,
    'S051': 52,  'S052': 53,  'S053': 54,  'S054': 55,  'S055': 56,
    'S056': 57,  'S057': 58,  'S058': 59,  'S059': 60,  'S060': 61,
    'S061': 62,  'S062': 63,  'S063': 64,  'S064': 65,  'S065': 66,
    'S066': 67,  'S067': 68,  'S068': 69,  'S069': 70,  'S070': 71,
    'S071': 72,  'S072': 73,  'S073': 74,  'S074': 75,  'S075': 76,
    'S076': 77,  'S077': 78,  'S078': 79,  'S079': 80,  'S080': 81,
    'S081': 82,  'S082': 83,  'S083': 84,  'S084': 85,  'S085': 86,
    'S086': 87,  'S087': 88,  'S088': 89,  'S089': 90,  'S090': 91,
    'S091': 92,  'S092': 93,  'S093': 94,  'S094': 95,  'S095': 96,
    'S096': 97,  'S097': 98,  'S098': 99,  'S099': 100, 'S100': 101,
    'S101': 102, 'S102': 103, 'S103': 104, 'S104': 105,
    'S105': 106, 'S106': 107, 'S107': 108, 'S108': 109, 'S109': 110,
    'S110': 111, 'S111': 112, 'S112': 113, 'S113': 114, 'S114': 115,
    'S115': 116, 'S116': 117, 'S117': 118, 'S118': 119, 'S119': 120,
    'S120': 121, 'S121': 122, 'S122': 123, 'S123': 124, 'S124': 125,
    'S125': 126,
}


def _compute_flags(ctx: StrategyContext) -> dict[str, str]:
    """
    Performance-gated strategy suggestions.

    Design principle:
      Every flag requires a KPI condition — not just structural presence/absence.
      Framework pillar owns structural presence checks.
      Strategy owns TIMING and PRIORITY based on where the account is:
      ACoS vs constraint, TACoS trend, YoY growth, objective, spend efficiency.

    Returns {control_id: 'FLAG'|'PARTIAL'} keyed by S-IDs.
    The writer translates to row numbers using _SID_TO_ROW.
    """
    flags: dict[str, str] = {}

    def flag(sid: str, level: str):
        if flags.get(sid) == 'FLAG':   # never downgrade
            return
        flags[sid] = level

    # ── normalised comparisons ────────────────────────────────────────────────
    # acos_actual / tacos_actual are decimal (0.60 = 60%)
    # acos_constraint / tacos_constraint are integer pp (60 = 60%)
    acos_pp        = ctx.acos_actual  * 100
    tacos_pp       = ctx.tacos_actual * 100
    constraint     = ctx.acos_constraint
    tacos_con      = ctx.tacos_constraint
    has_constraint = constraint > 0
    has_tacos_con  = tacos_con  > 0
    above_acos     = has_constraint and acos_pp  > constraint
    above_acos_10  = has_constraint and acos_pp  > constraint * 1.10
    above_tacos    = has_tacos_con  and tacos_pp > tacos_con
    above_tacos_10 = has_tacos_con  and tacos_pp > tacos_con  * 1.10
    non_qt_total   = ctx.pct_imported + ctx.pct_non_quartile
    declining_yoy  = ctx.yoy_ad_sales < -0.05
    growing_yoy    = ctx.yoy_ad_sales >  0.10
    tacos_rising   = ctx.tacos_trend == 'increasing' and ctx.tacos_trend_pp > 1.5
    spend_rising   = ctx.mom_spend_change > 0.10
    has_atc        = any(
        re.search(r'\bATC\b|SD_FLEX', n, re.IGNORECASE)
        for n in ctx.campaign_names
    )

    # ── ACOS AND TARGET ───────────────────────────────────────────────────────

    # S002 — ACoS target above constraint.
    if has_constraint and ctx.acos_gap_to_constraint > 5:
        flag('S002', 'FLAG')
    elif has_constraint and ctx.acos_gap_to_constraint > 2:
        flag('S002', 'PARTIAL')

    # S004 — ACoS reduction cadence. Keep as-is per review decision.

    # S006 — ACoS being loosened while performance is already below target.
    if ctx.acos_direction == 'increasing' and (above_acos or declining_yoy):
        flag('S006', 'FLAG')

    # S007 / S008 — keep as-is per review decision.

    # S009 — Framework meta-check. Only fires when gaps compound a perf problem.
    _gaps = sum([
        ctx.spend_sb   == 0,
        not ctx.has_cat_sp,
        not ctx.has_sbv and ctx.spend_sbv == 0,
        not ctx.has_sd  and ctx.spend_sd  == 0,
        ctx.spend_spt  > 0 and ctx.pct_atm < 0.03,
        ctx.campaigns_not_in_portfolio > 5,
        non_qt_total   > 0.40,
        not ctx.has_op and ctx.pct_op == 0,
        not has_atc,
    ])
    if _gaps >= 5 and (above_acos or declining_yoy):
        flag('S009', 'FLAG')
    elif _gaps >= 3 and above_acos:
        flag('S009', 'PARTIAL')

    # ── OVERALL STRUCTURE ─────────────────────────────────────────────────────

    # S005 — Portfolio completion: mid-way migration (40–90%).
    if 0.40 <= ctx.campaigns_in_portfolio_pct < 0.90 and ctx.total_campaign_count > 5:
        flag('S005', 'PARTIAL')

    # S014 — BA running but no BAK harvest layer. Gate: not critically above TACoS.
    if ctx.pct_ba > 0 and ctx.pct_bak == 0 and ctx.total_spend > 500 and not above_tacos_10:
        flag('S014', 'FLAG')
    elif ctx.pct_ba > 0 and 0 < ctx.pct_bak < ctx.pct_ba * 0.3:
        flag('S014', 'PARTIAL')

    # S017 — Single ASIN with multi-ASIN bulk structures.
    if ctx.parent_asin_count == 1 and (ctx.has_catchall or ctx.spend_spt > 0):
        flag('S017', 'FLAG')

    # ── OVERALL PARAMETERS AND KPIs ───────────────────────────────────────────

    # S018 — CPC up YoY. More urgent when efficiency is already suffering.
    if ctx.cpc_yoy_change_pct > 0.20 and above_acos:
        flag('S018', 'FLAG')
    elif ctx.cpc_yoy_change_pct > 0.20:
        flag('S018', 'PARTIAL')
    elif ctx.cpc_yoy_change_pct > 0.10 and above_acos:
        flag('S018', 'PARTIAL')

    # S019 — TACoS increasing trend approaching constraint (repurposed).
    if tacos_rising and ctx.tacos_trend_pp > 0.70 and above_tacos:
        flag('S019', 'FLAG')
    elif tacos_rising and ctx.tacos_trend_pp > 0.70:
        flag('S019', 'PARTIAL')

    # S020 — OOB + budget expansion. Urgency depends on efficiency state.
    if ctx.has_oob and not above_acos_10:
        flag('S020', 'FLAG')
    elif ctx.has_oob:
        flag('S020', 'PARTIAL')

    # S021 — No TACoS constraint documented.
    if tacos_con == 0:
        flag('S021', 'PARTIAL')

    # S023 — OOB scope/budget decision.
    if ctx.has_oob and above_acos_10:
        flag('S023', 'FLAG')
    elif ctx.has_oob:
        flag('S023', 'PARTIAL')

    # ── BASIC STRATEGY ────────────────────────────────────────────────────────

    # S029 — Unmanaged spend. More urgent when ACoS is already above constraint.
    if non_qt_total > 0.40 or (non_qt_total > 0.20 and above_acos_10):
        flag('S029', 'FLAG')
    elif non_qt_total > 0.20:
        flag('S029', 'PARTIAL')

    # S030/S031 — SPT structure review. Only when performance signals waste.
    if ctx.spend_spt > 0 and (above_acos or above_tacos or declining_yoy):
        flag('S030', 'PARTIAL')
        flag('S031', 'PARTIAL')

    # S032 — ATM expansion. Only when account has headroom to invest.
    if ctx.pct_atm < 0.03 and not above_tacos_10 and not (ctx.has_oob and above_acos_10):
        flag('S032', 'FLAG')
    elif ctx.pct_atm < 0.08 and not above_tacos and growing_yoy:
        flag('S032', 'PARTIAL')

    # S037 — BA refocus. BA slow movers are likely the cause when ACoS or YoY is bad.
    if ctx.spend_ba > 0 and (above_acos_10 or declining_yoy):
        flag('S037', 'FLAG')
    elif ctx.spend_ba > 0 and above_acos:
        flag('S037', 'PARTIAL')

    # S039 — BA not segmented. Only meaningful with enough spend volume.
    if 0 < ctx.ba_campaign_count < 2 and ctx.total_spend > 500:
        flag('S039', 'FLAG')

    # S045 — No SB. Always a suggestion; urgency adapts to performance.
    if ctx.spend_sb == 0:
        if declining_yoy:
            flag('S045', 'FLAG')
        elif above_acos_10:
            flag('S045', 'PARTIAL')
        else:
            flag('S045', 'FLAG')

    # S047 — Imported campaigns need import kickoff regardless of performance.
    if ctx.spend_imported > 0:
        flag('S047', 'FLAG')

    # ── NEW DEPLOYS ───────────────────────────────────────────────────────────

    # S062 — CAT_SP missing. Surface when account has headroom.
    if not ctx.has_cat_sp and ctx.total_spend > 500:
        if growing_yoy or not above_acos:
            flag('S062', 'FLAG')
        else:
            flag('S062', 'PARTIAL')

    # S063 — No SBV. Most urgent when SB is already active (natural next step).
    if not ctx.has_sbv and ctx.spend_sbv == 0:
        if ctx.spend_sb > 0:
            flag('S063', 'FLAG')
        elif growing_yoy:
            flag('S063', 'PARTIAL')

    # S071 — Multiple WATM without structural reason.
    if ctx.watm_campaign_count > 1:
        flag('S071', 'PARTIAL')

    # S078 — BAK high spend but ACoS above constraint.
    if ctx.pct_bak > 0.15 and above_acos_10:
        flag('S078', 'PARTIAL')

    # S082 — No SD at all. Surface when growing or SB already active.
    if not ctx.has_sd and ctx.spend_sd == 0 and ctx.total_spend > 500:
        if growing_yoy or ctx.spend_sb > 0:
            flag('S082', 'FLAG')
        elif not above_acos:
            flag('S082', 'PARTIAL')

    # S083 — No ATC retargeting. Only when ProSuite is active.
    if not has_atc and ctx.has_prosuite_audiences:
        flag('S083', 'PARTIAL')

    # S084 — SD active but no SD_PRD product-page coverage.
    if ctx.spend_sd > 0 and not ctx.has_sd_prd:
        flag('S084', 'PARTIAL')

    # ── GOVERNANCE ────────────────────────────────────────────────────────────

    # S086 — Portfolio governance.
    if ctx.portfolio_count > 3 and ctx.portfolios_with_budget_cap == 0:
        flag('S086', 'FLAG')
    elif ctx.portfolio_count > 0 and ctx.managed_portfolio_count == 0:
        flag('S086', 'PARTIAL')

    # S087 — >50% unmanaged. Aligned thresholds with S029.
    if non_qt_total > 0.50:
        flag('S087', 'FLAG')
    elif non_qt_total > 0.35 and above_acos:
        flag('S087', 'PARTIAL')

    # S088/S089 — Overrides. Only a concern when ACoS is already struggling.
    if ctx.has_campaign_acos_overrides and above_acos:
        flag('S088', 'PARTIAL')
    if ctx.has_product_acos_overrides and above_acos:
        flag('S089', 'PARTIAL')

    # S090 — VCPM overuse (share-based, not presence).
    if ctx.vcpm_spend_pct > 0.10:
        flag('S090', 'FLAG')
    elif ctx.vcpm_spend_pct > 0.05:
        flag('S090', 'PARTIAL')

    # S091 — Tagging/segmentation gap.
    if ctx.spend_ba > 0 and ctx.spend_spt > 0 and ctx.spend_atm > 0 and not ctx.has_op:
        flag('S091', 'PARTIAL')

    # S092 — RBO. Suggest only when there's a performance reason.
    if not ctx.has_rbo and (above_acos or declining_yoy):
        flag('S092', 'PARTIAL')

    # S093 — SBV naming convention.
    if ctx.has_sbv and not ctx.sbv_naming_compliant:
        flag('S093', 'PARTIAL')

    # S094/S095 — Campaigns outside portfolio.
    if ctx.campaigns_not_in_portfolio > 5:
        flag('S094', 'FLAG')
        flag('S095', 'FLAG')
    elif ctx.campaigns_not_in_portfolio > 0:
        flag('S094', 'PARTIAL')

    # S102 — ProSuite not deployed. Suggest at scale + growing.
    if not ctx.has_prosuite_audiences and len(ctx.campaign_names) > 20 and growing_yoy:
        flag('S102', 'FLAG')
    elif not ctx.has_prosuite_audiences and len(ctx.campaign_names) > 20:
        flag('S102', 'PARTIAL')

    # S104 — SB active but SBV missing.
    if ctx.sb_impressions > 0 and ctx.spend_sbv == 0:
        flag('S104', 'FLAG')

    # ── CLIENT DIRECTIONS ─────────────────────────────────────────────────────

    # S108 — Sales declining YoY while spend growing (efficiency trap).
    if declining_yoy and spend_rising:
        flag('S108', 'FLAG')
    elif declining_yoy:
        flag('S107', 'FLAG')

    # S113/S114 — Subscribe & Save. Suggest when no promo active.
    if not ctx.has_active_promo:
        if declining_yoy:
            flag('S113', 'FLAG')
            flag('S114', 'FLAG')
        else:
            flag('S113', 'PARTIAL')
            flag('S114', 'PARTIAL')

    # ── PROMO AND GGS ─────────────────────────────────────────────────────────

    # S117/S118 — Active promo — validate budget pacing.
    if ctx.has_active_promo:
        flag('S117', 'PARTIAL')
        flag('S118', 'PARTIAL')

    # S125 — No promo — channel expansion evaluation.
    if not ctx.has_active_promo:
        flag('S125', 'PARTIAL')

    # S120/S121 — GGS non-compliant and no SD spend.
    if ctx.ggs_status == 'No' and ctx.spend_sd == 0:
        flag('S120', 'FLAG')
        flag('S121', 'FLAG')

    # S122 — SD active but no remarketing.
    has_remarketing = any(
        re.search(r'SD_FLEX|SD_AUDI|remarketing', n, re.IGNORECASE)
        for n in ctx.campaign_names
    )
    if ctx.spend_sd > 0 and not has_remarketing:
        flag('S122', 'PARTIAL')

    # S123 — SD active but no ATC (ProSuite gated).
    if ctx.spend_sd > 0 and not has_atc and ctx.has_prosuite_audiences:
        flag('S123', 'PARTIAL')

    return flags


# ── dynamic What We Saw text ──────────────────────────────────────────────────

def _build_what_we_saw(ctx: StrategyContext, flags: dict[str, str]) -> dict[str, str]:
    """
    Returns {control_id: text} with plain-language What We Saw sentences
    built from real account numbers, for every control that fired.
    """
    texts: dict[str, str] = {}

    def pct(v: float) -> str:
        return f'{v:.0%}'

    def dollar(v: float) -> str:
        return f'${v:,.0f}'

    if 'S002' in flags:
        texts['S002'] = (
            f'The current ACoS target is {ctx.acos_current_target:.0f}%. '
            f'The account constraint is {ctx.acos_constraint:.0f}%. '
            f'The gap is +{ctx.acos_gap_to_constraint:.0f} percentage points. '
            f'The target needs to come down to align with the client objective.'
        )

    if 'S009' in flags:
        gap_labels = []
        if ctx.spend_sb == 0: gap_labels.append('no SB campaigns')
        if not ctx.has_cat_sp: gap_labels.append('no CAT_SP campaigns')
        if not ctx.has_sbv and ctx.spend_sbv == 0: gap_labels.append('no SBV campaigns')
        if not ctx.has_sd and ctx.spend_sd == 0: gap_labels.append('no SD campaigns')
        if ctx.spend_spt > 0 and ctx.pct_atm < 0.03: gap_labels.append('SPT active but ATM < 3%')
        if ctx.campaigns_not_in_portfolio > 5: gap_labels.append(f'{ctx.campaigns_not_in_portfolio} campaigns outside portfolios')
        if (ctx.pct_imported + ctx.pct_non_quartile) > 0.40: gap_labels.append(f'{pct(ctx.pct_imported + ctx.pct_non_quartile)} spend outside framework')
        if not ctx.has_op and ctx.pct_op == 0: gap_labels.append('no OP / product-target campaigns')
        if not any(re.search(r'\\bATC\\b|SD_FLEX', n, re.IGNORECASE) for n in ctx.campaign_names): gap_labels.append('no ATC retargeting')
        n_gaps = len(gap_labels)
        gaps_str = ', '.join(gap_labels[:5])
        suffix = f' (+{n_gaps - 5} more)' if n_gaps > 5 else ''
        texts['S009'] = (
            f'{n_gaps} structural framework gaps detected: {gaps_str}{suffix}. '
            f'A structured framework review is needed before the next QR.'
        )

    if 'S005' in flags:
        in_port = round(ctx.campaigns_in_portfolio_pct * ctx.total_campaign_count)
        not_in_port = ctx.total_campaign_count - in_port
        texts['S005'] = (
            f'{in_port} of {ctx.total_campaign_count} campaigns ({ctx.campaigns_in_portfolio_pct:.0%}) are already in portfolios. '
            f'{not_in_port} campaign(s) remain outside. Complete the portfolio assignment.'
        )

    if 'S007' in flags:
        texts['S007'] = (
            f'Branded spend is {ctx.branded_spend_pct:.0%} of total at {ctx.branded_acos:.0%} ACoS. '
            f'Non-branded is at {ctx.non_branded_acos:.0%} ACoS vs portal target {ctx.acos_current_target:.0f}%. '
            f'Revisit ACoS target to restore non-branded efficiency.'
        )

    if 'S006' in flags:
        texts['S006'] = (
            f'ACoS target increased {ctx.acos_changes_30d} time(s) in the last 30 days. '
            f'Current target: {ctx.acos_current_target:.0f}%. '
            f'Spend growth driven by loosening efficiency — not by structural improvements.'
        )

    if 'S008' in flags:
        texts['S008'] = (
            f'Account hit daily budget limits. ACoS target: {ctx.acos_current_target:.0f}% vs constraint {ctx.acos_constraint:.0f}%. '
            f'Reducing the ACoS target lowers CPC pressure and eases OOB.'
        )

    if 'S017' in flags:
        texts['S017'] = (
            f'The account has {ctx.parent_asin_count} parent ASIN. '
            f'Multi-ASIN bulk structures add complexity without value at this catalog size.'
        )

    if 'S014' in flags:
        texts['S014'] = (
            f'BA campaigns are active ({pct(ctx.pct_ba)} of spend / {dollar(ctx.spend_ba)}) '
            f'but no BAK harvest layer exists. '
            f'Discovery data is not being converted into manual precision targets.'
        )

    if 'S018' in flags:
        texts['S018'] = (
            f'CPC moved from ${ctx.cpc_last_year:.2f} last year to ${ctx.cpc_current:.2f} ({ctx.cpc_yoy_change_pct:+.0%}). '
            f'ACoS thresholds should be revisited to bring costs back under control.'
        )

    if 'S019' in flags:
        texts['S019'] = (
            f'TACoS has been {ctx.tacos_trend} for the last 3 months (+{ctx.tacos_trend_pp:.1f}pp). '
            f'Current TACoS: {ctx.tacos_actual:.0%} vs constraint {ctx.tacos_constraint:.0f}%. '
            f'Profitability is eroding — strategic action needed before the constraint is breached.'
        )

    if 'S020' in flags:
        texts['S020'] = (
            f'Account ran out of budget at least once. Total spend: {dollar(ctx.total_spend)}. '
            f'Budget expansion or scope reduction should be reviewed with the client.'
        )

    if 'S021' in flags:
        texts['S021'] = (
            f'No TACoS constraint documented for this account. '
            f'Without a TACoS target, profitability tracking has no reference point. '
            f'Agree a TACoS goal with the client and document it in Client Success.'
        )

    if 'S023' in flags:
        texts['S023'] = (
            f'Account hit OOB. Spend: {dollar(ctx.total_spend)}. '
            f'Budget should be expanded, or product scope reduced to concentrate on top ASINs.'
        )

    if 'S029' in flags:
        non_qt = ctx.pct_imported + ctx.pct_non_quartile
        texts['S029'] = (
            f'{pct(non_qt)} of spend is in Imported or Non-Quartile campaigns '
            f'({pct(ctx.pct_imported)} Imported, {pct(ctx.pct_non_quartile)} Non-Quartile). '
            f'The account is not fully operating within the Quartile framework.'
        )

    if 'S030' in flags:
        texts['S030'] = (
            f'SPT is active ({dollar(ctx.spend_spt)}, {pct(ctx.pct_spt)} of spend). '
            f'Defensive structure should be reviewed by category or brand segment.'
        )

    if 'S031' in flags:
        texts['S031'] = (
            f'SPT spend: {dollar(ctx.spend_spt)}. '
            f'Coverage should be narrowed to the strongest-selling products only.'
        )

    if 'S032' in flags:
        texts['S032'] = (
            f'ATM campaigns represent {pct(ctx.pct_atm)} of spend ({dollar(ctx.spend_atm)}). '
            + ('No ATM spend detected. ' if ctx.pct_atm == 0 else '')
            + f'Automatic targeting on best-selling ASINs should be expanded.'
        )

    if 'S037' in flags:
        texts['S037'] = (
            f'BA campaigns: {dollar(ctx.spend_ba)} ({pct(ctx.pct_ba)} of spend). '
            f'Review to remove slow-moving products and focus on best sellers.'
        )

    if 'S039' in flags:
        texts['S039'] = (
            f'Only {ctx.ba_campaign_count} BA campaign(s) detected. '
            f'Structure is not segmented by category — new BA campaigns by category needed.'
        )

    if 'S045' in flags:
        acos_pp  = ctx.acos_actual * 100
        declining = ctx.yoy_ad_sales < -0.05
        acos_high = ctx.acos_constraint > 0 and acos_pp > ctx.acos_constraint * 1.2
        prefix = f'No Sponsored Brands spend detected. '
        prefix += f'SBV is active ({dollar(ctx.spend_sbv)}) but SB is absent. ' if ctx.spend_sbv > 0 else ''
        if declining:
            suffix = f'Ad sales down {pct(abs(ctx.yoy_ad_sales))} YoY — SB is a direct lever for upper-funnel recovery.'
        elif acos_high:
            suffix = f'ACoS is {acos_pp:.0f}% vs {ctx.acos_constraint:.0f}% constraint — address efficiency before launching SB.'
        else:
            suffix = 'SB campaigns should be launched to build upper-funnel coverage.'
        texts['S045'] = prefix + suffix

    if 'S047' in flags:
        texts['S047'] = (
            f'Imported campaigns: {dollar(ctx.spend_imported)} ({pct(ctx.pct_imported)} of spend). '
            f'These run outside the Quartile system. An import kickoff CoE ticket is needed.'
        )

    if 'S062' in flags:
        texts['S062'] = (
            f'No CAT_SP campaigns detected. '
            f'Category-targeted SP campaigns should be launched for key product categories.'
        )

    if 'S063' in flags:
        prefix = 'No SBV campaigns detected. '
        if ctx.spend_sb > 0:
            prefix += f'SB is active ({ctx.sb_impressions:,} impressions) — SBV is the natural next deploy. '
        texts['S063'] = prefix + 'Launch SBV product-targeting campaigns.'

    if 'S071' in flags:
        texts['S071'] = (
            f'{ctx.watm_campaign_count} WATM campaigns active. '
            f'Multiple WATM campaigns add fragmentation without structural benefit.'
        )

    if 'S078' in flags:
        texts['S078'] = (
            f'BAK: {pct(ctx.pct_bak)} of spend ({dollar(ctx.spend_bak)}). '
            f'Account ACoS {ctx.acos_actual:.0%} vs {ctx.acos_constraint:.0f}% constraint. '
            f'High-spend BAK terms not meeting ACoS target should be identified and removed.'
        )

    if 'S082' in flags:
        texts['S082'] = (
            f'No Sponsored Display campaigns active. SD spend $0, impressions: {ctx.sd_impressions:,}. '
            f'Product-view remarketing and audience retargeting are not running.'
        )

    if 'S083' in flags:
        texts['S083'] = (
            f'No ATC retargeting campaigns detected. '
            f'ProSuite is active — SD_FLEX_ATC should be deployed for add-to-cart audiences.'
        )

    if 'S084' in flags:
        texts['S084'] = (
            f'SD is active ({dollar(ctx.spend_sd)}) but no SD_PRD product-page campaigns detected. '
            f'Product-page defensive coverage via SD_PRD is missing.'
        )

    if 'S086' in flags:
        texts['S086'] = (
            f'{ctx.portfolio_count} portfolios exist. '
            f'{ctx.managed_portfolio_count} managed. '
            f'{ctx.portfolios_with_budget_cap} have budget caps. '
            f'Portfolio governance needs to be tightened.'
        )

    if 'S087' in flags:
        non_qt = ctx.pct_imported + ctx.pct_non_quartile
        texts['S087'] = (
            f'{pct(non_qt)} of spend is in unmanaged campaigns. '
            f'More than half of spend is outside the Quartile system.'
        )

    if 'S088' in flags:
        texts['S088'] = (
            f'Campaign-level ACoS overrides are active. '
            f'With ACoS already above constraint, these overrides may be conflicting with system logic.'
        )

    if 'S089' in flags:
        texts['S089'] = (
            f'Product-level ACoS overrides are active. '
            f'With ACoS above constraint, per-ASIN overrides add inconsistency.'
        )

    if 'S090' in flags:
        texts['S090'] = (
            f'VCPM campaigns represent {pct(ctx.vcpm_spend_pct)} of SD spend. '
            f'Above the 10% threshold — impression-based buying is over-weighted. '
            f'Review Buy Box ownership before increasing VCPM investment.'
        )

    if 'S091' in flags:
        texts['S091'] = (
            f'BA ({dollar(ctx.spend_ba)}), SPT ({dollar(ctx.spend_spt)}), ATM ({dollar(ctx.spend_atm)}) all active '
            f'but no OP product-target campaigns detected. Product-page coverage is missing.'
        )

    if 'S092' in flags:
        texts['S092'] = (
            f'No RBO rules configured. Weekend bid management is not active. '
            f'With current performance pressure, an RBO rule should be considered.'
        )

    if 'S093' in flags:
        texts['S093'] = (
            f'SBV campaigns active but not all follow the SBV_ naming convention. '
            f'Non-standard naming reduces governance clarity.'
        )

    if 'S094' in flags:
        texts['S094'] = (
            f'{ctx.campaigns_not_in_portfolio} campaign(s) not assigned to any portfolio. '
            f'All active campaigns should be assigned consistently.'
        )

    if 'S095' in flags:
        texts['S095'] = (
            f'{ctx.campaigns_not_in_portfolio} campaign(s) outside the portfolio structure. '
            f'Review and rename to the correct Quartile naming convention.'
        )

    if 'S102' in flags:
        texts['S102'] = (
            f'{len(ctx.campaign_names)} campaigns active but no ProSuite AMC audiences applied. '
            f'Test Amazon native audiences on the strongest SP campaigns.'
        )

    if 'S104' in flags:
        texts['S104'] = (
            f'SB active ({ctx.sb_impressions:,} impressions) but SBV spend is $0. '
            f'Launch SBV product-targeting and branded campaigns.'
        )

    if 'S107' in flags:
        texts['S107'] = (
            f'Ad sales declined {abs(ctx.yoy_ad_sales):.0%} YoY. '
            f'Current ad sales: {dollar(ctx.ad_sales)}. '
            f'Recurring sales strategy and retention initiatives should be reviewed.'
        )

    if 'S108' in flags:
        texts['S108'] = (
            f'Ad sales declined {abs(ctx.yoy_ad_sales):.0%} YoY while spend increased {ctx.mom_spend_change:.0%} MoM. '
            f'More spend in, less revenue out — efficiency trap. '
            f'Budget levels and campaign scope must be reviewed before the next cycle.'
        )

    if 'S113' in flags:
        texts['S113'] = (
            f'No promotional activity is active. Subscribe & Save is not running. '
            f'For repurchasable products, S&S should be reviewed as a retention lever.'
        )

    if 'S114' in flags:
        texts['S114'] = (
            f'No recurring-purchase strategy currently active. '
            f'Subscribe & Save could increase customer lifetime value. Review with client.'
        )

    if 'S117' in flags:
        texts['S117'] = (
            f'{ctx.promo_asin_count} ASIN(s) in active promo. '
            + (f'Promo cost rate averaging {pct(ctx.promo_cost_rate)}. ' if ctx.promo_cost_rate > 0 else '')
            + f'Portfolio budgets should be reviewed to prevent intraday depletion.'
        )

    if 'S118' in flags:
        texts['S118'] = (
            f'Active promo campaigns running. '
            + (f'Avg promo cost rate: {pct(ctx.promo_cost_rate)}. ' if ctx.promo_cost_rate > 0 else '')
            + f'Campaign-level budget limits should be validated.'
        )

    if 'S125' in flags:
        texts['S125'] = (
            f'No Promo Management activity running. '
            f'Evaluate Promo Management as a channel expansion for recurring purchases.'
        )

    if 'S120' in flags:
        texts['S120'] = (
            f'GGS status: {ctx.ggs_status}. SD spend $0. '
            f'SD campaigns should be deployed to progress toward the 5% GGS target.'
        )

    if 'S121' in flags:
        texts['S121'] = (
            f'Display coverage is $0. No SD impressions in period. '
            f'No upper- or mid-funnel display activity running.'
        )

    if 'S122' in flags:
        texts['S122'] = (
            f'SD active ({dollar(ctx.spend_sd)}) but no SD_FLEX or SD_AUDI remarketing campaigns. '
            f'Product-view remarketing is not running.'
        )

    if 'S123' in flags:
        texts['S123'] = (
            f'SD active ({dollar(ctx.spend_sd)}) but no ATC retargeting in place. '
            f'Add-to-cart retargeting via ProSuite AMC is not activated.'
        )

    return texts


# ── grade calculator ──────────────────────────────────────────────────────────

def _calculate_grade(flags: dict[str, str]) -> tuple[str, str]:
    """Returns (grade_label, interpretation_text)."""
    n_flag    = sum(1 for v in flags.values() if v == 'FLAG')
    n_partial = sum(1 for v in flags.values() if v == 'PARTIAL')

    if n_flag == 0 and n_partial == 0:
        grade = 'Compliant'
        interp = (
            'The account reflects a well-defined strategic direction with no major gaps identified. '
            'Few or no changes are required — the current campaign structure, targeting approach, '
            'and client alignment are consistent with the account\'s objectives and roadmap.'
        )
    elif n_flag == 0 and n_partial <= 3:
        grade = 'Needs Review'
        interp = (
            f'The account has {n_partial} area(s) that require attention. '
            'No critical gaps were found, but several strategic items should be reviewed '
            'before the next client interaction to avoid them becoming larger issues.'
        )
    elif n_flag <= 2:
        grade = 'Needs Improvement'
        interp = (
            f'The account has {n_flag} critical gap(s) and {n_partial} item(s) needing attention. '
            'Action is required. Review the flagged controls below and align with the client '
            'or internal team on a clear plan before the next review cycle.'
        )
    else:
        grade = 'Non-Compliant'
        interp = (
            f'The account has {n_flag} critical strategic gaps. '
            'Significant structural or strategic work is required. '
            'Prioritise the flagged controls and escalate where client alignment is needed.'
        )

    return grade, interp


# ── main writer ───────────────────────────────────────────────────────────────

def write_strategy(pre_analysis_path: str, template_path: str, output_dir: str) -> str:

    # ── read context ─────────────────────────────────────────────────────────
    ctx = read_strategy_context(pre_analysis_path)

    # ── also read raw tabs for Questionnaire + ChildASIN ─────────────────────
    pa = openpyxl.load_workbook(pre_analysis_path, data_only=True, read_only=True)

    d55    = _tab_to_dict(pa['55_Salesforce_Consolidated_PreA'])
    d38    = _latest_record(_tab_to_records_full(pa['38_Client_Success_Insights_Repo']))
    gong_r = _tab_to_records(pa['37_Gong_Call_Insights_for_Sales'])
    gong   = gong_r[0] if gong_r else {}

    asin_records = _tab_to_records(pa['14_Campaign_Performance_by_Adve'])
    cat_records  = _tab_to_records(pa['22_Catalogue_Details'])
    cat_by_asin  = {r['asin']: r for r in cat_records if r.get('asin')}

    _d54_all      = _tab_to_records_full(pa['54_Project_Dataset_on_SF'])
    _d54_filtered = _filter_by_advertiser(_d54_all, ctx.profile_id)
    d54           = _latest_record(_d54_filtered)

    pa.close()

    # ── compute auto-flags ───────────────────────────────────────────────────
    flags = _compute_flags(ctx)
    grade, interp = _calculate_grade(flags)

    # ── load template ────────────────────────────────────────────────────────
    wb = openpyxl.load_workbook(template_path, keep_vba=True)

    # ════════════════════════════════════════════════════════════════════════════
    # TAB 1 — Questionnaire Survey - AMZ
    # ════════════════════════════════════════════════════════════════════════════
    ws1 = wb['Questionaire Survey - AMZ']

    def w1(coord, value):
        ws1[coord] = value

    w1('C6', ctx.member_id)
    w1('F6', ctx.profile_id)
    w1('J6', _safe(d55.get('CSP_Last_Modified_By')))

    w1('F7', ctx.profile_id)
    w1('J7', _safe(d55.get('Projected_Project_MRR__c')))

    w1('C8', _safe(d55.get('Account_Name')))
    ld = d55.get('Launch_Date__c')
    w1('F8', ld.strftime('%Y-%m-%d') if hasattr(ld, 'strftime') else _safe(ld))
    if hasattr(ld, 'strftime'):
        months = (datetime.now().year - ld.year) * 12 + (datetime.now().month - ld.month)
        w1('J8', f'{months} months')
    else:
        w1('J8', _safe(d38.get('Customer_Age_Months__c')))

    w1('C9', _safe(d55.get('Customer_Age_Months__c') or d38.get('Customer_Age_Months__c')))
    w1('F9', _safe(d38.get('Repeat_Purchase_Behavior__c')))
    w1('J9', _safe(d55.get('CSM_Churn_Risk__c')))

    w1('C10', _safe(d38.get('Commodity_Products_or_Branded_Products__c')))
    w1('F10', _safe(d55.get('Vertical__c')))
    w1('J10', _safe(d55.get('Contract_Term__c')))

    w1('C11', _safe(d38.get('Average_Order_Value__c') or d54.get('Average_Order_Value__c')))
    w1('F11', _safe(d55.get('Services_Sold__c')))
    w1('J11', _safe(d55.get('MRR__c')))

    # Row 12 — Director SF user ID not readable; write Active Products
    w1('C12', '')
    w1('F12', _safe(d55.get('Active_Products__c')))

    # Row 13
    w1('F13', _safe(d38.get('Customer_Feedback__c')))

    # Row 15
    w1('C15', _safe(d55.get('Current_Challenges__c')))
    w1('F15', _safe(d55.get('Primary_Objective__c')))
    w1('J15', _safe(d55.get('ACOS_Constraint__c')))

    # Row 16
    w1('C16', _safe(d55.get('Primary_Objective_Additional_Context__c')))
    w1('F16', _safe(d55.get('Primary_Spend_KPI__c')))
    w1('J16', _safe(d38.get('Customer_Acquisition_Cost_Target__c')))

    # Row 17
    w1('C17', _safe(d55.get('Top_Priority__c')))
    w1('J17', _safe(d55.get('TACOS_Constraint__c')))

    # Row 18
    w1('C18', _safe(d55.get('Second_Priority__c')))
    w1('F18', _safe(d54.get('CS_Notes__c')))
    w1('J18', _safe(d55.get('daily_target_spend__c')))

    # Row 19
    w1('C19', _safe(d55.get('Biggest_Expansion_Opportunity__c')))
    w1('F19', _safe(d55.get('Near_Term_3_Month_Considerations__c')))
    w1('J19', _safe(d55.get('Target_ROAS__c')))


    # CJM Salesforce strategy steps — fixed row positions in template
    stage_rows = {1: (24, 25), 2: (27, 28), 3: (30, 31), 4: (33, 34)}
    for s, (r_a, r_i) in stage_rows.items():
        w1(f'C{r_a}', _safe(d55.get(f'AdoptionOrUpsellS{s}__c')))
        w1(f'G{r_a}', _safe(d55.get(f'StrategyS{s}__c')))
        w1(f'J{r_a}', _safe(d55.get(f'StatusS{s}__c')))
        intro = d55.get(f'ExecutionDateS{s}__c')
        w1(f'C{r_i}', intro.strftime('%Y-%m-%d') if hasattr(intro, 'strftime') else _safe(intro))


    # Gong
    w1('C41', _safe(gong.get('Gong__Call_Brief__c') or d55.get('Call_Brief')))
    w1('C42', _safe(gong.get('Gong__Call_Key_Points__c') or d55.get('Key_Points')))
    w1('C43', _safe(gong.get('Gong__Call_Highlights_Next_Steps__c') or d55.get('Highlights_Next_Steps')))

    # ════════════════════════════════════════════════════════════════════════════
    # TAB — New Strategy Overview
    # Col 5  (E) = Auto Review (AUTO/MANUAL — static in template, not overwritten)
    # Col 6  (F) = STATUS written by agent: FLAG / PARTIAL / OK
    # Col 10 (J) = What We Saw — dynamic text built from real account numbers
    # ════════════════════════════════════════════════════════════════════════════
    ws_ov = wb['New Strategy Overview']

    # Write STATUS (col 6) for every control that fired
    for sid, level in flags.items():
        row_num = _SID_TO_ROW.get(sid)
        if row_num:
            ws_ov.cell(row=row_num, column=6, value=level)

    # Write dynamic What We Saw (col 10) for controls that fired
    dynamic_what = _build_what_we_saw(ctx, flags)
    for sid, text in dynamic_what.items():
        row_num = _SID_TO_ROW.get(sid)
        if row_num:
            ws_ov.cell(row=row_num, column=10, value=text)

    # ════════════════════════════════════════════════════════════════════════════
    # TAB — Account Strategy _Analysis
    # Header writes only — findings population stays manual per design decision
    # ════════════════════════════════════════════════════════════════════════════
    ws2 = wb['Account Strategy _Analysis']
    ws2['A1'] = f'{ctx.account_label} — Account Strategy Analysis'
    ws2['B3'] = f'Account: {ctx.account_label} | Tenant ID: {ctx.tenant_id} | Account ID: {ctx.profile_id}'
    ws2['B4'] = ctx.date_range
    ws2['B5'] = ctx.downloaded

    # TAB 3 — ChildASIN View
    # ════════════════════════════════════════════════════════════════════════════
    ws3 = wb['ChildASIN View']

    col14_map = {
        'Parent ASIN':       'ParentASIN',
        'ASIN':              'asin',
        'Total Sales':       'TotalSales',
        'Ad Spend':          'AdSpend',
        'TACoS':             'TACoS',
        'Ad Sales':          'AdSales',
        'Ads Units Ordered': 'Orders',
        'ACoS':              'ACoS',
        'Clicks':            'Clicks',
        'Tier':              'Tier',
        'ATM_Spend':         'ATM_Spend',
        'BA_Spend':          'BA_Spend',
        'Manual_Q1_Spend':   'Manual_Q1_Spend',
        'BAK_Spend':         'BAK_Spend',
        'OP_Spend':          'OP_Spend',
        'SPT_Spend':         'SPT_Spend',
        'CAT_SP_Spend':      'CAT_SP_Spend',
        'WATM_Spend':        'WATM_Spend',
        'SB_Spend':          'SB_Spend',
        'SBV_Spend':         'SBV_Spend',
        'SD_Spend':          'SD_Spend',
        'Imported_Spend':    'Imported_Spend',
        'NonQuartile_Spend': 'NonQuartile_Spend',
    }

    # find data start row — header is row 2 in the template
    data_row = 3
    for col_label, src_key in col14_map.items():
        # find which column in ws3 matches col_label
        header_row = list(ws3.iter_rows(min_row=2, max_row=2, values_only=True))[0]
        for ci, hval in enumerate(header_row, 1):
            if hval and str(hval).strip() == col_label:
                for ri, rec in enumerate(asin_records, data_row):
                    val = rec.get(src_key)
                    ws3.cell(row=ri, column=ci, value=val)
                break

    # ── save output ──────────────────────────────────────────────────────────
    import re as _re
    safe_label = _re.sub(r'[^\w\s\-]', '', ctx.account_label).strip().replace(' ', '_')
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    fname = f'{safe_label} - Account Strategy Analysis - {ts}.xlsm'
    fpath = os.path.join(output_dir, fname)
    wb.save(fpath)
    wb.close()
    return fpath
