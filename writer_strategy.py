"""
writer_strategy.py
──────────────────
Writes the Strategy Analysis output workbook.

What it does
────────────
1.  Reads the pre-analysis Databricks workbook via reader_databricks_strategy.
2.  Writes the Questionnaire Survey - AMZ tab (header fields, Salesforce data, Gong notes).
3.  Runs the strategy auto-flag logic and writes OK / PARTIAL / FLAG to column F
    of the STRATEGY OVERVIEW tab for every row that has a data signal.
    Rows with no automatable signal are left as-is (default OK in template).
4.  Reads back every row from STRATEGY OVERVIEW and copies the ones that are
    FLAG or PARTIAL into the Account Strategy _Analysis findings table.
    OK rows that belong to a section with at least one FLAG/PARTIAL are also
    included as context (marked OK).
5.  Writes header, grade, and interpretation to the Analysis tab.
6.  Writes the ChildASIN View tab.
7.  Saves the output .xlsm file.

Auto-flag logic (column F, STRATEGY OVERVIEW)
─────────────────────────────────────────────
Row  3  ACoS target above constraint by >5pp                         → FLAG
Row  4  ACoS consistently decreasing, changes ≥2 in 30 days         → PARTIAL
Row  6  ACoS being loosened (direction = increasing)                 → FLAG
Row  8  Account has OOB events in period                             → FLAG
Row 20  Account has OOB events in period (budget constraint note)    → FLAG
Row 29  Framework compliance gap: Imported+NonQT > 40% of spend     → FLAG
Row 29  Framework compliance gap: Imported+NonQT 20–40%             → PARTIAL
Row 30  SPT present (structure review)                               → PARTIAL
Row 31  SPT present (narrowing to best sellers)                      → PARTIAL
Row 32  ATM < 5% of spend (underweighted)                           → FLAG
Row 32  ATM 5–10% of spend                                          → PARTIAL
Row 37  BA present, review slow movers                               → PARTIAL
Row 39  BA campaign count < 2 (not developed by category)           → FLAG
Row 45  No SB spend at all                                          → FLAG
Row 62  No CAT_SP standard naming detected                          → FLAG
Row 62  CAT_P non-standard naming detected                          → PARTIAL
Row 63  No SBV campaigns                                            → FLAG
Row 82  No SD campaigns and SD impressions = 0                      → FLAG
Row 83  No SD_ATC / ProSuite ATC campaign                           → PARTIAL
Row 86  Portfolios present but 0 managed                            → PARTIAL
Row 86  0 portfolios with budget cap and >3 portfolios              → FLAG
Row 92  No RBO configured                                           → PARTIAL
Row 94  Campaigns not in portfolio > 0                              → FLAG
Row 95  Campaigns not in portfolio > 0 (rename/unmanaged note)      → FLAG
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

def _compute_flags(ctx: StrategyContext) -> dict[int, str]:
    """
    Returns {row_number: 'FLAG'|'PARTIAL'} for STRATEGY OVERVIEW rows
    that have a clear data signal from the Databricks file.
    Rows not returned keep column G at its template default (OK).
    """
    flags: dict[int, str] = {}

    def flag(row: int, level: str):
        if flags.get(row) == 'FLAG':  # never downgrade
            return
        flags[row] = level

    # ── ACoS and Target (rows 2–8) ────────────────────────────────────────────

    # Row 3: ACoS target above constraint — needs to close the gap
    if ctx.acos_constraint > 0 and ctx.acos_gap_to_constraint > 5:
        flag(3, 'FLAG')
    elif ctx.acos_constraint > 0 and ctx.acos_gap_to_constraint > 2:
        flag(3, 'PARTIAL')

    # Row 4: ACoS decreasing consistently — surface as positive momentum
    if ctx.acos_direction == 'decreasing' and ctx.acos_changes_30d >= 2:
        flag(4, 'PARTIAL')

    # Row 6: ACoS being loosened (increasing) to chase spend
    if ctx.acos_direction == 'increasing':
        flag(6, 'FLAG')

    # Rows 8 + 20: Account hit OOB in period
    if ctx.has_oob:
        flag(8, 'FLAG')
        flag(20, 'FLAG')

    # ── Live Strategy (rows 29–52) ────────────────────────────────────────────

    # Row 29: Framework gap — Imported + Non-Quartile spend
    non_qt_total = ctx.pct_imported + ctx.pct_non_quartile
    if non_qt_total > 0.40:
        flag(29, 'FLAG')
    elif non_qt_total > 0.20:
        flag(29, 'PARTIAL')

    # Rows 30 + 31: SPT active — review structure and narrow to best sellers
    if ctx.spend_spt > 0:
        flag(30, 'PARTIAL')
        flag(31, 'PARTIAL')

    # Row 32: ATM underweighted vs total spend
    if ctx.pct_atm < 0.03:
        flag(32, 'FLAG')
    elif ctx.pct_atm < 0.08:
        flag(32, 'PARTIAL')

    # Row 37: BA present — slow movers review
    if ctx.spend_ba > 0:
        flag(37, 'PARTIAL')

    # Row 39: BA active but fewer than 2 campaigns (no category segmentation)
    if 0 < ctx.ba_campaign_count < 2:
        flag(39, 'FLAG')

    # Row 45: No SB spend at all (SBV alone does not count)
    if ctx.spend_sb == 0:
        flag(45, 'FLAG')

    # ── Campaigns and Structure (rows 53–85) ─────────────────────────────────

    # Row 62: CAT_SP naming standard
    if not ctx.has_cat_sp and not ctx.has_cat_non_standard:
        flag(62, 'FLAG')
    elif ctx.has_cat_non_standard and not ctx.has_cat_sp:
        flag(62, 'PARTIAL')

    # Row 63: No SBV campaigns
    if not ctx.has_sbv and ctx.spend_sbv == 0:
        flag(63, 'FLAG')

    # Row 82: No SD spend and no SD impressions
    if not ctx.has_sd and ctx.sd_impressions == 0 and ctx.spend_sd == 0:
        flag(82, 'FLAG')

    # Row 83: No ATC / SD_FLEX retargeting campaigns
    has_atc = any(re.search(r'\bATC\b|SD_FLEX', n, re.IGNORECASE) for n in ctx.campaign_names)
    if not has_atc:
        flag(83, 'PARTIAL')

    # ── Account Management (rows 86–107) ─────────────────────────────────────

    # Row 86: Portfolio governance — no managed portfolios or no budget caps
    if ctx.portfolio_count > 3 and ctx.portfolios_with_budget_cap == 0:
        flag(86, 'FLAG')
    elif ctx.portfolio_count > 0 and ctx.managed_portfolio_count == 0:
        flag(86, 'PARTIAL')

    # Row 92: No RBO configured
    if not ctx.has_rbo:
        flag(92, 'PARTIAL')

    # Rows 94 + 95: Campaigns sitting outside portfolios
    if ctx.campaigns_not_in_portfolio > 0:
        flag(94, 'FLAG')
        flag(95, 'FLAG')

    return flags


# ── dynamic What We Saw text ──────────────────────────────────────────────────

def _build_what_we_saw(ctx: StrategyContext, flags: dict[int, str]) -> dict[int, str]:
    """
    Returns {row_number: text} with plain-language What We Saw sentences
    built from real account numbers, for every row that fired a flag.
    Only rows in `flags` are included — unflagged rows keep the template text.
    """
    texts: dict[int, str] = {}

    def pct(v: float) -> str:
        return f'{v:.0%}'

    def dollar(v: float) -> str:
        return f'${v:,.0f}'

    # Row 3 — ACoS target vs constraint
    if 3 in flags:
        texts[3] = (
            f'The current ACoS target is {ctx.acos_current_target:.0f}%. '
            f'The account constraint is {ctx.acos_constraint:.0f}%. '
            f'The gap is +{ctx.acos_gap_to_constraint:.0f} percentage points. '
            f'The target needs to come down to align with the client objective.'
        )

    # Row 4 — ACoS consistently decreasing
    if 4 in flags:
        texts[4] = (
            f'ACoS target has been reduced {ctx.acos_changes_30d} times in the last 30 days. '
            f'Current target is {ctx.acos_current_target:.0f}%. '
            f'The account is trending in the right direction. '
            f'Keep decreasing in line with framework governance.'
        )

    # Row 6 — ACoS increasing
    if 6 in flags:
        texts[6] = (
            f'ACoS target has been increased in the last 30 days ({ctx.acos_changes_30d} changes). '
            f'Current target is {ctx.acos_current_target:.0f}%. '
            f'Spend growth is being driven by loosening efficiency targets '
            f'rather than through campaign structure improvements.'
        )

    # Row 8 — OOB with ACoS action
    if 8 in flags:
        texts[8] = (
            f'The account has hit daily budget limits during the period. '
            f'Current ACoS target is {ctx.acos_current_target:.0f}% '
            f'against a constraint of {ctx.acos_constraint:.0f}%. '
            f'Reducing the target would lower CPC exposure and ease OOB pressure.'
        )

    # Row 20 — OOB budget constraint
    if 20 in flags:
        texts[20] = (
            f'The account ran out of budget on at least one day during the period. '
            f'Total spend was {dollar(ctx.total_spend)}. '
            f'Budget expansion or scope reduction should be reviewed with the client.'
        )

    # Row 29 — Framework gap
    if 29 in flags:
        non_qt = ctx.pct_imported + ctx.pct_non_quartile
        texts[29] = (
            f'{pct(non_qt)} of total spend is running through Imported or Non-Quartile campaigns '
            f'({pct(ctx.pct_imported)} Imported, {pct(ctx.pct_non_quartile)} Non-Quartile). '
            f'The account is not operating within the expected Quartile framework.'
        )

    # Row 30 — SPT structure
    if 30 in flags:
        texts[30] = (
            f'SPT campaigns are active with {dollar(ctx.spend_spt)} spend in the period '
            f'({pct(ctx.pct_spt)} of total). '
            f'The defensive structure should be reviewed by category or brand segment.'
        )

    # Row 31 — SPT narrow to best sellers
    if 31 in flags:
        texts[31] = (
            f'SPT campaigns cover {dollar(ctx.spend_spt)} in spend. '
            f'Coverage should be narrowed to the strongest-selling products only '
            f'to improve defensive efficiency.'
        )

    # Row 32 — ATM underweighted
    if 32 in flags:
        texts[32] = (
            f'ATM campaigns represent {pct(ctx.pct_atm)} of total spend ({dollar(ctx.spend_atm)}). '
            + ('No ATM spend detected. ' if ctx.pct_atm == 0 else '')
            + f'Automatic targeting on best-selling ASINs should be expanded.'
        )

    # Row 37 — BA slow movers
    if 37 in flags:
        texts[37] = (
            f'BA campaigns are active with {dollar(ctx.spend_ba)} spend '
            f'({pct(ctx.pct_ba)} of total). '
            f'The campaign should be reviewed to remove slow-moving products '
            f'and focus on best sellers and mid sellers.'
        )

    # Row 39 — BA not segmented by category
    if 39 in flags:
        texts[39] = (
            f'Only {ctx.ba_campaign_count} BA campaign(s) detected. '
            f'BA structure is not developed by category or product type. '
            f'New BA campaigns by category are needed to improve discovery coverage.'
        )

    # Row 45 — No SB spend
    if 45 in flags:
        texts[45] = (
            f'No Sponsored Brands spend detected in the period. '
            + (f'SBV is active ({dollar(ctx.spend_sbv)}) but SB campaigns are absent. '
               if ctx.spend_sbv > 0 else '')
            + f'SB campaigns should be launched to build upper-funnel coverage.'
        )

    # Row 62 — CAT_SP naming
    if 62 in flags:
        if ctx.has_cat_non_standard and not ctx.has_cat_sp:
            texts[62] = (
                f'Category-targeted campaigns are present but do not follow the CAT_SP_ naming standard. '
                f'Campaigns should be renamed to the correct convention so the system can manage them properly.'
            )
        else:
            texts[62] = (
                f'No CAT_SP campaigns detected. '
                f'Category-targeted Sponsored Products should be launched for key products '
                f'to capture category-level demand.'
            )

    # Row 63 — SBV missing
    if 63 in flags:
        texts[63] = (
            f'No SBV campaigns detected in the period. '
            f'Sponsored Brands Video is not active. '
            f'SBV should be tested for product-page defense and category targeting.'
        )

    # Row 82 — No SD
    if 82 in flags:
        texts[82] = (
            f'No Sponsored Display campaigns are active. '
            f'SD impressions in the period: 0. '
            f'Product-view remarketing and audience retargeting are not running.'
        )

    # Row 83 — No ATC / SD_FLEX
    if 83 in flags:
        texts[83] = (
            f'No SD_FLEX or ATC retargeting campaigns detected. '
            f'Add-to-cart retargeting through Pro Suite AMC audiences is not in place. '
            + (f'SD is active ({dollar(ctx.spend_sd)}) but ATC targeting is missing.'
               if ctx.spend_sd > 0 else '')
        )

    # Row 86 — Portfolio governance
    if 86 in flags:
        texts[86] = (
            f'{ctx.portfolio_count} portfolios exist. '
            f'{ctx.managed_portfolio_count} are managed by Quartile. '
            f'{ctx.portfolios_with_budget_cap} have a budget cap. '
            f'Portfolio governance needs to be reviewed and tightened.'
        )

    # Row 92 — No RBO
    if 92 in flags:
        texts[92] = (
            f'No RBO rules are configured on this account. '
            f'Weekend bid management is not active. '
            f'An RBO rule should be considered to reduce bidding during lower-efficiency periods.'
        )

    # Row 94 — Campaigns outside portfolios
    if 94 in flags:
        texts[94] = (
            f'{ctx.campaigns_not_in_portfolio} campaign(s) are not assigned to any portfolio. '
            f'If portfolio management is the chosen governance model, '
            f'all active campaigns should be assigned consistently.'
        )

    # Row 95 — Unmanaged/renamed
    if 95 in flags:
        texts[95] = (
            f'{ctx.campaigns_not_in_portfolio} campaign(s) are outside the portfolio structure '
            f'and may not be managed by the system. '
            f'These should be reviewed and renamed to the correct Quartile naming convention.'
        )

    return texts


# ── grade calculator ──────────────────────────────────────────────────────────

def _calculate_grade(flags: dict[int, str]) -> tuple[str, str]:
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
    # TAB — STRATEGY OVERVIEW
    # col F = AUTO (rows the agent can evaluate from Databricks) / blank (manual only)
    # col G = OK / PARTIAL / FLAG (written by agent for AUTO rows, stays OK for others)
    # col C = What We Saw (rewritten with real numbers for AUTO rows)
    # ════════════════════════════════════════════════════════════════════════════
    ws_ov = wb['STRATEGY OVERVIEW']

    # Write flags to column G (OK/PARTIAL/FLAG) for auto rows
    # Column F (AUTO markers) is static in the template — not written at runtime
    for row_num, level in flags.items():
        ws_ov.cell(row=row_num, column=7, value=level)

    # Write dynamic What We Saw (col C) for rows that fired a flag
    dynamic_what = _build_what_we_saw(ctx, flags)
    for row_num, text in dynamic_what.items():
        ws_ov.cell(row=row_num, column=3, value=text)

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
        'AOV':               'AOV',
        'TAG 1':             'Tag1',
        'TAG 2':             'Tag2',
        'TAG 3':             'Tag3',
        'TAG 4':             'Tag4',
        'TAG 5':             'Tag5',
    }

    col22_map = {
        'PriceTier':  'PriceTier',
        'Brand':      'Brand',
        'Department': 'Department',
        'Category':   'Category',
    }

    header_to_col = {}
    for cell in ws3[2]:
        if cell.value:
            header_to_col[cell.value] = cell.column

    for row in ws3.iter_rows(min_row=3, max_col=ws3.max_column):
        for cell in row:
            cell.value = None

    for row_idx, rec in enumerate(asin_records, start=3):
        asin = rec.get('asin', '')
        cat  = cat_by_asin.get(asin, {})

        for header, col_idx in header_to_col.items():
            val = None
            if header in col14_map:
                val = rec.get(col14_map[header])
            elif header == 'Total Units Ordered':
                total_sales = rec.get('TotalSales') or 0
                aov         = rec.get('AOV') or 0
                val = math.ceil(total_sales / aov) if aov else None
            elif header == 'Ad Sales (%)':
                ad_s  = rec.get('AdSales') or 0
                tot_s = rec.get('TotalSales') or 0
                val   = round(ad_s / tot_s, 4) if tot_s else None
            elif header == 'Organic Sales (%)':
                ad_s  = rec.get('AdSales') or 0
                tot_s = rec.get('TotalSales') or 0
                val   = round(1 - (ad_s / tot_s), 4) if tot_s else None
            elif header == 'Buy Box%':
                raw = rec.get('Weighted_BuyBoxPercentage')
                val = round(raw / 100, 4) if raw is not None else None
            elif header in col22_map:
                val = cat.get(col22_map[header])

            if val is not None:
                ws3.cell(row=row_idx, column=col_idx, value=val)

        ws3.cell(row=row_idx, column=28, value=f'=SUM(O{row_idx}+Q{row_idx}+S{row_idx})')
        ws3.cell(row=row_idx, column=29, value=f'=SUM(P{row_idx}+R{row_idx}+T{row_idx}+U{row_idx}+V{row_idx}+W{row_idx}+X{row_idx}+Y{row_idx}+Z{row_idx}+AA{row_idx})')

    # ── conditional formatting — ChildASIN View ───────────────────────────────
    last_row = 2 + len(asin_records)

    def _dxf_font(hex_color):
        return DifferentialStyle(font=Font(color=hex_color, bold=False))

    def _cell_is_rule(operator, formula, hex_color, priority):
        rule = Rule(type='cellIs', operator=operator, formula=formula, priority=priority)
        rule.dxf = _dxf_font(hex_color)
        return rule

    ws3.conditional_formatting.add(f'L3:L{last_row}', _cell_is_rule('greaterThanOrEqual', ['1'],       'FF9C0006', 22))
    ws3.conditional_formatting.add(f'M3:M{last_row}', _cell_is_rule('lessThanOrEqual',    ['0'],       'FF9C0006', 20))
    ws3.conditional_formatting.add(f'O3:AA{last_row}',_cell_is_rule('equal',              ['0'],       'FFFF0000', 15))
    ws3.conditional_formatting.add(f'AD3:AD{last_row}',_cell_is_rule('equal',             ['"TAG MISSING"'], 'FF9C5700', 17))
    ws3.conditional_formatting.add(f'AE3:AE{last_row}',_cell_is_rule('equal',             ['"TAG MISSING"'], 'FF9C0006', 18))
    ws3.conditional_formatting.add(f'AE3:AH{last_row}',_cell_is_rule('equal',             ['0'],       'FF9C0006', 19))
    ws3.conditional_formatting.add(f'AE3:AH{last_row}',_cell_is_rule('equal',             ['"Opportunity"'], 'FF006100', 21))
    ws3.conditional_formatting.add(f'AI3:AM{last_row}',_cell_is_rule('equal',             ['"Unavailable"'], 'FF9C0006', 12))

    ws3.conditional_formatting.add(
        f'K3:K{last_row}',
        ColorScaleRule(
            start_type='min',        start_color='63BE7B',
            mid_type='percentile',   mid_value=50, mid_color='FFEB84',
            end_type='max',          end_color='F8696B',
        )
    )
    ws3.conditional_formatting.add(
        f'N3:N{last_row}',
        IconSetRule(icon_style='3Symbols2', type='percent', values=[0, 70, 85])
    )

    # ── save ──────────────────────────────────────────────────────────────────
    filename = f'{ctx.account_label} — Strategy Analysis {ctx.date_range}.xlsm'
    filename = re.sub(r'[<>:"/\\|?*]', '-', filename)
    out_path = os.path.join(output_dir, filename)
    wb.save(out_path)
    wb.close()
    print(f'[strategy] Saved: {out_path}')
    return out_path


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage: python writer_strategy.py <pre_analysis.xlsx> <template.xlsm> [output_dir]')
        sys.exit(1)
    write_strategy(
        sys.argv[1],
        sys.argv[2],
        sys.argv[3] if len(sys.argv) > 3 else '/mnt/user-data/outputs'
    )
