"""
reader_databricks_strategy.py
─────────────────────────────
Reads the Databricks pre-analysis workbook and returns a StrategyContext
object with all signals needed by rules_engine_strategy.py.

Tabs consumed
─────────────
02  Date Range KPIs          → account-level ACoS, TACoS, CPC, spend, sales
08  Campaign Report          → campaign names, subtypes, spend, portfolio assignment
10  Campaigns Grouped by QT  → spend % per campaign subtype (ATM, BAK, BA, SPT, WATM, etc.)
15  Campaign Perf by Parent  → ASIN tier data
24  ACoS Changes History     → change frequency + direction in last 30 days
25  Portfolio Insights       → portfolio names, IsManaged, has budget cap
33  RBO Configuration        → whether any RBO rules exist
34  Product Level ACoS       → whether product-level ACoS overrides exist
35  Campaign Level ACoS      → whether campaign-level ACoS overrides exist
36  Account Out of Budget    → whether account hit OOB in period
38  Client Success Insights  → ACoS constraint, account details
42  Amazon GGS/Domo          → GGS status, SD impressions, SB impressions
50  Promo Management Trends  → promo activity (discount, coupon, deal)
55  Salesforce Consolidated  → account name, launch date, MRR

All values are stored as plain Python types — no DataFrames escape this module.
Callers receive a single StrategyContext dataclass.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Optional

import pandas as pd


# ── helpers ──────────────────────────────────────────────────────────────────

def _find_header_row(ws, max_scan: int = 10) -> Optional[int]:
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), 1):
        non_empty = [c for c in row if c is not None]
        if len(non_empty) > 3:
            return i
    return None


def _tab_to_records(ws) -> list[dict]:
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
        rec = {headers[j]: row[j] for j in range(len(headers)) if headers[j] is not None}
        records.append(rec)
    return records


def _tab_to_dict(ws) -> dict:
    """Single-data-row tab → dict (first data row only)."""
    records = _tab_to_records(ws)
    return records[0] if records else {}


def _latest_record(records: list[dict], modstamp_key: str = 'SystemModstamp') -> dict:
    if not records:
        return {}
    if len(records) == 1:
        return records[0]
    col = next(
        (k for k in records[0].keys()
         if re.sub(r'[\s_]', '', k).lower() in ('systemmodstamp', 'modstamp')),
        None
    )
    if col:
        try:
            return max(records, key=lambda r: pd.to_datetime(r.get(col), errors='coerce') or pd.NaT)
        except Exception:
            pass
    return records[0]


def _no_data(ws) -> bool:
    """Returns True if tab has 'NO DATA AVAILABLE' message."""
    for row in ws.iter_rows(min_row=5, max_row=10, values_only=True):
        for cell in row:
            if cell and 'NO DATA' in str(cell).upper():
                return True
    return False


def _safe_float(val, default: float = 0.0) -> float:
    try:
        return float(val) if val is not None else default
    except (ValueError, TypeError):
        return default


def _safe_str(val, default: str = '') -> str:
    return str(val).strip() if val is not None else default


# ── context dataclass ─────────────────────────────────────────────────────────

@dataclass
class StrategyContext:
    # ── identity ──────────────────────────────────────────────────────────────
    account_label: str = ''
    tenant_id: str = ''
    profile_id: str = ''
    member_id: str = ''
    date_range: str = ''
    downloaded: str = ''

    # ── account-level KPIs (tab 02) ───────────────────────────────────────────
    acos_actual: float = 0.0          # e.g. 0.64 = 64%
    tacos_actual: float = 0.0
    cpc_current: float = 0.0
    cpc_last_year: float = 0.0
    total_spend: float = 0.0
    total_sales: float = 0.0
    ad_sales: float = 0.0
    yoy_ad_sales: float = 0.0         # e.g. -0.47 = -47%

    # ── constraint (tab 38) ───────────────────────────────────────────────────
    acos_constraint: float = 0.0      # e.g. 25 = 25%

    # ── ACoS change history (tab 24) ─────────────────────────────────────────
    acos_changes_30d: int = 0         # number of changes in last 30 days
    acos_direction: str = 'stable'    # 'decreasing', 'increasing', 'mixed'
    acos_current_target: float = 0.0  # most recent IACoS target
    acos_gap_to_constraint: float = 0.0  # current_target - constraint (in pp)

    # ── campaign type spend mix (tab 10) ─────────────────────────────────────
    # Each value is % of total spend (0.0–1.0)
    pct_imported: float = 0.0
    pct_non_quartile: float = 0.0
    pct_atm: float = 0.0
    pct_ba: float = 0.0
    pct_bak: float = 0.0
    pct_spt: float = 0.0
    pct_watm: float = 0.0
    pct_sb: float = 0.0
    pct_sbv: float = 0.0
    pct_sd: float = 0.0
    pct_br: float = 0.0
    pct_ow: float = 0.0
    pct_op: float = 0.0

    # absolute spend per type
    spend_imported: float = 0.0
    spend_non_quartile: float = 0.0
    spend_atm: float = 0.0
    spend_ba: float = 0.0
    spend_bak: float = 0.0
    spend_spt: float = 0.0
    spend_watm: float = 0.0
    spend_sb: float = 0.0
    spend_sbv: float = 0.0
    spend_sd: float = 0.0

    # ── campaign names (tab 08) ───────────────────────────────────────────────
    campaign_names: list[str] = field(default_factory=list)
    campaigns_not_in_portfolio: int = 0
    has_cat_sp: bool = False           # CAT_SP_ naming convention
    has_cat_non_standard: bool = False # CAT_ but not CAT_SP_
    has_sbv: bool = False
    has_sd: bool = False
    has_watm: bool = False
    has_catchall: bool = False
    ba_campaign_count: int = 0
    unmanaged_campaign_count: int = 0  # tab 31

    # ── portfolio (tab 25) ────────────────────────────────────────────────────
    portfolio_count: int = 0
    managed_portfolio_count: int = 0
    portfolios_with_budget_cap: int = 0
    portfolio_names: list[str] = field(default_factory=list)

    # ── signals from presence/absence tabs ───────────────────────────────────
    has_rbo: bool = False              # tab 33 has data
    has_product_acos_overrides: bool = False   # tab 34
    has_campaign_acos_overrides: bool = False  # tab 35
    has_oob: bool = False              # tab 36

    # ── GGS / display (tab 42) ───────────────────────────────────────────────
    ggs_status: str = 'No'             # 'Yes' or 'No'
    sd_impressions: int = 0
    sb_impressions: int = 0

    # ── promo (tab 50) ────────────────────────────────────────────────────────
    has_active_promo: bool = False     # any PromotionDiscount > 0 in period

    # ── ASIN tiers (tab 15) ───────────────────────────────────────────────────
    tier1_asin_count: int = 0         # TIER 10–30
    tier1_with_atm: int = 0           # Tier1 ASINs that have ATM spend


# ── main reader ───────────────────────────────────────────────────────────────

def read_strategy_context(pre_analysis_path: str) -> StrategyContext:
    import openpyxl
    pa = openpyxl.load_workbook(pre_analysis_path, data_only=True, read_only=True)
    ctx = StrategyContext()

    # ── identity (tab 01) ────────────────────────────────────────────────────
    ws01 = pa['01_Advertiser_Name']
    account_str = date_range = downloaded = ''
    for row in ws01.iter_rows(min_row=1, max_row=4, values_only=True):
        for cell in row:
            if cell and isinstance(cell, str):
                if 'Account:' in cell:
                    account_str = cell
                elif 'Date Range:' in cell:
                    date_range = cell.replace('Date Range: ', '').strip()
                elif 'Downloaded:' in cell:
                    downloaded = cell.replace('Downloaded: ', '').strip()
    m = re.match(
        r'Account:\s*(.+?)\s*\|\s*Tenant ID:\s*(\S+)\s*\|\s*Account ID:\s*(\S+)',
        account_str
    )
    if m:
        ctx.account_label = m.group(1).strip()
        ctx.tenant_id     = m.group(2).strip()
        ctx.profile_id    = m.group(3).strip()
        ctx.member_id     = ctx.account_label.split(' - ')[0].strip()
    ctx.date_range  = date_range
    ctx.downloaded  = downloaded

    # ── tab 02 — account KPIs ────────────────────────────────────────────────
    d02 = _tab_to_dict(pa['02_Date_Range_KPIs__Date_Range_'])
    ctx.acos_actual   = _safe_float(d02.get('ACoS'))
    ctx.tacos_actual  = _safe_float(d02.get('TACoS'))
    ctx.cpc_current   = _safe_float(d02.get('CPC'))
    ctx.cpc_last_year = _safe_float(d02.get('LastYear_CPC'))
    ctx.total_spend   = _safe_float(d02.get('AdSpend'))
    ctx.total_sales   = _safe_float(d02.get('TotalSales'))
    ctx.ad_sales      = _safe_float(d02.get('AdSales'))
    ctx.yoy_ad_sales  = _safe_float(d02.get('YoY_AdSales'))

    # ── tab 38 — constraint ──────────────────────────────────────────────────
    d38_all = _tab_to_records(pa['38_Client_Success_Insights_Repo'])
    d38 = _latest_record(d38_all)
    ctx.acos_constraint = _safe_float(d38.get('ACOS_Constraint__c'))

    # ── tab 24 — ACoS change history ─────────────────────────────────────────
    changes = _tab_to_records(pa['24_Account_ACoS_Changes_History'])
    cutoff = datetime.now() - timedelta(days=30)
    recent = []
    for r in changes:
        dt = r.get('Change_Date')
        if dt and hasattr(dt, 'date') and dt >= cutoff:
            recent.append(r)
    ctx.acos_changes_30d = len(recent)

    if changes:
        newest = changes[0]  # already sorted newest-first by Databricks
        ctx.acos_current_target = _safe_float(newest.get('IACoS_Percent'))

        # direction: look at last 5 changes
        last5 = changes[:5]
        if len(last5) >= 2:
            deltas = []
            for r in last5:
                old = _safe_float(r.get('Old_IACoS_Target'))
                new = _safe_float(r.get('IACoS_Percent'))
                deltas.append(new - old)
            n_dec = sum(1 for d in deltas if d < 0)
            n_inc = sum(1 for d in deltas if d > 0)
            if n_dec >= 4:
                ctx.acos_direction = 'decreasing'
            elif n_inc >= 4:
                ctx.acos_direction = 'increasing'
            else:
                ctx.acos_direction = 'mixed'

    ctx.acos_gap_to_constraint = (
        ctx.acos_current_target - ctx.acos_constraint
        if ctx.acos_constraint > 0 else 0.0
    )

    # ── tab 10 — campaign type mix ───────────────────────────────────────────
    subtypes = _tab_to_records(pa['10_Campaigns_Grouped_by_QT_Camp'])
    subtype_map: dict[str, dict] = {}
    for r in subtypes:
        st = _safe_str(r.get('CampaignSubType')).upper()
        subtype_map[st] = r

    def _pct(key: str) -> float:
        return _safe_float(subtype_map.get(key.upper(), {}).get('Perc_Spend'))

    def _spend(key: str) -> float:
        return _safe_float(subtype_map.get(key.upper(), {}).get('Spend'))

    ctx.pct_imported      = _pct('Imported')
    ctx.pct_non_quartile  = _pct('Non-Quartile')
    ctx.pct_atm           = _pct('ATM')
    ctx.pct_ba            = _pct('BA')
    ctx.pct_bak           = _pct('BAK')
    ctx.pct_spt           = _pct('SPT')
    ctx.pct_watm          = _pct('WATM')
    ctx.pct_sb            = _pct('SB')
    ctx.pct_sbv           = _pct('SBV')
    ctx.pct_sd            = _pct('SD')
    ctx.pct_br            = _pct('BR')
    ctx.pct_ow            = _pct('OW')
    ctx.pct_op            = _pct('OP')

    ctx.spend_imported     = _spend('Imported')
    ctx.spend_non_quartile = _spend('Non-Quartile')
    ctx.spend_atm          = _spend('ATM')
    ctx.spend_ba           = _spend('BA')
    ctx.spend_bak          = _spend('BAK')
    ctx.spend_spt          = _spend('SPT')
    ctx.spend_watm         = _spend('WATM')
    ctx.spend_sb           = _spend('SB')
    ctx.spend_sbv          = _spend('SBV')
    ctx.spend_sd           = _spend('SD')

    # ── tab 08 — campaign names ───────────────────────────────────────────────
    camp_records = _tab_to_records(pa['08_Campaign_Report'])
    names = [_safe_str(r.get('CampaignName')) for r in camp_records if r.get('CampaignName')]
    ctx.campaign_names = names

    ctx.campaigns_not_in_portfolio = sum(
        1 for r in camp_records
        if _safe_str(r.get('PortfolioName')).startswith('Campaign Not in Portfolio')
    )
    ctx.has_cat_sp           = any(re.search(r'\bCAT_SP', n, re.IGNORECASE) for n in names)
    ctx.has_cat_non_standard = any(re.search(r'\bCAT_', n, re.IGNORECASE) for n in names) and not ctx.has_cat_sp
    ctx.has_sbv              = any(re.search(r'\bSBV', n, re.IGNORECASE) for n in names)
    ctx.has_sd               = any(re.search(r'\bSD_', n, re.IGNORECASE) for n in names)
    ctx.has_watm             = any(re.search(r'\bWATM', n, re.IGNORECASE) for n in names)
    ctx.has_catchall         = any(re.search(r'catch.?all|WATM', n, re.IGNORECASE) for n in names)
    ctx.ba_campaign_count    = sum(
        1 for r in camp_records
        if _safe_str(r.get('CampaignSubType')).upper() == 'BA'
    )

    # ── tab 25 — portfolios ───────────────────────────────────────────────────
    port_records = _tab_to_records(pa['25_Portfolio_Insights_and_Confi'])
    ctx.portfolio_count    = len(port_records)
    ctx.managed_portfolio_count = sum(
        1 for r in port_records if r.get('IsManaged') is True
    )
    ctx.portfolios_with_budget_cap = sum(
        1 for r in port_records if r.get('IsBudgetCap') is True
    )
    ctx.portfolio_names = [_safe_str(r.get('Portfolio_Name')) for r in port_records]

    # ── tabs 33/34/35/36 — presence/absence signals ───────────────────────────
    ctx.has_rbo                    = not _no_data(pa['33_RBO_Configuration_Insights'])
    ctx.has_product_acos_overrides = not _no_data(pa['34_Product_Level_ACoS'])
    ctx.has_campaign_acos_overrides = not _no_data(pa['35_Campaign_Level_ACoS'])
    ctx.has_oob                    = not _no_data(pa['36_Account_Out_of_Budget'])

    # ── tab 42 — GGS ─────────────────────────────────────────────────────────
    ggs_records = _tab_to_records(pa['42_Amazon_GGS_Domo'])
    if ggs_records:
        ggs_vals = [_safe_str(r.get('Amazon GGS')) for r in ggs_records]
        ctx.ggs_status = 'Yes' if any(v == 'Yes' for v in ggs_vals) else 'No'
        ctx.sd_impressions = sum(
            int(_safe_float(r.get('Impressions')))
            for r in ggs_records
            if _safe_str(r.get('CampaignType')) == 'Sponsored Display'
        )
        ctx.sb_impressions = sum(
            int(_safe_float(r.get('Impressions')))
            for r in ggs_records
            if _safe_str(r.get('CampaignType')) == 'Sponsored Brands'
        )

    # ── tab 50 — promo ────────────────────────────────────────────────────────
    promo_records = _tab_to_records(pa['50_Promo_Management___Account_T'])
    ctx.has_active_promo = any(
        _safe_float(r.get('PromotionDiscount')) > 0
        for r in promo_records
    )

    # ── tab 15 — ASIN tiers ───────────────────────────────────────────────────
    asin_records = _tab_to_records(pa['15_Campaign_Performance_by_PARE'])
    tier1 = [r for r in asin_records if _safe_str(r.get('Tier')) in ('TIER 10', 'TIER 20', 'TIER 30')]
    ctx.tier1_asin_count = len(tier1)
    ctx.tier1_with_atm   = sum(
        1 for r in tier1 if _safe_float(r.get('OP_Spend')) > 0  # proxy: OP_Spend present = ATM active
    )

    pa.close()
    return ctx
