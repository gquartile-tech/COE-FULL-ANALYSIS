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
14  Campaign Perf by Child ASIN → slow mover detection (<3 orders), ATM/BA overlap, SPT on Tier 100
15  Campaign Perf by Parent   → parent ASIN count (single-ASIN account detection)
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
    # NEW: additional campaign presence signals
    has_bak: bool = False              # any BAK campaign active
    has_op: bool = False               # any OP (product target) campaign active
    has_sd_prd: bool = False           # any SD_PRD campaign active
    has_vcpm: bool = False             # any VCPM campaign active
    watm_campaign_count: int = 0       # number of distinct WATM campaigns
    sbv_naming_compliant: bool = True  # all SBV campaigns use SBV_ prefix

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
    promo_asin_count: int = 0          # number of ASINs with active promo
    promo_cost_rate: float = 0.0       # avg PromoCostRate_pct over last 4 weeks

    # ── Pro Suite audiences (tab 51) ─────────────────────────────────────────
    has_prosuite_audiences: bool = False      # any campaign has HasAudience=True
    prosuite_audience_spend_pct: float = 0.0  # share of total spend with audience

    # ── SnS and Promo Management (tab 50) ────────────────────────────────────
    has_sns_active: bool = False       # SnS subscriptions active (ActiveSubscriptions > 0)
    has_promo_portfolio: bool = False  # portfolio named SD_QTL_AMZ or Promo_ exists (GGS gate)

    # ── TACoS constraint (tab 38) ─────────────────────────────────────────────
    tacos_constraint: float = 0.0      # 0 = not documented

    # ── ASIN tiers (tab 15) ───────────────────────────────────────────────────
    tier1_asin_count: int = 0         # TIER 10–30
    tier1_with_atm: int = 0           # Tier1 ASINs that have ATM spend

    # ── Tab 14 ASIN-level derived signals ─────────────────────────────────────
    slow_movers_with_ba: int = 0       # ASINs with <3 orders AND BA_Spend > 0
    slow_mover_asins_with_ba: list[str] = field(default_factory=list)   # ASIN IDs matching above
    slow_mover_asins_with_atm: list[str] = field(default_factory=list)  # ASINs with <3 orders AND ATM_Spend > 0
    tier100_with_spt_asins: list[str] = field(default_factory=list)     # Tier 100 ASINs with SPT_Spend > 0
    max_asin_orders_30d: float = 0.0   # highest order count for any single ASIN in period
    atm_ba_overlap_count: int = 0      # ASINs with ATM_Spend > 0 AND BA_Spend > 0 AND Orders > 80
    atm_ba_overlap_asins: list[str] = field(default_factory=list)       # ASIN IDs + orders for S012 what_we_saw
    spt_slow_mover_pct: float = 0.0    # pct of SPT spend on ASINs with <3 orders
    spt_avg_acos: float = 0.0          # spend-weighted avg ACoS across SPT campaigns (tab 08)
    bak_campaigns: list[dict] = field(default_factory=list)             # [{name, spend, pct_of_total, acos}]
    catalog_asin_count: int = 0        # total ASINs in tab 14 (catalog size)
    spending_asin_count: int = 0       # ASINs with AdSpend > 0 in period
    low_order_campaign_count: int = 0  # campaigns with 1-3 orders in 30d

    # ── Tab 08 campaign-type performance ─────────────────────────────────────
    # avg ACoS by campaign naming prefix — for S053-S056, S064, S066
    atm_avg_acos: float = 0.0          # avg ACoS of ATM_ campaigns
    br_avg_acos: float = 0.0           # avg ACoS of BR_ campaigns
    ph_avg_acos: float = 0.0           # avg ACoS of PH_ campaigns
    ow_avg_acos: float = 0.0           # avg ACoS of OW_ campaigns
    br_campaign_count: int = 0         # count of BR_ campaigns
    ow_campaign_count: int = 0         # count of OW_ campaigns
    ph_campaign_count: int = 0         # count of PH_ campaigns
    has_both_watm_and_catchall: bool = False  # both WATM and CatchAll active simultaneously
    bak_name_overlaps_ba: bool = False  # at least one BAK campaign name matches a BA campaign name

    # ── Campaign-type outperforming signals (tab 08) ──────────────────────────
    sd_flex_avg_acos: float = 0.0      # avg ACoS of SD_FLEX_ campaigns
    sd_audi_avg_acos: float = 0.0      # avg ACoS of SD_AUDI_ campaigns
    sd_prd_avg_acos: float = 0.0       # avg ACoS of SD_PRD_ campaigns
    sb_avg_acos: float = 0.0           # avg ACoS of SB_ campaigns
    sbv_avg_acos: float = 0.0          # avg ACoS of SBV_ campaigns
    ow_avg_acos: float = 0.0           # avg ACoS of OW_ (exact match) campaigns
    op_avg_acos: float = 0.0           # avg ACoS of OP_ (product target) campaigns
    op_campaign_count: int = 0         # count of OP_ campaigns
    catchall_orders: float = 0.0       # total orders from CatchAll campaigns
    catsp_avg_acos: float = 0.0        # avg ACoS of CAT_SP_ campaigns

    # ── ACoS change cadence (tab 24) ──────────────────────────────────────────
    days_since_last_acos_change: int = 999  # days since most recent portal ACoS change
    # NEW: single-ASIN account flag
    parent_asin_count: int = 0        # unique parent ASINs in tab 15

    # ── CPC trend ─────────────────────────────────────────────────────────────
    # NEW: derived from tab 02 cpc_current vs cpc_last_year
    cpc_yoy_change_pct: float = 0.0   # positive = increased, negative = decreased

    # ── Portfolio coverage (tab 08) ───────────────────────────────────────────
    campaigns_in_portfolio_pct: float = 0.0   # share of campaigns assigned to a portfolio
    total_campaign_count: int = 0              # total campaign count from tab 08

    # ── Search term category mix (tab 12) ─────────────────────────────────────
    branded_spend_pct: float = 0.0      # Branded row Spend_Pct
    branded_acos: float = 0.0           # Branded row acos
    branded_cpc: float = 0.0            # Branded row cpc
    non_branded_spend_pct: float = 0.0  # Non Branded row Spend_Pct
    non_branded_acos: float = 0.0       # Non Branded row acos
    non_branded_cpc: float = 0.0        # Non Branded row cpc
    vcpm_spend_pct: float = 0.0         # VCPM row Spend_Pct (of total search term spend)

    # ── Monthly trend signals (tab 04 — L24M) ────────────────────────────────
    tacos_trend: str = 'stable'         # 'increasing', 'decreasing', 'stable'
    tacos_trend_pp: float = 0.0         # total pp change over last 3 months
    mom_spend_change: float = 0.0       # MoM spend change last month (decimal)
    mom_sales_change: float = 0.0       # MoM total sales change last month (decimal)
    l3m_tacos_avg: float = 0.0          # average TACoS over last 3 full months

    # ── Salesforce account attributes (tab 38) ────────────────────────────────
    primary_objective: str = ''          # e.g. 'Growth' | 'Profit Maximization (Efficiency, Margin, TACOS)' | etc.
    primary_spend_kpi: str = ''          # 'ACOS' | 'ROAS' | 'TACOS'
    repeat_purchase: str = ''            # 'High' | 'Medium' | 'Low'
    commodity_or_brand: str = ''         # 'Commodity' | 'Brand'
    sales_concentration: str = ''        # 'Low Concentration' | 'Medium Concentration' | 'High Concentration'
    tacos_constraint_documented: bool = False  # True only when tab 38 has a real TACoS_Constraint__c value

    # ── Tab 18 — category performance signals ────────────────────────────────
    qualifying_category_count: int = 0   # categories with AsinCount>=30 AND TotalSalesPct>=5% (CAT_SP gate)

    # ── Structural signals for new controls S126–S136 ────────────────────────
    # S126 — branded + NB both significant in same auto layer
    branded_nb_mixed_in_ba: bool = False  # branded_spend_pct > 0.20 AND non_branded_spend_pct > 0.20

    # S127 — auto-to-manual ratio (tab 10 derived)
    auto_spend_pct: float = 0.0           # (ATM + BA + WATM) as pct of total
    manual_exact_pct: float = 0.0         # BAK pct (primary manual exact layer)

    # S128 — BAK harvest stalled (tab 10 derived)
    bak_underfed: bool = False            # pct_bak > 0 but pct_bak < pct_ba * 0.10

    # S129 — own product page undefended (tab 10 + tab 08 derived)
    has_ow: bool = False                  # any OW_ campaign in campaign names
    ow_campaign_count: int = 0            # count of OW_ campaigns

    # S130 — BR discovery layer
    has_br: bool = False                  # any BR_ campaign active with spend

    # S131 — OW own-page coverage (already has_ow above)

    # S133 — BAK branded/NB mixed signal (tab 12 + tab 10 derived)
    bak_branded_nb_mixed: bool = False    # branded_spend_pct > 0.40 AND non_branded_spend_pct > 0.20

    # S134 — TACoS/ACoS divergence (already in trend fields above)

    # S136 — CatchAll graduation overdue (catchall_orders already exists)

    # ── Tab 17 top search terms (S105) ───────────────────────────────────────
    unconverted_top_terms: int = 0        # top-30 terms with orders≥3 AND CVR≥10% (proxy for not-yet-in-BAK)

    # ── Tab 14 inefficient ASIN spend (S107) ─────────────────────────────────
    inefficient_asin_count: int = 0       # ASINs: AdSpend>$100 AND ACoS>1.5x constraint AND orders<5
    paused_sb_count: int = 0             # SB campaigns with state=paused AND spend > 0
    paused_sbv_count: int = 0            # SBV campaigns with state=paused AND spend > 0

    # ── Tab 15 top-seller type coverage (S102) ────────────────────────────────
    top_seller_type_gaps: int = 0         # Tier 10-30 ASINs missing ≥2 of (ATM, BAK, OP)

    # ── Tab 08 BAK inefficiency signals (S053) ────────────────────────────────
    inefficient_bak_count: int = 0        # BAK campaigns: spend>$200 AND ACoS>1.5x constraint AND orders<5

    # ── Tab 08/10 BR inefficiency signals (S057) ─────────────────────────────
    br_inefficiency_flag: bool = False    # pct_br>0.15 AND br_avg_acos>1.5x constraint AND above_acos

    # ── Tab 38 monthly budget (S113) ─────────────────────────────────────────
    monthly_budget: float = 0.0           # Monthly_Budget__c from tab 38 (0 = not documented)

    # ── S039 — category segmentation gate ────────────────────────────────────
    categories_above_10pct: int = 0       # categories contributing >10% of total sales (tab 18)

    # ── S053/S054/S055 — campaign-level ACoS by type (tab 08) ────────────────
    # SP type (Sponsored Products)
    sp_worst_campaign_name: str = ''      # campaign with highest ACoS among SP-type
    sp_worst_campaign_acos: float = 0.0   # its ACoS (decimal)
    sp_campaigns_above_threshold: int = 0 # count of SP campaigns above threshold
    # SB type (Sponsored Brands)
    sb_worst_campaign_name: str = ''
    sb_worst_campaign_acos: float = 0.0
    sb_campaigns_above_threshold: int = 0
    # SD type (Sponsored Display)
    sd_worst_campaign_name: str = ''
    sd_worst_campaign_acos: float = 0.0
    sd_campaigns_above_threshold: int = 0

    # ── S101 — tagging and segmentation gap (tab 14) ─────────────────────────
    tags: list[str] = field(default_factory=list)  # all unique tag values (Tag1–Tag5) from tab 14

    # ── S109 — inefficient ASIN spend (revised logic) ────────────────────────
    # Fires when ASIN has spend AND (zero sales OR ACoS > 2× constraint)
    # inefficient_asin_count already exists above — reusing it with new logic
    inefficient_asin_names: list[str] = field(default_factory=list)  # ASIN IDs for what_we_saw

    # ── S087/S088/S089 — bulk campaign type spend shares ─────────────────────
    pct_cat_sp: float = 0.0               # CAT_SP_ spend as fraction of total


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

    # CPC YoY change — only compute when both values are present and non-zero
    if ctx.cpc_last_year > 0 and ctx.cpc_current > 0:
        ctx.cpc_yoy_change_pct = (ctx.cpc_current - ctx.cpc_last_year) / ctx.cpc_last_year

    # ── tab 38 — constraint + Salesforce account attributes ─────────────────
    d38_all = _tab_to_records(pa['38_Client_Success_Insights_Repo'])
    d38 = _latest_record(d38_all)
    ctx.acos_constraint  = _safe_float(d38.get('ACOS_Constraint__c'))
    ctx.tacos_constraint = _safe_float(d38.get('TACoS_Constraint__c'))
    ctx.tacos_constraint_documented = ctx.tacos_constraint > 0
    # Fallback: if TACoS constraint is not documented, use actual TACoS as reference
    if ctx.tacos_constraint == 0.0 and ctx.tacos_actual > 0:
        ctx.tacos_constraint = ctx.tacos_actual * 100
    # Salesforce account attributes
    ctx.primary_objective  = _safe_str(d38.get('Primary_Objective__c'))
    ctx.primary_spend_kpi  = _safe_str(d38.get('Primary_Spend_KPI__c'))
    ctx.repeat_purchase    = _safe_str(d38.get('Repeat_Purchase_Behavior__c'))
    ctx.commodity_or_brand = _safe_str(d38.get('Commodity_Products_or_Branded_Products__c'))
    ctx.sales_concentration = _safe_str(d38.get('Sales_Concentration__c'))
    # Monthly budget — for S113 budget alignment check
    ctx.monthly_budget = _safe_float(d38.get('Monthly_Budget__c'))

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
    # has_catchall: only match campaigns with 'catch all' or 'catchall' in the name — NOT WATM
    ctx.has_catchall         = any(re.search(r'catch.?all', n, re.IGNORECASE) for n in names)
    ctx.ba_campaign_count    = sum(
        1 for r in camp_records
        if _safe_str(r.get('CampaignSubType')).upper() == 'BA'
    )
    # NEW signals derived from tab 08
    ctx.has_bak = any(
        _safe_str(r.get('CampaignSubType')).upper() == 'BAK'
        for r in camp_records
    )
    ctx.has_op = any(re.search(r'\bOP_|\bOP\b', n, re.IGNORECASE) for n in names)
    ctx.has_sd_prd = any(re.search(r'\bSD_PRD\b', n, re.IGNORECASE) for n in names)
    ctx.has_vcpm = any(re.search(r'\bVCPM\b', n, re.IGNORECASE) for n in names)
    ctx.watm_campaign_count = sum(
        1 for n in names if re.search(r'\bWATM\b', n, re.IGNORECASE)
    )
    # BR discovery and OW own-page coverage signals
    ctx.has_br = any(
        _safe_str(r.get('CampaignSubType')).upper() == 'BR'
        for r in camp_records
    )
    ctx.has_ow = any(re.search(r'\bOW_', n, re.IGNORECASE) for n in names)
    ctx.ow_campaign_count = sum(
        1 for n in names if re.search(r'\bOW_', n, re.IGNORECASE)
    )
    # SBV naming compliance: all SBV campaigns should start with SBV_
    sbv_names = [n for n in names if re.search(r'\bSBV\b', n, re.IGNORECASE)]
    ctx.sbv_naming_compliant = all(
        re.match(r'SBV_', n, re.IGNORECASE) for n in sbv_names
    ) if sbv_names else True

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
    # GGS gate: detect SD portfolio commitment.
    # Portfolio names vary: 'SD_QTL_AMZ', 'SD QT AMZ', 'SD QTL', etc.
    ctx.has_promo_portfolio = any(
        ('SD' in str(n).upper() and ('QT' in str(n).upper() or 'AMZ' in str(n).upper()))
        or 'PROMO' in str(n).upper()
        for n in ctx.portfolio_names
    )

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
    # Count ASINs with active promo and average promo cost rate over last 4 weeks
    if promo_records:
        last4 = promo_records[-4:]
        ctx.promo_asin_count = max(
            int(_safe_float(r.get('ActivePromoASINs'))) for r in last4
        )
        rates = [_safe_float(r.get('PromoCostRate_pct')) for r in last4
                 if _safe_float(r.get('PromoCostRate_pct')) > 0]
        ctx.promo_cost_rate = sum(rates) / len(rates) if rates else 0.0
        # SnS signal: any row with ActiveSubscriptions > 0
        ctx.has_sns_active = any(
            _safe_float(r.get('ActiveSubscriptions')) > 0
            for r in promo_records
        )

    # ── tab 51 — Pro Suite audience performance ───────────────────────────────
    try:
        prosuite_records = _tab_to_records(pa['51_Pro_Suite__Audience_Performa'])
        if prosuite_records and not any(
            'NO DATA' in str(list(r.values())).upper() for r in prosuite_records[:1]
        ):
            total_ps_spend = sum(_safe_float(r.get('TotalSpend')) for r in prosuite_records)
            audience_spend = sum(
                _safe_float(r.get('TotalSpend'))
                for r in prosuite_records
                if r.get('HasAudience') in (True, 'True', 'true', 1, '1')
            )
            ctx.has_prosuite_audiences = audience_spend > 0
            ctx.prosuite_audience_spend_pct = (
                audience_spend / total_ps_spend if total_ps_spend > 0 else 0.0
            )
    except Exception:
        pass  # tab may not exist in older exports — safe default is False / 0.0

    # ── tab 15 — ASIN tiers ───────────────────────────────────────────────────
    asin_records = _tab_to_records(pa['15_Campaign_Performance_by_PARE'])
    tier1 = [r for r in asin_records if _safe_str(r.get('Tier')) in ('TIER 10', 'TIER 20', 'TIER 30')]
    ctx.tier1_asin_count = len(tier1)
    ctx.tier1_with_atm   = sum(
        1 for r in tier1 if _safe_float(r.get('OP_Spend')) > 0  # proxy: OP_Spend present = ATM active
    )
    # NEW: count unique parent ASINs across all tiers to detect single-ASIN accounts
    parent_asins = set(
        _safe_str(r.get('ParentASIN'))
        for r in asin_records
        if not pd.isna(r.get('ParentASIN')) and _safe_str(r.get('ParentASIN'))
    )
    ctx.parent_asin_count = len(parent_asins)

    # Top-seller type coverage gaps — for S102
    # Check each Tier 10-30 ASIN: missing ≥2 of (ATM_Spend, BAK_Spend, OP_Spend) → counts as gap
    gap_count = 0
    for r in tier1:
        missing = sum([
            _safe_float(r.get('ATM_Spend')) == 0,
            _safe_float(r.get('BAK_Spend')) == 0,
            _safe_float(r.get('OP_Spend'))  == 0,
        ])
        if missing >= 2:
            gap_count += 1
    ctx.top_seller_type_gaps = gap_count

    # ── tab 14 — ASIN-level derived signals ─────────────────────────────────────
    try:
        import pandas as _pd14
        xl14 = _pd14.ExcelFile(pre_analysis_path, engine='calamine')
        df14 = xl14.parse('14_Campaign_Performance_by_Adve', header=5)
        df14 = df14.loc[:, ~df14.columns.astype(str).str.match(r'^Unnamed')]
        df14 = df14.dropna(subset=['asin']) if 'asin' in df14.columns else df14

        if not df14.empty and 'asin' in df14.columns:
            tier_col     = 'Tier' if 'Tier' in df14.columns else None
            ba_col       = 'BA_Spend' if 'BA_Spend' in df14.columns else None
            atm_col      = 'ATM_Spend' if 'ATM_Spend' in df14.columns else None
            spt_col      = 'SPT_Spend' if 'SPT_Spend' in df14.columns else None
            orders_col   = 'Orders' if 'Orders' in df14.columns else None
            spend_col    = 'AdSpend' if 'AdSpend' in df14.columns else None

            # max orders per ASIN — for S014 (no top seller check)
            if orders_col:
                ctx.max_asin_orders_30d = float(df14[orders_col].fillna(0).max())

            # Global slow mover definition: < 3 orders in the period
            if orders_col:
                is_slow_mover = df14[orders_col].fillna(0) < 3

                # Slow movers with BA spend — for S010/S011/S037
                if ba_col is not None:
                    has_ba = df14[ba_col].fillna(0) > 0
                    slow_ba_mask = is_slow_mover & has_ba
                    ctx.slow_movers_with_ba = int(slow_ba_mask.sum())
                    ctx.slow_mover_asins_with_ba = list(df14.loc[slow_ba_mask, 'asin'].astype(str))

                # Slow movers with ATM spend — informational, used in what_we_saw
                if atm_col is not None:
                    has_atm = df14[atm_col].fillna(0) > 0
                    ctx.slow_mover_asins_with_atm = list(
                        df14.loc[is_slow_mover & has_atm, 'asin'].astype(str)
                    )

                # Slow movers with SPT spend — for S031
                if spt_col is not None:
                    has_spt = df14[spt_col].fillna(0) > 0
                    spt_total  = df14[spt_col].fillna(0).sum()
                    spt_slow   = df14.loc[is_slow_mover, spt_col].fillna(0).sum()
                    ctx.spt_slow_mover_pct = spt_slow / spt_total if spt_total > 0 else 0.0

            # Tier 100 ASINs with SPT spend — for S031 (no Tier 100 in SPT)
            if tier_col is not None and spt_col is not None:
                is_tier100 = df14[tier_col].astype(str).str.upper() == 'TIER 100'
                has_spt    = df14[spt_col].fillna(0) > 0
                ctx.tier100_with_spt_asins = list(
                    df14.loc[is_tier100 & has_spt, 'asin'].astype(str)
                )

            # ATM + BA overlap on high-velocity ASINs (>80 orders) — for S012/S013
            if atm_col is not None and ba_col is not None and orders_col is not None:
                has_atm_spend  = df14[atm_col].fillna(0) > 0
                has_ba_spend   = df14[ba_col].fillna(0) > 0
                high_velocity  = df14[orders_col].fillna(0) > 80
                overlap_mask   = has_atm_spend & has_ba_spend & high_velocity
                ctx.atm_ba_overlap_count = int(overlap_mask.sum())
                if ctx.atm_ba_overlap_count > 0:
                    ctx.atm_ba_overlap_asins = [
                        f"{row['asin']} ({int(row[orders_col])} orders)"
                        for _, row in df14.loc[overlap_mask].iterrows()
                    ]

            # Catalog vs spending ASIN counts — for S022
            ctx.catalog_asin_count  = int(df14['asin'].dropna().nunique())
            if spend_col is not None:
                ctx.spending_asin_count = int(
                    df14[df14[spend_col].fillna(0) > 0]['asin'].dropna().nunique()
                )

            # Inefficient ASIN detection — S109 (revised logic)
            # Fires when: AdSpend > 0 AND (no sales = ACoS is 0/null OR ACoS > 2× constraint)
            acos_col14 = 'ACoS' if 'ACoS' in df14.columns else None
            sales_col14 = 'AdSales' if 'AdSales' in df14.columns else ('TotalSales' if 'TotalSales' in df14.columns else None)
            if spend_col is not None and ctx.acos_constraint > 0:
                has_spend14 = df14[spend_col].fillna(0) > 0
                if acos_col14 is not None:
                    threshold14 = (ctx.acos_constraint / 100) * 2.0
                    acos_num14 = _pd14.to_numeric(df14[acos_col14], errors='coerce')
                    # No sales = acos is NaN or 0 when spend > 0
                    no_sales = has_spend14 & (acos_num14.isna() | (acos_num14 == 0))
                    above_2x = has_spend14 & (acos_num14.fillna(0) > threshold14)
                    ineff14 = no_sales | above_2x
                    ctx.inefficient_asin_count = int(ineff14.sum())
                    if ctx.inefficient_asin_count > 0 and 'asin' in df14.columns:
                        ctx.inefficient_asin_names = list(
                            df14.loc[ineff14, 'asin'].astype(str).dropna().unique()[:10]
                        )
                else:
                    # Fallback: any ASIN with spend but no orders
                    if orders_col is not None:
                        ineff14 = has_spend14 & (df14[orders_col].fillna(0) == 0)
                        ctx.inefficient_asin_count = int(ineff14.sum())

            # Tags extraction — for S101 (Tagging and Segmentation Gap)
            tag_cols = [c for c in df14.columns if re.match(r'^Tag\d+$', str(c), re.IGNORECASE)]
            if tag_cols:
                tag_values = set()
                for tc in tag_cols:
                    tag_values.update(
                        df14[tc].dropna().astype(str).str.strip()
                        .loc[lambda s: s != ''].unique()
                    )
                ctx.tags = [t for t in tag_values if t and t.lower() not in {'nan', 'none', ''}]
    except Exception:
        pass

    # ── tab 24 — days since last ACoS change ─────────────────────────────────────
    try:
        import pandas as _pd24
        from datetime import datetime as _dt
        xl24 = _pd24.ExcelFile(pre_analysis_path, engine='calamine')
        df24 = xl24.parse('24_Account_ACoS_Changes_History', header=5)
        df24 = df24.loc[:, ~df24.columns.astype(str).str.match(r'^Unnamed')]
        if not df24.empty and 'Change_Date' in df24.columns:
            dates = _pd24.to_datetime(df24['Change_Date'], errors='coerce').dropna()
            if not dates.empty:
                last_change = dates.max()
                if ctx.downloaded is not None:
                    ref = _pd24.Timestamp(ctx.downloaded)
                elif ctx.ref_date is not None:
                    ref = _pd24.Timestamp(ctx.ref_date)
                else:
                    ref = _pd24.Timestamp(_dt.now())
                ctx.days_since_last_acos_change = max(0, int((ref - last_change).days))
    except Exception:
        pass

    # ── tab 08 — campaign-type ACoS and structural signals ──────────────────────
    try:
        import pandas as _pd08
        xl08 = _pd08.ExcelFile(pre_analysis_path, engine='calamine')
        df08 = xl08.parse('08_Campaign_Report', header=5)
        df08 = df08.loc[:, ~df08.columns.astype(str).str.match(r'^Unnamed')]

        if not df08.empty and 'CampaignName' in df08.columns:
            df08['_name'] = df08['CampaignName'].astype(str).str.strip()
            df08['_name_up'] = df08['_name'].str.upper()
            acos_col08 = '_ACOS' if '_ACOS' in df08.columns else ('ACoS' if 'ACoS' in df08.columns else None)

            def _prefix_acos(prefix):
                mask = df08['_name_up'].str.startswith(prefix)
                if acos_col08 and mask.any():
                    vals = _pd08.to_numeric(df08.loc[mask, acos_col08], errors='coerce').dropna()
                    return float(vals.mean()) if not vals.empty else 0.0
                return 0.0

            def _prefix_count(prefix):
                return int(df08['_name_up'].str.startswith(prefix).sum())

            ctx.atm_avg_acos = _prefix_acos('ATM_')
            ctx.br_avg_acos  = _prefix_acos('BR_')
            ctx.ph_avg_acos  = _prefix_acos('PH_')
            ctx.ow_avg_acos  = _prefix_acos('OW_')
            ctx.br_campaign_count = _prefix_count('BR_')
            ctx.ow_campaign_count = _prefix_count('OW_')
            ctx.ph_campaign_count = _prefix_count('PH_')

            # Low-order campaign count — for S041
            orders_col08 = 'Orders' if 'Orders' in df08.columns else None
            spend_col08  = 'Spend'  if 'Spend'  in df08.columns else None
            if orders_col08:
                lo = df08[orders_col08].fillna(0)
                ctx.low_order_campaign_count = int(((lo >= 1) & (lo <= 3)).sum())

            # SPT spend-weighted avg ACoS — for S030
            if sub_col and acos_col08 and spend_col08:
                spt_mask = df08[sub_col].astype(str).str.upper() == 'SPT'
                if spt_mask.any():
                    spt_df   = df08.loc[spt_mask].copy()
                    spt_acos = _pd08.to_numeric(spt_df[acos_col08], errors='coerce').fillna(0)
                    spt_spnd = _pd08.to_numeric(spt_df[spend_col08], errors='coerce').fillna(0)
                    total_spt_spend = spt_spnd.sum()
                    if total_spt_spend > 0:
                        ctx.spt_avg_acos = float((spt_acos * spt_spnd).sum() / total_spt_spend)

            # BAK campaigns list — for S078 what_we_saw
            if sub_col and acos_col08 and spend_col08:
                bak_mask = df08[sub_col].astype(str).str.upper() == 'BAK'
                if bak_mask.any():
                    bak_df = df08.loc[bak_mask].copy()
                    total_spend_acct = float(
                        _pd08.to_numeric(df08[spend_col08], errors='coerce').fillna(0).sum()
                    )
                    for _, row in bak_df.iterrows():
                        sp = float(_pd08.to_numeric(row.get(spend_col08), errors='coerce') or 0)
                        ac = float(_pd08.to_numeric(row.get(acos_col08), errors='coerce') or 0)
                        ctx.bak_campaigns.append({
                            'name':         str(row.get('CampaignName', '')),
                            'spend':        sp,
                            'pct_of_total': sp / total_spend_acct if total_spend_acct > 0 else 0.0,
                            'acos':         ac,
                        })

            # Both WATM and CatchAll active — for S076
            sub_col = 'CampaignSubType' if 'CampaignSubType' in df08.columns else None
            has_watm08 = (
                (sub_col and (df08[sub_col].astype(str).str.upper() == 'WATM').any())
                or df08['_name_up'].str.startswith('WATM_').any()
            )
            has_catchall08 = df08['_name'].str.lower().str.contains(r'catch.?all', regex=True).any()
            ctx.has_both_watm_and_catchall = bool(has_watm08 and has_catchall08)

            # BAK name overlaps BA name — for S038
            if sub_col:
                ba_names  = set(df08.loc[df08[sub_col].astype(str).str.upper() == 'BA', '_name'].tolist())
                bak_names = set(df08.loc[df08[sub_col].astype(str).str.upper() == 'BAK', '_name'].tolist())
                # check if any BAK shares the same parent-ASIN token as a BA
                def _extract_token(n):
                    # e.g. BA_P_APRQ8M786493_... → APRQ8M786493
                    parts = str(n).split('_')
                    return parts[2] if len(parts) > 2 else n
                ba_tokens  = {_extract_token(n) for n in ba_names}
                bak_tokens = {_extract_token(n) for n in bak_names}
                ctx.bak_name_overlaps_ba = bool(ba_tokens & bak_tokens)

            # Campaign-type outperforming signals — new controls S034/S035/S042/S043/S057-S060
            def _pfx_acos(pfx):
                mask = df08['_name_up'].str.startswith(pfx)
                if acos_col08 and mask.any():
                    vals = _pd08.to_numeric(df08.loc[mask, acos_col08], errors='coerce').dropna()
                    return float(vals.mean()) if not vals.empty else 0.0
                return 0.0

            ctx.sd_flex_avg_acos = _pfx_acos('SD_FLEX_')
            ctx.sd_audi_avg_acos = _pfx_acos('SD_AUDI_')
            ctx.sd_prd_avg_acos  = _pfx_acos('SD_PRD_')
            ctx.sb_avg_acos      = _pfx_acos('SB_')
            ctx.sbv_avg_acos     = _pfx_acos('SBV_')
            ctx.ow_avg_acos      = _pfx_acos('OW_')
            ctx.op_avg_acos      = _pfx_acos('OP_')
            ctx.op_campaign_count = _prefix_count('OP_')
            ctx.catsp_avg_acos   = _pfx_acos('CAT_SP_')

            # CatchAll orders
            if orders_col08:
                ca_mask = df08['_name'].str.lower().str.contains(r'catch.?all', regex=True, na=False)
                ctx.catchall_orders = float(df08.loc[ca_mask, orders_col08].fillna(0).sum())

            # Paused SB / SBV campaigns with historical spend — for S064/S066
            if 'State' in df08.columns and spend_col08:
                paused_mask = df08['State'].astype(str).str.lower() == 'paused'
                spend_num   = _pd08.to_numeric(df08[spend_col08], errors='coerce').fillna(0)
                had_spend   = spend_num > 0
                sb_mask  = df08['_name_up'].str.startswith('SB_')
                sbv_mask = df08['_name_up'].str.startswith('SBV_')
                ctx.paused_sb_count  = int((paused_mask & had_spend & sb_mask  & ~sbv_mask).sum())
                ctx.paused_sbv_count = int((paused_mask & had_spend & sbv_mask).sum())

            # BAK inefficiency — for S053: spend>$200, ACoS>1.5x constraint, orders<5
            if sub_col and acos_col08 and spend_col08 and orders_col08:
                bak_mask2 = df08[sub_col].astype(str).str.upper() == 'BAK'
                if bak_mask2.any() and ctx.acos_constraint > 0:
                    bak_sp   = _pd08.to_numeric(df08[spend_col08],  errors='coerce').fillna(0)
                    bak_ac   = _pd08.to_numeric(df08[acos_col08],   errors='coerce').fillna(0)
                    bak_ord  = _pd08.to_numeric(df08[orders_col08], errors='coerce').fillna(0)
                    threshold = (ctx.acos_constraint / 100) * 1.5
                    ineff_mask = bak_mask2 & (bak_sp > 200) & (bak_ac > threshold) & (bak_ord < 5)
                    ctx.inefficient_bak_count = int(ineff_mask.sum())

            # Campaign-level ACoS by type — for S053 (SP), S054 (SB), S055 (SD)
            # "worst offending" = highest ACoS campaign above the constraint threshold
            if acos_col08 and spend_col08 and ctx.acos_constraint > 0:
                acos_num = _pd08.to_numeric(df08[acos_col08], errors='coerce').fillna(0)
                spend_num08 = _pd08.to_numeric(df08[spend_col08], errors='coerce').fillna(0)
                has_spend_mask = spend_num08 > 0

                def _worst_campaign(type_mask):
                    m = type_mask & has_spend_mask
                    if not m.any():
                        return '', 0.0, 0
                    subset = df08.loc[m].copy()
                    subset['_acos_n'] = _pd08.to_numeric(subset[acos_col08], errors='coerce').fillna(0)
                    threshold_35 = (ctx.acos_constraint / 100) * 1.35
                    threshold_20 = (ctx.acos_constraint / 100) * 1.20
                    above = subset[subset['_acos_n'] > threshold_20]
                    if above.empty:
                        return '', 0.0, 0
                    worst = above.loc[above['_acos_n'].idxmax()]
                    count_above = int((above['_acos_n'] > threshold_20).sum())
                    return (
                        str(worst.get('CampaignName', '')),
                        float(worst['_acos_n']),
                        count_above,
                    )

                # SP = Sponsored Products (not SB, SBV, SD)
                sp_mask = ~(
                    df08['_name_up'].str.startswith('SB_') |
                    df08['_name_up'].str.startswith('SBV_') |
                    df08['_name_up'].str.startswith('SD_')
                )
                sb_mask_t = df08['_name_up'].str.startswith('SB_') & ~df08['_name_up'].str.startswith('SBV_')
                sd_mask_t = df08['_name_up'].str.startswith('SD_')

                ctx.sp_worst_campaign_name, ctx.sp_worst_campaign_acos, ctx.sp_campaigns_above_threshold = _worst_campaign(sp_mask)
                ctx.sb_worst_campaign_name, ctx.sb_worst_campaign_acos, ctx.sb_campaigns_above_threshold = _worst_campaign(sb_mask_t)
                ctx.sd_worst_campaign_name, ctx.sd_worst_campaign_acos, ctx.sd_campaigns_above_threshold = _worst_campaign(sd_mask_t)

            # CAT_SP spend share — for S087
            if spend_col08:
                total_sp_08 = float(_pd08.to_numeric(df08[spend_col08], errors='coerce').fillna(0).sum())
                cat_sp_spend = float(_pd08.to_numeric(
                    df08.loc[df08['_name_up'].str.startswith('CAT_SP_'), spend_col08],
                    errors='coerce'
                ).fillna(0).sum())
                ctx.pct_cat_sp = cat_sp_spend / total_sp_08 if total_sp_08 > 0 else 0.0
    except Exception:
        pass

    # ── tab 08 — portfolio coverage ratio ─────────────────────────────────────
    # Already read campaign_names above; re-use those records for portfolio count
    camp_records_08 = _tab_to_records(pa['08_Campaign_Report'])
    ctx.total_campaign_count = len(camp_records_08)
    if ctx.total_campaign_count > 0:
        in_port = sum(
            1 for r in camp_records_08
            if not _safe_str(r.get('PortfolioName')).startswith('Campaign Not in Portfolio')
            and _safe_str(r.get('PortfolioName'))
        )
        ctx.campaigns_in_portfolio_pct = in_port / ctx.total_campaign_count

    # ── tab 12 — branded vs non-branded keyword mix ────────────────────────────
    try:
        search_cat_records = _tab_to_records(pa['12_Search_Terms_by_Category'])
        for r in search_cat_records:
            cat = _safe_str(r.get('KeywordCategory')).lower().strip()
            if cat == 'branded':
                ctx.branded_spend_pct  = _safe_float(r.get('Spend_Pct'))
                ctx.branded_acos       = _safe_float(r.get('acos'))
                ctx.branded_cpc        = _safe_float(r.get('cpc'))
            elif cat == 'non branded':
                ctx.non_branded_spend_pct = _safe_float(r.get('Spend_Pct'))
                ctx.non_branded_acos      = _safe_float(r.get('acos'))
                ctx.non_branded_cpc       = _safe_float(r.get('cpc'))
            elif cat == 'vcpm':
                ctx.vcpm_spend_pct = _safe_float(r.get('Spend_Pct'))
    except Exception:
        pass

    # ── tab 18 — category performance (CAT_SP qualification gate + S039) ───────
    try:
        cat18_records = _tab_to_records(pa['18_Performance_by_Category'])
        ctx.qualifying_category_count = sum(
            1 for r in cat18_records
            if _safe_float(r.get('AsinCount')) >= 30
            and _safe_float(r.get('TotalSalesPct')) >= 0.05
        )
        # S039: count categories each contributing >10% of total sales
        ctx.categories_above_10pct = sum(
            1 for r in cat18_records
            if _safe_float(r.get('TotalSalesPct')) >= 0.10
        )
    except Exception:
        pass

    # ── tab 17 — top 30 search terms (S105 proxy) ────────────────────────────
    try:
        top17_records = _tab_to_records(pa['17_Top_30_Search_Terms'])
        ctx.unconverted_top_terms = sum(
            1 for r in top17_records
            if _safe_float(r.get('orders')) >= 3
            and _safe_float(r.get('cr')) >= 0.10
        )
    except Exception:
        pass

    # ── derived structural signals for new controls ───────────────────────────
    # S126 — branded + NB both > 20% of search term spend in the same auto bucket
    ctx.branded_nb_mixed_in_ba = (
        ctx.branded_spend_pct > 0.20 and ctx.non_branded_spend_pct > 0.20
    )
    # S127 — auto spend dominance vs manual exact
    ctx.auto_spend_pct   = ctx.pct_atm + ctx.pct_ba + ctx.pct_watm
    ctx.manual_exact_pct = ctx.pct_bak
    # S128 — BAK harvest stalled: BAK exists but is <10% of BA spend
    ctx.bak_underfed = (
        ctx.pct_bak > 0 and ctx.pct_ba > 0
        and ctx.pct_bak < ctx.pct_ba * 0.10
    )
    # S133 — BAK bucket carries both heavy branded and NB (not split by type)
    ctx.bak_branded_nb_mixed = (
        ctx.branded_spend_pct > 0.40 and ctx.non_branded_spend_pct > 0.20
        and ctx.pct_bak > 0
    )
    # S057 — BR strategy too broad: BR consuming >15% of spend at poor efficiency
    ctx.br_inefficiency_flag = (
        ctx.pct_br > 0.15
        and ctx.acos_constraint > 0
        and ctx.br_avg_acos > (ctx.acos_constraint / 100) * 1.5
        and ctx.acos_actual * 100 > ctx.acos_constraint
    )

    # ── tab 04 — L24M monthly trend signals ──────────────────────────────────
    try:
        import pandas as _pd
        ws04 = pa['04_L24M_Monthly_Performance_Sum']
        rows04 = [r for r in ws04.iter_rows(min_row=7, values_only=True) if any(v is not None for v in r)]
        if rows04:
            # Build minimal df: col 0 = Month, col 2 = AdSpend, col 3 = TACoS, col 4 = TotalSales
            months, tacos_vals, spend_vals, sales_vals = [], [], [], []
            for r in rows04:
                m = r[0]; tc = r[3]; sp = r[2]; sl = r[1]
                if m is None:
                    continue
                try:
                    m_ts = _pd.to_datetime(m, errors='coerce')
                    if _pd.isna(m_ts):
                        continue
                except Exception:
                    continue
                months.append(m_ts)
                tacos_vals.append(float(tc) if tc is not None else None)
                spend_vals.append(float(sp) if sp is not None else None)
                sales_vals.append(float(sl) if sl is not None else None)

            # Sort by month and take last 3 non-None tacos values
            paired = sorted(zip(months, tacos_vals, spend_vals, sales_vals), key=lambda x: x[0])
            valid_tacos = [(t, s, sl) for _, t, s, sl in paired if t is not None]

            if len(valid_tacos) >= 3:
                last3_tacos  = [v[0] for v in valid_tacos[-3:]]
                last3_spend  = [v[1] for v in valid_tacos[-3:] if v[1] is not None]
                last3_sales  = [v[2] for v in valid_tacos[-3:] if v[2] is not None]

                ctx.l3m_tacos_avg = sum(last3_tacos) / len(last3_tacos)
                pp_change = (last3_tacos[-1] - last3_tacos[0]) * 100  # decimal → pp
                ctx.tacos_trend_pp = round(pp_change, 2)

                if pp_change > 1.5:
                    ctx.tacos_trend = 'increasing'
                elif pp_change < -1.5:
                    ctx.tacos_trend = 'decreasing'
                else:
                    ctx.tacos_trend = 'stable'

                if len(last3_spend) >= 2 and last3_spend[-2] > 0:
                    ctx.mom_spend_change = (last3_spend[-1] - last3_spend[-2]) / last3_spend[-2]
                if len(last3_sales) >= 2 and last3_sales[-2] > 0:
                    ctx.mom_sales_change = (last3_sales[-1] - last3_sales[-2]) / last3_sales[-2]

            # VCPM spend share from tab 12 (already parsed above — read from search_cat_records if available)
    except Exception:
        pass

    pa.close()
    return ctx
