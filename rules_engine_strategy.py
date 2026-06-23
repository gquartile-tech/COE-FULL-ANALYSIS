"""
rules_engine_strategy.py
────────────────────────
Evaluates all 126 strategy controls and returns:
  flags      : dict[str, str]   → {control_id: 'FLAG'|'PARTIAL'}
  what_we_saw: dict[str, str]   → {control_id: dynamic plain-language text}

Called by writer_strategy.py — replaces the inline _compute_flags and
_build_what_we_saw that previously lived there.

Design principles
─────────────────
- Every auto-flag requires a KPI condition, not just structural presence/absence.
- Framework pillar owns structural presence checks.
- Strategy owns TIMING and PRIORITY based on where the account is:
  ACoS vs constraint, TACoS trend, YoY growth, objective, spend efficiency.
- Strategy flags are suggestions, not penalties.
  A FLAG can represent a positive signal, a concern, or an action item.
- When two outperforming signals combine to form a higher-order suggestion
  (S036 = ATM + BR both outperforming → Discovery-Performance Mix),
  the component signals (S056 ATM, S057 BR) are forced to OK.
- Never flag what the Quartile system already manages:
  bids, budgets, negatives, automation settings.

Changelog (applied here relative to prior writer_strategy.py version)
──────────────────────────────────────────────────────────────────────
S021: OOB flag also fires when ACoS or TACoS is BELOW constraint (OOB on
      clean account = negotiate higher budget with client).
S036: Auto-flag when both ATM AND BR are outperforming by >20% — positive
      composite suggestion.  When S036 fires, S056 (ATM) and S057 (BR) are
      silenced (forced OK) to avoid duplicate recommendations.
S039: BA segmentation gap also fires when only 1 BA campaign exists AND
      multiple categories each account for >10% of total sales.
S045: BAK Harvest Stalled only flags when objective is Growth or Expansion
      (not meaningful for Profit Maximization / Maintenance / Recovery).
S053/S054/S055: Campaign-level ACoS checks at SP/SB/SD campaign type level
      (not account average). what_we_saw names the worst-offending campaign
      and notes any additional campaigns above threshold.
S071: SBV Product Targeting Launch also requires CAT_SP outperforming OR
      OP outperforming (same gate as S067 CAT_SP Launch).
S077: PARTIAL/FLAG inverted — PARTIAL when CAT_SP avg ACoS above constraint
      and below 85%; FLAG when above 85%.
S082: BAK Branded/NB Mixed — check whether branded search terms are heavy
      inside BAK bucket (tab 12 branded_spend_pct > 40% AND BAK active AND
      non_branded_spend_pct > 20%).
S092/S093: SD Remarketing/ATC — add objective filter (Growth/Expansion) and
      SD spend threshold ($1,000); case-insensitive name matching for
      SD_FLEX_Remarketing, SD_FLEX_rmkt, SD_AUDI.
S096: SD PDP Maturity — OK when top ASIN already has SD spend (sufficient
      audience pool already exists).
S097: Portfolio governance gate changed to >15% of campaigns in portfolios.
S098/S099: Campaign/product ACoS override checks — confirmed comparing
      against constraint (no logic change needed; already correct).
S101: Tagging/Segmentation Gap — full tag-label logic using ctx.tags
      (bestseller + performance-tier dimensions).
S109 (old S107): Inefficient ASIN Spend — auto rule from tab 14:
      AdSpend > $0 (any spend) AND (ACoS is None/zero = no sales, OR
      ACoS > 2× constraint).  ctx.inefficient_asin_names populated in reader.
S110 (old S108): SB active — SBV missing — threshold extended to 10%.
S113/S119 (old S111/S117): Subscribe & Save / Recurring Sales — also fires
      when repeat_purchase is High regardless of YoY direction.
S124 (old S122): GGS SD Compliance — check account settings only (remove
      portfolio name gate).
S010/S011/S037: Minimum slow mover count gate added. Only flag when slow_movers_with_ba
      reaches max(2, 10% of catalog size). Prevents single tail-ASIN noise on small accounts.
      S037 now also suppressed when S010 is already FLAG.
S010/S037/S109: Slow mover definition now uses total orders proxy (TotalSales / AOV)
      instead of ad orders. ASINs selling organically are no longer classified as slow movers.
S014: Only evaluates bulk accounts (BA ≥ 15% of spend). Non-bulk accounts do not run
      the BA/BAK methodology so structural gaps are not meaningful for them.
S020: Suppressed for growth/expansion objective on the TACoS path. New CPC path added:
      PARTIAL at +20% CPC YoY, FLAG at +40% CPC YoY.
S032: Minimum gate raised to 15 slow movers with SPT spend before flagging.
S053/S054/S055: Suppressed for growth/expansion objective — overspending campaigns
      are expected when the account is actively scaling.
S063/S064/S065: Suppressed when >50% of the subtype's spend is VCPM campaigns.
      VCPM uses impression-based billing; ACoS comparisons are not valid.
S075: Requires OP campaigns with actual spend in the period (op_campaigns_with_spend).
      Accounts with zero OP campaigns at all are a framework gap, not a strategy signal.
S101: Converted to MANUAL. Already evaluated in Mastery and Framework pillars.
S109: Suppressed for growth AND expansion objectives (was expansion only).
S110: New PARTIAL path when branded search term spend < 5% target and SB > 5% of spend.
      Branded share note added to what_we_saw when below target.
"""

from __future__ import annotations

import re
from typing import Optional

from reader_databricks_strategy import StrategyContext


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def evaluate_strategy(ctx: StrategyContext) -> tuple[dict[str, str], dict[str, str], dict[str, str]]:
    """
    Returns (flags, what_we_saw, what_you_should_do).
      flags              : {sid: 'FLAG'|'PARTIAL'} — controls that fired
      what_we_saw        : {sid: plain-english text} — for every fired control
      what_you_should_do : {sid: actionable text} — for scoped controls only
    """
    flags = _compute_flags(ctx)
    texts = _build_what_we_saw(ctx, flags)
    how   = _build_what_you_should_do(ctx, flags)
    return flags, texts, how


def calculate_grade(flags: dict[str, str]) -> tuple[str, str]:
    """Returns (grade_label, interpretation_text)."""
    n_flag    = sum(1 for v in flags.values() if v == 'FLAG')
    n_partial = sum(1 for v in flags.values() if v == 'PARTIAL')

    if n_flag == 0 and n_partial == 0:
        return (
            'Compliant',
            'The account reflects a well-defined strategic direction with no major gaps identified. '
            'Few or no changes are required — the current campaign structure, targeting approach, '
            'and client alignment are consistent with the account\'s objectives and roadmap.',
        )
    if n_flag == 0 and n_partial <= 3:
        return (
            'Needs Review',
            f'The account has {n_partial} area(s) that require attention. '
            'No critical gaps were found, but several strategic items should be reviewed '
            'before the next client interaction.',
        )
    if n_flag <= 2:
        return (
            'Needs Improvement',
            f'The account has {n_flag} critical gap(s) and {n_partial} item(s) needing attention. '
            'Action is required. Review the flagged controls and align with the client or internal team '
            'on a clear plan before the next review cycle.',
        )
    return (
        'Non-Compliant',
        f'The account has {n_flag} critical strategic gaps. '
        'Significant structural or strategic work is required. '
        'Prioritise the flagged controls and escalate where client alignment is needed.',
    )


# ─────────────────────────────────────────────────────────────────────────────
# Internal helpers
# ─────────────────────────────────────────────────────────────────────────────

def _pct(v: float) -> str:
    return f'{v:.0%}'

def _dollar(v: float) -> str:
    return f'${v:,.0f}'


# ─────────────────────────────────────────────────────────────────────────────
# Flag engine
# ─────────────────────────────────────────────────────────────────────────────

def _compute_flags(ctx: StrategyContext) -> dict[str, str]:
    flags: dict[str, str] = {}

    def flag(sid: str, level: str) -> None:
        if flags.get(sid) == 'FLAG':   # never downgrade
            return
        flags[sid] = level

    # ── normalised comparisons ────────────────────────────────────────────────
    acos_pp        = ctx.acos_actual  * 100
    tacos_pp       = ctx.tacos_actual * 100
    constraint     = ctx.acos_constraint
    tacos_con      = ctx.tacos_constraint
    has_constraint = constraint > 0
    has_tacos_con  = tacos_con  > 0

    if not has_constraint and ctx.acos_actual > 0:
        constraint = acos_pp + 5.0
    if not has_tacos_con and ctx.tacos_actual > 0:
        tacos_con = tacos_pp + 5.0

    above_acos     = has_constraint and acos_pp  > constraint
    above_acos_10  = has_constraint and acos_pp  > constraint * 1.10
    above_tacos    = has_tacos_con  and tacos_pp > tacos_con
    above_tacos_10 = has_tacos_con  and tacos_pp > tacos_con  * 1.10
    non_qt_total   = ctx.pct_imported + ctx.pct_non_quartile
    declining_yoy  = ctx.yoy_ad_sales < -0.05
    growing_yoy    = ctx.yoy_ad_sales >  0.10
    tacos_rising   = ctx.tacos_trend == 'increasing' and ctx.tacos_trend_pp > 1.5
    spend_rising   = ctx.mom_spend_change > 0.10

    has_atc = any(
        re.search(r'\bATC\b|SD_FLEX_ATC|SD_FLEX_Add.?to.?cart', n, re.IGNORECASE)
        for n in ctx.campaign_names
    )
    has_sd_remarketing = any(
        re.search(r'SD_FLEX_Remarketing|SD_FLEX_rmkt|SD_AUDI', n, re.IGNORECASE)
        for n in ctx.campaign_names
    )

    # ── account state gates ───────────────────────────────────────────────────
    at_scale        = ctx.total_spend >= 1500
    base_built      = ctx.pct_ba > 0 and ctx.pct_bak > 0
    advanced_ready  = base_built and not above_acos_10 and at_scale
    efficiency_ok   = not above_acos and not declining_yoy

    # ── primary objective booleans ────────────────────────────────────────────
    obj_growth      = ctx.primary_objective == 'Growth'
    obj_expansion   = ctx.primary_objective == 'Expansion'
    obj_brand       = ctx.primary_objective == 'Brand Building'
    obj_profit      = 'Profit Maximization' in ctx.primary_objective
    obj_recovery    = ctx.primary_objective == 'Recovery/Stabilization'
    obj_maintenance = ctx.primary_objective == 'Maintenance (holding steady)'
    obj_ntb         = 'Aquisition' in ctx.primary_objective or 'Acquisition' in ctx.primary_objective
    repeat_high     = ctx.repeat_purchase == 'High'
    repeat_low      = ctx.repeat_purchase == 'Low'
    is_commodity    = ctx.commodity_or_brand == 'Commodity'
    high_concentration = 'High' in ctx.sales_concentration
    growth_or_expansion = obj_growth or obj_expansion

    # ── ACOS AND TARGET ───────────────────────────────────────────────────────

    # S002 — ACoS target above constraint
    if has_constraint and ctx.acos_gap_to_constraint > 10:
        flag('S002', 'FLAG')
    elif has_constraint and ctx.acos_gap_to_constraint > 5:
        flag('S002', 'PARTIAL')

    # S003 — TACoS alignment
    if has_tacos_con and tacos_pp > tacos_con + 5:
        flag('S003', 'FLAG')
    elif has_tacos_con and tacos_pp > tacos_con + 2:
        flag('S003', 'PARTIAL')

    # S004 — ACoS reduction cadence
    if above_acos and ctx.acos_changes_30d == 0:
        flag('S004', 'FLAG')
    elif above_acos and ctx.acos_changes_30d > 0 and ctx.acos_gap_to_constraint > 0:
        flag('S004', 'PARTIAL')

    # S005 — Portfolio migration progress
    if at_scale and 0.50 <= ctx.campaigns_in_portfolio_pct < 0.80 and ctx.total_campaign_count > 5:
        flag('S005', 'PARTIAL')

    # S006 — ACoS Target Loosening Risk
    # MANUAL — recommends ACoS target changes specific to the account. CSM reviews.

    # S007 — Branded vs Non-Branded ACoS imbalance
    if (ctx.branded_acos > 0
            and ctx.non_branded_acos > 0
            and ctx.non_branded_acos / ctx.branded_acos >= 3.0
            and ctx.acos_current_target > ctx.branded_acos * 100
            and ctx.branded_spend_pct >= 0.25):
        flag('S007', 'FLAG')

    # S008 — OOB ACoS Reduction to Ease Pressure
    # MANUAL — recommends account-specific ACoS reduction. CSM reviews.

    # ── OVERALL STRUCTURE ─────────────────────────────────────────────────────

    # S009 — Framework compliance review
    _gaps = sum([
        ctx.spend_sb      == 0,
        ctx.spend_spt     == 0,
        ctx.watm_campaign_count == 0 and not ctx.has_catchall,
        not any(re.search(r'SD_SPT', n, re.IGNORECASE) for n in ctx.campaign_names),
    ])
    if _gaps >= 3 and (above_acos or declining_yoy):
        flag('S009', 'FLAG')
    elif _gaps == 2 and (above_acos or declining_yoy):
        flag('S009', 'PARTIAL')

    # S010 — Slow movers in BA. Only fires when ATM-qualifying ASINs exist.
    # If no ASIN qualifies for ATM (no >1.5 orders/day), S011 owns this case instead.
    #
    # Minimum count gate: on small catalogs (< 10 ASINs) or when the number of
    # slow movers in BA is ≤ 10% of the total catalog, noise outweighs signal.
    # A couple of tail ASINs in BA on a 6-ASIN account is expected and not actionable.
    no_atm_qualifying = ctx.tier1_asin_count == 0
    _s010_catalog = max(ctx.catalog_asin_count, 1)
    _s010_min_slow = max(2, int(_s010_catalog * 0.10))  # at least 2, or 10% of catalog

    if not no_atm_qualifying:
        # Account has top sellers — slow movers in BA is a misallocation problem
        if ctx.slow_movers_with_ba >= _s010_min_slow and (ctx.tier1_with_atm == 0 or ctx.watm_campaign_count == 0):
            flag('S010', 'FLAG')
        elif ctx.slow_movers_with_ba >= _s010_min_slow and ctx.watm_campaign_count > 0:
            flag('S010', 'PARTIAL')
    else:
        # No ATM-qualifying ASIN exists — S011 owns this case. S010 is suppressed.
        # S011 — No top seller. Account runs on bulk methodology: BA + WATM for auto targets.
        if ctx.slow_movers_with_ba >= _s010_min_slow and ctx.watm_campaign_count == 0:
            flag('S011', 'FLAG')
        elif ctx.slow_movers_with_ba >= _s010_min_slow:
            flag('S011', 'PARTIAL')

    # S012 — ATM+BA overlap with CPC pressure.
    # Suppressed when no_atm_qualifying — if no ASIN qualifies for ATM, the overlap
    # is a structural issue already owned by S011. S012 only fires when a real ATM-qualifying
    # ASIN exists but is also covered by BA AND CPC pressure is evident.
    if not no_atm_qualifying and ctx.atm_ba_overlap_count > 0 and ctx.cpc_current > 1.20:
        flag('S012', 'FLAG')

    # S013 — ATM+BA overlap (general). Suppressed on bulk accounts — overlap is expected
    # when no ASIN qualifies for ATM and bulk methodology is intentional.
    if ctx.atm_ba_overlap_count > 0 and not no_atm_qualifying:
        flag('S013', 'PARTIAL')

    # S014 — Bulk structure completeness.
    # Three layers evaluated:
    #   Discovery  : BA + WATM/CatchAll + CAT_SP
    #   Precision  : BAK (harvest layer fed by BA)
    #   Defensive  : SPT (own product page protection)
    #
    # FLAG: BA is a dominant spend layer (≥15% spend) AND BAK completely missing
    # FLAG: BAK exists but severely underfed — BAK < 10% of BA spend
    # PARTIAL: CAT_SP missing (product-targeting discovery gap)
    # PARTIAL: WATM and CatchAll both absent (slow mover coverage gap)
    # PARTIAL: SPT missing (defensive layer absent)
    #
    # Non-bulk accounts (BA < 15% of spend) are not evaluated — they don't run
    # the bulk methodology so structural gaps in BA/BAK are not meaningful.
    has_watm_or_catchall = ctx.watm_campaign_count > 0 or ctx.has_catchall
    _is_bulk_account = ctx.pct_ba >= 0.15
    if _is_bulk_account and ctx.total_spend > 500:
        if ctx.pct_bak == 0 and not ctx.bak_name_overlaps_ba:
            flag('S014', 'FLAG')
        elif ctx.pct_bak > 0 and ctx.pct_bak < ctx.pct_ba * 0.10:
            flag('S014', 'FLAG')
        else:
            # Precision layer OK — check discovery and defensive gaps
            if not ctx.has_cat_sp:
                flag('S014', 'PARTIAL')
            if not has_watm_or_catchall:
                flag('S014', 'PARTIAL')
            if ctx.spend_spt == 0:
                flag('S014', 'PARTIAL')

    # S017 — Remove multi-ASIN bulk structures (single-ASIN account)
    if ctx.parent_asin_count == 1 and (ctx.has_catchall or ctx.spend_spt > 0):
        flag('S017', 'FLAG')

    # ── AUTO-TO-MANUAL / DISCOVERY ────────────────────────────────────────────

    # S018 — Auto-to-Manual Conversion Ratio
    if ctx.auto_spend_pct > 0.50 and ctx.manual_exact_pct < 0.15 and at_scale:
        if not growth_or_expansion and not obj_brand:
            flag('S018', 'FLAG')
        elif ctx.auto_spend_pct > 0.65:
            flag('S018', 'PARTIAL')

    # ── OVERALL PARAMETERS AND KPIs ───────────────────────────────────────────

    # S020 — TACoS increasing trend OR CPC rising.
    # TACoS path: suppress for growth/expansion — rising TACoS on a growing account
    # is expected and not actionable. Only flag when efficiency is at risk.
    # CPC path: added as a secondary signal when TACoS path does not fire.
    #   PARTIAL: CPC YoY +20% or more
    #   FLAG:    CPC YoY +40% or more
    if not growth_or_expansion:
        if tacos_rising and ctx.tacos_trend_pp > 0.70 and above_tacos and ctx.total_spend >= 1000:
            flag('S020', 'FLAG')
        elif tacos_rising and ctx.tacos_trend_pp > 0.70 and ctx.total_spend >= 1000:
            flag('S020', 'PARTIAL')

    # CPC path — independent of TACoS, fires even when S020 TACoS path is suppressed.
    # Suppressed for growth/expansion — rising CPCs are expected when scaling.
    if 'S020' not in flags and not growth_or_expansion and ctx.cpc_yoy_change_pct > 0 and ctx.total_spend >= 1000:
        if ctx.cpc_yoy_change_pct >= 0.40:
            flag('S020', 'FLAG')
        elif ctx.cpc_yoy_change_pct >= 0.20:
            flag('S020', 'PARTIAL')

    # S019 — CPC increase YoY
    # Suppressed when S020 (TACoS trend) already fires — S020 is the stronger signal.
    # CPC pressure is implicit in a rising TACoS — no need to flag both independently.
    if 'S020' not in flags:
        if ctx.cpc_yoy_change_pct > 0.20 and above_acos:
            flag('S019', 'FLAG')
        elif ctx.cpc_yoy_change_pct > 0.20 and not growing_yoy:
            flag('S019', 'PARTIAL')
        elif ctx.cpc_yoy_change_pct > 0.10 and above_acos:
            flag('S019', 'PARTIAL')

    # S021 — OOB — Budget Expansion Priority
    # MANUAL — recommends budget increase or efficiency action specific to the account. CSM reviews.

    # S022 — TACoS at risk level (absolute)
    if ctx.tacos_actual > 0.50:
        flag('S022', 'FLAG')
    elif ctx.tacos_actual > 0.30:
        flag('S022', 'PARTIAL')

    # S023 — Catalogue activation scope
    if ctx.catalog_asin_count >= 10:
        coverage = ctx.spending_asin_count / ctx.catalog_asin_count
        if coverage < 0.10:
            flag('S023', 'FLAG')
        elif coverage < 0.20:
            flag('S023', 'PARTIAL')

    # S024 — TACoS/ACoS divergence
    if (ctx.tacos_trend == 'increasing' and ctx.tacos_trend_pp > 1.5
            and ctx.acos_direction == 'decreasing' and ctx.total_spend >= 1000):
        flag('S024', 'FLAG')

    # ── BASIC STRATEGY ────────────────────────────────────────────────────────

    # S030 — Non-Quartile spend review
    if non_qt_total > 0.40 or (non_qt_total > 0.20 and above_acos_10):
        flag('S030', 'FLAG')
    elif non_qt_total > 0.20:
        flag('S030', 'PARTIAL')

    # S031 — SPT defensive structure review
    if ctx.spend_spt > 0 and has_constraint and ctx.spt_avg_acos > 0 and ctx.spt_avg_acos > constraint / 100:
        flag('S031', 'PARTIAL')

    # S032 — SPT covering slow movers
    # Only meaningful when SPT has enough coverage to be actionable.
    # Gate: at least 15 slow movers with SPT spend — small accounts will always
    # have some tail ASINs in SPT and that's not a problem worth flagging.
    if ctx.spend_spt > 0 and len(ctx.tier100_with_spt_asins) >= 15:
        flag('S032', 'PARTIAL')

    # S033 — ATM expansion on best sellers
    # Suppressed when no ASIN qualifies for ATM — bulk accounts don't need ATM expansion
    bulk_heavy = (ctx.pct_ba + ctx.pct_bak + ctx.pct_spt) > 0.60 and ctx.pct_bak > 0
    if not no_atm_qualifying:
        if ctx.pct_atm < 0.03 and not above_tacos_10 and not bulk_heavy and not (ctx.has_oob and above_acos_10):
            flag('S033', 'FLAG')
        elif ctx.pct_atm < 0.03 and bulk_heavy and declining_yoy:
            flag('S033', 'PARTIAL')
        elif ctx.pct_atm < 0.08 and not above_tacos and growing_yoy and not bulk_heavy:
            flag('S033', 'PARTIAL')

    # S034 — Best-Seller Campaigns Paused
    # MANUAL — requires human review of which top-seller campaigns are paused and why.

    # ── CAMPAIGNS STRATEGY ────────────────────────────────────────────────────

    # S036 — Discovery-Performance Mix (ATM + BR both outperforming)
    # CHANGELOG: auto-flag positive composite.  Silences S056 + S057 when fired.
    atm_outperforming = (
        ctx.atm_avg_acos > 0 and ctx.acos_actual > 0
        and ctx.atm_avg_acos < ctx.acos_actual * 0.80
    )
    br_outperforming = (
        ctx.br_avg_acos > 0 and ctx.acos_actual > 0
        and ctx.br_avg_acos < ctx.acos_actual * 0.80
    )
    if atm_outperforming and br_outperforming:
        flag('S036', 'FLAG')   # composite positive suggestion
        # Silence the individual outperforming signals — covered by S036
        flags.pop('S056', None)
        flags.pop('S057', None)

    # S037 — BA covering slow movers
    # Suppressed when S014 is already FLAG — S014 is the stronger structural signal.
    # Suppressed when S010 is already FLAG — same slow movers, S010 is more specific.
    # Minimum gate: use the same _s010_min_slow threshold so single-tail-ASIN
    # situations on small accounts don't generate noise here either.
    if flags.get('S014') != 'FLAG' and flags.get('S010') != 'FLAG':
        if ctx.slow_movers_with_ba >= _s010_min_slow:
            flag('S037', 'FLAG')
        elif ctx.spend_ba > 0 and above_acos:
            flag('S037', 'PARTIAL')

    # S038 — BAK harvest layer missing (BA > 30% spend, no BAK)
    # Suppressed when S014 is already FLAG — same root cause
    if flags.get('S014') != 'FLAG':
        if ctx.pct_ba > 0.30 and ctx.pct_bak == 0 and not ctx.bak_name_overlaps_ba:
            flag('S038', 'FLAG')
        elif ctx.pct_ba > 0.30 and ctx.pct_bak == 0 and ctx.bak_name_overlaps_ba:
            flag('S038', 'PARTIAL')

    # S039 — BA not segmented by category
    # CHANGELOG: also flag when only 1 BA campaign AND multiple categories each >10% of total sales
    cat_with_10pct = getattr(ctx, 'categories_above_10pct', 0)
    if 0 < ctx.ba_campaign_count < 2 and ctx.total_spend > 1500 and ctx.catalog_asin_count >= 5:
        flag('S039', 'FLAG')
    elif ctx.ba_campaign_count == 1 and cat_with_10pct > 1 and ctx.total_spend > 1500:
        # 1 BA but multiple material categories → needs segmentation
        flag('S039', 'FLAG')

    # S041 — Low-order campaign consolidation
    if at_scale and ctx.low_order_campaign_count > 80:
        flag('S041', 'FLAG')
    elif at_scale and ctx.low_order_campaign_count > 40:
        flag('S041', 'PARTIAL')

    # S044 — SB category target expansion
    # Requires: base built, no SB spend, product targeting exists.
    # Gate: account spend ≥ $5,000 AND either OP or CAT_SP is performing below constraint.
    has_product_targeting_base = (ctx.has_op and ctx.pct_op > 0) or (ctx.has_cat_sp and ctx.total_spend > 500)
    op_below_constraint = (
        ctx.op_avg_acos > 0 and ctx.acos_constraint > 0
        and ctx.op_avg_acos > (ctx.acos_constraint / 100)
    )
    catsp_below_constraint = (
        getattr(ctx, 'catsp_avg_acos', 0) > 0 and ctx.acos_constraint > 0
        and getattr(ctx, 'catsp_avg_acos', 0) > (ctx.acos_constraint / 100)
    )
    product_targeting_below_constraint = op_below_constraint or catsp_below_constraint
    if (base_built and ctx.spend_sb == 0 and has_product_targeting_base
            and ctx.total_spend >= 5000 and product_targeting_below_constraint):
        if declining_yoy:
            flag('S044', 'FLAG')
        elif above_acos_10:
            flag('S044', 'PARTIAL')
        else:
            flag('S044', 'FLAG')

    # S045 — BAK harvest stalled
    # CHANGELOG: only flag when objective is Growth or Expansion AND BA campaigns
    # have at least 80 orders in the period — accounts with BAK already launched
    # or insufficient BA order volume should not be flagged.
    if ctx.bak_underfed and at_scale and growth_or_expansion and ctx.ba_orders_30d >= 80:
        flag('S045', 'PARTIAL')

    # ── ADVANCED STRATEGIES ───────────────────────────────────────────────────

    # S047 — Import kickoff needed
    if ctx.total_spend > 0 and ctx.pct_imported > 0.30:
        flag('S047', 'FLAG')
    elif ctx.total_spend > 0 and ctx.pct_imported > 0.15:
        flag('S047', 'PARTIAL')

    # ── PERFORMANCE CAMPAIGN ─────────────────────────────────────────────────

    # S053 — SP Campaign ACoS Significantly Above Constraint (campaign-level)
    # Suppressed for growth/expansion objective — overspending campaigns are expected
    # when the account is actively scaling. Flag only on efficiency-focused accounts.
    if has_constraint and ctx.sp_worst_campaign_acos > 0 and not growth_or_expansion:
        gap_ratio = (ctx.sp_worst_campaign_acos * 100 - constraint) / constraint
        if gap_ratio > 0.35:
            flag('S053', 'FLAG')
        elif gap_ratio > 0.20:
            flag('S053', 'PARTIAL')

    # S054 — SB Campaign ACoS Significantly Above Constraint (campaign-level)
    if has_constraint and ctx.sb_worst_campaign_acos > 0 and not growth_or_expansion:
        gap_ratio = (ctx.sb_worst_campaign_acos * 100 - constraint) / constraint
        if gap_ratio > 0.35:
            flag('S054', 'FLAG')
        elif gap_ratio > 0.20:
            flag('S054', 'PARTIAL')

    # S055 — SD Campaign ACoS Significantly Above Constraint (campaign-level)
    if has_constraint and ctx.sd_worst_campaign_acos > 0 and not growth_or_expansion:
        gap_ratio = (ctx.sd_worst_campaign_acos * 100 - constraint) / constraint
        if gap_ratio > 0.35:
            flag('S055', 'FLAG')
        elif gap_ratio > 0.20:
            flag('S055', 'PARTIAL')

    # S056 — ATM campaigns outperforming (positive suggestion)
    # Only written if S036 did NOT fire (composite silences individual signals)
    if atm_outperforming and 'S036' not in flags:
        flag('S056', 'FLAG')

    # S072 — Broad Match Graduation Signal (BR outperforms OW)
    # Evaluated early so S058 can suppress when S072 is already active.
    if (ctx.br_campaign_count > 30 and ctx.ow_campaign_count > 30
            and ctx.ph_campaign_count < 10
            and ctx.br_avg_acos > 0 and ctx.ow_avg_acos > 0
            and ctx.br_avg_acos < ctx.ow_avg_acos):
        flag('S072', 'FLAG')

    # S057 — Broad Match Campaigns Outperforming (positive suggestion)
    # Only written if S036 did NOT fire
    if br_outperforming and 'S036' not in flags:
        flag('S057', 'FLAG')

    # S058 — Phrase Match Campaigns Outperforming
    # Suppressed when S072 is active — S072 already signals that BR outperforms OW,
    # surfacing a PH signal on top creates conflicting match-type recommendations.
    if (ctx.ph_avg_acos > 0 and ctx.acos_actual > 0 and ctx.ph_avg_acos < ctx.acos_actual * 0.80
            and 'S072' not in flags):
        flag('S058', 'FLAG')

    # S059 — Exact Match Campaigns Outperforming
    if ctx.ow_avg_acos > 0 and ctx.acos_actual > 0 and ctx.ow_avg_acos < ctx.acos_actual * 0.80:
        flag('S059', 'FLAG')

    # S060 — Product Targeting Campaigns Outperforming
    if ctx.op_avg_acos > 0 and ctx.acos_actual > 0 and ctx.op_avg_acos < ctx.acos_actual * 0.80:
        flag('S060', 'FLAG')

    # S062 — Paused SB Campaign Rebuild
    if ctx.paused_sb_count > 0 and ctx.spend_sb == 0 and not above_acos and at_scale:
        flag('S062', 'PARTIAL')

    # S063 — SD_FLEX Campaigns Outperforming
    # Suppressed when >50% of SD_FLEX spend is VCPM — impression-based billing,
    # ACoS is not comparable and the outperforming signal is misleading.
    if (ctx.sd_flex_avg_acos > 0 and ctx.acos_actual > 0
            and ctx.sd_flex_avg_acos < ctx.acos_actual * 0.80
            and getattr(ctx, 'sd_flex_vcpm_pct', 0.0) < 0.50):
        flag('S063', 'FLAG')

    # S064 — SD_AUDI Investment Opportunity
    if (ctx.sd_audi_avg_acos > 0 and ctx.acos_actual > 0
            and ctx.sd_audi_avg_acos < ctx.acos_actual * 0.80
            and getattr(ctx, 'sd_audi_vcpm_pct', 0.0) < 0.50):
        flag('S064', 'FLAG')

    # S065 — SD_PRD Investment Opportunity
    if (ctx.sd_prd_avg_acos > 0 and ctx.acos_actual > 0
            and ctx.sd_prd_avg_acos < ctx.acos_actual * 0.80
            and getattr(ctx, 'sd_prd_vcpm_pct', 0.0) < 0.50):
        flag('S065', 'FLAG')

    # S067 — SB Investment Opportunity
    if ctx.sb_avg_acos > 0 and ctx.acos_actual > 0 and ctx.sb_avg_acos < ctx.acos_actual * 0.80:
        flag('S067', 'FLAG')

    # S068 — SBV Investment Opportunity
    if ctx.sbv_avg_acos > 0 and ctx.acos_actual > 0 and ctx.sbv_avg_acos < ctx.acos_actual * 0.80:
        flag('S068', 'FLAG')

    # S069 — SBV Campaign Reactivation
    if ctx.paused_sbv_count > 0 and ctx.spend_sbv == 0 and ctx.spend_sb > 0 and not above_acos and base_built:
        flag('S069', 'PARTIAL')

    # ── NEW DEPLOYS ───────────────────────────────────────────────────────────

    # S070 — CAT_SP Launch
    # Spend gate uses OP spend specifically (pct_op × total_spend), not account total.
    # This avoids flagging small accounts where OP is minimal even if account spend is low.
    op_spend = ctx.pct_op * ctx.total_spend
    op_outperforming = (
        ctx.op_avg_acos > 0 and ctx.acos_actual > 0
        and ctx.op_avg_acos < ctx.acos_actual * 0.80
    )
    if not ctx.has_cat_sp and op_spend > 500:
        if op_outperforming and (growing_yoy or not above_acos):
            flag('S070', 'FLAG')
        elif op_outperforming and above_acos:
            flag('S070', 'PARTIAL')
        elif not ctx.has_op and (growing_yoy or not above_acos):
            flag('S070', 'FLAG')
        elif not ctx.has_op:
            flag('S070', 'PARTIAL')

    # S071 — SBV Product Targeting Launch
    # CHANGELOG: also requires CAT_SP or OP outperforming (same gate as S070)
    catsp_outperforming = (
        ctx.catsp_avg_acos > 0 and ctx.acos_actual > 0
        and ctx.catsp_avg_acos < ctx.acos_actual * 0.80
    )
    sb_well_established2 = ctx.pct_sb > 0.05 and not above_acos
    has_product_targeting_base2 = (ctx.has_op and ctx.pct_op > 0) or (ctx.has_cat_sp and ctx.total_spend > 500)
    product_signal = op_outperforming or catsp_outperforming
    if (base_built and not ctx.has_sbv and ctx.spend_sbv == 0
            and sb_well_established2 and has_product_targeting_base2
            and product_signal):
        flag('S071', 'FLAG')

    # S073 — Historical BAK Relaunch — MANUAL

    # S074 — Exact Match Graduation Signal (OW outperforms BR)
    if (ctx.br_campaign_count > 30 and ctx.ow_campaign_count > 30
            and ctx.ph_campaign_count < 10
            and ctx.br_avg_acos > 0 and ctx.ow_avg_acos > 0
            and ctx.ow_avg_acos < ctx.br_avg_acos):
        flag('S074', 'FLAG')

    # S075 — OP Target Expansion Opportunity
    # Only meaningful when the account has active keyword campaigns (OW/BR/PH) with spend.
    # If there are no keyword campaigns at all, OP underdevelopment is a framework gap,
    # not a strategy opportunity — suppress to avoid noise on pure-bulk accounts.
    kw_total = ctx.br_campaign_count + ctx.ow_campaign_count + ctx.ph_campaign_count
    _op_with_spend = getattr(ctx, 'op_campaigns_with_spend', ctx.op_campaign_count)
    if ctx.total_spend >= 1000 and kw_total > 50 and _op_with_spend < 10:
        flag('S075', 'FLAG')
    elif ctx.total_spend >= 1000 and kw_total > 30 and _op_with_spend < 10:
        flag('S075', 'PARTIAL')

    # S076 — CatchAll Graduation Overdue
    if ctx.catchall_orders > 100 and ctx.pct_bak < 0.10 and at_scale:
        flag('S076', 'FLAG')
    elif ctx.catchall_orders > 50 and ctx.pct_bak < 0.10 and at_scale:
        flag('S076', 'PARTIAL')

    # ── GOVERNANCE ────────────────────────────────────────────────────────────

    # S077 — CAT_SP Above ACoS Target
    # CHANGELOG: PARTIAL when above constraint AND below 85%; FLAG when above 85%
    if ctx.catsp_avg_acos > 0 and has_constraint:
        catsp_pp = ctx.catsp_avg_acos * 100
        if catsp_pp > constraint * 0.85:
            flag('S077', 'FLAG')
        elif catsp_pp > constraint:
            flag('S077', 'PARTIAL')

    # S079 — Multiple WATM without structural need
    if ctx.watm_campaign_count > 2 and ctx.spend_watm >= ctx.total_spend * 0.02:
        flag('S079', 'PARTIAL')

    # S082 — BAK Branded and Non-Branded Mixed
    # CHANGELOG: check branded search terms inside BAK bucket
    if ctx.bak_branded_nb_mixed and not is_commodity and at_scale:
        flag('S082', 'PARTIAL')

    # S083 — WATM/CatchAll Catalogue Coverage
    if (ctx.watm_campaign_count > 0 or ctx.has_catchall) and ctx.catalog_asin_count >= 8 and ctx.total_spend >= 1000:
        coverage = ctx.spending_asin_count / ctx.catalog_asin_count if ctx.catalog_asin_count > 0 else 1.0
        if coverage < 0.60:
            flag('S083', 'PARTIAL')

    # S084 — WATM and CatchAll active simultaneously
    if ctx.has_both_watm_and_catchall:
        flag('S084', 'FLAG')

    # S085 — WATM spend underweighted
    # Only meaningful when account has real spend and WATM has some budget allocated
    if ctx.watm_campaign_count > 0 and ctx.pct_watm < 0.03 and ctx.total_spend >= 1000 and ctx.spend_watm > 0:
        flag('S085', 'FLAG')

    # S086 — BAK high-spend with efficiency pressure
    # PARTIAL threshold tightened: BAK ACoS > 80% of constraint (was 50%)
    if has_constraint and ctx.bak_campaigns:
        for bak in ctx.bak_campaigns:
            if bak['pct_of_total'] > 0.15 and bak['acos'] > constraint / 100:
                flag('S086', 'FLAG')
                break
            elif bak['pct_of_total'] > 0.15 and bak['acos'] > (constraint / 100) * 0.80:
                flag('S086', 'PARTIAL')

    # S087 — CAT_SP high-spend with efficiency pressure
    if has_constraint and ctx.pct_cat_sp > 0.15 and acos_pp > constraint * 0.50:
        flag('S087', 'PARTIAL')

    # S088 — SB high-spend with efficiency pressure
    if has_constraint and ctx.pct_sb > 0.15 and acos_pp > constraint * 0.50:
        flag('S088', 'PARTIAL')

    # S089 — SBV high-spend with efficiency pressure
    if has_constraint and ctx.pct_sbv > 0.15 and acos_pp > constraint * 0.50:
        flag('S089', 'PARTIAL')

    # ── ADVANCED CAMPAIGNS ────────────────────────────────────────────────────

    # S092 — SD Remarketing — Product View
    # CHANGELOG: objective filter (Growth/Expansion) + SD spend threshold $1,000
    if (base_built and not ctx.has_sd and ctx.spend_sd == 0
            and ctx.total_spend > 500 and growth_or_expansion):
        if (growing_yoy or ctx.spend_sb > 0) and ctx.max_asin_orders_30d >= 50:
            flag('S092', 'FLAG')
        elif not above_acos and ctx.max_asin_orders_30d >= 50:
            flag('S092', 'PARTIAL')

    # S093 — SD ATC Retargeting — ProSuite
    # CHANGELOG: objective filter + SD spend threshold ($1,000); case-insensitive names
    if (not has_atc and ctx.has_prosuite_audiences
            and ctx.pct_sd >= 0.03 and ctx.spend_sd >= 1000
            and growth_or_expansion):
        flag('S093', 'PARTIAL')

    # S096 — SD Suggested — PDP Maturity Too Low
    # CHANGELOG: OK when top ASIN already has SD spend (audience pool present)
    if (not ctx.has_sd and ctx.total_spend > 500 and base_built
            and (growing_yoy or ctx.spend_sb > 0)
            and ctx.max_asin_orders_30d < 50 and ctx.max_asin_orders_30d > 0):
        flag('S096', 'PARTIAL')
    # If SD already has spend, this control is already addressed — no flag

    # ── GOVERNANCE ON FRAMEWORK ───────────────────────────────────────────────

    # S097 — Portfolio Governance — Unused Portfolios
    # MANUAL — portfolio governance requires human review of naming and structure.
    # Do not auto-flag. CSM reviews this manually during QR call.

    # S098 — Campaign-Level ACoS Overrides Active
    if ctx.has_campaign_acos_overrides and above_acos:
        flag('S098', 'PARTIAL')

    # S099 — Product-Level ACoS Overrides Active
    if ctx.has_product_acos_overrides and above_acos:
        flag('S099', 'PARTIAL')

    # S100 — VCPM Buy Box Requirement
    if ctx.vcpm_spend_pct > 0.10:
        flag('S100', 'FLAG')
    elif ctx.vcpm_spend_pct > 0.05:
        flag('S100', 'PARTIAL')

    # S101 — Tagging and Segmentation Gap
    # MANUAL — tagging is already evaluated in Mastery and Framework.
    # Duplicate auto-logic here generates redundant flags. CSM reviews manually.

    # S103 — SBV Naming Convention
    if ctx.has_sbv and not ctx.sbv_naming_compliant:
        flag('S103', 'PARTIAL')

    # ── SPENDING FOCUS ────────────────────────────────────────────────────────

    # S108 — ProSuite AMC Audience Testing
    # Only flags when ProSuite is active (tab 51 has real data) but audiences
    # are not yet applied. Accounts without ProSuite enabled cannot act on this.
    if ctx.prosuite_active and not ctx.has_prosuite_audiences and advanced_ready and growing_yoy:
        flag('S108', 'FLAG')
    elif ctx.prosuite_active and not ctx.has_prosuite_audiences and advanced_ready:
        flag('S108', 'PARTIAL')

    # S109 — Inefficient ASIN Spend Reduction
    # CHANGELOG: auto rule — ASIN spending with zero sales OR ACoS > 2× constraint
    # Suppressed for growth AND expansion objectives — on scaling accounts, some ASINs
    # will have poor efficiency while building visibility. Flagging them creates noise.
    if ctx.inefficient_asin_count > 0 and at_scale and not growth_or_expansion:
        flag('S109', 'FLAG')

    # S110 — SB Active — SBV Missing
    # CHANGELOG: threshold extended to 10%.
    # Also fires when branded search term spend is below 5% — SBV is the primary
    # tool to defend and grow branded search share; low branded spend reinforces urgency.
    _s110_low_branded = ctx.branded_spend_pct < 0.05 and ctx.branded_spend_pct > 0
    if base_built and ctx.spend_sbv == 0 and ctx.pct_sb > 0.10 and not above_acos:
        flag('S110', 'FLAG')
    elif base_built and ctx.spend_sbv == 0 and ctx.pct_sb > 0.05 and _s110_low_branded:
        flag('S110', 'PARTIAL')

    # ── CLIENT DIRECTIONS ─────────────────────────────────────────────────────

    # S111 — External Traffic Tracking — MANUAL

    # S113 — Recurring Sales Strategy
    # Suppressed when S119 (Subscribe & Save) already fires — same root cause, S119 is more specific
    # CHANGELOG: also fires when repeat_purchase is High (regardless of YoY)
    if 'S119' not in flags and not ctx.has_sns_active and not repeat_low and not obj_ntb:
        if declining_yoy or repeat_high:
            flag('S113', 'FLAG')
        elif not growing_yoy:
            flag('S113', 'PARTIAL')

    # S114 — Sales Declining While Spend Growing
    if declining_yoy and spend_rising:
        flag('S114', 'FLAG')

    # ── LISTING OPTIMIZATIONS ─────────────────────────────────────────────────

    # S119 — Subscribe & Save — Not Active
    # CHANGELOG: also fires when repeat_purchase is High
    if not ctx.has_sns_active and not repeat_low and not obj_ntb:
        if declining_yoy or repeat_high:
            flag('S119', 'FLAG')
        else:
            flag('S119', 'PARTIAL')

    # ── PROMO AND GGS ─────────────────────────────────────────────────────────

    # S122 — Promo Portfolio Budget Pacing
    has_named_promo_portfolio = any('PROMO' in str(n).upper() for n in ctx.portfolio_names)
    if ctx.has_active_promo and has_named_promo_portfolio and ctx.promo_cost_rate > 0.05:
        flag('S122', 'PARTIAL')

    # S124 — SD GGS Compliance
    # Only fires when account has a confirmed GGS commitment. Non-GGS accounts are fully suppressed.
    if ctx.ggs_status == 'Yes':
        if ctx.spend_sd == 0:
            flag('S124', 'FLAG')
        elif ctx.pct_sd < 0.05:
            flag('S124', 'PARTIAL')

    # S125 — SD Remarketing Missing
    if ctx.spend_sd > 0 and not has_sd_remarketing:
        flag('S125', 'PARTIAL')

    # S126 — SD ATC Retargeting — GGS Section
    # Only fire if account is GGS AND has ProSuite audiences active
    if ctx.spend_sd > 0 and not has_atc and ctx.ggs_status == 'Yes' and ctx.has_prosuite_audiences:
        flag('S126', 'PARTIAL')

    return flags


# ─────────────────────────────────────────────────────────────────────────────
# What We Saw builder
# ─────────────────────────────────────────────────────────────────────────────

def _build_what_we_saw(ctx: StrategyContext, flags: dict[str, str]) -> dict[str, str]:
    texts: dict[str, str] = {}
    constraint = ctx.acos_constraint if ctx.acos_constraint > 0 else (ctx.acos_actual * 100 + 5.0)

    def _t(sid: str) -> bool:
        return sid in flags

    if _t('S002'):
        texts['S002'] = (
            f'The current ACoS target is {ctx.acos_current_target:.0f}%. '
            f'The account constraint is {ctx.acos_constraint:.0f}%. '
            f'The gap is +{ctx.acos_gap_to_constraint:.0f} percentage points. '
            f'The target needs to come down to align with the client objective.'
        )

    if _t('S003'):
        texts['S003'] = (
            f'TACoS actual: {ctx.tacos_actual:.0%} vs constraint {ctx.tacos_constraint:.0f}%. '
            f'TACoS is above the agreed limit. ACoS reductions will bring TACoS down over time.'
        )

    if _t('S004'):
        texts['S004'] = (
            f'{ctx.acos_changes_30d} ACoS change(s) in the last 30 days. '
            f'Gap to constraint: {ctx.acos_gap_to_constraint:+.0f}pp. '
            + ('No changes made despite being above constraint. Act now.' if ctx.acos_changes_30d == 0
               else 'Changes are happening but the gap to constraint remains.')
        )

    if _t('S005'):
        in_port = round(ctx.campaigns_in_portfolio_pct * ctx.total_campaign_count)
        not_in_port = ctx.total_campaign_count - in_port
        texts['S005'] = (
            f'{in_port} of {ctx.total_campaign_count} campaigns ({ctx.campaigns_in_portfolio_pct:.0%}) are in portfolios. '
            f'{not_in_port} remain outside. Complete the portfolio assignment.'
        )

    if _t('S006'):
        texts['S006'] = (
            f'ACoS target increased {ctx.acos_changes_30d} time(s) in the last 30 days. '
            f'Current target: {ctx.acos_current_target:.0f}%. '
            f'Spend growth driven by loosening efficiency — not by structural improvements.'
        )

    if _t('S007'):
        texts['S007'] = (
            f'Branded spend is {ctx.branded_spend_pct:.0%} of total at {ctx.branded_acos:.0%} ACoS. '
            f'Non-branded is at {ctx.non_branded_acos:.0%} ACoS vs portal target {ctx.acos_current_target:.0f}%. '
            f'The target is calibrated to branded performance, leaving non-branded campaigns overspending.'
        )

    if _t('S008'):
        texts['S008'] = (
            f'Account hit daily budget limits. ACoS target: {ctx.acos_current_target:.0f}% vs constraint {ctx.acos_constraint:.0f}%. '
            f'Reducing the ACoS target lowers CPC pressure and eases out-of-budget events.'
        )

    if _t('S009'):
        gap_labels = []
        if ctx.spend_sb == 0: gap_labels.append('no SB campaigns')
        if not ctx.has_cat_sp: gap_labels.append('no CAT_SP campaigns')
        if not ctx.has_sbv and ctx.spend_sbv == 0: gap_labels.append('no SBV campaigns')
        if not ctx.has_sd and ctx.spend_sd == 0: gap_labels.append('no SD campaigns')
        if ctx.spend_spt > 0 and ctx.pct_atm < 0.03: gap_labels.append('SPT active but ATM < 3%')
        if ctx.campaigns_not_in_portfolio > 5: gap_labels.append(f'{ctx.campaigns_not_in_portfolio} campaigns outside portfolios')
        if (ctx.pct_imported + ctx.pct_non_quartile) > 0.40: gap_labels.append(f'{_pct(ctx.pct_imported + ctx.pct_non_quartile)} spend outside framework')
        n_gaps = len(gap_labels)
        gaps_str = ', '.join(gap_labels[:5])
        suffix = f' (+{n_gaps - 5} more)' if n_gaps > 5 else ''
        texts['S009'] = (
            f'{n_gaps} structural framework gaps detected: {gaps_str}{suffix}. '
            f'A structured framework review is needed before the next QR.'
        )

    if _t('S010'):
        asin_list = ', '.join(ctx.slow_mover_asins_with_ba[:5]) if ctx.slow_mover_asins_with_ba else ''
        suffix = f' ASINs: {asin_list}.' if asin_list else ''
        texts['S010'] = (
            f'{ctx.slow_movers_with_ba} ASIN(s) with fewer than 3 orders in the period have BA spend.{suffix} '
            f'Slow movers should only appear in WATM — not in BA campaigns.'
        )

    if _t('S011'):
        asin_list = ', '.join(ctx.slow_mover_asins_with_ba[:5]) if ctx.slow_mover_asins_with_ba else ''
        suffix = f' ASINs: {asin_list}.' if asin_list else ''
        texts['S011'] = (
            f'{ctx.slow_movers_with_ba} ASIN(s) with fewer than 3 orders are in BA campaigns '
            f'and no ASIN qualifies for ATM (less than 1.5 orders/day).{suffix} '
            f'This account runs on bulk methodology — BA and WATM for automatic targets. '
            f'Individual ATM products are not applicable. '
            f'Move slow movers to WATM and keep BA for mid-velocity ASINs.'
        )

    if _t('S012'):
        asin_list = ', '.join(ctx.atm_ba_overlap_asins[:5]) if ctx.atm_ba_overlap_asins else 'see tab 14'
        texts['S012'] = (
            f'{ctx.atm_ba_overlap_count} ASIN(s) have both ATM and BA spend with >80 orders. '
            f'CPC: ${ctx.cpc_current:.2f}. ASINs: {asin_list}. '
            f'ATM already covers these high-velocity ASINs — BA spend is redundant.'
        )

    if _t('S013'):
        asin_list = ', '.join(ctx.atm_ba_overlap_asins[:5]) if ctx.atm_ba_overlap_asins else 'see tab 14'
        texts['S013'] = (
            f'{ctx.atm_ba_overlap_count} ASIN(s) have both ATM and BA spend. '
            f'ASINs: {asin_list}. '
            f'ATM and BA are running on the same ASINs — review whether BA is still needed or can be reduced.'
        )

    if _t('S014'):
        has_watm_or_catchall = ctx.watm_campaign_count > 0 or ctx.has_catchall
        if ctx.pct_ba > 0 and ctx.pct_bak == 0:
            texts['S014'] = (
                f'BA campaigns active ({_pct(ctx.pct_ba)} of spend / {_dollar(ctx.spend_ba)}) '
                f'but no BAK harvest layer exists. '
                f'Discovery spend is running but converting terms are not being captured in manual campaigns. '
                f'Create BAK campaigns to harvest the best-performing BA search terms.'
            )
        elif ctx.pct_ba > 0 and ctx.pct_bak < ctx.pct_ba * 0.10:
            texts['S014'] = (
                f'BA is {_pct(ctx.pct_ba)} of spend ({_dollar(ctx.spend_ba)}) '
                f'but BAK is only {_pct(ctx.pct_bak)} ({_dollar(ctx.spend_bak)}). '
                f'BAK is less than 10% of BA — the harvest layer is severely underfed. '
                f'Review BA search term report and promote converting terms to BAK.'
            )
        else:
            gaps = []
            if not ctx.has_cat_sp:
                gaps.append('CAT_SP missing')
            if not has_watm_or_catchall:
                gaps.append('WATM/CatchAll missing')
            if ctx.spend_spt == 0:
                gaps.append('SPT defensive layer missing')
            texts['S014'] = (
                f'Bulk structure has gaps: {", ".join(gaps)}. '
                f'BA: {_pct(ctx.pct_ba)}, BAK: {_pct(ctx.pct_bak)}, SPT: {_pct(ctx.pct_spt)}. '
                f'Complete the bulk structure before expanding to advanced campaigns.'
            )

    if _t('S017'):
        texts['S017'] = (
            f'The account has {ctx.parent_asin_count} parent ASIN. '
            f'Multi-ASIN bulk structures add complexity without value at this catalog size.'
        )

    if _t('S018'):
        texts['S018'] = (
            f'Auto campaigns (BA + ATM + WATM) account for {ctx.auto_spend_pct:.0%} of total spend. '
            f'BAK (manual exact) is only {ctx.manual_exact_pct:.0%}. '
            f'Discovery is generating learnings that are not being converted into precision manual campaigns.'
        )

    if _t('S019'):
        texts['S019'] = (
            f'CPC moved from ${ctx.cpc_last_year:.2f} last year to ${ctx.cpc_current:.2f} ({ctx.cpc_yoy_change_pct:+.0%}). '
            f'ACoS thresholds should be revisited to bring costs back under control.'
        )

    if _t('S020'):
        texts['S020'] = (
            f'TACoS has been {ctx.tacos_trend} for the last 3 months (+{ctx.tacos_trend_pp:.1f}pp). '
            f'Current TACoS: {ctx.tacos_actual:.0%} vs constraint {ctx.tacos_constraint:.0f}%.'
        )

    if _t('S021'):
        oob_case = getattr(ctx, '_oob_case', 'inefficient')
        if oob_case == 'efficient':
            texts['S021'] = (
                f'Account hit daily budget limits. '
                f'ACoS is {ctx.acos_actual:.0%} — within the {ctx.acos_constraint:.0f}% constraint. '
                f'The account is running efficiently and fully utilising its budget. '
                f'This is the right time to negotiate a budget increase with the client.'
            )
        elif oob_case == 'inefficient':
            texts['S021'] = (
                f'Account hit daily budget limits while ACoS is {ctx.acos_actual:.0%} '
                f'vs {ctx.acos_constraint:.0f}% constraint. '
                f'Efficiency must improve before requesting more budget. '
                f'Reduce the ACoS target to lower CPC pressure and ease out-of-budget events.'
            )
        else:
            texts['S021'] = (
                f'Account hit daily budget limits. '
                f'Total sales declined {abs(ctx.mom_sales_change):.0%} MoM. '
                f'Review spend allocation before increasing budget — sales are declining despite full utilisation.'
            )

    if _t('S022'):
        texts['S022'] = (
            f'TACoS is {ctx.tacos_actual:.0%}. '
            + ('Severely high — profitability is heavily impacted.' if ctx.tacos_actual > 0.50
               else 'Above the 30% threshold — profitability risk is elevated.')
        )

    if _t('S023'):
        coverage = ctx.spending_asin_count / ctx.catalog_asin_count if ctx.catalog_asin_count > 0 else 0
        texts['S023'] = (
            f'{ctx.spending_asin_count} of {ctx.catalog_asin_count} catalog ASINs have ad spend ({coverage:.0%} coverage). '
            f'Portfolio activation is too narrow — most of the catalog is not receiving traffic.'
        )

    if _t('S024'):
        texts['S024'] = (
            f'ACoS is trending {ctx.acos_direction} while TACoS has risen '
            f'{ctx.tacos_trend_pp:+.1f}pp over the last 3 months. '
            f'When ACoS improves but TACoS rises, organic sales are declining or promotional activity is distorting total sales.'
        )

    if _t('S030'):
        non_qt = ctx.pct_imported + ctx.pct_non_quartile
        texts['S030'] = (
            f'{_pct(non_qt)} of spend is in Imported or Non-Quartile campaigns '
            f'({_pct(ctx.pct_imported)} Imported, {_pct(ctx.pct_non_quartile)} Non-Quartile). '
            f'The account is not fully operating within the Quartile framework.'
        )

    if _t('S031'):
        texts['S031'] = (
            f'SPT active ({_dollar(ctx.spend_spt)}, {_pct(ctx.pct_spt)} of spend). '
            f'SPT avg ACoS: {ctx.spt_avg_acos:.0%} vs constraint {ctx.acos_constraint:.0f}%. '
            f'Defensive structure should be split by category or brand segment.'
        )

    if _t('S032'):
        asin_list = ', '.join(ctx.tier100_with_spt_asins[:5]) if ctx.tier100_with_spt_asins else ''
        suffix = f' ASINs: {asin_list}.' if asin_list else ''
        texts['S032'] = (
            f'{len(ctx.tier100_with_spt_asins)} Tier 100 ASIN(s) have SPT spend.{suffix} '
            f'SPT spend: {_dollar(ctx.spend_spt)}. Tier 100 ASINs are slow movers — remove from SPT.'
        )

    if _t('S033'):
        texts['S033'] = (
            f'ATM campaigns represent {_pct(ctx.pct_atm)} of spend ({_dollar(ctx.spend_atm)}). '
            + ('No ATM spend detected. ' if ctx.pct_atm == 0 else '')
            + 'Automatic targeting on best-selling ASINs should be expanded.'
        )

    if _t('S034'):
        texts['S034'] = (
            f'{ctx.top_seller_type_gaps} of {ctx.tier1_asin_count} top-selling ASIN(s) (Tier 10–30) '
            f'are missing ≥2 key campaign types (ATM, BAK, OP). '
            f'Best-seller campaigns have likely been paused or were never fully deployed.'
        )

    if _t('S036'):
        texts['S036'] = (
            f'ATM avg ACoS: {_pct(ctx.atm_avg_acos)} — {(1 - ctx.atm_avg_acos / ctx.acos_actual):.0%} better than account avg. '
            f'BR avg ACoS: {_pct(ctx.br_avg_acos)} — {(1 - ctx.br_avg_acos / ctx.acos_actual):.0%} better than account avg. '
            f'Both discovery layers are outperforming. '
            f'Future launches should continue to prioritise ATM and Broad match targeting.'
        )

    if _t('S037'):
        asin_list = ', '.join(ctx.slow_mover_asins_with_ba[:5]) if ctx.slow_mover_asins_with_ba else ''
        suffix = f' ASINs with <3 orders in BA: {asin_list}.' if asin_list else ''
        texts['S037'] = (
            f'BA campaigns: {_dollar(ctx.spend_ba)} ({_pct(ctx.pct_ba)} of spend). '
            f'{ctx.slow_movers_with_ba} ASIN(s) with fewer than 3 orders have BA spend.{suffix} '
            f'Remove slow movers from BA and redirect spend to best sellers.'
        )

    if _t('S038'):
        texts['S038'] = (
            f'BA represents {_pct(ctx.pct_ba)} of spend ({_dollar(ctx.spend_ba)}) but BAK harvest layer is missing. '
            + ('BAK campaign names exist but have no current spend. '
               if ctx.bak_name_overlaps_ba else 'No BAK campaigns detected. ')
            + 'Graduate proven BA search terms into BAK exact match campaigns.'
        )

    if _t('S039'):
        cat_note = ''
        if getattr(ctx, 'categories_above_10pct', 0) > 1:
            cat_note = f' {ctx.categories_above_10pct} categories each contribute >10% of total sales.'
        texts['S039'] = (
            f'Only {ctx.ba_campaign_count} BA campaign(s) detected.{cat_note} '
            f'Structure is not segmented by category — new BA campaigns by category are needed.'
        )

    if _t('S041'):
        severity = 'severe fragmentation' if ctx.low_order_campaign_count > 80 else 'high fragmentation'
        texts['S041'] = (
            f'{ctx.low_order_campaign_count} campaigns have only 1–3 orders in the period ({severity}). '
            f'Consolidate converting terms into BAK campaigns by parent ASIN.'
        )

    if _t('S044'):
        acos_pp = ctx.acos_actual * 100
        declining = ctx.yoy_ad_sales < -0.05
        acos_high = ctx.acos_constraint > 0 and acos_pp > ctx.acos_constraint * 1.2
        prefix = 'No Sponsored Brands spend detected. '
        if declining:
            suffix = f'Ad sales down {_pct(abs(ctx.yoy_ad_sales))} YoY — SB is a direct lever for upper-funnel recovery.'
        elif acos_high:
            suffix = f'ACoS is {acos_pp:.0f}% vs {ctx.acos_constraint:.0f}% constraint — address efficiency before launching SB.'
        else:
            suffix = 'SB campaigns should be launched to build upper-funnel coverage.'
        texts['S044'] = prefix + suffix

    if _t('S045'):
        texts['S045'] = (
            f'BAK spend is {_pct(ctx.pct_bak)} vs BA spend at {_pct(ctx.pct_ba)}. '
            f'BAK exists but receives less than 10% of its BA feeder spend. '
            f'The harvest cycle has stalled — review BA search term report and promote converting terms to BAK.'
        )

    if _t('S047'):
        texts['S047'] = (
            f'Imported campaigns: {_dollar(ctx.spend_imported)} ({_pct(ctx.pct_imported)} of spend). '
            f'These run outside the Quartile system. An import kickoff CoE ticket is needed.'
        )

    # S053 — SP Campaign ACoS Significantly Above Constraint
    if _t('S053') and ctx.sp_worst_campaign_acos > 0:
        extra = f' {ctx.sp_campaigns_above_threshold - 1} additional SP campaign(s) are also above the threshold.' if ctx.sp_campaigns_above_threshold > 1 else ''
        texts['S053'] = (
            f'The worst SP campaign has an ACoS of {ctx.sp_worst_campaign_acos:.0%} '
            f'vs the {ctx.acos_constraint:.0f}% constraint. '
            f'{ctx.sp_campaigns_above_threshold} SP campaign(s) are above the threshold.{extra}'
        )

    # S054 — SB Campaign ACoS Significantly Above Constraint
    if _t('S054') and ctx.sb_worst_campaign_acos > 0:
        extra = f' {ctx.sb_campaigns_above_threshold - 1} additional SB campaign(s) are also above the threshold.' if ctx.sb_campaigns_above_threshold > 1 else ''
        texts['S054'] = (
            f'The worst SB campaign has an ACoS of {ctx.sb_worst_campaign_acos:.0%} '
            f'vs the {ctx.acos_constraint:.0f}% constraint. '
            f'{ctx.sb_campaigns_above_threshold} SB campaign(s) are above the threshold.{extra}'
        )

    # S055 — SD Campaign ACoS Significantly Above Constraint
    if _t('S055') and ctx.sd_worst_campaign_acos > 0:
        extra = f' {ctx.sd_campaigns_above_threshold - 1} additional SD campaign(s) are also above the threshold.' if ctx.sd_campaigns_above_threshold > 1 else ''
        texts['S055'] = (
            f'The worst SD campaign has an ACoS of {ctx.sd_worst_campaign_acos:.0%} '
            f'vs the {ctx.acos_constraint:.0f}% constraint. '
            f'{ctx.sd_campaigns_above_threshold} SD campaign(s) are above the threshold.{extra}'
        )

    if _t('S056') and 'S036' not in flags:
        texts['S056'] = (
            f'ATM campaigns avg ACoS: {_pct(ctx.atm_avg_acos)} vs account avg {_pct(ctx.acos_actual)}. '
            f'ATM outperforming by {(1 - ctx.atm_avg_acos / ctx.acos_actual):.0%} — consider expanding ATM coverage.'
        )

    if _t('S057') and 'S036' not in flags:
        texts['S057'] = (
            f'BR campaigns avg ACoS: {_pct(ctx.br_avg_acos)} vs account avg {_pct(ctx.acos_actual)}. '
            f'Broad match outperforming — consider expanding BR_ campaigns.'
        )

    if _t('S058'):
        texts['S058'] = (
            f'PH campaigns avg ACoS: {_pct(ctx.ph_avg_acos)} vs account avg {_pct(ctx.acos_actual)}. '
            f'Phrase match outperforming — prioritise for future launches.'
        )

    if _t('S059'):
        texts['S059'] = (
            f'OW campaigns avg ACoS: {_pct(ctx.ow_avg_acos)} vs account avg {_pct(ctx.acos_actual)}. '
            f'Exact match outperforming — continue supporting OW-focused launches.'
        )

    if _t('S060'):
        texts['S060'] = (
            f'OP product-targeting avg ACoS: {_pct(ctx.op_avg_acos)} vs account avg {_pct(ctx.acos_actual)}. '
            f'Product-targeting outperforming — consider expanding OP_ coverage.'
        )

    if _t('S062'):
        texts['S062'] = (
            f'{ctx.paused_sb_count} SB campaign(s) paused with historical spend. '
            f'Current SB spend: {_dollar(ctx.spend_sb)}. '
            f'Rebuild paused SB campaigns with updated branded keyword structures.'
        )

    if _t('S063'):
        texts['S063'] = (
            f'SD_FLEX campaigns avg ACoS: {_pct(ctx.sd_flex_avg_acos)} vs account avg {_pct(ctx.acos_actual)}. '
            f'SD_FLEX outperforming — consider expanding SD_FLEX coverage.'
        )

    if _t('S064'):
        texts['S064'] = (
            f'SD_AUDI campaigns avg ACoS: {_pct(ctx.sd_audi_avg_acos)} vs account avg {_pct(ctx.acos_actual)}. '
            f'SD audience campaigns outperforming — consider expanding SD_AUDI coverage.'
        )

    if _t('S065'):
        texts['S065'] = (
            f'SD_PRD campaigns avg ACoS: {_pct(ctx.sd_prd_avg_acos)} vs account avg {_pct(ctx.acos_actual)}. '
            f'Product-page defense outperforming — consider expanding SD_PRD coverage.'
        )

    if _t('S067'):
        texts['S067'] = (
            f'SB campaigns avg ACoS: {_pct(ctx.sb_avg_acos)} vs account avg {_pct(ctx.acos_actual)}. '
            f'Sponsored Brands outperforming — consider increasing SB investment.'
        )

    if _t('S068'):
        texts['S068'] = (
            f'SBV campaigns avg ACoS: {_pct(ctx.sbv_avg_acos)} vs account avg {_pct(ctx.acos_actual)}. '
            f'Sponsored Brand Video outperforming — consider expanding SBV category targets.'
        )

    if _t('S069'):
        texts['S069'] = (
            f'{ctx.paused_sbv_count} SBV campaign(s) paused with historical spend. '
            f'SB is active ({_dollar(ctx.spend_sb)}). '
            f'SBV should run alongside SB — evaluate reactivation.'
        )

    if _t('S070'):
        op_note = ''
        if ctx.op_avg_acos > 0 and ctx.acos_actual > 0:
            op_note = (
                f'OP campaigns avg ACoS: {ctx.op_avg_acos:.0%} vs account avg {ctx.acos_actual:.0%} '
                f'({(1 - ctx.op_avg_acos / ctx.acos_actual):.0%} better). '
            )
        texts['S070'] = f'No CAT_SP campaigns detected. {op_note}Category-targeted SP campaigns should be launched.'

    if _t('S071'):
        prefix = 'No SBV campaigns detected. '
        if ctx.spend_sb > 0:
            prefix += f'SB active ({ctx.sb_impressions:,} impressions). '
        product_note = ''
        if op_outperforming:
            product_note = f'OP outperforming ({_pct(ctx.op_avg_acos)} vs {_pct(ctx.acos_actual)} avg). '
        elif catsp_outperforming:
            product_note = f'CAT_SP outperforming ({_pct(ctx.catsp_avg_acos)} vs {_pct(ctx.acos_actual)} avg). '
        texts['S071'] = prefix + product_note + 'Launch SBV product-targeting campaigns.'

    if _t('S072'):
        texts['S072'] = (
            f'BR avg ACoS: {_pct(ctx.br_avg_acos)}, OW avg ACoS: {_pct(ctx.ow_avg_acos)}. '
            f'Broad match outperforming exact — consider expanding BR_ or graduating more terms to BAK.'
        )

    if _t('S074'):
        texts['S074'] = (
            f'OW avg ACoS: {_pct(ctx.ow_avg_acos)}, BR avg ACoS: {_pct(ctx.br_avg_acos)}. '
            f'Exact match outperforming broad — graduate more BR terms into OW campaigns.'
        )

    if _t('S075'):
        kw_total = ctx.br_campaign_count + ctx.ow_campaign_count + ctx.ph_campaign_count
        texts['S075'] = (
            f'{ctx.op_campaign_count} OP campaigns vs {kw_total} keyword campaigns (OW+BR+PH). '
            f'Product-targeting is underdeveloped relative to keyword volume. Expand OP_ coverage.'
        )

    if _t('S076'):
        texts['S076'] = (
            f'CatchAll campaigns generated {ctx.catchall_orders:.0f} orders in the period. '
            f'BAK is only {_pct(ctx.pct_bak)} of spend. '
            f'Graduate high-converting CatchAll search terms into BAK campaigns.'
        )

    if _t('S077'):
        texts['S077'] = (
            f'CAT_SP avg ACoS: {_pct(ctx.catsp_avg_acos)} vs constraint {ctx.acos_constraint:.0f}%. '
            + ('Significantly above — category targeting needs refinement.' if ctx.catsp_avg_acos * 100 > ctx.acos_constraint * 0.85
               else 'Above target — review targeting scope.')
        )

    if _t('S079'):
        texts['S079'] = (
            f'{ctx.watm_campaign_count} WATM campaigns active. '
            f'Multiple WATM campaigns add fragmentation without structural benefit.'
        )

    if _t('S082'):
        texts['S082'] = (
            f'Branded terms represent {ctx.branded_spend_pct:.0%} of spend and '
            f'non-branded {ctx.non_branded_spend_pct:.0%} — both are significant inside the same BAK bucket. '
            f'Split BAK into branded and non-branded campaigns for independent bid control.'
        )

    if _t('S083'):
        texts['S083'] = (
            f'{ctx.spending_asin_count} of {ctx.catalog_asin_count} ASINs have ad spend. '
            f'WATM/CatchAll is active but less than 60% of catalog ASINs are spending. '
            f'Ensure all catalog products are included in the WATM or CatchAll structure.'
        )

    if _t('S084'):
        texts['S084'] = (
            f'Both WATM and CatchAll campaigns are active simultaneously. '
            f'They serve the same purpose — only one should be active. Review and remove the redundant structure.'
        )

    if _t('S085'):
        texts['S085'] = (
            f'WATM campaigns exist but account for only {_pct(ctx.pct_watm)} of total spend. '
            f'WATM is not receiving meaningful budget. Consider switching to a CatchAll structure.'
        )

    if _t('S086'):
        over_threshold = [
            b for b in ctx.bak_campaigns
            if b['pct_of_total'] > 0.15 and b['acos'] > (ctx.acos_constraint / 100) * 0.50
        ]
        camp_lines = '; '.join(
            f"{b['name']} ({_pct(b['pct_of_total'])} of spend, {b['acos']:.0%} ACoS)"
            for b in over_threshold[:3]
        )
        texts['S086'] = (
            f'{len(over_threshold)} BAK campaign(s) exceed 15% of total spend with ACoS above constraint threshold. '
            + (f'Campaigns: {camp_lines}. ' if camp_lines else '')
            + 'Review top BAK terms and add negatives for wasteful keywords.'
        )

    if _t('S087'):
        texts['S087'] = (
            f'CAT_SP represents {_pct(getattr(ctx, "pct_cat_sp", 0.0))} of spend '
            f'with account ACoS at {ctx.acos_actual:.0%} vs {ctx.acos_constraint:.0f}% constraint. '
            f'Review CAT_SP campaigns — remove high-spend targets not meeting the efficiency target.'
        )

    if _t('S088'):
        texts['S088'] = (
            f'SB represents {_pct(ctx.pct_sb)} of spend '
            f'with account ACoS at {ctx.acos_actual:.0%} vs {ctx.acos_constraint:.0f}% constraint. '
            f'Review SB campaigns — remove high-spend keywords not meeting the efficiency target.'
        )

    if _t('S089'):
        texts['S089'] = (
            f'SBV represents {_pct(ctx.pct_sbv)} of spend '
            f'with account ACoS at {ctx.acos_actual:.0%} vs {ctx.acos_constraint:.0f}% constraint. '
            f'Review SBV campaigns — remove high-spend targets not meeting the efficiency target.'
        )

    if _t('S092'):
        texts['S092'] = (
            f'No SD campaigns active. SD spend $0. '
            f'Growth/Expansion objective with sufficient order velocity ({ctx.max_asin_orders_30d:.0f} orders on top ASIN). '
            f'Product-view remarketing and audience retargeting are not running.'
        )

    if _t('S093'):
        texts['S093'] = (
            f'No ATC retargeting campaigns detected. ProSuite is active. '
            f'SD spend: {_dollar(ctx.spend_sd)}. '
            f'Add-to-cart retargeting via SD_FLEX_ATC should be deployed.'
        )

    if _t('S096'):
        texts['S096'] = (
            f'SD expansion signal is present but top-selling ASIN has only '
            f'{ctx.max_asin_orders_30d:.0f} orders in the period. '
            f'Retargeting audience pool is too small to be effective. '
            f'Wait until top ASIN reaches ≥50 orders/month before launching SD.'
        )

    if _t('S097'):
        texts['S097'] = (
            f'{ctx.portfolio_count} portfolios active. '
            f'{ctx.managed_portfolio_count} managed. {ctx.portfolios_with_budget_cap} have budget caps. '
            f'Portfolio governance needs to be tightened.'
        )

    if _t('S098'):
        texts['S098'] = (
            f'Campaign-level ACoS overrides are active while ACoS is above constraint '
            f'({ctx.acos_actual:.0%} vs {ctx.acos_constraint:.0f}%). '
            f'Review each override — confirm it is intentional and still valid.'
        )

    if _t('S099'):
        texts['S099'] = (
            f'Product-level ACoS overrides are active while account ACoS is above constraint. '
            f'Review product overrides and confirm each is intentional.'
        )

    if _t('S100'):
        texts['S100'] = (
            f'VCPM represents {_pct(ctx.vcpm_spend_pct)} of SD spend. '
            f'VCPM on products without consistent Buy Box ownership wastes impressions.'
        )

    if _t('S101'):
        tags = [t.lower().strip() for t in (getattr(ctx, 'tags', None) or []) if t]
        has_bestseller = any(
            any(w in t for w in {'bestseller', 'best seller', 'hero', 'top', 'winner', 'core', 'priority'})
            for t in tags
        )
        has_segment = any(
            any(w in t for w in {'mid seller', 'slow mover', 'low perf', 'mid perf',
                                  'high traffic', 'low traffic'})
            for t in tags
        )
        if not has_bestseller and not has_segment:
            texts['S101'] = (
                f'No bestseller or performance-tier labels found in campaign tags. '
                f'At {_dollar(ctx.total_spend)}/month, the team has no visibility into how the portfolio is prioritised. '
                f'Defensive and tier-based strategy cannot be executed consistently.'
            )
        elif not has_bestseller:
            texts['S101'] = (
                f'Campaign tags include performance-tier labels but no bestseller or hero product label. '
                f'The highest-priority ASINs are not clearly identified — defensive coverage cannot be anchored.'
            )
        else:
            texts['S101'] = (
                f'Campaign tags include a bestseller label but no performance-tier segmentation. '
                f'Mid and slow movers are not separated — budget allocation relies on manual recall.'
            )

    if _t('S103'):
        texts['S103'] = (
            f'SBV campaigns active but not all follow the SBV_ naming convention. '
            f'Non-standard naming reduces governance clarity.'
        )

    if _t('S108'):
        texts['S108'] = (
            f'{getattr(ctx, "total_campaign_count", 0)} campaigns active but no ProSuite AMC audiences applied. '
            f'Test Amazon native audiences on the strongest SP campaigns.'
        )

    if _t('S109'):
        ineff_count = ctx.inefficient_asin_count
        asin_names  = getattr(ctx, 'inefficient_asin_names', [])
        shown       = asin_names[:3]
        more        = ineff_count - len(shown)
        asin_str    = ', '.join(shown)
        more_str    = f' (+{more} more)' if more > 0 else ''
        texts['S109'] = (
            f'{ineff_count} ASIN(s) are spending without generating meaningful sales '
            f'(no sales or ACoS >2× constraint). '
            f'ASINs: {asin_str}{more_str}. '
            f'Reduce or pause spend on these ASINs and reallocate budget to top performers.'
        )

    if _t('S110'):
        _s110_branded_note = ''
        if ctx.branded_spend_pct > 0 and ctx.branded_spend_pct < 0.05:
            _s110_branded_note = (
                f' Branded search term spend is {_pct(ctx.branded_spend_pct)} of total — '
                f'below the 5% target. SBV is key to defending and growing branded share.'
            )
        texts['S110'] = (
            f'SB active ({ctx.sb_impressions:,} impressions, {_pct(ctx.pct_sb)} of spend) but SBV spend is $0. '
            f'SBV is the natural next step — launch video campaigns on the same category targets as SB.'
            + _s110_branded_note
        )

    if _t('S113'):
        base = 'Subscribe & Save is not active. '
        if ctx.repeat_purchase == 'High':
            base += f'Repeat purchase behavior is High — SnS is a strong retention lever for this account. '
        base += _pct(ctx.yoy_ad_sales) + ' YoY ad sales. ' if ctx.yoy_ad_sales != 0 else ''
        texts['S113'] = base + 'Review SnS activation with the client.'

    if _t('S114'):
        texts['S114'] = (
            f'Ad sales declined {abs(ctx.yoy_ad_sales):.0%} YoY while spend increased {ctx.mom_spend_change:.0%} MoM. '
            f'More budget going in, less revenue coming out. Budget and campaign scope must be reviewed.'
        )

    if _t('S119'):
        base = 'Subscribe & Save is not active. '
        if ctx.repeat_purchase == 'High':
            base += 'Repeat purchase behavior is High — SnS is a strong retention lever. '
        base += (f'YoY ad sales: {ctx.yoy_ad_sales:+.0%}. ' if ctx.yoy_ad_sales != 0 else '')
        texts['S119'] = base + 'SnS should be evaluated as a retention and growth lever.'

    if _t('S122'):
        texts['S122'] = (
            f'{ctx.promo_asin_count} ASIN(s) in active promo. '
            + (f'Promo cost rate averaging {_pct(ctx.promo_cost_rate)}. ' if ctx.promo_cost_rate > 0 else '')
            + 'Portfolio budgets should be reviewed to prevent intraday depletion.'
        )

    if _t('S124'):
        sd_note = f'SD spend: {_dollar(ctx.spend_sd)} ({ctx.pct_sd:.0%} of total). ' if ctx.spend_sd > 0 else 'SD spend: $0. '
        texts['S124'] = (
            f'GGS status: {ctx.ggs_status}. {sd_note}'
            f'SD campaigns need to reach at least 5% of total spend to satisfy the GGS commitment.'
        )

    if _t('S125'):
        texts['S125'] = (
            f'SD active ({_dollar(ctx.spend_sd)}) but no SD_FLEX or SD_AUDI remarketing campaigns. '
            f'Product-view remarketing is not running.'
        )

    if _t('S126'):
        texts['S126'] = (
            f'SD active ({_dollar(ctx.spend_sd)}) but no ATC retargeting in place. '
            f'Add-to-cart retargeting via ProSuite AMC is not activated.'
        )

    return texts


def _build_what_you_should_do(ctx: StrategyContext, flags: dict[str, str]) -> dict[str, str]:
    """
    Builds dynamic 'What You Should Do' text for controls where we defined
    specific actionable instructions. Only covers controls explicitly scoped:
    S021, S053, S054, S055, S109.
    All other controls keep their static template text.
    """
    how: dict[str, str] = {}
    _t = lambda sid: sid in flags

    if _t('S021'):
        oob_case = getattr(ctx, '_oob_case', 'inefficient')
        if oob_case == 'efficient':
            how['S021'] = (
                'The account is clean and hitting budget limits. '
                'This is the right time to negotiate a budget increase with the client. '
                'Present the efficiency data (ACoS vs constraint) as justification.'
            )
        elif oob_case == 'inefficient':
            how['S021'] = (
                'Fix efficiency before requesting more budget. '
                'Reduce the ACoS target to lower CPC pressure and reduce out-of-budget events. '
                'Once ACoS is within constraint, revisit budget expansion with the client.'
            )
        else:
            how['S021'] = (
                'Review spend allocation before increasing budget. '
                'Sales are declining despite full budget utilisation — more budget will not fix the problem. '
                'Identify which campaigns are spending without delivering sales and reduce their scope first.'
            )

    if _t('S053'):
        how['S053'] = (
            'Revisit the SP campaigns listed above. '
            'These campaigns are running above the agreed ACoS constraint and need a CSM review.'
        )

    if _t('S054'):
        how['S054'] = (
            'Revisit the SB campaigns listed above. '
            'These campaigns are running above the agreed ACoS constraint and need a CSM review.'
        )

    if _t('S055'):
        how['S055'] = (
            'Revisit the SD campaigns listed above. '
            'These campaigns are running above the agreed ACoS constraint and need a CSM review.'
        )

    if _t('S109'):
        asin_names = getattr(ctx, 'inefficient_asin_names', [])
        shown      = asin_names[:3]
        more       = ctx.inefficient_asin_count - len(shown)
        asin_str   = ', '.join(shown) if shown else 'the ASINs listed'
        more_str   = f' (+{more} more)' if more > 0 else ''
        how['S109'] = (
            f'Reduce or pause spend on {asin_str}{more_str}. '
            f'Start with the highest-spend ASIN that has zero sales. '
            f'Reallocate that budget to top-performing ASINs. '
            f'Review each flagged ASIN for PDP issues, pricing, or review count before resuming spend.'
        )

    return how
