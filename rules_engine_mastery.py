from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple

import pandas as pd

from config_mastery import CONTROL_NAMES, IMPACT_LABEL, IMPORTANCE, PRIORITY_POINTS, SCORING_EXCLUDED, SOURCES, WHY, ControlResult
from reader_databricks_mastery import DatabricksContext, clean_text, money_str, monthly_budget_from_daily, norm_pct, pct_str, to_float, trim

OBJECTIVE_WORDS = {'objective', 'goal', 'grow', 'growth', 'scale', 'increase', 'improve', 'stabilize', 'maintain', 'reduce', 'defend', 'accelerate', 'awareness', 'sales', 'profit', 'profitability', 'ranking', 'market share'}
KPI_WORDS = {'roas', 'acos', 'tacos', 'spend', 'sales', 'cvr', 'ctr', 'cpc', 'ntb', 'rank', 'revenue'}
CONSTRAINT_WORDS = {'constraint', 'below', 'above', 'maintain', 'limit', 'threshold', 'guardrail', 'while', 'without', 'at or below'}
CHALLENGE_WORDS = {'challenge', 'issue', 'risk', 'inventory', 'out-of-stock', 'out of stock', 'slowdown', 'pressure', 'volatility', 'sensitive', 'buy box', 'listing', 'margin', 'competition', 'competitive', 'growth is not', 'not meeting', 'incomplete', 'dissatisfied', 'struggling', 'difficult', 'barrier', 'blocker', 'problem', 'concern', 'lack', 'limited', 'unable', 'falling', 'declined', 'losing'}

# Strong operational constraint signals — used by C005 cross-field check.
# Must be explicit enough to distinguish a real constraint from general commentary.
CONSTRAINT_SIGNALS = {
    'inventory', 'out of stock', 'out-of-stock', 'stock constraint', 'stock limit',
    'price increase', 'pricing constraint', 'price cap', 'margin constraint', 'margin limit',
    'budget cap', 'budget constraint', 'budget limit', 'spend cap', 'spend limit',
    'product restriction', 'category restriction', 'listing restriction', 'policy restriction',
    'restricted', 'ip constraint', 'ip restriction', 'intellectual property',
    'reseller', 'buy box constraint', 'buy box issue',
    'logistics', 'shipping constraint', 'fulfillment constraint',
    'cash flow', 'financial constraint',
    'cannot advertise', 'not allowed', 'compliance',
    'seasonal cap', 'seasonal constraint',
}

# Tactical campaign actions — used by C001 to catch AY7 filled with
# campaign tactics instead of a strategic business objective.
TACTICAL_ONLY_WORDS = {
    'increase budget', 'decrease budget', 'raise budget', 'lower budget',
    'increase bids', 'lower bids', 'adjust bids', 'bid adjustment',
    'pause campaign', 'launch campaign', 'add campaign', 'create campaign',
    'add keywords', 'add negatives', 'add negative keywords',
    'change match type', 'update targeting', 'fix targeting',
    'increase daily budget', 'decrease daily budget',
}

# Business outcome words — used by C001 to confirm the objective is strategic.
BUSINESS_OUTCOME_WORDS = {
    'sales', 'revenue', 'profit', 'profitability', 'growth', 'market share',
    'organic rank', 'organic ranking', 'brand awareness', 'new to brand',
    'customer acquisition', 'return', 'roas', 'tacos', 'acos',
    'yoy', 'year over year', 'year-over-year', 'quarter', 'quarterly',
    'margin', 'efficiency', 'ntb', 'new customers', 'brand growth',
}
TIME_WORDS = {'q1', 'q2', 'q3', 'q4', 'month', 'monthly', 'weekly', 'this period', 'near-term', 'near term', 'next', 'current period', 'prime day', 'holiday', 'bfcm', 'seasonal'}
CONFLICT_WORDS = {'but', 'however', 'while', 'tradeoff', 'trade-off', 'contrasting', 'despite', 'volatility', 'elevated', 'balancing'}
BESTSELLER_WORDS = {'bestseller', 'best seller', 'hero', 'top perf', 'top perf.', 'top', 'winner', 'core', 'priority', 'best-seller'}
SEGMENTATION_WORDS = {'mid seller', 'mid-seller', 'slow mover', 'slow-mover', 'low perf', 'low perf.', 'mid. perf.', 'high traffic', 'low traffic', 'high conversion', 'low conversion'}

# Positive category / product-type tag vocabulary — used by C012.
# These are intentional product grouping labels a CSM would apply.
# Rules:
#   1. Positive-match only — no "anything not in other sets" logic.
#   2. No short ambiguous substrings (e.g. 'us', 'ca', 'uk') that appear
#      inside common words. Geo codes use word-boundary regex instead (see C012).
#   3. Covers product types, sales tiers, bundles, and explicit segment labels.
CATEGORY_WORDS = {
    # Product type groupings
    'supplement', 'vitamin', 'protein', 'powder', 'capsule', 'gummy', 'softgel', 'liquid',
    'snack', 'food', 'beverage', 'coffee', 'tea',
    'skincare', 'haircare', 'bodycare', 'beauty', 'personal care',
    'tool', 'device', 'equipment', 'gear', 'accessory', 'accessories',
    'apparel', 'clothing', 'footwear',
    'kitchen', 'garden', 'outdoor', 'cleaning',
    'pet', 'baby', 'fitness',
    'electronics', 'office supplies',
    # Sales / performance tier labels  (written-out form to avoid substring hits)
    'tier 1', 'tier 2', 'tier 3', 'tier 4', 'tier 5',
    'tier1', 'tier2', 'tier3', 'tier4', 'tier5',
    # Bundle / variant groupings
    'bundle', 'multipack', 'variety pack', 'starter kit',
    # Explicit segment labels
    'private label', 'flagship', 'new launch',
    'limited edition', 'clearance', 'seasonal',
    # Generic intentional segment labels
    'niche', 'cross-sell',
}

# Short geo codes that must match as whole words (checked separately in C012)
CATEGORY_GEO_CODES = {'us', 'ca', 'uk', 'eu', 'au'}
MONTH_ALIASES = {'jan': 1, 'january': 1, 'feb': 2, 'february': 2, 'mar': 3, 'march': 3, 'apr': 4, 'april': 4, 'may': 5, 'jun': 6, 'june': 6, 'jul': 7, 'july': 7, 'aug': 8, 'august': 8, 'sep': 9, 'sept': 9, 'september': 9, 'oct': 10, 'october': 10, 'nov': 11, 'november': 11, 'dec': 12, 'december': 12}
NEGATIVE_EXCEPTIONS = ['deal', 'deals', 'discount', 'black friday', 'cyber monday', 'prime day', 'holiday']
PERSONALIZATION_KEYWORDS = {
    'unmanaged_asin': ['unmanaged asin', 'asin excluded', 'excluded asin', 'unmanaged product'],
    'timeframe_boost': ['timeframe boost', 'boost period', 'boosted timeframe', 'temporary boost'],
    'unmanaged_budget': ['unmanaged budget', 'budget override', 'budget unmanaged'],
    'negative_keywords': ['negative keyword', 'global negative', 'negative terms'],
    'unmanaged_campaigns': ['unmanaged campaign', 'campaign unmanaged'],
    'unmanaged_campaign_budget': ['campaign budget override', 'unmanaged campaign budget', 'campaign budget unmanaged'],
    'rbo_config': ['rbo', 'rule based optimization', 'rule-based optimization'],
    'product_level_acos': ['product level acos', 'asin level acos', 'product acos override'],
    'campaign_level_acos': ['campaign level acos', 'campaign acos override'],
}


def has_any(text: str, words: set[str]) -> bool:
    t = clean_text(text).lower()
    return any(w in t for w in words)


def parse_months_from_text(text: str) -> set[int]:
    t = clean_text(text).lower()
    months: set[int] = set()
    if not t:
        return months
    q_map = {'q1': {1, 2, 3}, 'q2': {4, 5, 6}, 'q3': {7, 8, 9}, 'q4': {10, 11, 12}}
    for q, ms in q_map.items():
        if q in t:
            months |= ms
    for k, v in MONTH_ALIASES.items():
        if re.search(rf'\b{k}\b', t):
            months.add(v)
    for m in re.finditer(r'\b(1[0-2]|0?[1-9])\s*[-/]\s*(1[0-2]|0?[1-9])\b', t):
        a = int(m.group(1)); b = int(m.group(2))
        if a <= b:
            months |= set(range(a, b + 1))
    month_keys = '|'.join(sorted(MONTH_ALIASES.keys(), key=len, reverse=True))
    for m in re.finditer(rf'\b({month_keys})\b\s*(?:-|to|through|thru)\s*\b({month_keys})\b', t):
        a = MONTH_ALIASES[m.group(1)]; b = MONTH_ALIASES[m.group(2)]
        if a <= b:
            months |= set(range(a, b + 1))
    if 'prime day' in t:
        months.add(7)
    return months


def classify_concentration(top1: float, top3: float, top5: float) -> str:
    if top1 > 0.5 or top3 > 0.75 or top5 > 0.8:
        return 'high'
    if top1 >= 0.25 or top3 >= 0.55 or top5 >= 0.60:
        return 'medium'
    return 'low'


def _find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    norm = {str(c).strip().lower().replace(' ', '').replace('_', ''): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower().replace(' ', '').replace('_', '')
        if key in norm:
            return norm[key]
    return None


def _nonempty_df(df: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
    if df is None or df.empty:
        return None
    tmp = df.copy()
    tmp = tmp.dropna(how='all')
    tmp = tmp.loc[:, ~tmp.columns.astype(str).str.contains('^Unnamed', case=False, na=False)]
    if tmp.empty:
        return None
    return tmp


def _is_exception_negative(term: str) -> bool:
    t = clean_text(term).lower()
    return bool(t) and any(k in t for k in NEGATIVE_EXCEPTIONS)


def _active_end_date_rows(df: Optional[pd.DataFrame], ref_date, idx: int) -> int:
    df = _nonempty_df(df)
    if df is None or ref_date is None or df.shape[1] <= idx:
        return 0
    end_dates = pd.to_datetime(df.iloc[:, idx], errors='coerce')
    return int((end_dates > pd.Timestamp(ref_date)).sum())


def detect_personalizations(ctx: DatabricksContext) -> List[str]:
    active: List[str] = []
    if _nonempty_df(ctx.df34) is not None:
        active.append('product_level_acos')
    if _nonempty_df(ctx.df35) is not None:
        active.append('campaign_level_acos')

    tf = _nonempty_df(ctx.df27)
    if tf is not None:
        status_col = _find_col(tf, ['status', 'statusname'])
        if status_col:
            statuses = tf[status_col].astype(str).fillna('').str.strip().str.lower()
            active_mask = (statuses != '') & (statuses != 'expired')
            if active_mask.any():
                active.append('timeframe_boost')
        else:
            active.append('timeframe_boost')

    neg = _nonempty_df(ctx.df29)
    if neg is not None:
        neg_col = _find_col(neg, ['negative_word', 'negative word', 'negative', 'keyword'])
        prod_col = _find_col(neg, ['product', 'asin', 'targetasin'])
        if neg_col:
            tmp = neg.copy()
            tmp['_neg'] = tmp[neg_col].astype(str).fillna('').str.strip()
            tmp = tmp[tmp['_neg'] != '']
            if not tmp.empty:
                if prod_col:
                    tmp['_prod'] = tmp[prod_col].astype(str).fillna('').str.strip()
                    acct = tmp[tmp['_prod'] == '']
                    prod = tmp[tmp['_prod'] != '']
                else:
                    acct = tmp
                    prod = pd.DataFrame()
                if any(not _is_exception_negative(x) for x in acct['_neg'].tolist()):
                    active.append('negative_keywords')
                elif not prod.empty and any(not _is_exception_negative(x) for x in prod['_neg'].tolist()):
                    active.append('negative_keywords')

    if _active_end_date_rows(ctx.df26, ctx.ref_date, 4) > 0:
        active.append('unmanaged_asin')
    if _active_end_date_rows(ctx.df28, ctx.ref_date, 6) > 0:
        active.append('unmanaged_budget')
    if _active_end_date_rows(ctx.df31, ctx.ref_date, 11) > 0:
        active.append('unmanaged_campaigns')
    if _active_end_date_rows(ctx.df32, ctx.ref_date, 6) > 0:
        active.append('unmanaged_campaign_budget')
    if _nonempty_df(ctx.df33) is not None:
        active.append('rbo_config')
    return sorted(set(active))


def documented_personalizations(note_text: str, active_types: List[str]) -> Tuple[int, List[str]]:
    note = clean_text(note_text).lower()
    if not active_types:
        return 0, []
    matched = []
    generic = any(x in note for x in ['custom', 'exception', 'manual', 'override', 'testing', 'temporary', 'special handling', 'out of framework'])
    for key in active_types:
        kws = PERSONALIZATION_KEYWORDS.get(key, [])
        if any(kw in note for kw in kws):
            matched.append(key)
    if generic and not matched and active_types:
        matched.append(active_types[0])
    return len(set(matched)), sorted(set(matched))


def build_primary_objective(ctx: DatabricksContext, results: Dict[str, ControlResult]) -> str:
    ay = clean_text(ctx.ay)
    am = clean_text(ctx.am)
    if results['C001'].status == 'FLAG':
        return 'Primary objective is not clearly documented.'
    if not ay and not am:
        return 'Primary objective is not clearly documented.'
    if results['C002'].status == 'FLAG':
        return f"Primary objective is documented as {trim(ay or am, 180)}, but strategic context is incomplete."
    if results['C002'].status == 'PARTIAL':
        return f"The primary objective is to {trim(ay or am, 160)}, but the supporting KPI, timeframe, or constraint context is incomplete."
    if ay and am:
        return f"The primary objective is to {trim(ay, 140)}, with supporting context that {trim(am, 220)}"
    return f"The primary objective is to {trim(ay or am, 220)}"


def _fallback_results() -> Dict[str, ControlResult]:
    """Returns a fully-flagged result set used when evaluate_all() fails mid-run."""
    return {
        cid: ControlResult('FLAG', 'Evaluation failed — check input file and re-run.', WHY[cid], SOURCES[cid])
        for cid in CONTROL_NAMES
    }


def evaluate_all(ctx: DatabricksContext) -> Dict[str, ControlResult]:
    try:
        return _evaluate_all_inner(ctx)
    except Exception as exc:
        import traceback
        print(f"[rules_engine] evaluate_all() failed: {exc}\n{traceback.format_exc()}")
        return _fallback_results()


def _evaluate_all_inner(ctx: DatabricksContext) -> Dict[str, ControlResult]:
    r: Dict[str, ControlResult] = {}

    # -------------------------------------------------------------------------
    # C001 — Objective Clearly Defined
    # Primary source: AY7 narrative (ctx.ay).
    # Cross-reference: sf_primary_objective from tab 55 (structured CSP field).
    # OK:      narrative has outcome language + measurable anchor; SF field aligned.
    # PARTIAL: narrative OK but SF field blank (CSP completeness gap), OR
    #          narrative absent but SF field has usable content (use SF, note gap).
    # FLAG:    both sources empty, or neither has valid business outcome language.
    # -------------------------------------------------------------------------
    txt = ctx.ay
    sf_obj = ctx.sf_primary_objective  # structured CSP field from tab 55
    t_lower = clean_text(txt).lower() if txt else ''
    sf_lower = sf_obj.lower() if sf_obj else ''

    def _score_objective_text(t: str) -> str:
        """Return 'ok', 'partial', or 'flag' for a given objective text."""
        tl = t.lower()
        has_outcome = has_any(tl, BUSINESS_OUTCOME_WORDS)
        has_number = bool(re.search(r'\d+\s*%|\$\s*\d+|\d+\s*x\b|\d+[kKmM]\b|\bROAS\b|\bACoS\b|\bTACoS\b', t, re.I))
        outcome_count = sum(1 for w in BUSINESS_OUTCOME_WORDS if w in tl)
        is_purely_tactical = has_any(tl, TACTICAL_ONLY_WORDS) and not has_outcome
        if is_purely_tactical:
            return 'partial'
        if has_outcome and (has_number or outcome_count >= 2):
            return 'ok'
        if has_outcome:
            return 'partial'
        return 'flag'

    if not txt and not sf_obj:
        r['C001'] = ControlResult('FLAG', 'No primary objective is written in the account notes (AY7), and the CSP Primary Objective field is also empty.', WHY['C001'], SOURCES['C001'])
    elif not txt and sf_obj:
        # AY7 empty but SF has content — use SF, note the narrative gap
        sf_score = _score_objective_text(sf_obj)
        if sf_score == 'ok':
            r['C001'] = ControlResult('PARTIAL', f'AY7 narrative is empty, but the CSP Primary Objective field is documented: "{trim(sf_obj, 180)}". Copy this into AY7 to close the gap.', WHY['C001'], SOURCES['C001'])
        else:
            r['C001'] = ControlResult('PARTIAL', f'AY7 narrative is empty. The CSP Primary Objective field has content but it is not anchored to a measurable target: "{trim(sf_obj, 180)}". Update both sources.', WHY['C001'], SOURCES['C001'])
    else:
        # AY7 has content — evaluate it first
        score = _score_objective_text(txt)
        if score == 'flag':
            r['C001'] = ControlResult('FLAG', 'The objective field does not contain a clear business goal. It needs to explain what the account is trying to achieve and why.', WHY['C001'], SOURCES['C001'])
        elif score == 'partial':
            if has_any(t_lower, TACTICAL_ONLY_WORDS) and not has_any(t_lower, BUSINESS_OUTCOME_WORDS):
                r['C001'] = ControlResult('PARTIAL', 'The objective field describes campaign actions, not a business goal. Rewrite it to focus on the business outcome (e.g. growth, profitability, market share).', WHY['C001'], SOURCES['C001'])
            else:
                r['C001'] = ControlResult('PARTIAL', 'Objective is written, but it is not anchored to a measurable target or specific KPI. Add a number or metric to make it actionable.', WHY['C001'], SOURCES['C001'])
        else:
            # AY7 is OK — cross-check CSP field
            if not sf_obj:
                r['C001'] = ControlResult('PARTIAL', f'Objective is documented in AY7 and is clear. However, the CSP Primary Objective field (Salesforce) is empty — the Salesforce record is incomplete.', WHY['C001'], SOURCES['C001'])
            else:
                r['C001'] = ControlResult('OK', 'Primary objective is documented and linked to a clear business outcome. CSP field is also populated.', WHY['C001'], SOURCES['C001'])

    # -------------------------------------------------------------------------
    # C002 — Objective vs Near-Term Alignment
    # Primary source: AM7 narrative (ctx.am).
    # Enhancements from tab 55:
    #   - sf_near_term_conflict: explicit Yes/No field — 6th check, no keyword inference needed.
    #   - sf_primary_spend_kpi: gates the 'kpi' dimension (structured > keyword).
    #   - sf_near_term: fallback source if AM7 is empty.
    # Timeframe is still a hard gate — without it the result cannot be OK.
    # Requires all dimensions for OK; at least 3 for PARTIAL.
    # -------------------------------------------------------------------------
    txt = ctx.am
    sf_near = ctx.sf_near_term
    sf_conflict = ctx.sf_near_term_conflict.strip() if ctx.sf_near_term_conflict else ''
    sf_kpi = ctx.sf_primary_spend_kpi  # 'ACOS', 'ROAS', 'TACOS', or ''

    # If AM7 is empty, use sf_near_term as fallback source
    eval_text = txt if txt else sf_near
    source_note = '' if txt else ' (from CSP Salesforce field — AY7 narrative is empty)'

    if not eval_text:
        if sf_conflict in ('Yes', 'No'):
            r['C002'] = ControlResult('FLAG', f'The objective context field (AM7) is empty and no near-term text is documented in the CSP. However, the Conflict field is set to "{sf_conflict}". Add supporting context to explain the near-term situation.', WHY['C002'], SOURCES['C002'])
        else:
            r['C002'] = ControlResult('FLAG', 'The objective context field (AM7) is empty and the CSP Near-Term Considerations field is also blank. There is no supporting detail for the primary objective.', WHY['C002'], SOURCES['C002'])
    else:
        # KPI dimension: prefer structured sf_primary_spend_kpi, fall back to keyword match
        kpi_ok = sf_kpi in ('ACOS', 'ROAS', 'TACOS') or has_any(eval_text, KPI_WORDS)
        dims = {
            'objective':  has_any(eval_text, OBJECTIVE_WORDS),
            'kpi':        kpi_ok,
            'constraint': has_any(eval_text, CONSTRAINT_WORDS),
            'context':    len(eval_text.split()) >= 15,
            'timeframe':  has_any(eval_text, TIME_WORDS),
        }
        # 6th check: near-term conflict assessment (structured field)
        conflict_assessed = sf_conflict in ('Yes', 'No')
        n = sum(dims.values()) + (1 if conflict_assessed else 0)
        total_possible = 6
        missing = [k for k, v in dims.items() if not v]
        if not conflict_assessed:
            missing.append('conflict assessment')
        has_timeframe = dims['timeframe']
        conflict_note = f' Conflict with primary objective: {sf_conflict}.' if conflict_assessed else ''

        if n == total_possible:
            r['C002'] = ControlResult('OK', f'Objective context covers all elements: goal, KPI, constraint, timeframe, narrative depth, and conflict assessment.{source_note}{conflict_note}', WHY['C002'], SOURCES['C002'])
        elif n >= 4 and has_timeframe:
            r['C002'] = ControlResult('PARTIAL', f'Objective context is written but some elements are missing: {", ".join(missing)}.{source_note}{conflict_note}', WHY['C002'], SOURCES['C002'])
        elif n >= 3 and not has_timeframe:
            r['C002'] = ControlResult('PARTIAL', f'Objective context is written but has no timeframe or near-term reference. Also missing: {", ".join([m for m in missing if m != "timeframe"])}.{source_note}{conflict_note}', WHY['C002'], SOURCES['C002'])
        else:
            r['C002'] = ControlResult('FLAG', f'Objective context does not have enough detail. Missing elements: {", ".join(missing)}.{source_note}', WHY['C002'], SOURCES['C002'])

    # -------------------------------------------------------------------------
    # C003 — Account Challenges Documented
    # Primary source: BN7 narrative (ctx.bn).
    # Cross-reference: sf_current_challenges from tab 55 (structured CSP field).
    # If BN7 is empty but SF has content — evaluate SF text instead of auto-FLAG.
    # If BN7 is OK but SF is blank — PARTIAL (CSP completeness gap).
    # -------------------------------------------------------------------------
    txt = ctx.bn
    sf_chal = ctx.sf_current_challenges

    eval_text = txt if txt else sf_chal
    source_note = '' if txt else ' (from CSP Salesforce field — BN7 narrative is empty)'

    if not eval_text:
        r['C003'] = ControlResult('FLAG', 'No current challenges are documented in the account notes (BN7) or the CSP Salesforce record.', WHY['C003'], SOURCES['C003'])
    else:
        t_lower = clean_text(eval_text).lower()
        metric_target_count = len(re.findall(
            r'\b(acos|tacos|roas|spend|sales|revenue)\b.{0,30}(\d+\s*%|\$\s*\d+|\d+\s*x\b)',
            t_lower, re.I
        ))
        has_barrier = has_any(t_lower, CHALLENGE_WORDS)
        has_specific = len(eval_text.split()) >= 12 and has_barrier
        if metric_target_count >= 2 and not has_barrier:
            r['C003'] = ControlResult('FLAG', f'The challenges field contains performance targets, not challenges. Replace the content with the actual blockers and issues the account is facing.{source_note}', WHY['C003'], SOURCES['C003'])
        elif has_specific:
            if txt and not sf_chal:
                # BN7 is OK but CSP field empty
                r['C003'] = ControlResult('PARTIAL', f'Current challenges are documented with enough detail. However, the CSP Current Challenges field (Salesforce) is empty — the Salesforce record is incomplete.', WHY['C003'], SOURCES['C003'])
            else:
                r['C003'] = ControlResult('OK', f'Current challenges are documented with enough detail to understand the active account blockers.{source_note}', WHY['C003'], SOURCES['C003'])
        elif len(eval_text.split()) >= 6:
            r['C003'] = ControlResult('PARTIAL', f'Challenges are written, but the description is too general. It does not clearly explain what is blocking the account today.{source_note}', WHY['C003'], SOURCES['C003'])
        else:
            r['C003'] = ControlResult('FLAG', f'The challenges field has very little content. More detail is needed for a proper review.{source_note}', WHY['C003'], SOURCES['C003'])

    # -------------------------------------------------------------------------
    # C004 — Seasonality Awareness
    # -------------------------------------------------------------------------
    source_months = parse_months_from_text(ctx.am)
    mention_months = set()
    for text in [ctx.ay, ctx.bn]:
        mention_months |= parse_months_from_text(text)
    if source_months and mention_months:
        r['C004'] = ControlResult('OK', f'Seasonality is documented and consistent across account fields. Seasonal months detected: {sorted(source_months)}.', WHY['C004'], SOURCES['C004'])
    elif source_months and not mention_months:
        r['C004'] = ControlResult('FLAG', f'Seasonality was detected in the account context (months: {sorted(source_months)}), but it is not referenced in the main narrative fields.', WHY['C004'], SOURCES['C004'])
    elif not source_months and mention_months:
        r['C004'] = ControlResult('PARTIAL', f'Seasonality is mentioned in the narrative (months: {sorted(mention_months)}), but no matching signal was found in the account context source.', WHY['C004'], SOURCES['C004'])
    else:
        r['C004'] = ControlResult('OK', 'No seasonality detected. This is expected for non-seasonal accounts.', WHY['C004'], SOURCES['C004'])

    # -------------------------------------------------------------------------
    # C005 — Operational Constraints Awareness
    # Reads AL7 (Yes/No) and cross-checks AY7 + AM7 + BN7 for strong constraint signals.
    #
    # Four scenarios:
    # 1. AL7 = No  + no signals in other fields  → OK  (consistent, no constraints)
    # 2. AL7 = No  + strong signals in other fields → PARTIAL (default No, constraints exist)
    # 3. AL7 = Yes + signals found or detail present → OK (acknowledged and supported)
    # 4. AL7 = Yes + no signals and no detail → PARTIAL (acknowledged but not documented)
    #
    # Contradiction (AL7 = No + explicit constraint language like "product restriction") → FLAG
    # -------------------------------------------------------------------------
    al7 = clean_text(ctx.al).lower().strip() if ctx.al else ''
    al7_yes = al7 == 'yes'
    al7_no  = al7 in ('no', '')

    # Scan narrative fields for strong constraint signals
    narrative = ' '.join([
        clean_text(ctx.ay).lower(),
        clean_text(ctx.am).lower(),
        clean_text(ctx.bn).lower(),
    ])
    signals_found = [sig for sig in CONSTRAINT_SIGNALS if sig in narrative]
    strong_signals = [s for s in signals_found if any(w in s for w in [
        'restriction', 'restricted', 'cannot advertise', 'not allowed',
        'compliance', 'intellectual property', 'logistics', 'cash flow',
    ])]

    if al7_yes:
        if signals_found:
            detail_preview = signals_found[0]
            r['C005'] = ControlResult(
                'OK',
                f'Operational constraints are acknowledged in AL7. Supporting signals found in account notes: {detail_preview}.',
                WHY['C005'], SOURCES['C005']
            )
        else:
            r['C005'] = ControlResult(
                'PARTIAL',
                'AL7 marks constraints as present, but no constraint detail was found in the objective, context, or challenges fields. Document what the constraints are.',
                WHY['C005'], SOURCES['C005']
            )
    elif al7_no:
        if strong_signals:
            r['C005'] = ControlResult(
                'FLAG',
                f'AL7 says no constraints, but the account notes mention: {strong_signals[0]}. Update AL7 and document the constraint properly.',
                WHY['C005'], SOURCES['C005']
            )
        elif signals_found:
            r['C005'] = ControlResult(
                'PARTIAL',
                f'AL7 says no constraints, but the account notes mention: {signals_found[0]}. Check if this is an operational constraint and update AL7 if needed.',
                WHY['C005'], SOURCES['C005']
            )
        else:
            r['C005'] = ControlResult(
                'OK',
                'No operational constraints are documented or detected in the account notes. AL7 is consistent.',
                WHY['C005'], SOURCES['C005']
            )
    else:
        r['C005'] = ControlResult(
            'PARTIAL',
            'AL7 (operational constraints field) has not been filled in. Mark it as Yes or No.',
            WHY['C005'], SOURCES['C005']
        )

    # -------------------------------------------------------------------------
    # C006 — Client Journey Map
    # Binary gate: tab 39 H7 (journey_h7) OR cjm_id from tab 55 must be present.
    # If CJM is found, evaluate 4 sub-check groups from tab 55 stage data:
    #   Sub-1 (stage count):    ≥3 stages defined
    #   Sub-2 (status dist.):   exactly 1 In Progress, exactly 1 Next, ≥1 Planned
    #   Sub-3 (strategy fill):  Strategy field populated for every active stage
    #   Sub-4 (date logic):     intro date < exec date within each stage;
    #                           no Next/Planned stages with exec date in the past
    # 0 sub-checks fail → OK with stage summary
    # 1–2 fail          → PARTIAL listing issues
    # 3–4 fail          → FLAG listing all issues
    # No CJM linked at all → FLAG (same as before)
    # -------------------------------------------------------------------------
    has_cjm = bool(ctx.journey_h7) or bool(ctx.cjm_id)

    if not has_cjm:
        r['C006'] = ControlResult('FLAG', 'No Client Journey Map was found for this account. It needs to be created and linked.', WHY['C006'], SOURCES['C006'])
    else:
        # Check whether we have stage data to evaluate
        cjm_statuses  = ctx.cjm_status or [None, None, None, None]
        cjm_strategies = ctx.cjm_strategy or [None, None, None, None]
        cjm_intros    = ctx.cjm_intro_date or [None, None, None, None]
        cjm_execs     = ctx.cjm_exec_date or [None, None, None, None]
        cjm_completions = ctx.cjm_actual_completion or [None, None, None, None]

        # Active stages = those with a Status value
        active_idx = [i for i, s in enumerate(cjm_statuses) if s]
        stage_count = len(active_idx)

        # If tab 55 has no stage data, fall back to the old binary check
        if stage_count == 0:
            r['C006'] = ControlResult('OK', 'A Client Journey Map is linked to this account. Stage detail was not available for deeper evaluation.', WHY['C006'], SOURCES['C006'])
        else:
            issues_flag = []
            issues_partial = []

            # Sub-1: stage count ≥ 3
            if stage_count < 3:
                issues_flag.append(f'Only {stage_count} stage(s) defined — minimum 3 required.')

            # Sub-2: status distribution
            in_prog  = sum(1 for s in cjm_statuses if s == 'In Progress')
            nxt      = sum(1 for s in cjm_statuses if s == 'Next')
            planned  = sum(1 for s in cjm_statuses if s == 'Planned')
            status_issues = []
            if in_prog != 1:
                status_issues.append(f'{in_prog} In Progress stage(s) — exactly 1 expected')
            if nxt != 1:
                status_issues.append(f'{nxt} Next stage(s) — exactly 1 expected')
            if planned < 1:
                status_issues.append('no Planned stage — at least 1 required')
            if status_issues:
                issues_flag.append('Status distribution: ' + '; '.join(status_issues) + '.')

            # Sub-3: strategy fields populated for every active stage
            missing_strategy = [i + 1 for i in active_idx if not cjm_strategies[i]]
            if missing_strategy:
                issues_partial.append(f'Strategy field is blank for stage(s): {missing_strategy}.')

            # Sub-4: date logic
            from datetime import date as _date
            date_issues = []
            today = _date.today()
            for i in active_idx:
                intro = cjm_intros[i]
                exec_ = cjm_execs[i]
                status = cjm_statuses[i]
                # intro must be before exec date
                if intro is not None and exec_ is not None:
                    intro_d = intro.date() if hasattr(intro, 'date') else intro
                    exec_d  = exec_.date() if hasattr(exec_, 'date') else exec_
                    try:
                        if intro_d >= exec_d:
                            date_issues.append(f'Stage {i+1}: Introduction date ({intro_d}) is not before Completion date ({exec_d}).')
                    except Exception:
                        pass
                # Next/Planned stages should not have a completion date in the past
                if status in ('Next', 'Planned') and exec_ is not None:
                    exec_d = exec_.date() if hasattr(exec_, 'date') else exec_
                    try:
                        if exec_d < today:
                            date_issues.append(f'Stage {i+1} ({status}) has a past Completion date ({exec_d}) — update the timeline.')
                    except Exception:
                        pass
            if date_issues:
                issues_partial.extend(date_issues)

            # Summarise for what_we_saw
            status_summary = f'{stage_count} stage(s): {in_prog}× In Progress, {nxt}× Next, {planned}× Planned.'
            all_issues = issues_flag + issues_partial
            fail_count = len(issues_flag) + len(issues_partial)

            if fail_count == 0:
                r['C006'] = ControlResult('OK', f'Client Journey Map is complete. {status_summary} Strategy fields and date logic are consistent.', WHY['C006'], SOURCES['C006'])
            elif len(issues_flag) >= 3 or fail_count >= 3:
                r['C006'] = ControlResult('FLAG', f'Client Journey Map has multiple issues. {status_summary} Issues: ' + ' | '.join(all_issues), WHY['C006'], SOURCES['C006'])
            else:
                r['C006'] = ControlResult('PARTIAL', f'Client Journey Map is linked but has gaps. {status_summary} Issues: ' + ' | '.join(all_issues), WHY['C006'], SOURCES['C006'])

    # -------------------------------------------------------------------------
    # C007 — Narrative Consistency
    # Reads: ACoS constraint (O7), TACoS constraint (AX7), ACoS target (J7), TACoS target (K7).
    # Enhancement from tab 55:
    #   - sf_primary_spend_kpi gates which constraint pair is required:
    #       ACOS/ROAS → ACoS constraint required; TACoS missing is not penalised.
    #       TACOS     → TACoS constraint required; ACoS missing is not penalised.
    #       blank     → require both (original behaviour).
    #   - sf_acos_constraint / sf_tacos_constraint cross-checked against O7 / AX7.
    #     A material mismatch (>2pp) between tab 55 and tab 38 sources is flagged.
    # TACoS must be strictly lower than ACoS when both are present.
    # All issues are listed in the what message. Worst-case status wins.
    # -------------------------------------------------------------------------
    acos_c     = norm_pct(ctx.o7)
    tacos_c    = norm_pct(ctx.ax7)
    proj_acos  = norm_pct(ctx.proj_j)
    proj_tacos = norm_pct(ctx.proj_k)
    sf_kpi     = ctx.sf_primary_spend_kpi  # 'ACOS', 'ROAS', 'TACOS', or ''

    # Determine which constraints are actually required given the primary KPI
    acos_required  = sf_kpi in ('', 'ACOS', 'ROAS') or sf_kpi == ''
    tacos_required = sf_kpi in ('', 'TACOS') or sf_kpi == ''
    if sf_kpi == 'ACOS' or sf_kpi == 'ROAS':
        tacos_required = False
    elif sf_kpi == 'TACOS':
        acos_required = False
    # sf_kpi blank → require both (original behaviour)
    if not sf_kpi:
        acos_required = True
        tacos_required = True

    issues_flag    = []
    issues_partial = []

    # — Missing field checks (gated on KPI relevance) —
    field_labels = []
    if acos_required:
        field_labels.append((acos_c, 'ACoS constraint (O7)'))
    if tacos_required:
        field_labels.append((tacos_c, 'TACoS constraint (AX7)'))
    field_labels.append((proj_acos,  'ACoS target (J7)'))
    field_labels.append((proj_tacos, 'TACoS target (K7)'))

    missing_fields = [label for value, label in field_labels if value is None]
    if len(missing_fields) >= 2:
        issues_flag.append(f'Missing fields: {", ".join(missing_fields)}.')
    elif len(missing_fields) == 1:
        issues_partial.append(f'Missing field: {missing_fields[0]}.')

    # — Target vs constraint checks —
    if proj_acos is not None and acos_c is not None:
        if proj_acos > acos_c + 1e-9:
            issues_flag.append(
                f'ACoS target ({pct_str(proj_acos)}) is higher than the agreed constraint ({pct_str(acos_c)}).'
            )
    if proj_tacos is not None and tacos_c is not None:
        if proj_tacos > tacos_c + 1e-9:
            issues_flag.append(
                f'TACoS target ({pct_str(proj_tacos)}) is higher than the agreed constraint ({pct_str(tacos_c)}).'
            )

    # — TACoS vs ACoS ordering (skip if either is missing) —
    if proj_tacos is not None and proj_acos is not None:
        if proj_tacos >= proj_acos - 1e-9:
            issues_flag.append(
                f'TACoS target ({pct_str(proj_tacos)}) is not lower than ACoS target ({pct_str(proj_acos)}). TACoS must always be below ACoS.'
            )
    if tacos_c is not None and acos_c is not None:
        if tacos_c >= acos_c - 1e-9:
            issues_flag.append(
                f'TACoS constraint ({pct_str(tacos_c)}) is not lower than ACoS constraint ({pct_str(acos_c)}). TACoS must always be below ACoS.'
            )

    # — Cross-source mismatch: tab 55 vs tab 38 (>2pp = noteworthy) —
    sf_acos_c  = norm_pct(ctx.sf_acos_constraint)
    sf_tacos_c = norm_pct(ctx.sf_tacos_constraint)
    if sf_acos_c is not None and acos_c is not None:
        if abs(sf_acos_c - acos_c) > 0.02:
            issues_partial.append(
                f'ACoS constraint mismatch: Salesforce CSP says {pct_str(sf_acos_c)} but the project record (O7) shows {pct_str(acos_c)}. Reconcile the two sources.'
            )
    if sf_tacos_c is not None and tacos_c is not None:
        if abs(sf_tacos_c - tacos_c) > 0.02:
            issues_partial.append(
                f'TACoS constraint mismatch: Salesforce CSP says {pct_str(sf_tacos_c)} but the project record (AX7) shows {pct_str(tacos_c)}. Reconcile the two sources.'
            )

    # — Resolve status and build message —
    kpi_note = f' (KPI: {sf_kpi} — only {("ACoS" if not tacos_required else "TACoS" if not acos_required else "both")} constraint(s) required)' if sf_kpi else ''
    all_issues = issues_flag + issues_partial
    if not all_issues:
        what = (
            f'All documented fields are consistent.{kpi_note} '
            f'ACoS: target {pct_str(proj_acos)} within constraint {pct_str(acos_c)}. '
            f'TACoS: target {pct_str(proj_tacos)} within constraint {pct_str(tacos_c)}.'
        )
        r['C007'] = ControlResult('OK', what, WHY['C007'], SOURCES['C007'])
    elif issues_flag:
        what = ' | '.join(all_issues) + kpi_note
        r['C007'] = ControlResult('FLAG', what, WHY['C007'], SOURCES['C007'])
    else:
        what = ' | '.join(all_issues) + kpi_note
        r['C007'] = ControlResult('PARTIAL', what, WHY['C007'], SOURCES['C007'])

    # -------------------------------------------------------------------------
    # C008 — Sales Concentration Matches Account Story
    # Primary source: AU7 narrative (ctx.au) — free-text field.
    # Cross-reference: sf_sales_concentration from tab 55 (structured CSP field).
    #   SF values: 'Low Concentration' | 'Medium Concentration' | 'High Concentration'
    #   → normalised to 'low' | 'medium' | 'high' for comparison.
    # Resolution logic:
    #   - AU7 populated → classify and compare to actual data.
    #   - AU7 empty but SF field populated → use SF value directly (no FLAG for missing AU7).
    #   - Both populated but diverge from each other → note the source mismatch.
    #   - Both empty → FLAG (not documented anywhere).
    # -------------------------------------------------------------------------
    if ctx.top1 is None:
        r['C008'] = ControlResult('FLAG', 'Sales concentration could not be checked because parent-ASIN sales data was not available.', WHY['C008'], SOURCES['C008'])
    else:
        actual_class = classify_concentration(ctx.top1, ctx.top3, ctx.top5)
        conc_detail = f'Top 1 ASIN: {pct_str(ctx.top1)}, top 3: {pct_str(ctx.top3)}, top 5: {pct_str(ctx.top5)}.'

        # Classify AU7 narrative
        narr = ctx.au.lower()
        narr_class = (
            'high' if 'high' in narr
            else 'medium' if ('medium' in narr or 'moderate' in narr)
            else 'low' if ('low' in narr or 'diversified' in narr)
            else None
        )

        # Classify SF structured field (Sales_Concentration__c)
        sf_raw = ctx.sf_sales_concentration.lower() if ctx.sf_sales_concentration else ''
        sf_class = (
            'high' if 'high' in sf_raw
            else 'medium' if 'medium' in sf_raw
            else 'low' if 'low' in sf_raw
            else None
        )

        # Pick the best documented class (AU7 first, SF fallback)
        doc_class = narr_class if narr_class is not None else sf_class
        doc_source = 'AU7' if narr_class is not None else 'CSP Salesforce field'

        # Cross-source consistency note
        source_conflict = (
            narr_class is not None and sf_class is not None and narr_class != sf_class
        )

        if doc_class is None:
            r['C008'] = ControlResult('FLAG', f'Sales concentration is not documented in AU7 or the CSP record. Actual concentration is {actual_class}. {conc_detail}', WHY['C008'], SOURCES['C008'])
        elif doc_class == actual_class:
            if source_conflict:
                r['C008'] = ControlResult('PARTIAL', f'Sales concentration ({doc_source}) is documented as {doc_class} and matches actual data. However, AU7 says "{narr_class}" and the CSP says "{sf_class}" — reconcile the two sources. {conc_detail}', WHY['C008'], SOURCES['C008'])
            else:
                r['C008'] = ControlResult('OK', f'Sales concentration is documented as {doc_class} ({doc_source}) and matches the actual data. {conc_detail}', WHY['C008'], SOURCES['C008'])
        else:
            r['C008'] = ControlResult('FLAG', f'Sales concentration documented as "{doc_class}" ({doc_source}) but actual data shows "{actual_class}". {conc_detail} Update the documentation.', WHY['C008'], SOURCES['C008'])

    # -------------------------------------------------------------------------
    # C009 — Client Contact Cadence (last 6 months)
    # -------------------------------------------------------------------------
    if ctx.gap is None:
        if ctx.last_call is not None:
            r['C009'] = ControlResult('PARTIAL', f'Only one Gong meeting was found ({ctx.last_call.date()}). Two meetings are needed to measure the contact cadence.', WHY['C009'], SOURCES['C009'])
        else:
            r['C009'] = ControlResult('FLAG', 'No Gong meetings were found for this account. Client contact cadence cannot be confirmed.', WHY['C009'], SOURCES['C009'])
    else:
        if ctx.gap <= 30:
            r['C009'] = ControlResult('OK', f'Last two meetings were {ctx.gap} days apart ({ctx.prev_call.date()} → {ctx.last_call.date()}). Cadence is within the 30-day target.', WHY['C009'], SOURCES['C009'])
        elif ctx.gap <= 60:
            r['C009'] = ControlResult('PARTIAL', f'Last two meetings were {ctx.gap} days apart ({ctx.prev_call.date()} → {ctx.last_call.date()}). This is above the 30-day target.', WHY['C009'], SOURCES['C009'])
        else:
            r['C009'] = ControlResult('FLAG', f'Last two meetings were {ctx.gap} days apart ({ctx.prev_call.date()} → {ctx.last_call.date()}). This is a long gap — the account story may be out of date.', WHY['C009'], SOURCES['C009'])

    # -------------------------------------------------------------------------
    # C010 — Customizations Documented & Justified
    # -------------------------------------------------------------------------
    active_types = detect_personalizations(ctx)
    documented_count, matched = documented_personalizations(ctx.proj_cs_notes, active_types)
    active_count = len(active_types)
    if active_count == 0:
        r['C010'] = ControlResult('OK', 'No active framework customizations were detected. Nothing to document.', WHY['C010'], SOURCES['C010'])
    else:
        ratio = documented_count / active_count if active_count else 0
        labels = ', '.join(active_types)
        if documented_count >= active_count:
            r['C010'] = ControlResult('OK', f'{active_count} active customization(s) detected ({labels}) and all are documented in CS Notes.', WHY['C010'], SOURCES['C010'])
        elif ratio >= 0.5:
            r['C010'] = ControlResult('PARTIAL', f'{active_count} active customization(s) detected ({labels}), but only {documented_count} of them are documented in CS Notes.', WHY['C010'], SOURCES['C010'])
        else:
            r['C010'] = ControlResult('FLAG', f'{active_count} active customization(s) detected ({labels}), but most are not documented in CS Notes. The CoE cannot tell if these are intentional.', WHY['C010'], SOURCES['C010'])

    # -------------------------------------------------------------------------
    # C011 — Target Spend / KPI Targets Documented
    # Sub-check 1 (spend pacing): daily spend target from proj_h (tab 54),
    #   with sf_daily_target_spend (tab 55) as fallback if proj_h is blank.
    # Sub-check 2 (ROAS target): sf_target_roas (tab 55) vs actual ROAS from tab 02.
    #   Tiers: ≤20% deviation OK, ≤40% PARTIAL, else FLAG.
    # -------------------------------------------------------------------------
    checks = []
    msgs   = []

    # Sub-check 1: spend pacing
    daily_target = to_float(ctx.proj_h)
    target_source = 'project record'
    if daily_target is None and ctx.sf_daily_target_spend is not None:
        daily_target = to_float(ctx.sf_daily_target_spend)
        target_source = 'CSP Salesforce field'

    if daily_target is not None and ctx.window_days and ctx.metrics.get('AdSpend') is not None:
        actual_daily = float(ctx.metrics['AdSpend']) / ctx.window_days
        gap = abs(actual_daily - daily_target) / daily_target if daily_target else None
        deviation_pct = f'{gap * 100:.0f}%' if gap is not None else 'unknown'
        direction = 'below' if actual_daily < daily_target else 'above'
        checks.append('OK' if gap is not None and gap <= 0.20 else 'PARTIAL' if gap is not None and gap <= 0.40 else 'FLAG')
        msgs.append(f'Spend target ${daily_target:.0f}/day ({target_source}) vs actual ${actual_daily:.0f}/day ({deviation_pct} {direction} target)')

    # Sub-check 2: ROAS target vs actual
    sf_target_roas = to_float(ctx.sf_target_roas)
    actual_roas = ctx.metrics.get('ROAS') if ctx.metrics else None
    if sf_target_roas is not None and actual_roas is not None and sf_target_roas > 0:
        roas_gap = abs(actual_roas - sf_target_roas) / sf_target_roas
        direction_r = 'below' if actual_roas < sf_target_roas else 'above'
        roas_status = 'OK' if roas_gap <= 0.20 else 'PARTIAL' if roas_gap <= 0.40 else 'FLAG'
        checks.append(roas_status)
        msgs.append(f'ROAS target {sf_target_roas:.2f}x vs actual {actual_roas:.2f}x ({roas_gap * 100:.0f}% {direction_r} target)')

    if not checks:
        r['C011'] = ControlResult('OK', 'No spend or ROAS target is documented. Spend pacing and ROAS alignment were not evaluated.', WHY['C011'], SOURCES['C011'])
    elif all(x == 'OK' for x in checks):
        r['C011'] = ControlResult('OK', f'{" | ".join(msgs)} — within acceptable range.', WHY['C011'], SOURCES['C011'])
    elif 'FLAG' in checks:
        r['C011'] = ControlResult('FLAG', f'{" | ".join(msgs)} — significant deviation from the documented target.', WHY['C011'], SOURCES['C011'])
    else:
        r['C011'] = ControlResult('PARTIAL', f'{" | ".join(msgs)} — moderate deviation from the documented target.', WHY['C011'], SOURCES['C011'])

    # -------------------------------------------------------------------------
    # C012 — Tagging / Segmentation Logic Clear
    # Requires both a bestseller label AND a category/product-type label for OK.
    # PARTIAL: one dimension present, one missing.
    # FLAG: neither present.
    #
    # has_category previously used a negative match (anything not in the other
    # word sets). That caused false positives on filler tags like "test", "Q2",
    # "new", etc. Replaced with a positive vocabulary (CATEGORY_WORDS) that
    # covers known intentional grouping patterns — product types, tiers, bundles.
    # Surface the matched tag values in the finding for transparency.
    # -------------------------------------------------------------------------
    tags = [t.lower() for t in ctx.tags if t]

    matched_best = [t for t in tags if any(w in t for w in BESTSELLER_WORDS)]
    matched_cat  = [t for t in tags if any(w in t for w in CATEGORY_WORDS)
                    or any(re.search(rf'\b{re.escape(g)}\b', t) for g in CATEGORY_GEO_CODES)]
    matched_seg  = [t for t in tags if any(w in t for w in SEGMENTATION_WORDS)]

    has_best      = bool(matched_best)
    has_cat_or_seg = bool(matched_cat) or bool(matched_seg)

    # Build readable tag previews (up to 3 examples each)
    best_preview = ', '.join(f'"{x}"' for x in matched_best[:3])
    cat_preview  = ', '.join(f'"{x}"' for x in (matched_cat + matched_seg)[:3])

    if has_best and has_cat_or_seg:
        r['C012'] = ControlResult(
            'OK',
            f'Campaign tags show clear product segmentation. Bestseller label(s): {best_preview}. Category/tier label(s): {cat_preview}.',
            WHY['C012'], SOURCES['C012']
        )
    elif has_best and not has_cat_or_seg:
        r['C012'] = ControlResult(
            'PARTIAL',
            f'Bestseller label found ({best_preview}), but no category or performance tier label was detected. Add a product-type or tier tag to complete the segmentation.',
            WHY['C012'], SOURCES['C012']
        )
    elif has_cat_or_seg and not has_best:
        r['C012'] = ControlResult(
            'PARTIAL',
            f'Category/tier label found ({cat_preview}), but no bestseller label was detected. Add a hero/winner/core tag to complete the segmentation.',
            WHY['C012'], SOURCES['C012']
        )
    else:
        total_tags = len(set(tags))
        tag_note = f' ({total_tags} tag value(s) found but none matched known segmentation patterns).' if total_tags else ' No tag values were found.'
        r['C012'] = ControlResult(
            'FLAG',
            f'Neither a bestseller label nor a category or performance tier label was found in the campaign tags.{tag_note} The team cannot tell how the portfolio is being prioritized.',
            WHY['C012'], SOURCES['C012']
        )

    # -------------------------------------------------------------------------
    # C013 / C014 — Manual on-call controls
    # -------------------------------------------------------------------------
    r['C013'] = ControlResult('OK', 'To be reviewed during the QR presentation call.', WHY['C013'], SOURCES['C013'])
    r['C014'] = ControlResult('OK', 'To be reviewed during the QR presentation call.', WHY['C014'], SOURCES['C014'])

    return r


def build_summary(ctx: DatabricksContext, results: Dict[str, ControlResult]) -> dict:
    return {
        'primary_objective': build_primary_objective(ctx, results),
        'customization_context': ctx.proj_cs_notes if ctx.proj_cs_notes else 'No notes documented.',
        'acos_objective': norm_pct(ctx.proj_j),
        'tacos_objective': norm_pct(ctx.proj_k),
        'acos_constraint': norm_pct(ctx.o7),
        'tacos_constraint': norm_pct(ctx.ax7),
        'budget_constraint': _extract_budget_constraint(ctx),
        'primary_kpi': ctx.bw if ctx.bw else 'Not documented',
    }


def _extract_budget_constraint(ctx: DatabricksContext):
    import warnings
    text = ' '.join([ctx.ay, ctx.am, ctx.bn])
    m = re.search(r'([0-9]{1,3}(?:,[0-9]{3})+|[0-9]+(?:\.[0-9]+)?k)\s*(?:monthly|/month|per month)', text, re.I)
    if m:
        return to_float(m.group(1))
    warnings.warn(
        f"build_summary: budget_constraint could not be extracted from narrative fields for {ctx.hash_name}. "
        "Budget will show as 'Not documented' in the output.",
        stacklevel=2,
    )
    return None


def score_grade(score: float) -> str:
    if score >= 75:
        return 'Compliant'
    if score >= 40:
        return 'Needs Attention'
    return 'Not Compliant'


def interpretation(grade: str) -> str:
    return {
        'Compliant': 'Account mastery signals are largely documented and internally consistent based on the currently available sources.',
        'Needs Attention': 'Some mastery elements are present, but important documentation or consistency gaps still need follow-up.',
        'Not Compliant': 'Key mastery signals are missing or inconsistent, which limits confidence in account ownership and account-story accuracy.',
    }[grade]


def compute_score(results: Dict[str, ControlResult]):
    findings = []
    total_penalty = 0.0
    for cid, res in results.items():
        imp = IMPORTANCE[cid]
        pen = 0.0
        if cid not in SCORING_EXCLUDED:
            if res.status == 'FLAG':
                pen = PRIORITY_POINTS[imp]
            elif res.status == 'PARTIAL':
                pen = PRIORITY_POINTS[imp] * 0.5
        total_penalty += pen
        # C013 and C014 are manual controls — exclude from findings list entirely
        if cid in SCORING_EXCLUDED:
            continue
        findings.append({'cid': cid, 'name': CONTROL_NAMES[cid], 'status': res.status, 'what': res.what, 'why': res.why, 'importance': imp, 'impact': IMPACT_LABEL[imp], 'penalty': pen})
    score = 100 + total_penalty
    grade = score_grade(score)
    findings.sort(key=lambda x: (0 if x['status'] == 'FLAG' else 1, x['penalty']))
    return total_penalty, score, grade, findings
