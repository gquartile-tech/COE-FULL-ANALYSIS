from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple

import pandas as pd

from config_mastery import CONTROL_NAMES, IMPACT_LABEL, IMPORTANCE, PRIORITY_POINTS, SCORING_EXCLUDED, SOURCES, WHY, ControlResult
from reader_databricks_mastery import DatabricksContext, clean_text, money_str, monthly_budget_from_daily, norm_pct, pct_str, to_float, trim

OBJECTIVE_WORDS = {'objective', 'goal', 'grow', 'growth', 'scale', 'increase', 'improve', 'stabilize', 'maintain', 'reduce', 'defend', 'accelerate', 'awareness', 'sales', 'profit', 'profitability', 'ranking', 'market share'}
KPI_WORDS = {'roas', 'acos', 'tacos', 'spend', 'sales', 'cvr', 'ctr', 'cpc', 'ntb', 'rank', 'revenue'}
CONSTRAINT_WORDS = {'constraint', 'below', 'above', 'maintain', 'limit', 'threshold', 'guardrail', 'while', 'without', 'at or below'}
CHALLENGE_WORDS = {'challenge', 'issue', 'risk', 'inventory', 'out-of-stock', 'out of stock', 'slowdown', 'pressure', 'volatility', 'sensitive', 'buy box', 'listing', 'margin', 'competition', 'competitive', 'growth is not', 'not meeting', 'incomplete', 'dissatisfied'}
TIME_WORDS = {'q1', 'q2', 'q3', 'q4', 'month', 'monthly', 'weekly', 'this period', 'near-term', 'near term', 'next', 'current period', 'prime day', 'holiday', 'bfcm', 'seasonal'}
CONFLICT_WORDS = {'but', 'however', 'while', 'tradeoff', 'trade-off', 'contrasting', 'despite', 'volatility', 'elevated', 'balancing'}
BESTSELLER_WORDS = {'bestseller', 'best seller', 'hero', 'top perf', 'top perf.', 'top', 'winner', 'core', 'priority', 'best-seller'}
SEGMENTATION_WORDS = {'mid seller', 'mid-seller', 'slow mover', 'slow-mover', 'low perf', 'low perf.', 'mid. perf.', 'high traffic', 'low traffic', 'high conversion', 'low conversion'}
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
        return f"Primary objective is documented as {trim(am or ay, 180)}, but strategic context is incomplete."
    if results['C002'].status == 'PARTIAL':
        return f"The primary objective is to {trim(am or ay, 160)}, but the supporting KPI, timeframe, or constraint context is incomplete."
    if ay and am:
        return f"The primary objective is to {trim(am, 140)}, with supporting context that {trim(ay, 220)}"
    return f"The primary objective is to {trim(am or ay, 220)}"


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
    # Reads AM7 — the objective context field.
    # Timeframe is a hard gate — without it the result cannot be OK.
    # Requires all 5 dimensions for OK; at least 3 for PARTIAL.
    # -------------------------------------------------------------------------
    txt = ctx.am
    if not txt:
        r['C001'] = ControlResult('FLAG', 'The objective context field (AM7) is empty. There is no supporting detail for the primary objective.', WHY['C001'], SOURCES['C001'])
    else:
        dims = {
            'objective':  has_any(txt, OBJECTIVE_WORDS),
            'kpi':        has_any(txt, KPI_WORDS),
            'constraint': has_any(txt, CONSTRAINT_WORDS),
            'context':    len(txt.split()) >= 15,
            'timeframe':  has_any(txt, TIME_WORDS),
        }
        n = sum(dims.values())
        missing = [k for k, v in dims.items() if not v]
        has_timeframe = dims['timeframe']
        if n == 5:
            r['C001'] = ControlResult('OK', 'Objective context covers all required elements: goal, KPI, constraint, timeframe, and narrative depth.', WHY['C001'], SOURCES['C001'])
        elif n >= 3 and has_timeframe:
            r['C001'] = ControlResult('PARTIAL', f'Objective context is written but some elements are missing: {", ".join(missing)}.', WHY['C001'], SOURCES['C001'])
        elif n >= 3 and not has_timeframe:
            r['C001'] = ControlResult('PARTIAL', f'Objective context is written but has no timeframe or near-term reference. Also missing: {", ".join([m for m in missing if m != "timeframe"])}.', WHY['C001'], SOURCES['C001'])
        else:
            r['C001'] = ControlResult('FLAG', f'Objective context does not have enough detail to review near-term alignment. Missing elements: {", ".join(missing)}.', WHY['C001'], SOURCES['C001'])

    # -------------------------------------------------------------------------
    # C002 — Objective vs Near-Term Alignment
    # Reads AY7 — the primary objective field.
    # -------------------------------------------------------------------------
    txt = ctx.ay
    if not txt:
        r['C002'] = ControlResult('FLAG', 'No primary objective is written in the account notes (AY7).', WHY['C002'], SOURCES['C002'])
    else:
        score = sum([
            has_any(txt, OBJECTIVE_WORDS),
            has_any(txt, KPI_WORDS | {'awareness', 'ranking', 'market share'}),
            any(w in txt.lower() for w in ['gain', 'grow', 'increase', 'improve', 'stabilize', 'maintain', 'scale', 'accelerate']),
        ])
        if score >= 3:
            r['C002'] = ControlResult('OK', 'Primary objective is documented and linked to a clear business outcome.', WHY['C002'], SOURCES['C002'])
        elif score == 2:
            r['C002'] = ControlResult('PARTIAL', 'Objective is written, but it is not clearly linked to a measurable KPI or a specific business result.', WHY['C002'], SOURCES['C002'])
        else:
            r['C002'] = ControlResult('FLAG', 'Objective is written, but the language is too vague to use as a clear direction for this account.', WHY['C002'], SOURCES['C002'])

    # -------------------------------------------------------------------------
    # C003 — Account Challenges Documented
    # -------------------------------------------------------------------------
    txt = ctx.bn
    if not txt:
        r['C003'] = ControlResult('FLAG', 'No current challenges are documented in the account notes (BN7).', WHY['C003'], SOURCES['C003'])
    else:
        specific = len(txt.split()) >= 12 and has_any(txt, CHALLENGE_WORDS)
        if specific:
            r['C003'] = ControlResult('OK', 'Current challenges are documented with enough detail to understand the active account blockers.', WHY['C003'], SOURCES['C003'])
        elif len(txt.split()) >= 6:
            r['C003'] = ControlResult('PARTIAL', 'Challenges are written, but the description is too general. It does not clearly explain what is blocking the account today.', WHY['C003'], SOURCES['C003'])
        else:
            r['C003'] = ControlResult('FLAG', 'The challenges field has very little content. More detail is needed for a proper review.', WHY['C003'], SOURCES['C003'])

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
    # Hardcoded to OK — field data is not reliable enough to evaluate.
    # -------------------------------------------------------------------------
    r['C005'] = ControlResult('OK', 'Operational constraints check is not evaluated automatically.', WHY['C005'], SOURCES['C005'])

    # -------------------------------------------------------------------------
    # C006 — Client Journey Map
    # -------------------------------------------------------------------------
    if ctx.journey_h7:
        r['C006'] = ControlResult('OK', 'A Client Journey Map is linked to this account.', WHY['C006'], SOURCES['C006'])
    else:
        r['C006'] = ControlResult('FLAG', 'No Client Journey Map was found for this account. It needs to be created and linked.', WHY['C006'], SOURCES['C006'])

    # -------------------------------------------------------------------------
    # C007 — Narrative Consistency
    # Validates 4 fields: ACoS constraint (O7), TACoS constraint (AX7),
    # ACoS target (J7), TACoS target (K7).
    # Missing fields: 1 missing → PARTIAL; 2+ missing → FLAG.
    # Target vs constraint: any target > its constraint → FLAG regardless of how many fail.
    # TACoS vs ACoS: TACoS must be strictly lower than ACoS for both the target pair
    #   and the constraint pair. TACoS >= ACoS → FLAG.
    #   Skipped silently if either field in the pair is missing (already captured above).
    # All issues are listed in the what message. Worst-case status wins.
    # -------------------------------------------------------------------------
    acos_c     = norm_pct(ctx.o7)
    tacos_c    = norm_pct(ctx.ax7)
    proj_acos  = norm_pct(ctx.proj_j)
    proj_tacos = norm_pct(ctx.proj_k)

    issues_flag    = []
    issues_partial = []

    # — Missing field checks —
    field_labels = [
        (acos_c,     'ACoS constraint (O7)'),
        (tacos_c,    'TACoS constraint (AX7)'),
        (proj_acos,  'ACoS target (J7)'),
        (proj_tacos, 'TACoS target (K7)'),
    ]
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

    # — TACoS vs ACoS checks (both pairs, skip silently if either field missing) —
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

    # — Resolve status and build message —
    all_issues = issues_flag + issues_partial
    if not all_issues:
        what = (
            f'All four fields are documented and consistent. '
            f'ACoS: target {pct_str(proj_acos)} within constraint {pct_str(acos_c)}. '
            f'TACoS: target {pct_str(proj_tacos)} within constraint {pct_str(tacos_c)}. '
            f'TACoS is correctly below ACoS across both pairs.'
        )
        r['C007'] = ControlResult('OK', what, WHY['C007'], SOURCES['C007'])
    elif issues_flag:
        what = ' | '.join(all_issues)
        r['C007'] = ControlResult('FLAG', what, WHY['C007'], SOURCES['C007'])
    else:
        what = ' | '.join(all_issues)
        r['C007'] = ControlResult('PARTIAL', what, WHY['C007'], SOURCES['C007'])

    # -------------------------------------------------------------------------
    # C008 — Sales Concentration Matches Account Story
    # -------------------------------------------------------------------------
    if ctx.top1 is None:
        r['C008'] = ControlResult('FLAG', 'Sales concentration could not be checked because parent-ASIN sales data was not available.', WHY['C008'], SOURCES['C008'])
    else:
        actual_class = classify_concentration(ctx.top1, ctx.top3, ctx.top5)
        narr = ctx.au.lower()
        narr_class = 'high' if 'high' in narr else 'medium' if ('medium' in narr or 'moderate' in narr) else 'low' if ('low' in narr or 'diversified' in narr) else None
        conc_detail = f'Top 1 ASIN: {pct_str(ctx.top1)}, top 3: {pct_str(ctx.top3)}, top 5: {pct_str(ctx.top5)}.'
        if narr_class == actual_class:
            r['C008'] = ControlResult('OK', f'Sales concentration is documented as {actual_class} and matches the actual data. {conc_detail}', WHY['C008'], SOURCES['C008'])
        elif narr_class is None:
            r['C008'] = ControlResult('FLAG', f'Sales concentration is not documented in AU7. Actual concentration is {actual_class}. {conc_detail}', WHY['C008'], SOURCES['C008'])
        else:
            r['C008'] = ControlResult('FLAG', f'Sales concentration in the notes says "{narr_class}" but the actual data shows "{actual_class}". {conc_detail} The notes need to be updated.', WHY['C008'], SOURCES['C008'])

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
    # Only spend deviation is evaluated (KPI targets excluded per design decision).
    # -------------------------------------------------------------------------
    checks = []
    msgs = []
    daily_target = to_float(ctx.proj_h)
    if daily_target is not None and ctx.window_days and ctx.metrics.get('AdSpend') is not None:
        actual_daily = float(ctx.metrics['AdSpend']) / ctx.window_days
        gap = abs(actual_daily - daily_target) / daily_target if daily_target else None
        deviation_pct = f'{gap * 100:.0f}%' if gap is not None else 'unknown'
        direction = 'below' if actual_daily < daily_target else 'above'
        checks.append('OK' if gap is not None and gap <= 0.20 else 'PARTIAL' if gap is not None and gap <= 0.40 else 'FLAG')
        msgs.append(f'Spend target ${daily_target:.0f}/day vs actual ${actual_daily:.0f}/day ({deviation_pct} {direction} target)')
    if not checks:
        r['C011'] = ControlResult('OK', 'No spend target is documented in the project dataset. Spend pacing was not evaluated.', WHY['C011'], SOURCES['C011'])
    elif all(x == 'OK' for x in checks):
        r['C011'] = ControlResult('OK', f'{" | ".join(msgs)} — within acceptable range.', WHY['C011'], SOURCES['C011'])
    elif 'FLAG' in checks:
        r['C011'] = ControlResult('FLAG', f'{" | ".join(msgs)} — significant deviation from the documented target.', WHY['C011'], SOURCES['C011'])
    else:
        r['C011'] = ControlResult('PARTIAL', f'{" | ".join(msgs)} — moderate deviation from the documented target.', WHY['C011'], SOURCES['C011'])

    # -------------------------------------------------------------------------
    # C012 — Tagging / Segmentation Logic Clear
    # Requires both a bestseller label and a category/segment label for OK.
    # PARTIAL: one of the two is missing — message names which one.
    # FLAG: neither is present.
    # -------------------------------------------------------------------------
    tags = [t.lower() for t in ctx.tags if t]
    has_best = any(any(w in t for w in BESTSELLER_WORDS) for t in tags)
    has_category = any(t not in {'', 'none'} and not any(w in t for w in BESTSELLER_WORDS | SEGMENTATION_WORDS) for t in tags)
    has_segment = any(any(w in t for w in SEGMENTATION_WORDS) for t in tags)
    has_cat_or_seg = has_category or has_segment

    if has_best and has_cat_or_seg:
        r['C012'] = ControlResult('OK', 'Campaign tags show clear product segmentation with both a bestseller label and a category or performance tier label identified.', WHY['C012'], SOURCES['C012'])
    elif has_best and not has_cat_or_seg:
        r['C012'] = ControlResult('PARTIAL', 'A bestseller label was found in the campaign tags, but no category or performance tier label was detected. The second tagging dimension is missing.', WHY['C012'], SOURCES['C012'])
    elif has_cat_or_seg and not has_best:
        r['C012'] = ControlResult('PARTIAL', 'A category or performance tier label was found in the campaign tags, but no bestseller label was detected. The bestseller tagging dimension is missing.', WHY['C012'], SOURCES['C012'])
    else:
        r['C012'] = ControlResult('FLAG', 'Neither a bestseller label nor a category or performance tier label was found in the campaign tag fields. Both tagging dimensions are missing — the team cannot tell how the portfolio is being prioritized.', WHY['C012'], SOURCES['C012'])

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
