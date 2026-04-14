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

    txt = ctx.ay
    if not txt:
        r['C001'] = ControlResult('FLAG', 'Objective is not documented in 38_Client_Success_Insights_Repo!AY7.', WHY['C001'], SOURCES['C001'])
    else:
        score = sum([has_any(txt, OBJECTIVE_WORDS), has_any(txt, KPI_WORDS | {'awareness', 'ranking', 'market share'}), any(w in txt.lower() for w in ['gain', 'grow', 'increase', 'improve', 'stabilize', 'maintain', 'scale', 'accelerate'])])
        if score >= 3:
            r['C001'] = ControlResult('OK', '', WHY['C001'], SOURCES['C001'])
        elif score == 2:
            r['C001'] = ControlResult('PARTIAL', 'Objective is documented, but it is not clearly anchored to a measurable KPI or business outcome.', WHY['C001'], SOURCES['C001'])
        else:
            r['C001'] = ControlResult('FLAG', 'Objective is documented, but the definition is too vague to use as a clear account anchor.', WHY['C001'], SOURCES['C001'])

    txt = ctx.am
    if not txt:
        r['C002'] = ControlResult('FLAG', 'Primary objective description is not documented in 38_Client_Success_Insights_Repo!AM7.', WHY['C002'], SOURCES['C002'])
    else:
        dims = {'objective': has_any(txt, OBJECTIVE_WORDS), 'kpi': has_any(txt, KPI_WORDS), 'constraint': has_any(txt, CONSTRAINT_WORDS), 'context': len(txt.split()) >= 15, 'timeframe': has_any(txt, TIME_WORDS)}
        n = sum(dims.values())
        if n >= 4:
            r['C002'] = ControlResult('OK', '', WHY['C002'], SOURCES['C002'])
        elif n >= 2:
            r['C002'] = ControlResult('PARTIAL', 'Primary objective description is documented, but KPI, timeframe, constraint, or context coverage is incomplete.', WHY['C002'], SOURCES['C002'])
        else:
            r['C002'] = ControlResult('FLAG', 'Primary objective description is documented, but it does not contain enough strategic detail to assess the account objective.', WHY['C002'], SOURCES['C002'])

    txt = ctx.bn
    if not txt:
        r['C003'] = ControlResult('FLAG', 'Current challenges are not documented in 38_Client_Success_Insights_Repo!BN7.', WHY['C003'], SOURCES['C003'])
    else:
        specific = len(txt.split()) >= 12 and has_any(txt, CHALLENGE_WORDS)
        if specific:
            r['C003'] = ControlResult('OK', '', WHY['C003'], SOURCES['C003'])
        elif len(txt.split()) >= 6:
            r['C003'] = ControlResult('PARTIAL', 'Current challenges are documented, but the description is generic and does not clearly define the active account blockers.', WHY['C003'], SOURCES['C003'])
        else:
            r['C003'] = ControlResult('FLAG', 'Current challenges are not clearly documented.', WHY['C003'], SOURCES['C003'])

    source_months = parse_months_from_text(ctx.am)
    mention_months = set()
    for text in [ctx.ay, ctx.bn]:
        mention_months |= parse_months_from_text(text)
    if source_months and mention_months:
        r['C004'] = ControlResult('OK', '', WHY['C004'], SOURCES['C004'])
    elif source_months and not mention_months:
        r['C004'] = ControlResult('FLAG', 'Seasonality is detected based on account context, but it is not documented in the approved narrative fields outside the Health seasonality source.', WHY['C004'], SOURCES['C004'])
    elif not source_months and mention_months:
        r['C004'] = ControlResult('PARTIAL', 'Seasonality is referenced in narrative, but no seasonality signal was detected from the Health seasonality source.', WHY['C004'], SOURCES['C004'])
    else:
        r['C004'] = ControlResult('OK', '', WHY['C004'], SOURCES['C004'])

    r['C005'] = ControlResult('OK', '', WHY['C005'], SOURCES['C005'])

    if ctx.journey_h7:
        r['C006'] = ControlResult('OK', '', WHY['C006'], SOURCES['C006'])
    else:
        r['C006'] = ControlResult('FLAG', 'A Client Journey Map is not present in 39_Client_Journey_Insights_Data.', WHY['C006'], SOURCES['C006'])

    acos_c = norm_pct(ctx.o7); tacos_c = norm_pct(ctx.ax7); proj_acos = norm_pct(ctx.proj_j); proj_tacos = norm_pct(ctx.proj_k)
    checks = []
    for target, constraint in [(proj_acos, acos_c), (proj_tacos, tacos_c)]:
        if target is None or constraint is None:
            checks.append(None)
        else:
            checks.append(target <= constraint + 1e-9)
    if checks == [True, True]:
        r['C007'] = ControlResult('OK', '', WHY['C007'], SOURCES['C007'])
    elif any(v is False for v in checks) and any(v is True for v in checks):
        r['C007'] = ControlResult('PARTIAL', f'Constraint alignment is mixed. ACoS constraint {pct_str(acos_c)} vs project target {pct_str(proj_acos)}; TACoS constraint {pct_str(tacos_c)} vs project target {pct_str(proj_tacos)}.', WHY['C007'], SOURCES['C007'])
    elif any(v is False for v in checks):
        r['C007'] = ControlResult('FLAG', f'Project KPI targets are looser than the current account constraints. ACoS constraint {pct_str(acos_c)} vs project target {pct_str(proj_acos)}; TACoS constraint {pct_str(tacos_c)} vs project target {pct_str(proj_tacos)}.', WHY['C007'], SOURCES['C007'])
    else:
        r['C007'] = ControlResult('PARTIAL', f'Constraint alignment could not be fully validated. ACoS constraint {pct_str(acos_c)} vs project target {pct_str(proj_acos)}; TACoS constraint {pct_str(tacos_c)} vs project target {pct_str(proj_tacos)}.', WHY['C007'], SOURCES['C007'])

    if ctx.top1 is None:
        r['C008'] = ControlResult('FLAG', 'Sales concentration could not be evaluated because parent-ASIN sales data was unavailable.', WHY['C008'], SOURCES['C008'])
    else:
        actual_class = classify_concentration(ctx.top1, ctx.top3, ctx.top5)
        narr = ctx.au.lower()
        narr_class = 'high' if 'high' in narr else 'medium' if ('medium' in narr or 'moderate' in narr) else 'low' if ('low' in narr or 'diversified' in narr) else None
        if narr_class == actual_class:
            r['C008'] = ControlResult('OK', '', WHY['C008'], SOURCES['C008'])
        elif narr_class is None:
            r['C008'] = ControlResult('FLAG', f'Sales concentration is not documented in AU7. Actual concentration is {actual_class} (top1 {pct_str(ctx.top1)}, top3 {pct_str(ctx.top3)}, top5 {pct_str(ctx.top5)}).', WHY['C008'], SOURCES['C008'])
        else:
            r['C008'] = ControlResult('FLAG', f'Sales concentration narrative does not match actual concentration. AU7 says {ctx.au or "not documented"}; actual concentration is {actual_class} (top1 {pct_str(ctx.top1)}, top3 {pct_str(ctx.top3)}, top5 {pct_str(ctx.top5)}).', WHY['C008'], SOURCES['C008'])

    if ctx.gap is None:
        if ctx.last_call is not None:
            r['C009'] = ControlResult('PARTIAL', f'Only one valid Gong meeting date was found ({ctx.last_call.date()}); cadence could not be assessed from two meetings.', WHY['C009'], SOURCES['C009'])
        else:
            r['C009'] = ControlResult('FLAG', 'No valid Gong meeting dates were found in column P.', WHY['C009'], SOURCES['C009'])
    else:
        if ctx.gap <= 30:
            r['C009'] = ControlResult('OK', '', WHY['C009'], SOURCES['C009'])
        elif ctx.gap <= 60:
            r['C009'] = ControlResult('PARTIAL', f'Latest Gong meeting spacing is {ctx.gap} days ({ctx.prev_call.date()} to {ctx.last_call.date()}).', WHY['C009'], SOURCES['C009'])
        else:
            r['C009'] = ControlResult('FLAG', f'Latest Gong meeting spacing is {ctx.gap} days ({ctx.prev_call.date()} to {ctx.last_call.date()}).', WHY['C009'], SOURCES['C009'])

    active_types = detect_personalizations(ctx)
    documented_count, matched = documented_personalizations(ctx.t7 if hasattr(ctx, "t7") else "", active_types)
    active_count = len(active_types)
    if active_count == 0:
        r['C010'] = ControlResult('OK', '', WHY['C010'], SOURCES['C010'])
    else:
        ratio = documented_count / active_count if active_count else 0
        labels = ', '.join(active_types)
        if documented_count >= active_count:
            r['C010'] = ControlResult('OK', '', WHY['C010'], SOURCES['C010'])
        elif ratio >= 0.5:
            r['C010'] = ControlResult('PARTIAL', f'Active personalization was detected across {active_count} area(s) ({labels}), with partial documentation in CS Notes.', WHY['C010'], SOURCES['C010'])
        else:
            r['C010'] = ControlResult('FLAG', f'Active personalization was detected across {active_count} area(s) ({labels}), but CS Notes do not sufficiently document the setup.', WHY['C010'], SOURCES['C010'])

    checks = []
    msgs = []
    daily_target = to_float(ctx.proj_h)
    if daily_target is not None and ctx.window_days and ctx.metrics.get('AdSpend') is not None:
        actual_daily = float(ctx.metrics['AdSpend']) / ctx.window_days
        gap = abs(actual_daily - daily_target) / daily_target if daily_target else None
        checks.append('OK' if gap is not None and gap <= 0.20 else 'PARTIAL' if gap is not None and gap <= 0.40 else 'FLAG')
        msgs.append(f'spend target {daily_target:.1f} vs actual daily spend {actual_daily:.1f}')
    acos_t = norm_pct(ctx.proj_j)
    if acos_t is not None and ctx.metrics.get('ACoS') is not None:
        gap = abs(float(ctx.metrics['ACoS']) - acos_t) / acos_t if acos_t else None
        checks.append('OK' if gap is not None and gap <= 0.10 else 'PARTIAL' if gap is not None and gap <= 0.25 else 'FLAG')
        msgs.append(f'ACoS target {pct_str(acos_t)} vs current {pct_str(float(ctx.metrics["ACoS"]))}')
    tacos_t = norm_pct(ctx.proj_k)
    if tacos_t is not None and ctx.metrics.get('TACoS') is not None:
        gap = abs(float(ctx.metrics['TACoS']) - tacos_t) / tacos_t if tacos_t else None
        checks.append('OK' if gap is not None and gap <= 0.10 else 'PARTIAL' if gap is not None and gap <= 0.25 else 'FLAG')
        msgs.append(f'TACoS target {pct_str(tacos_t)} vs current {pct_str(float(ctx.metrics["TACoS"]))}')
    if not checks:
        r['C011'] = ControlResult('OK', '', WHY['C011'], SOURCES['C011'])
    elif all(x == 'OK' for x in checks):
        r['C011'] = ControlResult('OK', '', WHY['C011'], SOURCES['C011'])
    elif 'FLAG' in checks:
        r['C011'] = ControlResult('FLAG', 'Target vs actual performance shows significant deviation across key KPIs.', WHY['C011'], SOURCES['C011'])
    else:
        r['C011'] = ControlResult('PARTIAL', 'Target vs actual performance shows moderate deviation across key KPIs.', WHY['C011'], SOURCES['C011'])

    tags = [t.lower() for t in ctx.tags if t]
    has_best = any(any(w in t for w in BESTSELLER_WORDS) for t in tags)
    has_category = any(t not in {'', 'none'} and not any(w in t for w in BESTSELLER_WORDS | SEGMENTATION_WORDS) for t in tags)
    has_segment = any(any(w in t for w in SEGMENTATION_WORDS) for t in tags)
    if has_best and has_category:
        r['C012'] = ControlResult('OK', '', WHY['C012'], SOURCES['C012'])
    elif has_category or has_segment:
        r['C012'] = ControlResult('PARTIAL', 'Product tagging is present, but bestseller or category segmentation is only partially identifiable.', WHY['C012'], SOURCES['C012'])
    else:
        r['C012'] = ControlResult('FLAG', 'No clear product tagging or segmentation logic was identified in the campaign tag fields.', WHY['C012'], SOURCES['C012'])

    r['C013'] = ControlResult('OK', '', WHY['C013'], SOURCES['C013'])
    r['C014'] = ControlResult('OK', '', WHY['C014'], SOURCES['C014'])

    return r


def build_summary(ctx: DatabricksContext, results: Dict[str, ControlResult]) -> dict:
    import warnings
    budget_constraint = None
    text = ' '.join([ctx.ay, ctx.am, ctx.bn])
    m = re.search(r'([0-9]{1,3}(?:,[0-9]{3})+|[0-9]+(?:\.[0-9]+)?k)\s*(?:monthly|/month|per month)', text, re.I)
    if m:
        budget_constraint = to_float(m.group(1))
    else:
        warnings.warn(
            f"build_summary: budget_constraint could not be extracted from narrative fields for {ctx.hash_name}. "
            "Budget will show as 'Not documented' in the output.",
            stacklevel=2,
        )
    if m:
        budget_constraint = to_float(m.group(1))
    return {
        'primary_objective': build_primary_objective(ctx, results),
        'customization_context': 'Not evaluated — CS Notes assessed during QR call.',
        'monthly_budget': monthly_budget_from_daily(ctx),
        'acos_objective': norm_pct(ctx.proj_j),
        'tacos_objective': norm_pct(ctx.proj_k),
        'acos_constraint': norm_pct(ctx.o7),
        'tacos_constraint': norm_pct(ctx.ax7),
        'budget_constraint': budget_constraint,
        'primary_kpi': ctx.bw if ctx.bw else 'Not documented',
    }


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
