from __future__ import annotations

from openpyxl import load_workbook
from openpyxl.styles import Alignment

from config_mastery import CONTROL_NAMES, IMPORTANCE, MAX_FINDINGS, PRIORITY_POINTS
from reader_databricks_mastery import money_str, pct_str, clean_text
from rules_engine_mastery import interpretation


def write_mastery_output(template_path, output_path, summary, results, penalty, score, grade, findings, ctx):
    wb = load_workbook(template_path, keep_vba=True)
    ws_main = wb['Account Mastery_Analysis']
    ws_ref = wb['Account Mastery_Reference']
    ws_logic = wb['Logic and Calculation']

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass

    ws_main['A1'] = f"{ctx.hash_name} — Account Mastery Analysis"
    ws_main['B3'] = f"Account: {ctx.hash_name} | Tenant ID: {ctx.tenant_id} | Account ID: {ctx.account_id}"
    if ctx.window_start and ctx.window_end and ctx.window_days:
        ws_main['B4'] = f"{ctx.window_start} to {ctx.window_end} ({ctx.window_days} days)"
    ws_main['B5'] = ctx.downloaded
    ws_main['B5'].number_format = 'yyyy-mm-dd hh:mm:ss'

    ws_main['C11'] = summary['primary_objective']
    ws_main['C13'] = summary.get('customization_context', '')
    ws_main['C16'] = money_str(summary['monthly_budget']) if summary['monthly_budget'] is not None else 'Monthly budget target not available.'
    ws_main['B18'] = pct_str(summary['acos_objective'])
    ws_main['B19'] = pct_str(summary['tacos_objective'])
    ws_main['B22'] = pct_str(summary['acos_constraint'])
    ws_main['B23'] = pct_str(summary['tacos_constraint'])
    ws_main['B24'] = money_str(summary['budget_constraint'])
    ws_main['E23'] = summary['primary_kpi']
    ws_main['D7'] = grade
    ws_main['D8'] = interpretation(grade)

    for cell in ['C11', 'C13', 'C16', 'D8']:
        ws_main[cell].alignment = Alignment(wrap_text=True, vertical='top')

    cid_to_row = {}
    for r in range(2, ws_ref.max_row + 1):
        cid = clean_text(ws_ref[f'B{r}'].value).upper()
        if cid:
            cid_to_row[cid] = r

    for cid, res in results.items():
        if cid not in cid_to_row:
            print(f"[writer] WARNING: control {cid} not found in Account Mastery_Reference sheet — skipping.")
            continue
        rr = cid_to_row[cid]
        ws_ref[f'C{rr}'] = CONTROL_NAMES.get(cid, '')
        ws_ref[f'D{rr}'] = res.status
        ws_ref[f'H{rr}'] = res.what
        ws_ref[f'I{rr}'] = res.why
        ws_ref[f'J{rr}'] = res.source
        ws_ref[f'M{rr}'] = PRIORITY_POINTS.get(IMPORTANCE.get(cid, 1), 0)
        for cell in [f'H{rr}', f'I{rr}', f'J{rr}']:
            ws_ref[cell].alignment = Alignment(wrap_text=True, vertical='top')

    ws_logic['B3'] = penalty
    ws_logic['B4'] = score
    ws_logic['B5'] = grade
    for idx, (cid, res) in enumerate(results.items(), start=8):
        row_ref = cid_to_row.get(cid)
        ws_logic[f'A{idx}'] = cid
        ws_logic[f'B{idx}'] = clean_text(ws_ref[f'A{row_ref}'].value) if row_ref else ''
        ws_logic[f'C{idx}'] = CONTROL_NAMES.get(cid, '')
        ws_logic[f'D{idx}'] = res.status
        ws_logic[f'E{idx}'] = PRIORITY_POINTS.get(IMPORTANCE.get(cid, 1), 0)
        ws_logic[f'F{idx}'] = PRIORITY_POINTS.get(IMPORTANCE.get(cid, 1), 0) if res.status == 'FLAG' else PRIORITY_POINTS.get(IMPORTANCE.get(cid, 1), 0) * 0.5 if res.status == 'PARTIAL' else 0

    start_row = 28
    priority_items = [f for f in findings if f['status'] in ('FLAG', 'PARTIAL')]
    for i in range(MAX_FINDINGS):
        r = start_row + i
        item = priority_items[i] if i < len(priority_items) else None
        for c in 'ABCDEF':
            ws_main[f'{c}{r}'] = ''
        if item:
            ws_main[f'A{r}'] = item['cid']
            ws_main[f'B{r}'] = item['name']
            ws_main[f'C{r}'] = item['status']
            ws_main[f'D{r}'] = item['what']
            ws_main[f'E{r}'] = item['why']
            ws_main[f'F{r}'] = item['impact']
            for c in ['B', 'D', 'E']:
                ws_main[f'{c}{r}'].alignment = Alignment(wrap_text=True, vertical='top')

    wb.save(output_path)
    try:
        wb.close()
    except Exception:
        pass
