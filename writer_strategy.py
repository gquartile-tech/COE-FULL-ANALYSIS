import openpyxl
import re
import sys
import os
from datetime import datetime


# ── helpers ───────────────────────────────────────────────────────────────────

def read_header(ws):
    account_str = date_range = downloaded = ""
    for row in ws.iter_rows(min_row=1, max_row=4, values_only=True):
        for cell in row:
            if cell and isinstance(cell, str):
                if "Account:" in cell:
                    account_str = cell
                elif "Date Range:" in cell:
                    date_range = cell.replace("Date Range: ", "").strip()
                elif "Downloaded:" in cell:
                    downloaded = cell.replace("Downloaded: ", "").strip()
    return account_str, date_range, downloaded


def find_header_row(ws, max_scan=10):
    """Find the row index where actual column headers live."""
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), 1):
        non_empty = [c for c in row if c is not None]
        if len(non_empty) > 3:
            return i
    return None


def tab_to_dict(ws):
    """Single-data-row tab → dict. Auto-detects header row."""
    header_row = find_header_row(ws)
    if header_row is None:
        return {}
    rows = list(ws.iter_rows(min_row=header_row, max_row=header_row + 1, values_only=True))
    if len(rows) < 2:
        return {}
    headers, data_row = rows[0], rows[1]
    result = {}
    for i, h in enumerate(headers):
        if h is not None and i < len(data_row):
            result[h] = data_row[i]
    return result


def tab_to_records(ws):
    """Multi-row tab → list of dicts. Auto-detects header row."""
    header_row = find_header_row(ws)
    if header_row is None:
        return []
    headers = None
    records = []
    for i, row in enumerate(ws.iter_rows(min_row=header_row, values_only=True), header_row):
        if headers is None:
            headers = list(row)
            continue
        if not any(row):
            continue
        rec = {headers[j]: row[j] for j in range(len(headers)) if headers[j] is not None}
        records.append(rec)
    return records


def safe(val, default=""):
    return default if val is None else val


# ── main writer ───────────────────────────────────────────────────────────────

def write_strategy(pre_analysis_path: str, template_path: str, output_dir: str):

    pa = openpyxl.load_workbook(pre_analysis_path, data_only=True, read_only=True)

    # ── header ────────────────────────────────────────────────────────────────
    ws01 = pa['01_Advertiser_Name']
    account_str, date_range, downloaded = read_header(ws01)

    m = re.match(r"Account:\s*(.+?)\s*\|\s*Tenant ID:\s*(\S+)\s*\|\s*Account ID:\s*(\S+)", account_str)
    if not m:
        raise ValueError(f"Could not parse account string: {account_str}")
    account_label = m.group(1).strip()
    tenant_id     = m.group(2).strip()
    profile_id    = m.group(3).strip()
    member_id     = account_label.split(" - ")[0].strip()

    # ── tab 55 ────────────────────────────────────────────────────────────────
    d55 = tab_to_dict(pa['55_Salesforce_Consolidated_PreA'])

    # ── tab 38 (fallback fields) ──────────────────────────────────────────────
    d38 = tab_to_dict(pa['38_Client_Success_Insights_Repo'])

    # ── tab 37 gong ───────────────────────────────────────────────────────────
    gong_records = tab_to_records(pa['37_Gong_Call_Insights_for_Sales'])
    gong = gong_records[0] if gong_records else {}

    # ── tab 14 ASIN data ──────────────────────────────────────────────────────
    asin_records = tab_to_records(pa['14_Campaign_Performance_by_Adve'])

    # ── tab 54 project dataset ──────────────────────────────────────────────
    d54 = tab_to_dict(pa['54_Project_Dataset_on_SF'])

    # ── tab 22 catalogue ─────────────────────────────────────────────────────
    cat_records  = tab_to_records(pa['22_Catalogue_Details'])
    cat_by_asin  = {r['asin']: r for r in cat_records if r.get('asin')}

    pa.close()

    # ── load template ─────────────────────────────────────────────────────────
    wb = openpyxl.load_workbook(template_path, keep_vba=True)

    # ════════════════════════════════════════════════════════════════════════════
    # TAB 1 — Questionaire Survey - AMZ
    # ════════════════════════════════════════════════════════════════════════════
    ws1 = wb['Questionaire Survey - AMZ']

    def w1(coord, value):
        ws1[coord] = value

    # Row 6
    w1('C6', member_id)
    w1('F6', profile_id)
    w1('J6', safe(d55.get('CSP_Last_Modified_By')))

    # Row 7
    w1('F7', profile_id)
    w1('J7', safe(d55.get('Projected_Project_MRR__c')))

    # Row 8
    w1('C8', safe(d55.get('Account_Name')))
    ld = d55.get('Launch_Date__c')
    w1('F8', ld.strftime('%Y-%m-%d') if hasattr(ld, 'strftime') else safe(ld))
    if hasattr(ld, 'strftime'):
        months = (datetime.now().year - ld.year) * 12 + (datetime.now().month - ld.month)
        w1('J8', f"{months} months")
    else:
        w1('J8', safe(d38.get('Customer_Age_Months__c')))

    # Row 9
    w1('C9', safe(d55.get('Customer_Age_Months__c') or d38.get('Customer_Age_Months__c')))
    w1('F9', safe(d38.get('Repeat_Purchase_Behavior__c')))
    w1('J9', safe(d55.get('CSM_Churn_Risk__c')))

    # Row 10
    w1('C10', safe(d38.get('Commodity_Products_or_Branded_Products__c')))
    w1('F10', safe(d38.get('Sales_Concentration__c')))
    w1('J10', safe(d55.get('Director_Churn_Risk__c')))

    # Row 11
    w1('C11', safe(d55.get('CSM')))
    w1('F11', safe(d38.get('CSM_Tenure__c')))
    w1('J11', safe(d55.get('Account_Risk_Score__c')))

    # Row 12 — Director SF user ID not readable, leave blank
    w1('C12', '')
    w1('F12', safe(d55.get('Active_Products__c')))

    # Row 13
    w1('F13', safe(d38.get('Customer_Feedback__c')))

    # Row 15
    w1('C15', safe(d55.get('Current_Challenges__c')))
    w1('F15', safe(d55.get('Primary_Objective__c')))
    w1('J15', safe(d55.get('ACOS_Constraint__c')))

    # Row 16
    w1('C16', safe(d55.get('Primary_Objective_Additional_Context__c')))
    w1('F16', safe(d55.get('Primary_Spend_KPI__c')))
    w1('J16', safe(d38.get('Customer_Acquisition_Cost_Target__c')))

    # Row 17
    w1('C17', safe(d55.get('Top_Priority__c')))
    w1('J17', safe(d55.get('TACOS_Constraint__c')))

    # Row 18
    w1('C18', safe(d55.get('Second_Priority__c')))
    w1('J18', safe(d55.get('daily_target_spend__c')))

    # F18 CS Notes value from tab 54 (E18 label 'CS_NOTES' already in template)
    w1('F18', safe(d54.get('CS_Notes__c')))

    # Row 19
    w1('C19', safe(d55.get('Biggest_Expansion_Opportunity__c')))
    w1('F19', safe(d55.get('Near_Term_3_Month_Considerations__c')))
    w1('J19', safe(d55.get('Target_ROAS__c')))

    # CJM Stages
    stage_rows = {1: (24, 25), 2: (27, 28), 3: (30, 31), 4: (33, 34)}
    for s, (r_a, r_i) in stage_rows.items():
        w1(f'C{r_a}', safe(d55.get(f'AdoptionOrUpsellS{s}__c')))
        w1(f'G{r_a}', safe(d55.get(f'StrategyS{s}__c')))
        w1(f'J{r_a}', safe(d55.get(f'StatusS{s}__c')))
        intro = d55.get(f'ExecutionDateS{s}__c')
        w1(f'C{r_i}', intro.strftime('%Y-%m-%d') if hasattr(intro, 'strftime') else safe(intro))

    # Gong
    w1('C41', safe(gong.get('Gong__Call_Brief__c') or d55.get('Call_Brief')))
    w1('C42', safe(gong.get('Gong__Call_Key_Points__c') or d55.get('Key_Points')))
    w1('C43', safe(gong.get('Gong__Call_Highlights_Next_Steps__c') or d55.get('Highlights_Next_Steps')))

    # ════════════════════════════════════════════════════════════════════════════
    # TAB 2 — Account Strategy _Analysis header
    # ════════════════════════════════════════════════════════════════════════════
    ws2 = wb['Account Strategy _Analysis']
    ws2['A1'] = f"{account_label} — Account Strategy Analysis"
    ws2['B3'] = f"Account: {account_label} | Tenant ID: {tenant_id} | Account ID: {profile_id}"
    ws2['B4'] = date_range
    ws2['B5'] = downloaded

    # ════════════════════════════════════════════════════════════════════════════
    # TAB 3 — ChildASIN View
    # ════════════════════════════════════════════════════════════════════════════
    ws3 = wb['ChildASIN View']

    col14_map = {
        'Parent ASIN':        'ParentASIN',
        'ASIN':               'asin',
        'Total Sales':        'TotalSales',
        'Total Units Ordered':'UnitsOrdered',
        'Ad Spend':           'AdSpend',
        'TACoS':              'TACoS',
        'Ad Sales':           'AdSales',
        'Ads Units Ordered':  'Orders',
        'ACoS':               'ACoS',
        'Clicks':             'Clicks',
        'Tier':               'Tier',
        'Buy Box%':           'Weighted_BuyBoxPercentage',
        'ATM_Spend':          'ATM_Spend',
        'BA_Spend':           'BA_Spend',
        'Manual_Q1_Spend':    'Manual_Q1_Spend',
        'BAK_Spend':          'BAK_Spend',
        'OP_Spend':           'OP_Spend',
        'SPT_Spend':          'SPT_Spend',
        'CAT_SP_Spend':       'CAT_SP_Spend',
        'WATM_Spend':         'WATM_Spend',
        'SB_Spend':           'SB_Spend',
        'SBV_Spend':          'SBV_Spend',
        'SD_Spend':           'SD_Spend',
        'Imported_Spend':     'Imported_Spend',
        'NonQuartile_Spend':  'NonQuartile_Spend',
        'TAG 1':              'Tag1',
        'TAG 2':              'Tag2',
        'TAG 3':              'Tag3',
        'TAG 4':              'Tag4',
        'TAG 5':              'Tag5',
    }

    col22_map = {
        'AOV':        'AOV',
        'PriceTier':  'PriceTier',
        'Brand':      'Brand',
        'Department': 'Department',
        'Category':   'Category',
    }

    # Build header → column index from row 2
    header_to_col = {}
    for cell in ws3[2]:
        if cell.value:
            header_to_col[cell.value] = cell.column

    # Clear existing data/formulas in Tab 3 from row 3 down
    for row in ws3.iter_rows(min_row=3, max_col=ws3.max_column):
        for cell in row:
            cell.value = None

    # TotalSalesAll denominator from first ASIN record
    total_sales_all = (asin_records[0].get('TotalSalesAll') or 1) if asin_records else 1

    for row_idx, rec in enumerate(asin_records, start=3):
        asin = rec.get('asin', '')
        cat  = cat_by_asin.get(asin, {})

        for header, col_idx in header_to_col.items():
            val = None

            if header in col14_map:
                val = rec.get(col14_map[header])

            elif header == 'Ad Sales (%)':
                ad_s  = rec.get('AdSales') or 0
                tot_s = rec.get('TotalSalesAll') or 1
                val   = round(ad_s / tot_s, 4)

            elif header == 'Organic Sales (%)':
                ad_s  = rec.get('AdSales') or 0
                tot_s = rec.get('TotalSalesAll') or 1
                val   = round(1 - (ad_s / tot_s), 4)

            elif header in col22_map:
                val = cat.get(col22_map[header])

            # Quartile One / Quartile Bulk — leave blank (calculated in sheet)
            if val is not None:
                ws3.cell(row=row_idx, column=col_idx, value=val)

    # ── save ──────────────────────────────────────────────────────────────────
    filename = f"{account_label} — Strategy Analysis {date_range}.xlsm"
    filename = re.sub(r'[<>:"/\\|?*]', '-', filename)
    out_path = os.path.join(output_dir, filename)
    wb.save(out_path)
    print(f"Saved: {out_path}")
    return out_path


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python writer_strategy.py <pre_analysis.xlsx> <template.xlsm> [output_dir]")
        sys.exit(1)
    write_strategy(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "/mnt/user-data/outputs")
# ── patch already applied inline in next version ──
