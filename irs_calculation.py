import os
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def calculate_30_360_days(start_date, end_date):
    """Calculate day count between two dates using 30/360 US (ISDA) convention."""
    y1, m1, d1 = start_date.year, start_date.month, start_date.day
    y2, m2, d2 = end_date.year, end_date.month, end_date.day

    if d1 == 31:
        d1 = 30
    if d2 == 31 and d1 >= 30:
        d2 = 30

    return 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)


def find_column_name(df, candidates):
    """Helper to find column by multiple possible header names (case-insensitive)."""
    col_map = {str(c).strip().lower(): c for c in df.columns}
    for name in candidates:
        if name.lower() in col_map:
            return col_map[name.lower()]
    return None


def match_benchmark_interest(end_date, period_idx, benchmark_flows, used_indices):
    """
    Matches benchmark interest using a 3-tier strategy:
    1. Exact date match (FLOW_PAYMENT_DATE == INTEREST_END_DATE)
    2. Date window match (closest payment date within 10 days of end date)
    3. Positional fallback (chronological period index / FLOW_ID order)
    """
    # 1. Exact date match
    for idx, flow in enumerate(benchmark_flows):
        if idx not in used_indices and flow['date'] == end_date:
            used_indices.add(idx)
            return flow['amount']

    # 2. Window match (handles payment lag / BDC roll up to 10 days)
    candidates = []
    for idx, flow in enumerate(benchmark_flows):
        if idx not in used_indices:
            diff_days = abs((flow['date'] - end_date).days)
            if diff_days <= 10:
                candidates.append((diff_days, idx, flow['amount']))

    if candidates:
        candidates.sort(key=lambda x: x[0])
        _, best_idx, best_amt = candidates[0]
        used_indices.add(best_idx)
        return best_amt

    # 3. Positional fallback by period index
    if period_idx < len(benchmark_flows) and period_idx not in used_indices:
        used_indices.add(period_idx)
        return benchmark_flows[period_idx]['amount']

    return 0.0


def calculate_irs_schedule(product_type, input_df, bstrat_df):
    """Calculate interest accrual schedule and match against BSTRAT benchmark."""
    if input_df.empty:
        return pd.DataFrame()

    start_col = find_column_name(input_df, ['INTEREST_START_DATE', 'START_DATE', 'Start Date'])
    end_col = find_column_name(input_df, ['INTEREST_END_DATE', 'END_DATE', 'MATURITY_DATE', 'End Date'])
    principal_col = find_column_name(input_df, ['CURRENT_PRINCIPAL', 'ORIGINAL_PRINCIPAL', 'PRINCIPAL', 'Notional'])
    rate_col = find_column_name(input_df, ['CURVE_POINT', 'CURRENT_INTEREST_RATE', 'FEE_RATE', 'INTEREST_RATE', 'Rate'])
    offset_col = find_column_name(input_df, ['INTEREST_OFFSET_PAYMENT', 'OFFSET_PAYMENT'])

    df = input_df.copy()
    df['parsed_start'] = pd.to_datetime(df[start_col])
    df['parsed_end'] = pd.to_datetime(df[end_col])
    df = df.sort_values(by=['parsed_start', 'parsed_end']).reset_index(drop=True)

    # Build chronological list of benchmark flows from Output_BSTRAT
    benchmark_flows = []
    if not bstrat_df.empty:
        b_df = bstrat_df.copy()
        prod_col = find_column_name(b_df, ['PRODUCT', 'Product'])
        if prod_col:
            matched_b = b_df[b_df[prod_col].astype(str).str.strip().str.upper() == product_type.upper()]
            if not matched_b.empty:
                b_df = matched_b

        pay_date_col = find_column_name(b_df, ['FLOW_PAYMENT_DATE', 'PAYMENT_DATE', 'Flow Date', 'Payment Date'])
        amount_col = find_column_name(b_df, ['AMOUNT', 'AMOUNT_USD', 'REPORTED_INTEREST', 'Interest Amount'])

        if pay_date_col and amount_col:
            b_df['parsed_pay_date'] = pd.to_datetime(b_df[pay_date_col])
            flow_id_col = find_column_name(b_df, ['FLOW_ID', 'Flow ID', 'FLOWID'])
            if flow_id_col:
                b_df = b_df.sort_values(by=[flow_id_col, 'parsed_pay_date']).reset_index(drop=True)
            else:
                b_df = b_df.sort_values(by='parsed_pay_date').reset_index(drop=True)

            for _, brow in b_df.iterrows():
                benchmark_flows.append({
                    'date': brow['parsed_pay_date'],
                    'amount': float(brow[amount_col])
                })

    is_receive = (product_type.upper() == 'IRS_RECEIVE')
    records = []
    used_indices = set()

    for idx, row in df.iterrows():
        s_date = row['parsed_start']
        e_date = row['parsed_end']
        days = calculate_30_360_days(s_date, e_date)
        notional = float(row[principal_col]) if principal_col and pd.notna(row[principal_col]) else 1000000.0
        rate = float(row[rate_col]) if rate_col and pd.notna(row[rate_col]) else 0.0

        bstrat_interest = match_benchmark_interest(e_date, idx, benchmark_flows, used_indices)

        start_str = s_date.strftime('%m/%d/%Y').lstrip('0').replace('/0', '/')
        end_str = e_date.strftime('%m/%d/%Y').lstrip('0').replace('/0', '/')

        if is_receive:
            calc_interest = -round(notional * (rate / 100.0) * (days / 360.0), 4)
            bstrat_val = round(bstrat_interest, 4)
            diff = round(calc_interest - bstrat_val, 4)
            records.append({
                'INTEREST_START_DATE': start_str,
                'INTEREST_END_DATE': end_str,
                'Notional': notional,
                'Interest Rate': rate,
                'Interest Amount Calc': calc_interest,
                'BSTRAT Interest': bstrat_val,
                'Difference': diff
            })
        else:
            calc_interest = round(notional * (rate / 100.0) * (days / 360.0), 2)
            offset_payment = int(row[offset_col]) if offset_col and pd.notna(row[offset_col]) else 0
            bstrat_val = round(bstrat_interest, 2)
            diff = round(calc_interest - bstrat_val, 2)
            records.append({
                'INTEREST_START_DATE': start_str,
                'INTEREST_END_DATE': end_str,
                'Notional': notional,
                'Days': days,
                'INTEREST_OFFSET_PAYMENT': offset_payment,
                'Interest Rate': rate,
                'Interest Amount': calc_interest,
                'BSTRAT Interest': bstrat_val,
                'Difference': diff
            })

    return pd.DataFrame(records)


def format_excel_worksheet(ws, schedule_df, product_type):
    """Apply styling, formulas, number formatting, and borders to the worksheet."""
    is_receive = (product_type.upper() == 'IRS_RECEIVE')

    headers = [
        'INTEREST_START_DATE', 'INTEREST_END_DATE', 'Notional',
        'Interest Rate', 'Interest Amount Calc', 'BSTRAT Interest', 'Difference'
    ] if is_receive else [
        'INTEREST_START_DATE', 'INTEREST_END_DATE', 'Notional', 'Days',
        'INTEREST_OFFSET_PAYMENT', 'Interest Rate', 'Interest Amount',
        'BSTRAT Interest', 'Difference'
    ]

    header_font = Font(name="Calibri", size=11, bold=True)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_font = Font(name="Calibri", size=11, bold=True)
    total_border = Border(top=Side(style="thin"), bottom=Side(style="double"))
    alert_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    alert_font = Font(name="Calibri", size=11, bold=True, color="C00000")

    currency_fmt = '" "* #,##0.00" ";" "* (#,##0.00);" "* "-"??" '

    # Header row
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 24

    # Data rows
    for row_idx, data in enumerate(schedule_df.to_dict('records'), start=2):
        ws.cell(row=row_idx, column=1, value=data['INTEREST_START_DATE']).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=2, value=data['INTEREST_END_DATE']).alignment = Alignment(horizontal="center")

        if is_receive:
            c_notional = ws.cell(row=row_idx, column=3, value=data['Notional'])
            c_notional.number_format = '#,##0'
            c_notional.alignment = Alignment(horizontal="right")

            c_rate = ws.cell(row=row_idx, column=4, value=data['Interest Rate'])
            c_rate.number_format = '#,##0.00000'
            c_rate.alignment = Alignment(horizontal="right")

            c_calc = ws.cell(row=row_idx, column=5, value=data['Interest Amount Calc'])
            c_calc.number_format = currency_fmt
            c_calc.alignment = Alignment(horizontal="right")

            c_bstrat = ws.cell(row=row_idx, column=6, value=data['BSTRAT Interest'])
            c_bstrat.number_format = currency_fmt
            c_bstrat.alignment = Alignment(horizontal="right")

            # Formula for difference: Calc - BSTRAT
            c_diff = ws.cell(row=row_idx, column=7, value=f"=E{row_idx}-F{row_idx}")
            c_diff.number_format = currency_fmt
            c_diff.alignment = Alignment(horizontal="right")

            if abs(data['Difference']) >= 0.01:
                c_diff.fill = alert_fill
                c_diff.font = alert_font
        else:
            c_notional = ws.cell(row=row_idx, column=3, value=data['Notional'])
            c_notional.number_format = '#,##0.00'
            c_notional.alignment = Alignment(horizontal="right")

            # Formula for days: End Date - Start Date
            c_days = ws.cell(row=row_idx, column=4, value=f"=B{row_idx}-A{row_idx}")
            c_days.number_format = 'General'
            c_days.alignment = Alignment(horizontal="right")

            c_offset = ws.cell(row=row_idx, column=5, value=data['INTEREST_OFFSET_PAYMENT'])
            c_offset.number_format = '#,##0'
            c_offset.alignment = Alignment(horizontal="right")

            c_rate = ws.cell(row=row_idx, column=6, value=data['Interest Rate'])
            c_rate.number_format = '0.000'
            c_rate.alignment = Alignment(horizontal="right")

            c_calc = ws.cell(row=row_idx, column=7, value=data['Interest Amount'])
            c_calc.number_format = currency_fmt
            c_calc.alignment = Alignment(horizontal="right")

            c_bstrat = ws.cell(row=row_idx, column=8, value=data['BSTRAT Interest'])
            c_bstrat.number_format = currency_fmt
            c_bstrat.alignment = Alignment(horizontal="right")

            # Formula for difference: Calc - BSTRAT
            c_diff = ws.cell(row=row_idx, column=9, value=f"=G{row_idx}-H{row_idx}")
            c_diff.number_format = currency_fmt
            c_diff.alignment = Alignment(horizontal="right")

            if abs(data['Difference']) >= 0.01:
                c_diff.fill = alert_fill
                c_diff.font = alert_font

    # Total summary row
    num_rows = len(schedule_df)
    tot_row = num_rows + 2
    ws.cell(row=tot_row, column=1, value="Total").font = total_font
    ws.cell(row=tot_row, column=1).alignment = Alignment(horizontal="center")

    if is_receive:
        ws.cell(row=tot_row, column=5, value=f"=SUM(E2:E{tot_row-1})").number_format = currency_fmt
        ws.cell(row=tot_row, column=6, value=f"=SUM(F2:F{tot_row-1})").number_format = currency_fmt
        ws.cell(row=tot_row, column=7, value=f"=SUM(G2:G{tot_row-1})").number_format = currency_fmt
        cols = range(1, 8)
    else:
        ws.cell(row=tot_row, column=7, value=f"=SUM(G2:G{tot_row-1})").number_format = currency_fmt
        ws.cell(row=tot_row, column=8, value=f"=SUM(H2:H{tot_row-1})").number_format = currency_fmt
        ws.cell(row=tot_row, column=9, value=f"=SUM(I2:I{tot_row-1})").number_format = currency_fmt
        cols = range(1, 10)

    for c in cols:
        cell = ws.cell(row=tot_row, column=c)
        cell.font = total_font
        cell.border = total_border

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)


def process_irs_file(file_path, contract_id):
    """Main processing workflow for IRS calculation and reconciliation."""
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' does not exist.")
        sys.exit(1)

    excel = pd.ExcelFile(file_path)

    # Locate worksheets
    sheet_names_lower = {s.strip().lower(): s for s in excel.sheet_names}
    input_sheet = sheet_names_lower.get('input_customize') or sheet_names_lower.get('input')
    output_sheet = sheet_names_lower.get('output_bstrat') or sheet_names_lower.get('output')

    if not input_sheet:
        print(f"Error: 'Input_Customize' sheet not found in '{file_path}'.")
        sys.exit(1)

    input_df = excel.parse(input_sheet)
    bstrat_df = excel.parse(output_sheet) if output_sheet else pd.DataFrame()

    # Find contract ID column
    cid_col = find_column_name(input_df, ['CONTRACT_ID', 'Contract ID', 'Contract', 'ID', 'UNIQUE_ID'])
    if not cid_col:
        print(f"Error: Contract ID column not found in '{input_sheet}'.")
        sys.exit(1)

    # Normalize contract IDs for matching
    clean_target_id = str(contract_id).strip().removesuffix('.0')
    input_df['_norm_cid'] = input_df[cid_col].astype(str).str.strip().str.removesuffix('.0')

    if clean_target_id not in input_df['_norm_cid'].values:
        print("Record not found")
        sys.exit(0)

    matched_input = input_df[input_df['_norm_cid'] == clean_target_id].copy()

    # Filter benchmark data for this contract ID
    matched_bstrat = pd.DataFrame()
    if not bstrat_df.empty:
        b_cid_col = find_column_name(bstrat_df, ['UNIQUE_ID', 'CONTRACT_ID', 'Contract ID', 'ID'])
        if b_cid_col:
            bstrat_df['_norm_cid'] = bstrat_df[b_cid_col].astype(str).str.strip().str.removesuffix('.0')
            matched_bstrat = bstrat_df[bstrat_df['_norm_cid'] == clean_target_id].copy()
        else:
            matched_bstrat = bstrat_df.copy()

    # Separate by Product
    prod_col = find_column_name(matched_input, ['PRODUCT', 'Product'])
    if not prod_col:
        print("Error: 'PRODUCT' column not found in Input_Customize.")
        sys.exit(1)

    products = matched_input[prod_col].astype(str).str.strip().str.upper()
    pay_records = matched_input[products == 'IRS_PAY'].copy()
    receive_records = matched_input[products == 'IRS_RECEIVE'].copy()

    print(f"Found {len(pay_records)} IRS_PAY and {len(receive_records)} IRS_RECEIVE records for Contract {clean_target_id}.")

    # Calculate schedules
    pay_schedule = calculate_irs_schedule('IRS_PAY', pay_records, matched_bstrat) if not pay_records.empty else pd.DataFrame()
    receive_schedule = calculate_irs_schedule('IRS_RECEIVE', receive_records, matched_bstrat) if not receive_records.empty else pd.DataFrame()

    if pay_schedule.empty and receive_schedule.empty:
        print("No valid calculation records found.")
        return

    # Create new workbook and write sheets
    output_filename = f"IRS_Calculation_Output_{clean_target_id}.xlsx"
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    if not receive_schedule.empty:
        ws_rec = wb.create_sheet(title=f"{clean_target_id}_IRS Receive")
        format_excel_worksheet(ws_rec, receive_schedule, 'IRS_RECEIVE')
        print(f"Created sheet: '{clean_target_id}_IRS Receive' ({len(receive_schedule)} rows)")

    if not pay_schedule.empty:
        ws_pay = wb.create_sheet(title=f"{clean_target_id}_IRS Pay")
        format_excel_worksheet(ws_pay, pay_schedule, 'IRS_PAY')
        print(f"Created sheet: '{clean_target_id}_IRS Pay' ({len(pay_schedule)} rows)")

    wb.save(output_filename)
    print(f"Successfully saved output to '{output_filename}'")


def main():
    if len(sys.argv) >= 3:
        source_file = sys.argv[1].strip('"').strip("'")
        contract_id = sys.argv[2].strip()
    elif len(sys.argv) == 2:
        source_file = sys.argv[1].strip('"').strip("'")
        contract_id = input("Enter Contract ID: ").strip()
    else:
        source_file = input("Enter the source Excel file path: ").strip().strip('"').strip("'")
        contract_id = input("Enter Contract ID: ").strip()

    if not source_file or not contract_id:
        print("Error: Both source file path and Contract ID are required.")
        sys.exit(1)

    process_irs_file(source_file, contract_id)


if __name__ == "__main__":
    main()
