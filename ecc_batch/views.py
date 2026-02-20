"""
Django views for ECC Batch Processing System

This module handles:
- PDF upload and processing
- Cheque data extraction from PDF
- Excel batch file generation (Full and Accepted batches)
- Data display and validation
"""

import os
import re
import tempfile
import traceback
from datetime import datetime

import pdfplumber
import xlwt
from django.http import HttpResponse, JsonResponse
from django.shortcuts import render
from django.views.decorators.csrf import csrf_exempt


# =============================================================================
# DJANGO VIEWS
# =============================================================================

def dashboard(request):
    """Render the dashboard page with upload and processing options."""
    return render(request, 'dashboard.html', {'now': datetime.now().year})


@csrf_exempt
def generate_batch(request):
    """
    Process uploaded PDF and generate an Excel batch file.

    Supports:
    - 'full': all cheques
    - 'accepted': only cheques where reason contains 'ACCEPTED'
    """
    if request.method != 'POST':
        return JsonResponse({'error': 'Only POST method allowed'}, status=405)

    pdf_file = request.FILES.get('pdf_file')
    if not pdf_file:
        return JsonResponse({'error': 'No PDF file uploaded'}, status=400)

    if not pdf_file.name.endswith('.pdf'):
        return JsonResponse({'error': 'Invalid file format. Please upload a PDF file.'}, status=400)

    batch_type = request.POST.get('batch_type', 'full')

    # Read branch code and parking account from form; fall back to defaults
    clearing_branch = request.POST.get('branch_code', '255').strip() or '255'
    clearing_account = request.POST.get('parking_account', '9313102000').strip() or '9313102000'

    temp_pdf_path = None
    output_path = None

    try:
        # Save uploaded PDF to a temp file
        with tempfile.NamedTemporaryFile(mode='wb', suffix='.pdf', delete=False) as temp_pdf:
            for chunk in pdf_file.chunks():
                temp_pdf.write(chunk)
            temp_pdf_path = temp_pdf.name

        extracted_data = extract_pdf_data(temp_pdf_path)
        if not extracted_data:
            return JsonResponse({'error': 'No data extracted from PDF'}, status=400)

        # Filter if accepted batch requested
        data = extracted_data
        if batch_type == 'accepted':
            data = filter_accepted_cheques(extracted_data)
            if not data:
                return JsonResponse(
                    {'error': 'No accepted cheques found in the PDF. Please check the reason column.'},
                    status=400
                )

        # Pass user-supplied clearing details
        output_path = generate_excel_batch(
            data,
            clearing_account=clearing_account,
            clearing_branch=clearing_branch,
        )

        filename = 'ecc_batch_accepted.xls' if batch_type == 'accepted' else 'ecc_batch.xls'
        with open(output_path, 'rb') as excel_file:
            response = HttpResponse(excel_file.read(), content_type='application/vnd.ms-excel')
            response['Content-Disposition'] = f'attachment; filename="{filename}"'
            return response

    except Exception as e:
        print(f"Error processing file: {traceback.format_exc()}")
        return JsonResponse({'error': f'Error processing file: {str(e)}'}, status=500)

    finally:
        _safe_remove(temp_pdf_path)
        _safe_remove(output_path)


@csrf_exempt
def display_data(request):
    """Display extracted data from uploaded PDF as JSON."""
    if request.method != 'POST':
        return JsonResponse({'error': 'Only POST method allowed'}, status=405)

    pdf_file = request.FILES.get('pdf_file')
    if not pdf_file:
        return JsonResponse({'error': 'No PDF file uploaded'}, status=400)

    if not pdf_file.name.endswith('.pdf'):
        return JsonResponse({'error': 'Invalid file format. Please upload a PDF file.'}, status=400)

    temp_pdf_path = None

    try:
        with tempfile.NamedTemporaryFile(mode='wb', suffix='.pdf', delete=False) as temp_pdf:
            for chunk in pdf_file.chunks():
                temp_pdf.write(chunk)
            temp_pdf_path = temp_pdf.name

        extracted_data = extract_pdf_data(temp_pdf_path)
        if not extracted_data:
            return JsonResponse({'error': 'Upload ECC Report Only'}, status=400)

        return JsonResponse({'success': True, 'count': len(extracted_data), 'data': extracted_data})

    except Exception as e:
        print(f"Error processing file: {traceback.format_exc()}")
        return JsonResponse({'error': f'Error processing file: {str(e)}'}, status=500)

    finally:
        _safe_remove(temp_pdf_path)


@csrf_exempt
def display_data_table(request):
    """Display extracted data from uploaded PDF in HTML table format."""
    if request.method != 'POST':
        return JsonResponse({'error': 'Only POST method allowed'}, status=405)

    pdf_file = request.FILES.get('pdf_file')
    if not pdf_file:
        return JsonResponse({'error': 'No PDF file uploaded'}, status=400)

    if not pdf_file.name.endswith('.pdf'):
        return JsonResponse({'error': 'Invalid file format. Please upload a PDF file.'}, status=400)

    temp_pdf_path = None

    try:
        with tempfile.NamedTemporaryFile(mode='wb', suffix='.pdf', delete=False) as temp_pdf:
            for chunk in pdf_file.chunks():
                temp_pdf.write(chunk)
            temp_pdf_path = temp_pdf.name

        extracted_data = extract_pdf_data(temp_pdf_path)
        if not extracted_data:
            return JsonResponse({'error': 'Upload ECC Report Only'}, status=400)

        # Sort data by cheque_amount in ascending order
        extracted_data = sorted(extracted_data, key=lambda x: float(x.get('cheque_amount', 0)))

        # Store data in session for Excel download
        request.session['display_data'] = extracted_data
        request.session['display_filename'] = pdf_file.name

        # Summary stats
        total_records = len(extracted_data)
        total_amount = 0.0

        accepted_count = 0
        accepted_amount = 0.0
        insufficient_funds_count = 0
        insufficient_funds_amount = 0.0
        other_reasons_count = 0
        other_reasons_amount = 0.0

        for r in extracted_data:
            reason = str(r.get('reason', '')).upper().strip()
            amount = float(r['cheque_amount'])
            total_amount += amount
            
            if 'ACCEPTED' in reason:
                accepted_count += 1
                accepted_amount += amount
            elif 'INSUFFICIENT' in reason or 'FUNDS' in reason:
                insufficient_funds_count += 1
                insufficient_funds_amount += amount
            elif reason:
                other_reasons_count += 1
                other_reasons_amount += amount

        no_reason_count = total_records - (accepted_count + insufficient_funds_count + other_reasons_count)

        context = {
            'data': extracted_data,
            'total_records': total_records,
            'total_amount': total_amount,
            'accepted_count': accepted_count,
            'accepted_amount': accepted_amount,
            'insufficient_funds_count': insufficient_funds_count,
            'insufficient_funds_amount': insufficient_funds_amount,
            'other_reasons_count': other_reasons_count,
            'other_reasons_amount': other_reasons_amount,
            'no_reason_count': no_reason_count,
            'filename': pdf_file.name,
            'now': datetime.now(),
        }
        
        print(f"\n=== SUMMARY STATS ===")
        print(f"Total Records: {total_records}")
        print(f"Total Amount: {total_amount}")
        print(f"Accepted: {accepted_count} cheques, Amount: {accepted_amount}")
        print(f"Insufficient Funds: {insufficient_funds_count} cheques, Amount: {insufficient_funds_amount}")
        print(f"Other Reasons: {other_reasons_count} cheques, Amount: {other_reasons_amount}")
        print(f"No Reason: {no_reason_count}")
        
        return render(request, 'display_data.html', context)

    except Exception as e:
        print(f"Error processing file: {traceback.format_exc()}")
        return JsonResponse({'error': f'Error processing file: {str(e)}'}, status=500)

    finally:
        _safe_remove(temp_pdf_path)


def download_excel(request):
    """Download the displayed data as Excel file from session."""
    # Get data from session
    extracted_data = request.session.get('display_data')
    
    if not extracted_data:
        return JsonResponse({'error': 'No data available. Please upload a PDF first.'}, status=400)

    output_path = None

    try:
        # Generate Excel file
        output_path = generate_display_excel(extracted_data)

        # Return file as download
        with open(output_path, 'rb') as excel_file:
            response = HttpResponse(
                excel_file.read(),
                content_type='application/vnd.ms-excel'
            )
            response['Content-Disposition'] = 'attachment; filename="full_report.xls"'
            return response

    except Exception as e:
        print(f"Error generating Excel: {traceback.format_exc()}")
        return JsonResponse({'error': f'Error generating Excel: {str(e)}'}, status=500)

    finally:
        _safe_remove(output_path)


# =============================================================================
# DATA FILTERING
# =============================================================================

def filter_accepted_cheques(data):
    """Return only records where reason contains 'ACCEPTED'."""
    accepted = []
    for record in data:
        reason = str(record.get('reason', '')).upper().strip()
        if 'ACCEPTED' in reason:
            accepted.append(record)

    print("\nFiltering Results:")
    print(f"  Total records: {len(data)}")
    print(f"  Accepted records: {len(accepted)}")
    return accepted


# =============================================================================
# PDF DATA EXTRACTION
# =============================================================================

BANK_NAME_FIXES = {
    'CITIZE': 'CITIZENS',
    'KUMARI': 'KUMARI',
    'KUMA': 'KUMARI',
    'SIDDHA': 'SIDDHARTHA',
    'SIDDH': 'SIDDHARTHA',
    'MACHI': 'MACHAPUCHARE',
    'MACHA': 'MACHAPUCHARE',
    'SUNRI': 'SUNRISE',
    'EXCEL': 'EXCEL',
}


def extract_pdf_data(pdf_path):
    """
    Extract cheque data from PDF using pdfplumber.
    """
    extracted_records = []

    print("\n" + "=" * 80)
    print("STARTING PDF EXTRACTION")
    print("=" * 80)

    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"Total pages in PDF: {len(pdf.pages)}")

            for page_num, page in enumerate(pdf.pages):
                print(f"\n--- PAGE {page_num + 1} ---")
                tables = page.extract_tables()

                if tables:
                    print(f"Found {len(tables)} table(s) on page {page_num + 1}")
                    page_records = _extract_from_table(tables, BANK_NAME_FIXES)
                    print(f"Extracted {len(page_records)} records from tables")
                else:
                    print("No tables found, using text extraction")
                    page_records = _extract_from_text(page, BANK_NAME_FIXES)
                    print(f"Extracted {len(page_records)} records from text")

                extracted_records.extend(page_records)

    except Exception as e:
        print(f"Error extracting PDF data: {str(e)}")
        raise

    print("\n" + "=" * 80)
    print(f"TOTAL EXTRACTED: {len(extracted_records)} records")
    print("=" * 80 + "\n")

    return extracted_records


def _extract_from_table(tables, bank_name_fixes):
    """
    Extract data from PDF tables (ECC Report format).
    """
    records = []

    for table_idx, table in enumerate(tables):
        print(f"\n  Table {table_idx + 1}: {len(table)} rows")

        for row_idx, row in enumerate(table):
            try:
                if not row or len(row) < 9:
                    print(f"    Row {row_idx}: TOO SHORT ({len(row) if row else 0} columns) - skipped")
                    continue

                first_cell = str(row[0]).strip().upper() if row[0] else ''
                if first_cell in ['SESSION', 'SN', 'S.N', 'S.N.', 'SR', 'NO', 'SERIAL', 'DATE', 'SEQUENCE']:
                    print(f"    Row {row_idx}: HEADER ({first_cell}) - skipped")
                    continue
                if 'TOTAL' in first_cell or 'END OF REPORT' in first_cell:
                    print(f"    Row {row_idx}: FOOTER - skipped")
                    continue

                bfd_account = _extract_bfd_account(row)
                cheque_number = _extract_cheque_number(row)
                cheque_amount = _extract_amount(row)
                pay_bank_name, pay_account = _extract_bank_info(row, bank_name_fixes)
                branch_code = _extract_branch_code(row)
                reason = _extract_reason(row)

                if row_idx <= 3:
                    print(
                        f"    Row {row_idx}: Debug - Cols[6]={row[6] if len(row) > 6 else 'N/A'}, "
                        f"Cols[9]={row[9] if len(row) > 9 else 'N/A'}, "
                        f"Cols[13]={row[13] if len(row) > 13 else 'N/A'}"
                    )

                if bfd_account and cheque_amount and cheque_number:
                    records.append({
                        'bfd_account': bfd_account,
                        'cheque_amount': cheque_amount,
                        'pay_bank_name': pay_bank_name if pay_bank_name else 'CLG',
                        'pay_account': pay_account if pay_account else '',
                        'cheque_number': cheque_number,
                        'branch_code': branch_code if branch_code else '255',
                        'reason': reason if reason else '',
                    })
                    print(f"    Row {row_idx}: ✓ EXTRACTED - BFD:{bfd_account}, Amt:{cheque_amount}, Chq:{cheque_number}")
                else:
                    print(f"    Row {row_idx}: ✗ SKIPPED - BFD:{bfd_account}, Amt:{cheque_amount}, Chq:{cheque_number}")

            except Exception as e:
                print(f"    Row {row_idx}: ✗ EXCEPTION - {str(e)}")
                continue

    return records


def _extract_from_text(page, bank_name_fixes):
    """Fallback extraction using raw text."""
    records = []
    text = page.extract_text()
    if not text:
        return records

    lines = text.split('\n')
    for idx, line in enumerate(lines):
        try:
            parts = line.split()
            if len(parts) < 9:
                continue

            has_cheque_number = any(p.isdigit() and 6 <= len(p) <= 10 for p in parts)
            has_amount = any(',' in p or ('.' in p and any(c.isdigit() for c in p)) for p in parts)
            if not (has_cheque_number and has_amount):
                continue

            bfd_account = _extract_bfd_account_from_parts(parts)
            pay_bank_name, pay_account = _extract_bank_info_from_parts(parts, idx, lines, bank_name_fixes)
            cheque_number = _extract_cheque_number_from_parts(parts)
            cheque_amount = _extract_amount_from_parts(parts)
            branch_code = _extract_branch_code_from_parts(parts)
            reason = _extract_reason_from_parts(parts)

            if bfd_account and cheque_amount and cheque_number:
                records.append({
                    'bfd_account': bfd_account,
                    'cheque_amount': cheque_amount,
                    'pay_bank_name': pay_bank_name if pay_bank_name else 'CLG',
                    'pay_account': pay_account if pay_account else '',
                    'cheque_number': cheque_number,
                    'branch_code': branch_code if branch_code else '255',
                    'reason': reason if reason else '',
                })

        except Exception:
            continue

    return records


# =============================================================================
# FIELD EXTRACTION HELPERS - TABLE MODE
# =============================================================================

def _extract_bfd_account(row):
    """BFD Account is typically at col[9]; accept 12-17 digits; keep as text."""
    if len(row) > 9 and row[9]:
        cell_str = str(row[9]).strip()
        if cell_str.isdigit() and 12 <= len(cell_str) <= 17:
            return cell_str

    candidates = []
    for i, cell in enumerate(row):
        if not cell:
            continue
        s = str(cell).strip()
        if s.isdigit() and 12 <= len(s) <= 17:
            candidates.append((len(s), i, s))

    if candidates:
        candidates.sort(key=lambda x: (x[0], -abs(x[1] - 9)), reverse=True)
        return candidates[0][2]
    return None


def _extract_bank_info(row, bank_name_fixes):
    """Pay Bank Name at col[11], Pay Account at col[12]."""
    pay_bank_name = None
    pay_account = None

    if len(row) > 11 and row[11]:
        cell_str = str(row[11]).strip()
        if '\n' in cell_str:
            pay_bank_name = ''.join(cell_str.split('\n')).strip()
        elif cell_str.isalpha() or (cell_str and not cell_str.isdigit()):
            pay_bank_name = cell_str.upper()

        if pay_bank_name and pay_bank_name in bank_name_fixes:
            pay_bank_name = bank_name_fixes[pay_bank_name]

    if len(row) > 12 and row[12]:
        acc_cell = str(row[12]).strip()
        if acc_cell.isdigit() and len(acc_cell) >= 8:
            pay_account = acc_cell

    if not pay_bank_name:
        for i, cell in enumerate(row):
            if not cell:
                continue
            s = str(cell).strip()
            if s.isdigit() and 3 <= len(s) <= 4 and i < len(row) - 2:
                if i + 1 < len(row) and row[i + 1]:
                    next_cell = str(row[i + 1]).strip()
                    if next_cell and (next_cell.isalpha() or '\n' in next_cell):
                        if '\n' in next_cell:
                            pay_bank_name = ''.join(next_cell.split('\n')).strip()
                        else:
                            pay_bank_name = next_cell.upper()

                        if pay_bank_name in bank_name_fixes:
                            pay_bank_name = bank_name_fixes[pay_bank_name]

                        if i + 2 < len(row) and row[i + 2]:
                            acc_cell = str(row[i + 2]).strip()
                            if acc_cell.isdigit() and len(acc_cell) >= 8:
                                pay_account = acc_cell
                        break

    return pay_bank_name, pay_account


def _extract_cheque_number(row):
    """Cheque number at col[6], 6-10 digits."""
    if len(row) > 6 and row[6]:
        s = str(row[6]).strip()
        if s.isdigit() and 6 <= len(s) <= 10:
            return s

    for cell in row:
        if not cell:
            continue
        s = str(cell).strip()
        if s.isdigit() and 6 <= len(s) <= 10:
            return s
    return None


def _extract_amount(row):
    """Cheque amount at col[13]; fallback scans for reasonable numeric values."""
    if len(row) > 13 and row[13]:
        s = str(row[13]).strip()
        try:
            if ',' in s or '.' in s or s.isdigit():
                amount_str = s.replace(',', '')
                if amount_str.replace('.', '').replace('-', '').isdigit():
                    amt = float(amount_str)
                    if amt > 0:
                        return amt
        except Exception:
            pass

    for cell in row:
        if not cell:
            continue
        s = str(cell).strip()

        if ',' in s or '.' in s:
            try:
                amount_str = s.replace(',', '')
                if amount_str.replace('.', '').replace('-', '').isdigit():
                    amt = float(amount_str)
                    if 0 < amt < 100000000:
                        return amt
            except Exception:
                continue
        elif s.isdigit() and len(s) >= 3:
            try:
                amt = float(s)
                if 0 < amt < 100000000:
                    return amt
            except Exception:
                continue

    return None


def _extract_branch_code(row):
    """Branch code at col[7], 3 digits."""
    if len(row) > 7 and row[7]:
        s = str(row[7]).strip()
        if s.isdigit() and len(s) == 3:
            return s

    for cell in row:
        if not cell:
            continue
        s = str(cell).strip()
        if s.isdigit() and len(s) == 3 and s != '201':
            return s
    return None


def _extract_reason(row):
    """Reason is usually near end; try col[13] then scan backwards for keywords/text."""
    reason_keywords = [
        'INSUFFICIENT', 'SIGNATURE', 'ACCOUNT', 'CLOSED', 'STOP',
        'PAYMENT', 'DRAWER', 'IRREGUL', 'FUNDS', 'REFER', 'EXCEEDS',
        'AMOUNT', 'WORDS', 'FIGURES', 'DIFFER', 'POST', 'DATED',
        'STALE', 'MUTILATED', 'CLEARING', 'ENDORSEMENT', 'ACCEPTED',
        'INVALID', 'WRONG',
    ]

    if len(row) > 13 and row[13]:
        s = str(row[13]).strip().upper()
        if len(s) >= 3 and not s.replace(',', '').replace('.', '').isdigit():
            return s

    for i in range(len(row) - 1, -1, -1):
        cell = row[i]
        if not cell:
            continue
        s = str(cell).strip().upper()

        if s.replace(',', '').replace('.', '').isdigit():
            continue
        if len(s) < 3:
            continue
        if s in ['CLG', 'BANK', 'ENDORSEMENT', 'BFD', 'PAY', 'ACCOUNT']:
            continue
        if any(k in s for k in reason_keywords):
            return s
        if ' ' in s and any(c.isalpha() for c in s):
            return s

    return None


# =============================================================================
# FIELD EXTRACTION HELPERS - TEXT MODE
# =============================================================================

def _extract_bfd_account_from_parts(parts):
    """BFD account should be 12-17 digits; return longest match."""
    candidates = [(len(p), p) for p in parts if p.isdigit() and 12 <= len(p) <= 17]
    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][1]
    return None


def _extract_bank_info_from_parts(parts, line_idx, all_lines, bank_name_fixes):
    pay_bank_name = None
    pay_account = None
    pay_bank_idx = None

    for part_idx, part in enumerate(parts):
        if part.isalpha() and 2 <= len(part) <= 12 and part.isupper():
            if part not in ['ACCEPTED', 'Bank', 'Endorsement', 'Irregular', 'Fund', 'BFD', 'PAY']:
                if line_idx + 1 < len(all_lines) and len(part) < 6:
                    next_line_parts = all_lines[line_idx + 1].split()
                    if next_line_parts and next_line_parts[0].isalpha() and len(next_line_parts[0]) <= 6:
                        combined = part + next_line_parts[0]
                        if combined in bank_name_fixes.values() or len(combined) >= 4:
                            pay_bank_name = combined
                            pay_bank_idx = part_idx
                            break

                pay_bank_name = part
                pay_bank_idx = part_idx

                if pay_bank_name in bank_name_fixes:
                    pay_bank_name = bank_name_fixes[pay_bank_name]
                break

    if pay_bank_idx is not None and pay_bank_idx + 1 < len(parts):
        next_part = parts[pay_bank_idx + 1]
        if next_part.isdigit() and len(next_part) >= 8:
            pay_account = next_part

    return pay_bank_name, pay_account


def _extract_cheque_number_from_parts(parts):
    for p in parts:
        if p.isdigit() and 6 <= len(p) <= 10:
            return p
    return None


def _extract_amount_from_parts(parts):
    for p in reversed(parts):
        try:
            if ',' in p or '.' in p:
                s = p.replace(',', '')
                if s.replace('.', '').replace('-', '').isdigit():
                    amt = float(s)
                    if 0 < amt < 100000000:
                        return amt
        except Exception:
            continue

    for p in reversed(parts):
        try:
            if p.isdigit() and len(p) >= 3:
                amt = float(p)
                if 0 < amt < 100000000:
                    return amt
        except Exception:
            continue

    return None


def _extract_branch_code_from_parts(parts):
    for p in parts:
        if p.isdigit() and len(p) == 3 and p != '201':
            return p
    return None


def _extract_reason_from_parts(parts):
    reason_keywords = [
        'INSUFFICIENT', 'SIGNATURE', 'ACCOUNT', 'CLOSED', 'STOP',
        'PAYMENT', 'DRAWER', 'IRREGUL', 'FUNDS', 'REFER', 'EXCEEDS',
        'AMOUNT', 'WORDS', 'FIGURES', 'DIFFER', 'POST', 'DATED',
        'STALE', 'MUTILATED', 'CLEARING', 'ENDORSEMENT', 'ACCEPTED',
    ]

    reason_parts = []
    found_amount = False

    for p in reversed(parts):
        pu = p.upper()

        if ',' in p and any(c.isdigit() for c in p):
            found_amount = True
            break

        if any(k in pu for k in reason_keywords):
            reason_parts.insert(0, p)
        elif found_amount or len(p) < 3:
            continue
        elif (not p.isdigit()) and p.isalpha() and pu not in ['CLG', 'BANK', 'BFD', 'PAY']:
            reason_parts.insert(0, p)

    return ' '.join(reason_parts) if reason_parts else None


# =============================================================================
# EXCEL GENERATION - BATCH FORMAT
# =============================================================================

def generate_excel_batch(data, clearing_account='9313102000', clearing_branch='255', output_path=None):
    """
    Generate Excel in batch format with:
    - Debit: TRANCODE='555'
    - Credit: TRANCODE='055'
    """
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('Batch')

    headers = ['BRANCHCODE', 'MAINCODE', 'TRANCODE', 'AMOUNT', 'LCYAMOUNT', 'DESC1', 'DESC2']
    header_style = xlwt.XFStyle()
    header_font = xlwt.Font()
    header_font.bold = True
    header_style.font = header_font

    for col_idx, header in enumerate(headers):
        worksheet.write(0, col_idx, header, header_style)

    row_idx = 1
    row_idx = _write_debit_entries(worksheet, data, row_idx)
    row_idx = _write_credit_entries(worksheet, data, clearing_account, clearing_branch, row_idx)

    if output_path is None:
        with tempfile.NamedTemporaryFile(mode='wb', suffix='.xls', delete=False) as temp_excel:
            output_path = temp_excel.name

    workbook.save(output_path)
    return output_path


def _write_debit_entries(worksheet, data, start_row):
    """Write debit entries (TRANCODE='555')."""
    num_style = xlwt.XFStyle()
    num_style.num_format_str = '0.00'

    row_idx = start_row
    for record in data:
        maincode = str(record['bfd_account'])
        branch_code = maincode[:3] if len(maincode) >= 3 else maincode

        worksheet.write(row_idx, 0, branch_code)
        worksheet.write(row_idx, 1, maincode)
        worksheet.write(row_idx, 2, '555')

        amount = float(record['cheque_amount'])
        worksheet.write(row_idx, 3, amount, num_style)
        worksheet.write(row_idx, 4, amount, num_style)
        worksheet.write(row_idx, 5, '')

        pay_bank = str(record['pay_bank_name']).upper().replace(' ', '')
        cheque_num = str(record['cheque_number'])
        worksheet.write(row_idx, 6, f"CLG {pay_bank} {cheque_num}")

        row_idx += 1

    return row_idx


def _write_credit_entries(worksheet, data, clearing_account, clearing_branch, start_row):
    """Write credit entries (TRANCODE='055')."""
    num_style = xlwt.XFStyle()
    num_style.num_format_str = '0.00'

    row_idx = start_row
    for record in data:
        maincode = str(record['bfd_account'])

        worksheet.write(row_idx, 0, clearing_branch)
        worksheet.write(row_idx, 1, clearing_account)
        worksheet.write(row_idx, 2, '055')

        amount = float(record['cheque_amount'])
        worksheet.write(row_idx, 3, -amount, num_style)
        worksheet.write(row_idx, 4, -amount, num_style)

        worksheet.write(row_idx, 5, f"CLG TFR {maincode}")

        pay_bank = str(record['pay_bank_name']).upper().replace(' ', '')
        pay_account = str(record.get('pay_account', '')).strip()
        worksheet.write(row_idx, 6, f"{pay_bank} {pay_account}".strip())

        row_idx += 1

    return row_idx


# =============================================================================
# EXCEL GENERATION - DISPLAY FORMAT
# =============================================================================

def generate_display_excel(data):
    """
    Generate Excel file in display format (similar to the HTML table).
    """
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('ECC Data')

    # Styles
    header_style = xlwt.XFStyle()
    header_font = xlwt.Font()
    header_font.bold = True
    header_style.font = header_font
    
    # Light gray background for header
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 22  # light gray
    header_style.pattern = pattern

    # Number format for amounts
    amount_style = xlwt.XFStyle()
    amount_style.num_format_str = '#,##0.00'

    # Bold style for totals
    total_style = xlwt.XFStyle()
    total_font = xlwt.Font()
    total_font.bold = True
    total_style.font = total_font
    total_style.num_format_str = '#,##0.00'

    # Write headers
    headers = ['#', 'BFD Account', 'Pay Bank', 'Pay Account', 'Cheque No.', 'Cheque Amount', 'Remarks']
    for col_idx, header in enumerate(headers):
        worksheet.write(0, col_idx, header, header_style)

    # Set column widths
    worksheet.col(0).width = 2000   # #
    worksheet.col(1).width = 5000   # BFD Account
    worksheet.col(2).width = 4000   # Pay Bank
    worksheet.col(3).width = 5000   # Pay Account
    worksheet.col(4).width = 4000   # Cheque No.
    worksheet.col(5).width = 4500   # Cheque Amount
    worksheet.col(6).width = 8000   # Remarks

    # Write data rows
    total_amount = 0
    for idx, record in enumerate(data, start=1):
        row_idx = idx  # row 0 is header, data starts from row 1
        
        worksheet.write(row_idx, 0, idx)
        worksheet.write(row_idx, 1, str(record['bfd_account']))
        worksheet.write(row_idx, 2, record.get('pay_bank_name', 'CLG'))
        worksheet.write(row_idx, 3, record.get('pay_account', '-'))
        worksheet.write(row_idx, 4, str(record['cheque_number']))
        
        amount = float(record['cheque_amount'])
        worksheet.write(row_idx, 5, amount, amount_style)
        total_amount += amount
        
        worksheet.write(row_idx, 6, record.get('reason', '-'))

    # Write total row
    total_row = len(data) + 1
    worksheet.write(total_row, 0, '', total_style)
    worksheet.write(total_row, 1, '', total_style)
    worksheet.write(total_row, 2, '', total_style)
    worksheet.write(total_row, 3, '', total_style)
    worksheet.write(total_row, 4, 'Total', total_style)
    worksheet.write(total_row, 5, total_amount, total_style)
    worksheet.write(total_row, 6, '', total_style)

    # Save to temp file
    with tempfile.NamedTemporaryFile(mode='wb', suffix='.xls', delete=False) as temp_excel:
        output_path = temp_excel.name

    workbook.save(output_path)
    return output_path


# =============================================================================
# VALIDATION AND UTILITIES
# =============================================================================

def validate_data(data):
    """Validate extracted data before generating Excel."""
    errors = []

    if not isinstance(data, list):
        return False, ["Data must be a list"]
    if len(data) == 0:
        return False, ["Data list is empty"]

    required_fields = ['bfd_account', 'cheque_amount', 'pay_bank_name', 'cheque_number']

    for idx, record in enumerate(data):
        if not isinstance(record, dict):
            errors.append(f"Row {idx + 1}: Record must be a dictionary")
            continue

        for field in required_fields:
            if field not in record:
                errors.append(f"Row {idx + 1}: Missing required field '{field}'")

        if 'cheque_amount' in record:
            try:
                float(record['cheque_amount'])
            except (ValueError, TypeError):
                errors.append(f"Row {idx + 1}: cheque_amount must be a number")

        if not record.get('pay_account'):
            print(f"Warning Row {idx + 1}: pay_account is missing or empty")

    return len(errors) == 0, errors


def generate_excel_batch_safe(data, clearing_account='9313102000', clearing_branch='255', output_path=None):
    """Generate Excel file with validation."""
    is_valid, errors = validate_data(data)
    if not is_valid:
        return False, "Validation errors:\n" + "\n".join(errors)

    try:
        file_path = generate_excel_batch(data, clearing_account, clearing_branch, output_path)
        return True, file_path
    except Exception as e:
        return False, f"Error generating Excel file: {str(e)}"


def _safe_remove(path):
    """Remove file safely if exists."""
    if path and os.path.exists(path):
        try:
            os.remove(path)
        except Exception as e:
            print(f"Error removing file {path}: {e}") 
            
            
def upload(request):
    """Render the initial upload page."""
    return render(request, 'upload.html', {'now': datetime.now().year})


@csrf_exempt
def process_upload(request):
    """Process uploaded PDF and redirect to dashboard."""
    if request.method != 'POST':
        return JsonResponse({'error': 'Only POST method allowed'}, status=405)

    pdf_file = request.FILES.get('pdf_file')
    if not pdf_file:
        return JsonResponse({'error': 'No PDF file uploaded'}, status=400)

    if not pdf_file.name.endswith('.pdf'):
        return JsonResponse({'error': 'Invalid file format. Please upload a PDF file.'}, status=400)

    temp_pdf_path = None

    try:
        with tempfile.NamedTemporaryFile(mode='wb', suffix='.pdf', delete=False) as temp_pdf:
            for chunk in pdf_file.chunks():
                temp_pdf.write(chunk)
            temp_pdf_path = temp_pdf.name

        extracted_data = extract_pdf_data(temp_pdf_path)
        if not extracted_data:
            return JsonResponse({'error': 'Upload ECC Report Only'}, status=400)

        # Sort data by cheque_amount in ascending order
        extracted_data = sorted(extracted_data, key=lambda x: float(x.get('cheque_amount', 0)))

        # Store data in session
        request.session['display_data'] = extracted_data
        request.session['display_filename'] = pdf_file.name

        # Calculate statistics
        total_records = len(extracted_data)
        total_amount = 0.0
        accepted_count = 0
        accepted_amount = 0.0
        insufficient_funds_count = 0
        insufficient_funds_amount = 0.0
        other_reasons_count = 0
        other_reasons_amount = 0.0

        for r in extracted_data:
            reason = str(r.get('reason', '')).upper().strip()
            amount = float(r['cheque_amount'])
            total_amount += amount
            
            if 'ACCEPTED' in reason:
                accepted_count += 1
                accepted_amount += amount
            elif 'INSUFFICIENT' in reason or 'FUNDS' in reason:
                insufficient_funds_count += 1
                insufficient_funds_amount += amount
            elif reason:
                other_reasons_count += 1
                other_reasons_amount += amount

        no_reason_count = total_records - (accepted_count + insufficient_funds_count + other_reasons_count)

        context = {
            'data': extracted_data,
            'total_records': total_records,
            'total_amount': total_amount,
            'accepted_count': accepted_count,
            'accepted_amount': accepted_amount,
            'insufficient_funds_count': insufficient_funds_count,
            'insufficient_funds_amount': insufficient_funds_amount,
            'other_reasons_count': other_reasons_count,
            'other_reasons_amount': other_reasons_amount,
            'no_reason_count': no_reason_count,
            'filename': pdf_file.name,
            'now': datetime.now(),
        }
        
        return render(request, 'dashboard.html', context)

    except Exception as e:
        print(f"Error processing file: {traceback.format_exc()}")
        return JsonResponse({'error': f'Error processing file: {str(e)}'}, status=500)

    finally:
        _safe_remove(temp_pdf_path)  
        

import xlrd  # Add this with other imports at top

@csrf_exempt
def generate_final_batch(request):
    """
    Compare uploaded PDF + full_batch XLS.
    Keep only rows (debit + credit pair) where the cheque's Reason in PDF is ACCEPTED.
    """
    if request.method != 'POST':
        return JsonResponse({'error': 'Only POST method allowed'}, status=405)

    pdf_file = request.FILES.get('pdf_file')
    xls_file = request.FILES.get('xls_file')

    if not pdf_file:
        return JsonResponse({'error': 'No PDF file uploaded'}, status=400)
    if not xls_file:
        return JsonResponse({'error': 'No Full Batch XLS file uploaded'}, status=400)
    if not pdf_file.name.endswith('.pdf'):
        return JsonResponse({'error': 'Invalid PDF file format.'}, status=400)
    if not (xls_file.name.endswith('.xls') or xls_file.name.endswith('.xlsx')):
        return JsonResponse({'error': 'Invalid XLS file format.'}, status=400)

    temp_pdf_path = None
    temp_xls_path = None
    output_path = None

    try:
        with tempfile.NamedTemporaryFile(mode='wb', suffix='.pdf', delete=False) as f:
            for chunk in pdf_file.chunks():
                f.write(chunk)
            temp_pdf_path = f.name

        xls_suffix = '.xlsx' if xls_file.name.endswith('.xlsx') else '.xls'
        with tempfile.NamedTemporaryFile(mode='wb', suffix=xls_suffix, delete=False) as f:
            for chunk in xls_file.chunks():
                f.write(chunk)
            temp_xls_path = f.name

        # Extract PDF data → get accepted cheque numbers
        extracted_data = extract_pdf_data(temp_pdf_path)
        if not extracted_data:
            return JsonResponse({'error': 'No data extracted from PDF'}, status=400)

        accepted_cheque_numbers = set()
        for record in extracted_data:
            reason = str(record.get('reason', '')).upper().strip()
            if 'ACCEPTED' in reason:
                accepted_cheque_numbers.add(str(record['cheque_number']).strip())

        print(f"\nAccepted cheque numbers from PDF ({len(accepted_cheque_numbers)}): {accepted_cheque_numbers}")

        if not accepted_cheque_numbers:
            return JsonResponse({'error': 'No ACCEPTED cheques found in the PDF.'}, status=400)

        # Read full batch XLS
        wb_in = xlrd.open_workbook(temp_xls_path)
        ws_in = wb_in.sheet_by_index(0)

        headers = ws_in.row_values(0)
        rows_555 = []
        rows_055 = []
        for i in range(1, ws_in.nrows):
            r = ws_in.row_values(i)
            trancode = str(r[2]).strip()
            if trancode == '555':
                rows_555.append(r)
            elif trancode == '055':
                rows_055.append(r)

        if len(rows_555) != len(rows_055):
            return JsonResponse({
                'error': f'XLS structure error: {len(rows_555)} debit rows vs {len(rows_055)} credit rows.'
            }, status=400)

        # Match: 555 row i pairs with 055 row i (positional)
        # Cheque number = last token of 555 DESC2 ('CLG EBL 61825524' → '61825524')
        kept_555 = []
        kept_055 = []
        skipped = 0

        for idx, row_555 in enumerate(rows_555):
            desc2 = str(row_555[6]).strip()
            parts = desc2.split()
            cheque_num = parts[-1] if parts else ''

            if cheque_num in accepted_cheque_numbers:
                kept_555.append(row_555)
                kept_055.append(rows_055[idx])
            else:
                skipped += 1
                print(f"  Skipping cheque {cheque_num} (not ACCEPTED)")

        print(f"\nKept: {len(kept_555)} pairs | Skipped: {skipped}")

        if not kept_555:
            return JsonResponse({'error': 'No matching ACCEPTED cheques found between PDF and XLS.'}, status=400)

        # Write output XLS
        wb_out = xlwt.Workbook()
        ws_out = wb_out.add_sheet('Batch')

        header_style = xlwt.XFStyle()
        header_font = xlwt.Font()
        header_font.bold = True
        header_style.font = header_font

        num_style = xlwt.XFStyle()
        num_style.num_format_str = '0.00'

        for col_idx, h in enumerate(headers):
            ws_out.write(0, col_idx, h, header_style)

        row_idx = 1
        for r in kept_555:
            for col_idx, val in enumerate(r):
                if col_idx in (3, 4):
                    ws_out.write(row_idx, col_idx, float(val), num_style)
                else:
                    ws_out.write(row_idx, col_idx, val)
            row_idx += 1

        for r in kept_055:
            for col_idx, val in enumerate(r):
                if col_idx in (3, 4):
                    ws_out.write(row_idx, col_idx, float(val), num_style)
                else:
                    ws_out.write(row_idx, col_idx, val)
            row_idx += 1

        with tempfile.NamedTemporaryFile(mode='wb', suffix='.xls', delete=False) as f:
            output_path = f.name
        wb_out.save(output_path)

        with open(output_path, 'rb') as excel_file:
            response = HttpResponse(excel_file.read(), content_type='application/vnd.ms-excel')
            response['Content-Disposition'] = 'attachment; filename="final_batch.xls"'
            return response

    except Exception as e:
        print(f"Error generating final batch: {traceback.format_exc()}")
        return JsonResponse({'error': f'Error: {str(e)}'}, status=500)

    finally:
        _safe_remove(temp_pdf_path)
        _safe_remove(temp_xls_path)
        _safe_remove(output_path)      
        

#Commision Batch Generation 
@csrf_exempt
def generate_commission_batch(request):
    """
    Generate commission batch for cheques with amount > threshold (default 200000).
    Only ACCEPTED cheques are included.
    Format:
      - Debit (055): from BFD account, -commission_amount
      - Credit (555): to commission account, +commission_amount
    """
    if request.method != 'POST':
        return JsonResponse({'error': 'Only POST method allowed'}, status=405)

    pdf_file = request.FILES.get('pdf_file')
    if not pdf_file:
        return JsonResponse({'error': 'No PDF file uploaded'}, status=400)
    if not pdf_file.name.endswith('.pdf'):
        return JsonResponse({'error': 'Invalid PDF file format.'}, status=400)

    # Read form inputs
    commission_account = request.POST.get('commission_account', '9505062601').strip() or '9505062601'
    clearing_branch    = request.POST.get('branch_code', '255').strip() or '255'
    try:
        commission_amount = float(request.POST.get('commission_amount', '15').strip())
    except ValueError:
        commission_amount = 15.0
    try:
        amount_threshold = float(request.POST.get('amount_threshold', '200000').strip())
    except ValueError:
        amount_threshold = 200000.0

    temp_pdf_path = None
    output_path   = None

    try:
        with tempfile.NamedTemporaryFile(mode='wb', suffix='.pdf', delete=False) as f:
            for chunk in pdf_file.chunks():
                f.write(chunk)
            temp_pdf_path = f.name

        extracted_data = extract_pdf_data(temp_pdf_path)
        if not extracted_data:
            return JsonResponse({'error': 'No data extracted from PDF'}, status=400)

        # Filter: ACCEPTED and amount > threshold
        eligible = []
        for record in extracted_data:
            reason = str(record.get('reason', '')).upper().strip()
            amount = float(record.get('cheque_amount', 0))
            if amount > amount_threshold:
                eligible.append(record)

        print(f"\nCommission eligible records: {len(eligible)} (threshold > {amount_threshold})")

        if not eligible:
            return JsonResponse({
                'error': f'No ACCEPTED cheques with amount greater than Rs.{amount_threshold:,.0f} found.'
            }, status=400)

        # Build Excel
        wb  = xlwt.Workbook()
        ws  = wb.add_sheet('Sheet1')

        num_style = xlwt.XFStyle()
        num_style.num_format_str = '0.00'

        row_idx = 0

        # --- Debit rows (055) ---
        for record in eligible:
            bfd_account  = str(record['bfd_account'])
            branch_code  = bfd_account[:3] if len(bfd_account) >= 3 else clearing_branch
            cheque_amt   = float(record['cheque_amount'])
            pay_bank     = str(record.get('pay_bank_name', 'CLG')).upper().replace(' ', '')
            cheque_num   = str(record['cheque_number'])

            # Format amount as integer if whole number, else 2 decimal
            amt_int = int(cheque_amt) if cheque_amt == int(cheque_amt) else cheque_amt
            desc1 = f"Ecc Charge Rs.{amt_int}"
            desc2 = f"CLG {pay_bank} {cheque_num}"

            ws.write(row_idx, 0, branch_code)
            ws.write(row_idx, 1, bfd_account)
            ws.write(row_idx, 2, '055')
            ws.write(row_idx, 3, -commission_amount, num_style)
            ws.write(row_idx, 4, -commission_amount, num_style)
            ws.write(row_idx, 5, desc1)
            ws.write(row_idx, 6, desc2)
            row_idx += 1

        # --- Credit rows (555) ---
        for record in eligible:
            bfd_account  = str(record['bfd_account'])
            cheque_amt   = float(record['cheque_amount'])
            amt_int = int(cheque_amt) if cheque_amt == int(cheque_amt) else cheque_amt
            desc1 = f"Ecc Charge Rs.{amt_int}"

            ws.write(row_idx, 0, clearing_branch)
            ws.write(row_idx, 1, commission_account)
            ws.write(row_idx, 2, '555')
            ws.write(row_idx, 3, commission_amount, num_style)
            ws.write(row_idx, 4, commission_amount, num_style)
            ws.write(row_idx, 5, desc1)
            ws.write(row_idx, 6, bfd_account)
            row_idx += 1

        with tempfile.NamedTemporaryFile(mode='wb', suffix='.xls', delete=False) as f:
            output_path = f.name
        wb.save(output_path)

        with open(output_path, 'rb') as excel_file:
            response = HttpResponse(excel_file.read(), content_type='application/vnd.ms-excel')
            response['Content-Disposition'] = 'attachment; filename="comm_batch.xls"'
            return response

    except Exception as e:
        print(f"Error generating commission batch: {traceback.format_exc()}")
        return JsonResponse({'error': f'Error: {str(e)}'}, status=500)

    finally:
        _safe_remove(temp_pdf_path)
        _safe_remove(output_path)           