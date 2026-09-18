"""
FORMS - FAB OPEX Reconciliation & Monitoring System
Parser Engine: Dynamic, robust extraction for raw Oracle Financials / IBS OPEX reports (PDF/TXT).
Zero hardcoded accounts, dynamic unique account discovery, smart period detection, and accurate net saldo calculation.
"""

import re
import io
from typing import Dict, Any, Tuple, Optional, List
import pandas as pd
import pdfplumber


def clean_text(text: str) -> str:
    """Normalize whitespace and strip unnecessary characters."""
    if not text:
        return ""
    return re.sub(r'\s+', ' ', str(text)).strip()


def parse_saldo(value_str: Any) -> float:
    """
    Parse a numeric string into float with support for:
    - Minus sign before: -1,250,000.00 or -1.250.000,00
    - Minus sign after: 1,250,000.00- or 1.250.000,00-
    - Parentheses (negative): (1,250,000.00) or (1.250.000,00)
    - Credit/Debit suffixes: 1,250,000.00 CR
    """
    if value_str is None or value_str == "":
        return 0.0
    if isinstance(value_str, (int, float)):
        return float(value_str)
    
    val = str(value_str).strip()
    if not val:
        return 0.0

    is_negative = False

    # Check for minus at beginning or end
    if val.startswith('-'):
        is_negative = True
        val = val[1:].strip()
    elif val.endswith('-'):
        is_negative = True
        val = val[:-1].strip()

    # Check for parentheses: (1,000.00)
    if val.startswith('(') and val.endswith(')'):
        is_negative = True
        val = val[1:-1].strip()

    # Check for CR / DR indicators
    if val.upper().endswith('CR'):
        # In expense accounts, CR is a credit/reversal (negative)
        is_negative = True
        val = val[:-2].strip()
    elif val.upper().endswith('DR'):
        val = val[:-2].strip()

    # Strip currency symbols and spaces
    val = re.sub(r'[RpRP\s$]', '', val)

    # Detect comma vs dot decimal separators
    # Case A: 1.234.567,89 (Indonesian/European format)
    if ',' in val and '.' in val:
        if val.rfind(',') > val.rfind('.'):
            # Comma is decimal
            val = val.replace('.', '').replace(',', '.')
        else:
            # Dot is decimal
            val = val.replace(',', '')
    elif ',' in val and '.' not in val:
        # Check if comma is decimal (e.g. 1250,50) or thousands (1,250,000)
        parts = val.split(',')
        if len(parts) == 2 and len(parts[1]) in (1, 2):
            val = val.replace(',', '.')
        else:
            val = val.replace(',', '')
    elif '.' in val and ',' not in val:
        parts = val.split('.')
        # If multiple dots (e.g. 1.250.000), it's thousands separator
        if len(parts) > 2:
            val = val.replace('.', '')
        elif len(parts) == 2 and len(parts[1]) == 3:
            # Likely thousands separator without decimal e.g. 150.000
            val = val.replace('.', '')

    try:
        num = float(val)
        return -num if is_negative else num
    except ValueError:
        return 0.0


def extract_period(text: str, fallback_dates: Optional[List[str]] = None) -> str:
    """
    Automatically detects the report month/period from document header or transaction dates.
    Examples: 'Aug-26', 'AUG-2026', 'Periode YTD : Aug-26', 'Period: 08-2026'.
    """
    # Pattern 1: Explicit Period label (including YTD / MTD)
    # e.g., "Periode YTD : Aug-26", "Period: Aug-26", "PERIODE : AGUSTUS 2026"
    m = re.search(r'(?:PERIOD(?:E)?(?:\s+YTD|\s+MTD)?(?:\s+NAME)?)\s*[:=]\s*([A-Za-z]{3,9}[-\s_]\d{2,4})', text, re.IGNORECASE)
    if m:
        return clean_text(m.group(1))

    # Pattern 2: Date range in header
    # e.g., "01-AUG-2026 TO 31-AUG-2026" or "01/08/2026 - 31/08/2026"
    m2 = re.search(r'(?:TO|-)\s*\d{1,2}[-/ ]([A-Za-z]{3,9}|\d{1,2})[-/ ](\d{2,4})', text, re.IGNORECASE)
    if m2:
        month_part = m2.group(1)
        year_part = m2.group(2)
        if len(year_part) == 4:
            year_part = year_part[-2:]
        return f"{month_part.capitalize()}-{year_part}"

    # Pattern 3: Standalone "Month-YY" header pattern
    m3 = re.search(r'\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-/ ](20\d{2}|\d{2})\b', text, re.IGNORECASE)
    if m3:
        month = m3.group(1).capitalize()
        year = m3.group(2)
        if len(year) == 4:
            year = year[-2:]
        return f"{month}-{year}"

    # Pattern 4: Derivation from transaction dates if available
    if fallback_dates:
        for d in fallback_dates:
            # Try 01-Aug-2026 or 2026-08-01 or 01/08/2026
            dm = re.search(r'(\d{1,2})[-/]([A-Za-z]{3})[-/](\d{2,4})', str(d))
            if dm:
                yr = dm.group(3)
                if len(yr) == 4:
                    yr = yr[-2:]
                return f"{dm.group(2).capitalize()}-{yr}"
            
            dm2 = re.search(r'(\d{4})[-/](\d{1,2})[-/]\d{1,2}', str(d))
            if dm2:
                month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                mo_idx = int(dm2.group(2)) - 1
                if 0 <= mo_idx < 12:
                    return f"{month_names[mo_idx]}-{dm2.group(1)[-2:]}"

    return "Current-Period"


def parse_raw_text_content(full_text: str) -> Tuple[List[Dict[str, Any]], str]:
    """
    Parses unstructured / semi-structured text lines from Oracle Financials / IBS reports.
    Identifies rows containing transaction date, account code, account name, saldo, and description.
    """
    records: List[Dict[str, Any]] = []
    lines = full_text.splitlines()

    # State variables for hierarchical/contextual parsing
    current_batch = "OPEX-BATCH"
    current_code_account = ""
    current_account_name = ""
    
    # Regex building blocks
    date_pattern = r'(\d{1,2}[-/](?:[A-Za-z]{3}|\d{1,2})[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})'
    # Account code typically 4 to 15 alphanumeric or segmented with dots/dashes e.g. 61010101 or 01-100-61010
    account_code_pattern = r'(\b(?:6\d{5,9}|[0-9]{2,4}[-.]\d{2,4}[-.]\d{2,4}(?:[-.]\d{2,4})?|\d{6,10})\b)'
    # Saldo amount: -?\(?\d+...\)?-?
    amount_pattern = r'(-?\(?\s*\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?\s*\)?-?)'

    # Check for Pipe-separated format (| Batch | Date | Account Code | ...)
    has_pipes = any(l.count('|') >= 4 for l in lines[:100])

    if has_pipes:
        col_indices = {}
        for line in lines:
            line_str = line.strip()
            if not line_str or line_str.startswith('---') or line_str.startswith('==='):
                continue
            
            parts = [p.strip() for p in line_str.split('|')]
            # Discard leading/trailing empty elements from outer pipes
            if parts and parts[0] == '':
                parts = parts[1:]
            if parts and parts[-1] == '':
                parts = parts[:-1]
            
            # Check if this line is a header row (e.g. contains 'Batch Name', 'Trans.Date', 'Seg4' or 'Saldo')
            line_lower = " ".join([p.lower() for p in parts])
            if ('batch' in line_lower and 'saldo' in line_lower) or ('trans' in line_lower and 'account' in line_lower):
                # Dynamically map headers to column indices
                col_indices = {}
                for idx, col_name in enumerate(parts):
                    cn = col_name.lower()
                    if 'batch' in cn:
                        col_indices['batch'] = idx
                    elif 'trans' in cn or 'date' in cn or 'tanggal' in cn:
                        col_indices['date'] = idx
                    elif 'seg4' in cn or 'code' in cn or 'coa' in cn or 'kode' in cn:
                        col_indices['code'] = idx
                    elif 'account name' in cn or 'nama akun' in cn or ('account' in cn and 'code' not in cn):
                        col_indices['name'] = idx
                    elif 'saldo' in cn or 'amount' in cn or 'nominal' in cn:
                        col_indices['saldo'] = idx
                    elif 'description' in cn or 'keterangan' in cn or 'uraian' in cn:
                        col_indices['desc'] = idx
                continue

            # Check if this is a summary / total line e.g. "| |TOTAL : | 652,845,928 | |"
            if 'total' in line_lower and any(c.isdigit() for c in line_lower):
                continue

            if len(parts) >= 5:
                batch_val = ""
                date_val = ""
                code_val = ""
                name_val = ""
                saldo_val = 0.0
                desc_val = ""

                # If dynamic header indices were found, use them
                if col_indices and 'code' in col_indices and 'saldo' in col_indices:
                    code_idx = col_indices['code']
                    saldo_idx = col_indices['saldo']
                    
                    if code_idx < len(parts) and saldo_idx < len(parts):
                        code_val = parts[code_idx]
                        saldo_val = parse_saldo(parts[saldo_idx])
                        
                        b_idx = col_indices.get('batch')
                        if b_idx is not None and b_idx < len(parts):
                            batch_val = parts[b_idx]
                            
                        d_idx = col_indices.get('date')
                        if d_idx is not None and d_idx < len(parts):
                            date_val = parts[d_idx]
                            
                        n_idx = col_indices.get('name')
                        if n_idx is not None and n_idx < len(parts):
                            name_val = parts[n_idx]
                            
                        desc_idx = col_indices.get('desc')
                        if desc_idx is not None and desc_idx < len(parts):
                            desc_val = " | ".join(parts[desc_idx:])
                elif len(parts) >= 11:
                    # Specific standard 11-column Oracle Financials layout:
                    # [0:Batch, 1:Date, 2:COY, 3:Seg3, 4:Dept, 5:Seg4(Code), 6:Name, 7:Seg5, 8:Project, 9:Saldo, 10:Desc]
                    batch_val = parts[0]
                    date_val = parts[1]
                    code_val = parts[5]
                    name_val = parts[6]
                    saldo_val = parse_saldo(parts[9])
                    desc_val = parts[10]
                elif len(parts) >= 6:
                    batch_val = parts[0]
                    date_val = parts[1]
                    code_val = parts[2]
                    name_val = parts[3]
                    saldo_val = parse_saldo(parts[4])
                    desc_val = parts[5]
                elif len(parts) == 5:
                    date_val = parts[0]
                    code_val = parts[1]
                    name_val = parts[2]
                    saldo_val = parse_saldo(parts[3])
                    desc_val = parts[4]
                    batch_val = current_batch

                # Only add if code_val looks like an account code and not empty
                if code_val and (saldo_val != 0.0 or any(char.isdigit() for char in str(parts[saldo_idx if 'saldo' in col_indices else 4]))):
                    records.append({
                        "Batch Name": clean_text(batch_val) or "OPEX-GEN",
                        "Trans. Date": clean_text(date_val),
                        "Code Account": clean_text(code_val),
                        "Account Name": clean_text(name_val),
                        "Saldo": saldo_val,
                        "Description": clean_text(desc_val)
                    })
        if records:
            extracted_dates = [r["Trans. Date"] for r in records if r["Trans. Date"]]
            period = extract_period(full_text, extracted_dates)
            return records, period

    # Strategy 2: Line-by-line regex & context parsing
    # Oracle Financials reports frequently print an Account Header section, then multiple journal lines:
    # Example:
    # "Account: 61010101 - Beban Gaji & Tunjangan"
    # Followed by lines: "JE_0801  01-AUG-2026   Pembayaran Gaji Bulanan   45,000,000.00"
    # OR all in single rows: "JE_0801 01-AUG-2026 61010101 Beban Gaji 45,000,000.00 Pembayaran Gaji"
    for line in lines:
        raw_line = line.strip()
        if not raw_line:
            continue

        # Skip page headers / footers / separator lines
        if re.match(r'^[=\-_*]{4,}$', raw_line):
            continue
        if re.search(r'Page\s+\d+\s+of|\bTotal\s+Halaman\b|\bReport\s+ID\b', raw_line, re.IGNORECASE):
            continue

        # Check for Batch context
        batch_match = re.search(r'\b(?:Batch(?:\s+Name)?|JE_HEADER|VOUCHER)\s*[:=]\s*([A-Za-z0-9_\-]+)', raw_line, re.IGNORECASE)
        if batch_match:
            current_batch = batch_match.group(1).strip()

        # Check for Account Header context
        # e.g., "Account: 61010101 - Beban Gaji" or "Akun: 61010101 Beban Operasional"
        acc_header_match = re.search(r'(?:Account|Akun|Code\s*Account|Kode\s*Akun)\s*[:=]?\s*' + account_code_pattern + r'\s*[-:\s]\s*([A-Za-z0-9\s&,/.\-()]+)', raw_line, re.IGNORECASE)
        if acc_header_match:
            current_code_account = acc_header_match.group(1).strip()
            current_account_name = acc_header_match.group(2).strip()
            continue

        # Look for a standard 6-column single line transaction
        # Standard: Batch / Date / Code / Name / Saldo / Desc
        date_m = re.search(date_pattern, raw_line)
        if date_m:
            trans_date = date_m.group(1)
            
            # Find numbers with currency / decimals
            amounts = re.findall(r'(?:^|\s)' + amount_pattern + r'(?:\s|$)', raw_line)
            # Find account code if in line
            acc_m = re.search(account_code_pattern, raw_line)

            line_code = acc_m.group(1).strip() if acc_m else current_code_account
            
            # If we found an account code and an amount:
            if line_code and amounts:
                # Last or prominent amount is saldo
                target_amount_str = amounts[-1].strip()
                saldo = parse_saldo(target_amount_str)

                # Extract line batch name or default
                line_batch = current_batch
                batch_cand = re.search(r'\b([A-Za-z0-9_]{3,15}[-_][A-Za-z0-9_]+)\b', raw_line)
                if batch_cand and batch_cand.group(1) != line_code and batch_cand.group(1) != trans_date:
                    line_batch = batch_cand.group(1)

                # Extract description & account name
                line_name = current_account_name
                desc = ""

                # If account name was not in header, try extracting from between code and amount
                if not line_name and acc_m:
                    after_code = raw_line[acc_m.end():].strip()
                    # Remove the amount from after_code
                    text_rem = after_code.replace(target_amount_str, '').strip()
                    # If there's text left, split into Name and Description or use as Name
                    parts = re.split(r'\s{2,}|\t', text_rem)
                    if len(parts) >= 2:
                        line_name = parts[0].strip()
                        desc = parts[1].strip()
                    elif len(parts) == 1:
                        line_name = parts[0].strip()
                else:
                    # Description is whatever remains
                    rem = raw_line.replace(trans_date, '').replace(line_code, '').replace(target_amount_str, '')
                    if line_batch != current_batch:
                        rem = rem.replace(line_batch, '')
                    desc = clean_text(rem)

                if not line_name:
                    line_name = f"Akun {line_code}"

                records.append({
                    "Batch Name": line_batch,
                    "Trans. Date": trans_date,
                    "Code Account": line_code,
                    "Account Name": line_name,
                    "Saldo": saldo,
                    "Description": desc or "Realisasi Beban OPEX"
                })

    extracted_dates = [r["Trans. Date"] for r in records if r["Trans. Date"]]
    period = extract_period(full_text, extracted_dates)
    return records, period


def parse_pdf_report(file_bytes_or_path: Any) -> Tuple[List[Dict[str, Any]], str]:
    """
    Parses an OPEX report PDF using pdfplumber, handling both vector tables and raw text layouts.
    """
    records: List[Dict[str, Any]] = []
    full_text_list: List[str] = []

    # If bytes or buffer, ensure it's a file-like object with seek
    if isinstance(file_bytes_or_path, bytes):
        file_bytes_or_path = io.BytesIO(file_bytes_or_path)

    with pdfplumber.open(file_bytes_or_path) as pdf:
        # Check if pages contain structured tables
        for page_idx, page in enumerate(pdf.pages):
            page_text = page.extract_text(layout=True) or page.extract_text() or ""
            full_text_list.append(page_text)

            # Try table extraction first
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    
                    # Detect header
                    header_row_idx = -1
                    for idx, row in enumerate(table[:3]):
                        row_str = " ".join([str(c or '').lower() for c in row])
                        if any(k in row_str for k in ['batch', 'date', 'code', 'account', 'saldo', 'debet', 'kredit', 'amount']):
                            header_row_idx = idx
                            break
                    
                    if header_row_idx != -1:
                        headers = [str(c or '').strip() for c in table[header_row_idx]]
                        for row in table[header_row_idx + 1:]:
                            if not row or all(c is None or str(c).strip() == '' for c in row):
                                continue
                            
                            row_dict = {}
                            for h_idx, col in enumerate(row):
                                if h_idx < len(headers):
                                    row_dict[headers[h_idx]] = str(col or '').strip()
                            
                            # Map columns dynamically
                            batch_val = ""
                            date_val = ""
                            code_val = ""
                            name_val = ""
                            saldo_val = 0.0
                            desc_val = ""

                            for k, v in row_dict.items():
                                k_low = k.lower()
                                if 'batch' in k_low:
                                    batch_val = v
                                elif 'date' in k_low or 'tanggal' in k_low:
                                    date_val = v
                                elif 'code' in k_low or 'kode' in k_low or 'coa' in k_low or 'account code' in k_low:
                                    code_val = v
                                elif 'name' in k_low or 'nama' in k_low or 'deskripsi akun' in k_low:
                                    name_val = v
                                elif 'saldo' in k_low or 'amount' in k_low or 'nominal' in k_low or 'debet' in k_low:
                                    saldo_val = parse_saldo(v)
                                elif 'desc' in k_low or 'keterangan' in k_low or 'uraian' in k_low:
                                    desc_val = v

                            # Fallback if positional
                            if not code_val and len(row) >= 5:
                                # Standard order: [Batch, Date, Code, Name, Saldo, Desc]
                                batch_val = str(row[0] or '').strip()
                                date_val = str(row[1] or '').strip()
                                code_val = str(row[2] or '').strip()
                                name_val = str(row[3] or '').strip()
                                saldo_val = parse_saldo(row[4])
                                desc_val = str(row[5] or '').strip() if len(row) > 5 else ""

                            if code_val and (saldo_val != 0.0 or any(char.isdigit() for char in str(row))):
                                records.append({
                                    "Batch Name": batch_val or "OPEX-SYS",
                                    "Trans. Date": date_val,
                                    "Code Account": code_val,
                                    "Account Name": name_val or f"Akun {code_val}",
                                    "Saldo": saldo_val,
                                    "Description": desc_val or "Realisasi Beban OPEX"
                                })

    full_document_text = "\n".join(full_text_list)
    
    # If the document has pipe delimiters (typical for Oracle Financials FNDWRR reports),
    # text-based pipe parsing is significantly more accurate and comprehensive than vector tables
    if any(line.count('|') >= 4 for line in full_document_text.splitlines()[:50]):
        text_records, period = parse_raw_text_content(full_document_text)
        if text_records:
            return text_records, period

    # If table extraction yielded records, return them with detected period
    if records:
        extracted_dates = [r["Trans. Date"] for r in records if r["Trans. Date"]]
        period = extract_period(full_document_text, extracted_dates)
        return records, period

    # Otherwise, fallback to full text parser engine
    return parse_raw_text_content(full_document_text)


def process_opex_file(file_content: Any, file_name: str) -> Dict[str, Any]:
    """
    Master pipeline entry point.
    Receives file bytes / buffer and file name, processes it dynamically, and generates:
    - df_detail (Batch Name, Trans. Date, Code Account, Account Name, Saldo, Description)
    - df_rekap (Code Account, Account Name, Total Saldo)
    - Key metrics (total saldo, unique accounts, total transactions, reversals)
    - Top 5 spending accounts
    """
    file_ext = file_name.split('.')[-1].lower() if '.' in file_name else ""
    
    records: List[Dict[str, Any]] = []
    period = "Aug-26"

    # Handle PDF
    if file_ext == 'pdf':
        records, period = parse_pdf_report(file_content)
    else:
        # Handle TXT, PRN, REP, CSV, or raw text streams
        if isinstance(file_content, bytes):
            # Try utf-8 then latin-1
            try:
                text = file_content.decode('utf-8')
            except UnicodeDecodeError:
                text = file_content.decode('latin-1', errors='ignore')
        elif hasattr(file_content, 'read'):
            raw_bytes = file_content.read()
            try:
                text = raw_bytes.decode('utf-8')
            except UnicodeDecodeError:
                text = raw_bytes.decode('latin-1', errors='ignore')
        else:
            text = str(file_content)

        records, period = parse_raw_text_content(text)

    # Standardize column structure
    required_cols = ["Batch Name", "Trans. Date", "Code Account", "Account Name", "Saldo", "Description"]
    
    if not records:
        df_detail = pd.DataFrame(columns=required_cols)
        df_rekap = pd.DataFrame(columns=["Code Account", "Account Name", "Total Saldo"])
        return {
            "success": False,
            "period": period,
            "df_detail": df_detail,
            "df_rekap": df_rekap,
            "metrics": {
                "total_saldo": 0.0,
                "total_accounts": 0,
                "total_transactions": 0,
                "reversal_count": 0,
                "reversal_amount": 0.0
            },
            "top_5": pd.DataFrame(columns=["Account Name", "Total Saldo"]),
            "error": "Tidak ada baris transaksi OPEX yang valid ditemukan pada dokumen."
        }

    df_detail = pd.DataFrame(records)
    # Ensure all 6 required columns exist
    for c in required_cols:
        if c not in df_detail.columns:
            df_detail[c] = ""

    # Reorder exactly as required by specification
    df_detail = df_detail[required_cols]

    # Clean Code Account & Account Name
    df_detail["Code Account"] = df_detail["Code Account"].astype(str).str.strip()
    df_detail["Account Name"] = df_detail["Account Name"].astype(str).str.strip()
    df_detail["Saldo"] = pd.to_numeric(df_detail["Saldo"], errors="coerce").fillna(0.0)

    # Generate Summary Rekapitulasi (100% Dynamic - grouped by unique Code Account & Account Name)
    df_rekap = df_detail.groupby(["Code Account", "Account Name"], as_index=False)["Saldo"].sum()
    df_rekap.rename(columns={"Saldo": "Total Saldo"}, inplace=True)
    df_rekap.sort_values(by="Total Saldo", ascending=False, inplace=True)
    df_rekap.reset_index(drop=True, inplace=True)

    # Calculate metrics
    total_saldo = float(df_detail["Saldo"].sum())
    total_accounts = int(df_rekap["Code Account"].nunique())
    total_transactions = len(df_detail)
    
    # Reversals / minus transactions
    minus_txs = df_detail[df_detail["Saldo"] < 0]
    reversal_count = len(minus_txs)
    reversal_amount = float(minus_txs["Saldo"].sum())

    # Top 5 OPEX accounts
    top_5 = df_rekap.head(5).copy()

    return {
        "success": True,
        "period": period,
        "df_detail": df_detail,
        "df_rekap": df_rekap,
        "metrics": {
            "total_saldo": total_saldo,
            "total_accounts": total_accounts,
            "total_transactions": total_transactions,
            "reversal_count": reversal_count,
            "reversal_amount": reversal_amount
        },
        "top_5": top_5
    }
