"""
Patient Admission Service Module
Handles all patient admission billing logic and Google Sheets operations
"""

from typing import List, Dict, Any, Optional
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import calendar
import gspread
from google.oauth2.service_account import Credentials
import os
from fastapi import HTTPException
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configuration
CREDENTIALS_FILE = os.getenv("CREDENTIALS_FILE", "google_credentials.json")
if not os.path.exists(CREDENTIALS_FILE):
    for candidate in [
        "google_credentials.json",
        os.path.join("backend", "google_credentials.json"),
        os.path.join(os.path.dirname(__file__), "google_credentials.json"),
    ]:
        if os.path.exists(candidate):
            CREDENTIALS_FILE = candidate
            break
PATIENTADMISSION_SHEET_ID = os.getenv("PATIENT_ADMISSION_SHEET_ID")
CRM_ADMISSION_SHEET_ID = os.getenv("PATIENT_ADMISSION_SHEET_ID")
ADMISSION_CREDENTIALS_FILE = "CRM-admission.json"


def get_google_sheet_client(credentials_file: str = None):
    """Get authenticated gspread client"""
    if credentials_file is None:
        credentials_file = CREDENTIALS_FILE
    if not os.path.exists(credentials_file):
        raise HTTPException(status_code=404, detail=f"Credentials file not found: {credentials_file}")
    
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]
    creds = Credentials.from_service_account_file(credentials_file, scopes=scope)
    client = gspread.authorize(creds)
    return client


def get_patientadmission_sheet():
    """Get SNF worksheet for patient admission data"""
    client = get_google_sheet_client()
    if not PATIENTADMISSION_SHEET_ID:
        raise HTTPException(status_code=500, detail="PATIENT_ADMISSION_SHEET_ID not configured in .env")
    spreadsheet = client.open_by_key(PATIENTADMISSION_SHEET_ID)
    worksheet = spreadsheet.worksheet("SNF")  # Changed from CRM_PatientAdmission to SNF
    return worksheet


def get_accounts_receivable_sheet():
    """Get CRM_Admission → Invoice Table for invoice storage"""
    # Try to use separate credentials for admission sheet if available
    creds_file = ADMISSION_CREDENTIALS_FILE if os.path.exists(ADMISSION_CREDENTIALS_FILE) else CREDENTIALS_FILE
    client = get_google_sheet_client(creds_file)
    
    if not CRM_ADMISSION_SHEET_ID:
        raise HTTPException(status_code=500, detail="CRM_Admission Sheet ID not configured")
    spreadsheet = client.open_by_key(CRM_ADMISSION_SHEET_ID)
    worksheet = spreadsheet.worksheet("Invoice Table")
    return worksheet


def parse_date(date_str: str) -> Optional[datetime]:
    """
    Parse date string in various formats (DD/MM/YYYY, DD-MM-YYYY, YYYY-MM-DD)
    Returns datetime object or None if parsing fails
    """
    if not date_str or not str(date_str).strip():
        return None
    
    date_str = str(date_str).strip()
    
    # Try different date formats
    formats = [
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d/%m/%y",
        "%d-%m-%y",
    ]
    
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    
    return None


def format_date(dt: datetime, format_str: str = "%d/%m/%Y") -> str:
    """Format datetime object to string"""
    return dt.strftime(format_str)


def calculate_next_billing_date(service_start_date: datetime, reference_date: datetime = None) -> datetime:
    """
    Calculate next billing date based on service start date.
    Handles month-end edge cases (28/29/30/31 days).
    
    Args:
        service_start_date: Date when service started
        reference_date: Current date (defaults to today)
    
    Returns:
        Next billing date
    
    Example:
        Service started: 22/12/2025
        Reference date: 22/12/2025
        Next billing: 22/01/2026
        
        Service started: 31/01/2025
        Reference date: 31/01/2025
        Next billing: 28/02/2025 (or 29 in leap year)
    """
    if reference_date is None:
        reference_date = datetime.now()
    
    # Get the day of month from service start date
    billing_day = service_start_date.day
    
    # Calculate next month
    next_month_date = reference_date + relativedelta(months=1)
    
    # Get the last day of the target month
    last_day_of_month = calendar.monthrange(next_month_date.year, next_month_date.month)[1]
    
    # Use MIN(original day, last day of target month)
    actual_billing_day = min(billing_day, last_day_of_month)
    
    # Create the next billing date
    next_billing = datetime(next_month_date.year, next_month_date.month, actual_billing_day)
    
    return next_billing


def is_billing_due_today(service_start_date: datetime, last_billed_date: Optional[datetime] = None) -> bool:
    """
    Check if billing is due today for a client.
    
    Args:
        service_start_date: Date when service started
        last_billed_date: Date of last billing (None if never billed)
    
    Returns:
        True if billing is due today, False otherwise
    """
    today = datetime.now().date()
    
    # If never billed, check if we've passed the first billing cycle
    if last_billed_date is None:
        # First billing is one month after service start
        first_billing = calculate_next_billing_date(service_start_date, service_start_date)
        return first_billing.date() == today
    
    # Calculate next billing date from last billed date
    next_billing = calculate_next_billing_date(service_start_date, last_billed_date)
    return next_billing.date() == today


def calculate_next_future_billing_date(service_start_date: datetime, last_billed_date: datetime = None) -> datetime:
    """
    Calculate the next billing date that is TODAY or in the FUTURE.
    This handles cases where service started years ago and we need to find the next upcoming billing.
    
    Args:
        service_start_date: Date when service started
        last_billed_date: Date of last billing (None if never billed)
    
    Returns:
        Next billing date that is today or in the future
    """
    today = datetime.now().date()
    
    # Start from last billed date or service start date
    reference_date = last_billed_date if last_billed_date else service_start_date
    
    # Calculate next billing
    next_billing = calculate_next_billing_date(service_start_date, reference_date)
    
    # If next billing is in the past, fast-forward to find the next future billing
    max_iterations = 100  # Safety limit to prevent infinite loops
    iterations = 0
    
    while next_billing.date() < today and iterations < max_iterations:
        # Move forward one month
        next_billing = calculate_next_billing_date(service_start_date, next_billing)
        iterations += 1
    
    if iterations >= max_iterations:
        print(f"[Patient Admission] Warning: Hit max iterations calculating future billing for service started {service_start_date}")
    
    return next_billing


def calculate_patientadmission_revenue(home_care_revenue: float, additional_nursing: float, discount: float) -> float:
    """
    Calculate total revenue for patient admission service.
    Formula: (Patient Admission Revenue + Additional Nursing Charges) - Discount
    
    Args:
        home_care_revenue: Base patient admission revenue
        additional_nursing: Additional nursing charges
        discount: Discount amount
    
    Returns:
        Total revenue
    """
    try:
        revenue = float(home_care_revenue or 0)
        nursing = float(additional_nursing or 0)
        disc = float(discount or 0)
        
        total = (revenue + nursing) - disc
        return max(0, total)  # Ensure non-negative
    except (ValueError, TypeError):
        return 0.0



def get_all_patientadmission_clients() -> List[Dict[str, Any]]:
    """
    Get all patient admission clients from CRM_PatientAdmission sheet.
    Returns ALL clients (both ACTIVE and INACTIVE) - filtering happens at API level.
    
    Returns:
        List of client records
    """
    if not PATIENTADMISSION_SHEET_ID:
        print("[Patient Admission] PATIENT_ADMISSION_SHEET_ID not configured - returning empty list")
        return []
    try:
        worksheet = get_patientadmission_sheet()
        all_values = worksheet.get_all_values()
        
        if not all_values or len(all_values) < 2:
            print("[Patient Admission] No clients found in CRM_PatientAdmission sheet")
            return []
        
        headers = all_values[0]
        rows = all_values[1:]
        
        # DEBUG: Print headers and first row
        print(f"[Patient Admission DEBUG] Headers: {headers}")
        if rows:
            print(f"[Patient Admission DEBUG] First row: {rows[0]}")
        
        # Create header map for case-insensitive lookup
        header_map = {header.strip(): idx for idx, header in enumerate(headers)}
        
        clients = []
        for row_idx, row in enumerate(rows, start=2):  # Start from row 2 for _row_number
            if not any(row):  # Skip empty rows
                continue
            
            # Build client record
            client = {}
            for header, idx in header_map.items():
                if idx < len(row):
                    client[header] = row[idx]
                else:
                    client[header] = ""
            
            # Add row number for updates
            client['_row_number'] = row_idx
            
            # Only include clients with a patient name (check both possible column names)
            patient_name = client.get("Patient Name", "").strip() or client.get("PATIENT NAME", "").strip()
            if patient_name:
                clients.append(client)
        
        print(f"[Patient Admission] Retrieved {len(clients)} total clients from sheet")
        return clients
        
    except Exception as e:
        print(f"[Patient Admission] Error fetching clients: {e}")
        raise




def get_patientadmission_client_by_id(patient_name: str) -> Optional[Dict[str, Any]]:
    """
    Get specific patient admission client by patient name.
    
    Args:
        patient_name: Patient name to search for
    
    Returns:
        Client record or None if not found
    """
    try:
        worksheet = get_patientadmission_sheet()
        all_values = worksheet.get_all_values()
        
        if not all_values or len(all_values) < 2:
            return None
        
        headers = all_values[0]
        header_map = {}
        for idx, header in enumerate(headers):
            if header and header not in header_map:
                header_map[header] = idx
        
        for row_idx, row in enumerate(all_values[1:], start=2):
            if not any(row):
                continue
            
            record = {}
            for header, idx in header_map.items():
                record[header] = row[idx] if idx < len(row) else ""
            
            record['_row_number'] = row_idx
            
            # Check if patient name matches (try both column name variations)
            patient_name_value = record.get("Patient Name", "") or record.get("PATIENT NAME", "")
            if str(patient_name_value).strip().lower() == patient_name.lower():
                return record
        
        return None
        
    except Exception as e:
        print(f"Error fetching patient admission client: {e}")
        return None


def get_billing_history(patient_name: str) -> List[Dict[str, Any]]:
    """
    Get billing history for a patient admission client from Invoice Table sheet.
    
    Args:
        patient_name: Patient name
    
    Returns:
        List of invoice records
    """
    try:
        worksheet = get_accounts_receivable_sheet()
        all_values = worksheet.get_all_values()
        
        if not all_values or len(all_values) < 2:
            return []
        
        headers = all_values[0]
        header_map = {}
        for idx, header in enumerate(headers):
            if header and header not in header_map:
                header_map[header] = idx
        
        results = []
        for row in all_values[1:]:
            if not any(row):
                continue
            
            record = {}
            for header, idx in header_map.items():
                record[header] = row[idx] if idx < len(row) else ""
            
            # Check if this invoice is for the patient and is patient admission
            record_patient = str(record.get("Patient Name", "")).strip().lower()
            service_type = str(record.get("Service Type", "")).strip().lower()
            
            if record_patient == patient_name.lower() and "patient admission" in service_type:
                results.append({
                    "invoice_ref": record.get("Invoice Ref", ""),
                    "invoice_date": record.get("Invoice Date", ""),
                    "amount": record.get("Total Amount", 0),
                    "status": record.get("Status", ""),
                })
        
        # Sort by date (newest first)
        results.sort(key=lambda x: x.get("invoice_date", ""), reverse=True)
        
        return results
        
    except Exception as e:
        print(f"Error fetching billing history: {e}")
        return []


def generate_invoice_ref() -> str:
    """Generate unique invoice reference number"""
    try:
        worksheet = get_accounts_receivable_sheet()
        all_values = worksheet.get_all_values()
        
        if not all_values or len(all_values) < 2:
            return "INV000001"
        
        # Get headers
        headers = all_values[0]
        inv_ref_idx = next((i for i, h in enumerate(headers) if h == "Invoice Ref"), None)
        
        if inv_ref_idx is None:
            return "INV000001"
        
        # Find the highest invoice number
        max_num = 0
        for row in all_values[1:]:
            if inv_ref_idx < len(row):
                inv_ref = str(row[inv_ref_idx])
                if inv_ref.startswith("INV"):
                    try:
                        num = int(inv_ref.replace("INV", ""))
                        max_num = max(max_num, num)
                    except:
                        pass
        
        # Generate new invoice number
        new_num = max_num + 1
        return f"INV{new_num:06d}"  # INV000001, INV000002, etc.
        
    except Exception as e:
        # Fallback to timestamp-based
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        return f"INV{timestamp}"


def generate_patientadmission_invoice(client_record: Dict[str, Any]) -> Dict[str, Any]:
    """
    Generate invoice for a patient admission client and save to Invoice Table sheet.
    
    Args:
        client_record: Home care client record from CRM_PatientAdmission sheet
    
    Returns:
        Invoice details
    """
    try:
        patient_name = client_record.get("PATIENT NAME", "")
        
        # CRITICAL: Check for duplicate invoices today
        # Prevent generating multiple invoices on the same day
        today = datetime.now().date()
        billing_history = get_billing_history(patient_name)
        
        for invoice in billing_history:
            invoice_date_str = invoice.get("invoice_date", "")
            # Parse invoice date (format: DD-MM-YYYY HH:MM)
            try:
                if invoice_date_str:
                    invoice_date_parts = invoice_date_str.split()[0]  # Get date part only
                    invoice_date = datetime.strptime(invoice_date_parts, "%d-%m-%Y").date()
                    if invoice_date == today:
                        print(f"[Patient Admission Billing] Duplicate invoice prevented for {patient_name} - Invoice already exists today")
                        return {
                            "invoice_ref": invoice.get("invoice_ref", ""),
                            "invoice_date": invoice_date_str,
                            "patient_name": patient_name,
                            "amount": invoice.get("amount", 0),
                            "status": "duplicate_prevented",
                            "message": "Invoice already generated today"
                        }
            except:
                continue
        
        worksheet = get_accounts_receivable_sheet()
        
        # Generate invoice reference
        invoice_ref = generate_invoice_ref()
        invoice_date = datetime.now().strftime("%d-%m-%Y %H:%M")
        current_date = datetime.now().strftime("%Y-%m-%d")
        
        # Calculate revenue
        home_care_revenue = float(client_record.get("Patient Admission Revenue", 0) or 0)
        additional_nursing = float(client_record.get("Additional Nursing Charges", 0) or 0)
        discount = float(client_record.get("Discount", 0) or 0)
        
        total_amount = calculate_patientadmission_revenue(home_care_revenue, additional_nursing, discount)
        
        # Prepare row data
        row_data = {
            "Date": current_date,
            "Invoice Date": invoice_date,
            "Invoice Ref": invoice_ref,
            "Patient Name": client_record.get("PATIENT NAME", ""),
            "Gender": client_record.get("GENDER", ""),
            "Age": client_record.get("AGE", ""),
            "Location": client_record.get("CARE CENTER", ""),
            "Pain Point": client_record.get("PAIN POINT", ""),
            "Service Type": "Patient Admission",
            "Service Name": f"Patient Admission - {client_record.get('SHIFT', 'Regular')}",
            "Patient Admission Revenue": home_care_revenue,
            "Additional Nursing Charges": additional_nursing,
            "Discount": discount,
            "Total Amount": total_amount,
            "Status": "Invoiced",
            "Created At": datetime.now().isoformat(),
            "Updated At": datetime.now().isoformat(),
            "Notes": f"Auto-generated monthly patient admission invoice. Service started: {client_record.get('SERVICE STARTED ON', 'N/A')}",
        }
        
        # Get headers and append row
        headers = worksheet.row_values(1)
        row_values = []
        for header in headers:
            row_values.append(str(row_data.get(header, "")))
        
        worksheet.append_row(row_values)
        
        print(f"[Patient Admission Billing] Generated invoice {invoice_ref} for {client_record.get('PATIENT NAME')} - Amount: ₹{total_amount}")
        
        # NOTE: LAST BILLED DATE is updated by the calling function (generate_invoice_manual in routes)
        # This avoids duplicate rows in SNF sheet
        
        return {
            "invoice_ref": invoice_ref,
            "invoice_date": invoice_date,
            "patient_name": client_record.get("PATIENT NAME", ""),
            "amount": total_amount,
            "status": "success",
            "message": "Invoice created successfully"
        }
        
    except Exception as e:
        print(f"Error generating patient admission invoice: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to generate invoice: {str(e)}")


def process_daily_billing() -> Dict[str, Any]:
    """
    Process daily billing for all active patient admission clients.
    This function is called by the scheduler.
    
    Returns:
        Summary of billing operations
    """
    print(f"\n[Patient Admission Billing] Starting daily billing check at {datetime.now()}")
    
    try:
        all_clients = get_all_patientadmission_clients()
        # Filter for ACTIVE clients only for billing
        active_clients = [c for c in all_clients if c.get("ACTIVE / INACTIVE", "").strip().upper() == "ACTIVE" and not c.get("SERVICE STOPPED ON", "").strip()]
        print(f"[Patient Admission Billing] Found {len(active_clients)} active clients out of {len(all_clients)} total")
        
        billed_count = 0
        skipped_count = 0
        error_count = 0
        billed_clients = []
        errors = []
        
        for client in active_clients:
            try:
                patient_name = client.get("PATIENT NAME", "Unknown")
                service_start_str = client.get("SERVICE STARTED ON", "")
                
                if not service_start_str:
                    print(f"[Patient Admission Billing] Skipping {patient_name} - No service start date")
                    skipped_count += 1
                    continue
                
                service_start_date = parse_date(service_start_str)
                if not service_start_date:
                    print(f"[Patient Admission Billing] Skipping {patient_name} - Invalid service start date: {service_start_str}")
                    skipped_count += 1
                    continue
                
                # Check for discharge date - skip if patient has been discharged
                discharge_date_str = client.get("Discharge Date", "") or client.get("SERVICE STOPPED ON", "")
                if discharge_date_str:
                    discharge_date = parse_date(discharge_date_str)
                    if discharge_date:
                        today = datetime.now().date()
                        discharge_date_only = discharge_date.date() if isinstance(discharge_date, datetime) else discharge_date
                        
                        # If discharge date is today or in the past, skip auto-billing
                        if discharge_date_only <= today:
                            print(f"[Patient Admission Billing] Skipping {patient_name} - Patient discharged on {discharge_date_str}")
                            skipped_count += 1
                            continue
                
                # Get last billing date from history
                billing_history = get_billing_history(patient_name)
                last_billed_date = None
                
                if billing_history:
                    last_invoice_date_str = billing_history[0].get("invoice_date", "")
                    last_billed_date = parse_date(last_invoice_date_str)
                
                # Check if billing is due today
                if is_billing_due_today(service_start_date, last_billed_date):
                    print(f"[Patient Admission Billing] Billing due for {patient_name}")
                    invoice = generate_patientadmission_invoice(client)
                    
                    # Check if invoice was actually created or duplicate was prevented
                    if invoice.get("status") == "duplicate_prevented":
                        print(f"[Patient Admission Billing] Duplicate prevented for {patient_name}")
                        skipped_count += 1
                    else:
                        billed_count += 1
                        billed_clients.append({
                            "patient_name": patient_name,
                            "invoice_ref": invoice["invoice_ref"],
                            "amount": invoice["amount"]
                        })
                else:
                    skipped_count += 1
                    
            except Exception as e:
                error_count += 1
                error_msg = f"Error billing {client.get('PATIENT NAME', 'Unknown')}: {str(e)}"
                print(f"[Patient Admission Billing] {error_msg}")
                errors.append(error_msg)
                continue
        
        summary = {
            "timestamp": datetime.now().isoformat(),
            "total_active_clients": len(active_clients),
            "billed_count": billed_count,
            "skipped_count": skipped_count,
            "error_count": error_count,
            "billed_clients": billed_clients,
            "errors": errors
        }
        
        print(f"[Patient Admission Billing] Daily billing complete - Billed: {billed_count}, Skipped: {skipped_count}, Errors: {error_count}")
        
        return summary
        
    except Exception as e:
        print(f"[Patient Admission Billing] Critical error in daily billing: {e}")
        return {
            "timestamp": datetime.now().isoformat(),
            "error": str(e),
            "status": "failed"
        }


def create_patientadmission_client(client_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Create a new patient admission client in the Google Sheet.
    
    Args:
        client_data: Dictionary containing client information
    
    Returns:
        Result dictionary with status and message
    """
    try:
        worksheet = get_patientadmission_sheet()
        
        # Get headers
        headers = worksheet.row_values(1)
        
        # Convert date from YYYY-MM-DD to DD/MM/YYYY
        service_date = client_data.get('service_started_on', '')
        if service_date and '-' in service_date:
            parts = service_date.split('-')
            if len(parts) == 3:
                service_date = f"{parts[2]}/{parts[1]}/{parts[0]}"
        
        # Calculate revenue
        home_care_revenue = float(client_data.get('home_care_revenue', 0) or 0)
        additional_nursing = float(client_data.get('additional_nursing_charges', 0) or 0)
        discount = float(client_data.get('discount', 0) or 0)
        total_revenue = (home_care_revenue + additional_nursing) - discount
        
        # Calculate number of occupied bed days (from admission date to today)
        occupied_bed_days = 0
        if service_date:
            try:
                # Parse the admission date (DD/MM/YYYY format)
                admission_date = parse_date(service_date)
                if admission_date:
                    today = datetime.now().date()
                    admission_date_only = admission_date.date() if isinstance(admission_date, datetime) else admission_date
                    days_diff = (today - admission_date_only).days
                    occupied_bed_days = max(0, days_diff)  # Ensure non-negative
            except Exception as e:
                print(f"[Patient Admission] Error calculating bed days: {e}")
                occupied_bed_days = 0
        
        # Get current row count for SI NO
        all_values = worksheet.get_all_values()
        si_no = len(all_values)  # Current row count (including header)
        
        # Prepare row data matching the SNF sheet structure
        row_data = {
            "SI NO": si_no,
            "Date": datetime.now().strftime("%d/%m/%Y"),
            "Patient Name": client_data.get('patient_name', ''),
            "Gender": client_data.get('gender', ''),
            "Admission Date": service_date,
            "Discharge Date": client_data.get('service_stopped_on', ''),
            "Room Type": client_data.get('room_type', 'Twin'),
            "Room No": "",  # Not collected in form
            "Bed Count": "",  # Not collected in form
            "No of occupied bed days By Patient": occupied_bed_days,
            "No of bed days avaiable": "",  # Not collected in form
            "Occupants as on last day of last month": "",  # Not collected in form
            "Expected Revenue": total_revenue,
            "Revenue": total_revenue,
            "Additional Nursing Charges": additional_nursing,
            "Other Charges (Amenities)": 0,  # Not collected in form
            "Discount": discount,
            "Last Month Revenue": "",  # Not collected in form
            "Pain point": client_data.get('pain_point', ''),
            "Type of complaint": "",  # Not collected in form
            "Date_1": "",  # Not collected in form
            "Resolved": "",  # Not collected in form
            "Location": client_data.get('care_center', ''),
        }
        
        # Build row values in the same order as headers
        row_values = []
        for header in headers:
            row_values.append(str(row_data.get(header, "")))
        
        # Append the row
        worksheet.append_row(row_values)
        
        print(f"[Patient Admission] Created new client: {client_data.get('patient_name')}")
        
        return {
            "status": "success",
            "message": "Home care client created successfully",
            "client_name": client_data.get('patient_name')
        }
        
    except Exception as e:
        print(f"[Patient Admission] Error creating client: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to create client: {str(e)}")


def update_patientadmission_client(patient_name: str, client_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Update an existing patient admission client in the Google Sheet.
    
    Args:
        patient_name: Name of the patient to update
        client_data: Dictionary containing updated client information
    
    Returns:
        Result dictionary with status and message
    """
    try:
        worksheet = get_patientadmission_sheet()
        all_values = worksheet.get_all_values()
        
        if not all_values or len(all_values) < 2:
            raise HTTPException(status_code=404, detail="Client not found")
        
        headers = all_values[0]
        header_map = {header: idx for idx, header in enumerate(headers)}
        
        # Find the client row
        client_row_number = None
        for row_idx, row in enumerate(all_values[1:], start=2):
            if not any(row):
                continue
            
            # Try both column name variations
            current_name = ""
            if "Patient Name" in header_map:
                current_name = row[header_map["Patient Name"]]
            elif "PATIENT NAME" in header_map:
                current_name = row[header_map["PATIENT NAME"]]
            
            if current_name.strip().lower() == patient_name.lower():
                client_row_number = row_idx
                break
        
        if not client_row_number:
            raise HTTPException(status_code=404, detail=f"Client '{patient_name}' not found")
        
        # Convert date from YYYY-MM-DD to DD/MM/YYYY
        service_date = client_data.get('service_started_on', '')
        if service_date and '-' in service_date:
            parts = service_date.split('-')
            if len(parts) == 3:
                service_date = f"{parts[2]}/{parts[1]}/{parts[0]}"
        
        # Convert discharge date if provided
        discharge_date = client_data.get('discharge_date', '') or client_data.get('service_stopped_on', '')
        if discharge_date and '-' in discharge_date:
            parts = discharge_date.split('-')
            if len(parts) == 3:
                discharge_date = f"{parts[2]}/{parts[1]}/{parts[0]}"
        
        # Calculate revenue
        home_care_revenue = float(client_data.get('home_care_revenue', 0) or 0)
        additional_nursing = float(client_data.get('additional_nursing_charges', 0) or 0)
        discount = float(client_data.get('discount', 0) or 0)
        total_revenue = (home_care_revenue + additional_nursing) - discount
        
        # Prepare updated data using SNF sheet column names
        updated_data = {
            "Gender": client_data.get('gender', ''),
            "AGE": client_data.get('age', ''),
            "Pain point": client_data.get('pain_point', ''),
            "Location": client_data.get('care_center', ''),
            "Admission Date": service_date,
            "Discharge Date": discharge_date,
            "Room Type": client_data.get('room_type', 'Twin'),
            "Room No": client_data.get('room_no', ''),
            "Bed Count": client_data.get('bed_count', ''),
            "Occupants as on last day of last month": client_data.get('occupants_status', ''),
            "Expected Revenue": home_care_revenue,
            "Revenue": total_revenue,
            "Additional Nursing Charges": additional_nursing,
            "Other Charges (Amenities)": float(client_data.get('other_charges', 0) or 0),
            "Discount": discount,
            "Last Month Revenue": float(client_data.get('last_month_revenue', 0) or 0),
            "Type of complaint": client_data.get('type_of_complaint', ''),
            "Resolved": client_data.get('resolved', ''),
            "Date_1": datetime.now().strftime("%d/%m/%Y"),  # Auto-populate with current date
        }
        
        # Update each cell
        for field, value in updated_data.items():
            if field in header_map:
                col_idx = header_map[field] + 1  # gspread uses 1-indexed columns
                worksheet.update_cell(client_row_number, col_idx, str(value))
        
        print(f"[Patient Admission] Updated client: {patient_name}")
        
        return {
            "status": "success",
            "message": "Home care client updated successfully",
            "client_name": patient_name
        }
        
    except HTTPException:
        raise
    except Exception as e:
        print(f"[Patient Admission] Error updating client: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to update client: {str(e)}")

