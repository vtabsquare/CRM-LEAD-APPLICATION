"""
Invoice Service Module
Handles all invoice-related business logic and Google Sheets operations
"""

from typing import List, Dict, Any, Optional
from datetime import datetime
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
ADMISSION_CREDENTIALS_FILE = "CRM-admission.json"
CRM_LEAD_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
CRM_ADMISSION_SHEET_ID = os.getenv("PATIENT_ADMISSION_SHEET_ID")



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


def get_crm_lead_sheet():
    """Get CRM_Lead → Sheet1 for patient data"""
    client = get_google_sheet_client()
    if not CRM_LEAD_SHEET_ID:
        raise HTTPException(status_code=500, detail="CRM_Lead Sheet ID not configured")
    spreadsheet = client.open_by_key(CRM_LEAD_SHEET_ID)
    worksheet = spreadsheet.worksheet("Sheet1")
    return worksheet


def get_crm_admission_sheet():
    """Get CRM_Admission → Invoice Table for invoice data"""
    # Try to use separate credentials for admission sheet if available
    creds_file = ADMISSION_CREDENTIALS_FILE if os.path.exists(ADMISSION_CREDENTIALS_FILE) else CREDENTIALS_FILE
    client = get_google_sheet_client(creds_file)
    
    if not CRM_ADMISSION_SHEET_ID:
        raise HTTPException(status_code=500, detail="CRM_Admission Sheet ID not configured")
    spreadsheet = client.open_by_key(CRM_ADMISSION_SHEET_ID)
    worksheet = spreadsheet.worksheet("Invoice Table")
    return worksheet


def search_patients(query: str) -> List[Dict[str, Any]]:
    """
    Search patients in CRM_Lead → Sheet1
    Searches by Patient Name, Patient ID, or Mobile Number
    """
    try:
        worksheet = get_crm_lead_sheet()
        all_records = worksheet.get_all_records()
        
        if not query or not query.strip():
            return []
        
        query_lower = query.strip().lower()
        results = []
        
        for record in all_records:
            # Search in Patient Name, Member ID, Mobile Number
            patient_name = str(record.get("Patient Name", "")).strip().lower()
            member_id = str(record.get("Member ID Key", "") or record.get("Member ID key", "")).strip().lower()
            mobile = str(record.get("Mobile Number", "")).strip().lower()
            
            if (query_lower in patient_name or 
                query_lower in member_id or 
                query_lower in mobile):
                
                results.append({
                    "patient_id": record.get("Member ID Key") or record.get("Member ID key", ""),
                    "patient_name": record.get("Patient Name", ""),
                    "gender": record.get("Gender", ""),
                    "age": record.get("Age", ""),
                    "mobile": record.get("Mobile Number", ""),
                    "address": record.get("Patient Location", "") or record.get("Address", ""),
                    "email": record.get("Email Id", "") or record.get("Email ID", ""),
                })
        
        return results[:50]  # Limit to 50 results
        
    except Exception as e:
        print(f"Error searching patients: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to search patients: {str(e)}")


def get_invoices(patient_id: Optional[str] = None, 
                status: Optional[str] = None,
                care_center: Optional[str] = None,
                provider: Optional[str] = None,
                invoice_ref: Optional[str] = None,
                date_from: Optional[str] = None,
                date_to: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Get invoices from CRM_Admission sheet with optional filtering
    Returns empty list if no results found (not an error)
    """
    try:
        worksheet = get_crm_admission_sheet()
        all_values = worksheet.get_all_values()
        
        if not all_values or len(all_values) < 2:
            return []
        
        # Get headers from first row
        headers = all_values[0]
        
        # Create a mapping of header names to indices
        header_map = {}
        for idx, header in enumerate(headers):
            if header and header not in header_map:  # Skip empty headers and duplicates
                header_map[header] = idx
        
        results = []
        for row in all_values[1:]:  # Skip header row
            if not any(row):  # Skip empty rows
                continue
            
            # Create record dict using header mapping
            record = {}
            for header, idx in header_map.items():
                record[header] = row[idx] if idx < len(row) else ""
            
            # Apply filters (AND logic)
            # Patient ID filter
            if patient_id:
                record_patient_id = str(record.get("Patient ID", "")) or str(record.get("Member ID Key", ""))
                if record_patient_id != patient_id:
                    continue
            
            # Status filter
            if status and status.lower() != 'all':
                if str(record.get("Status", "")).lower() != status.lower():
                    continue
            
            # Care Center filter
            if care_center and care_center.lower() != 'all':
                if str(record.get("Care Center", "")).lower() != care_center.lower():
                    continue
            
            # Provider filter
            if provider and provider.lower() not in ['all', 'all providers']:
                if str(record.get("Provider", "")).lower() != provider.lower():
                    continue
            
            # Invoice Ref filter (contains search)
            if invoice_ref:
                record_invoice_ref = str(record.get("Invoice Ref", "")).lower()
                if invoice_ref.lower() not in record_invoice_ref:
                    continue
            
            # Date range filtering
            if date_from or date_to:
                invoice_date_str = record.get("Invoice Date", "")
                if invoice_date_str:
                    try:
                        # Try to parse the date (assuming format: DD-MM-YYYY or YYYY-MM-DD)
                        if '-' in invoice_date_str:
                            parts = invoice_date_str.split('-')
                            if len(parts) == 3:
                                # Check if it's DD-MM-YYYY or YYYY-MM-DD
                                if len(parts[0]) == 4:  # YYYY-MM-DD
                                    invoice_date = datetime.strptime(invoice_date_str, "%Y-%m-%d")
                                else:  # DD-MM-YYYY
                                    invoice_date = datetime.strptime(invoice_date_str, "%d-%m-%Y")
                                
                                if date_from:
                                    from_date = datetime.strptime(date_from, "%Y-%m-%d")
                                    if invoice_date < from_date:
                                        continue
                                
                                if date_to:
                                    to_date = datetime.strptime(date_to, "%Y-%m-%d")
                                    if invoice_date > to_date:
                                        continue
                    except ValueError:
                        # If date parsing fails, skip date filtering for this record
                        pass
            
            results.append({
                "invoice_id": record.get("Invoice Ref", ""),
                "invoice_date": record.get("Invoice Date", ""),
                "patient_id": record.get("Patient ID", "") or record.get("Member ID Key", ""),
                "patient_name": record.get("Patient Name", ""),
                "visit_id": record.get("Visit ID", ""),
                "care_center": record.get("Care Center", ""),
                "provider": record.get("Provider", ""),
                "status": record.get("Status", "Invoiced"),
                "total_amount": record.get("Total Amount", 0),
            })
        
        return results
        
    except Exception as e:
        print(f"Error fetching invoices: {e}")
        # Return empty list instead of raising exception for better UX
        return []




def generate_invoice_ref() -> str:
    """Generate unique invoice reference number"""
    try:
        worksheet = get_crm_admission_sheet()
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



def create_invoice(invoice_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Create new invoice in CRM_Admission → Invoice Table sheet
    Fetches complete patient data from CRM_Lead to maintain consistency
    """
    try:
        worksheet = get_crm_admission_sheet()
        
        # Generate invoice reference
        invoice_ref = generate_invoice_ref()
        invoice_date = datetime.now().strftime("%d-%m-%Y %H:%M")
        current_date = datetime.now().strftime("%Y-%m-%d")
        
        # Fetch complete patient data from CRM_Lead to get all fields
        patient_id = invoice_data.get("patient_id", "")
        patient_data = None
        
        try:
            lead_sheet = get_crm_lead_sheet()
            all_patients = lead_sheet.get_all_records()
            
            for patient in all_patients:
                member_id = str(patient.get("Member ID Key", "") or patient.get("Member ID key", "")).strip()
                if member_id == patient_id:
                    patient_data = patient
                    break
        except Exception as e:
            print(f"Warning: Could not fetch patient data from CRM_Lead: {e}")
        
        # Prepare row data with complete patient information
        row_data = {
            "Date": current_date,
            "Invoice Date": invoice_date,
            "Invoice Ref": invoice_ref,
            "Member ID Key": patient_id,  # Use same format as CRM_Lead
            "Patient ID": patient_id,
            "Patient Name": invoice_data.get("patient_name", ""),
            "Visit ID": invoice_data.get("visit_id", ""),
            "Care Center": invoice_data.get("care_center", ""),
            "Corporate Customer": "Yes" if invoice_data.get("corporate_customer") else "No",
            "Status": "Invoiced",
            "Total Amount": invoice_data.get("total_amount", 0),
            "Created At": datetime.now().isoformat(),
            "Updated At": datetime.now().isoformat(),
        }
        
        # Add patient details from CRM_Lead if available
        if patient_data:
            row_data.update({
                "Patient Last Name": patient_data.get("Patient Last Name", ""),
                "Gender": patient_data.get("Gender", ""),
                "Date of Birth": patient_data.get("Date of Birth", ""),
                "Age": patient_data.get("Age", ""),
                "Blood Group": patient_data.get("Blood Group", ""),
                "Patient Marital": patient_data.get("Patient Marital", ""),
                "Nationality": patient_data.get("Nationality", ""),
                "Religion": patient_data.get("Religion", ""),
                "Aadhaar": patient_data.get("Aadhaar", ""),
                "ID Proof Type": patient_data.get("ID Proof Type", ""),
                "ID Proof Number": patient_data.get("ID Proof Number", ""),
                "Door Number": patient_data.get("Door Number", ""),
                "Street": patient_data.get("Street", ""),
                "City": patient_data.get("City", ""),
                "District": patient_data.get("District", ""),
                "State": patient_data.get("State", ""),
                "Pin Code": patient_data.get("Pin Code", ""),
                "Mobile Number": patient_data.get("Mobile Number", ""),
                "Email Id": patient_data.get("Email Id", "") or patient_data.get("Email ID", ""),
            })
        
        # Add service items
        services = invoice_data.get("services", [])
        if services:
            # Store first service in main row
            first_service = services[0]
            row_data.update({
                "Service Name": first_service.get("service_name", ""),
                "Provider": first_service.get("provider", ""),
                "Perform Date": first_service.get("perform_date", ""),
                "Price": first_service.get("price", 0),
                "Quantity": first_service.get("quantity", 1),
                "Discount": first_service.get("discount", 0),
                "Tax Type": first_service.get("tax_type", "Non-taxable"),
                "Tax Amount": first_service.get("tax_amount", 0),
                "Amount": first_service.get("amount", 0),
                "Sold By": first_service.get("sold_by", ""),
                "External Provider": first_service.get("external_provider", ""),
                "Notes": first_service.get("notes", ""),
            })
        
        # Get headers and append row
        headers = worksheet.row_values(1)
        row_values = []
        for header in headers:
            row_values.append(row_data.get(header, ""))
        
        worksheet.append_row(row_values)
        
        # If multiple services, append additional rows with same invoice_ref
        if len(services) > 1:
            for service in services[1:]:
                service_row_data = row_data.copy()
                service_row_data.update({
                    "Service Name": service.get("service_name", ""),
                    "Provider": service.get("provider", ""),
                    "Perform Date": service.get("perform_date", ""),
                    "Price": service.get("price", 0),
                    "Quantity": service.get("quantity", 1),
                    "Discount": service.get("discount", 0),
                    "Tax Type": service.get("tax_type", "Non-taxable"),
                    "Tax Amount": service.get("tax_amount", 0),
                    "Amount": service.get("amount", 0),
                    "Sold By": service.get("sold_by", ""),
                    "External Provider": service.get("external_provider", ""),
                    "Notes": service.get("notes", ""),
                })
                
                service_row_values = []
                for header in headers:
                    service_row_values.append(service_row_data.get(header, ""))
                
                worksheet.append_row(service_row_values)
        
        return {
            "invoice_id": invoice_ref,
            "invoice_date": invoice_date,
            "status": "success",
            "message": "Invoice created successfully"
        }
        
    except Exception as e:
        print(f"Error creating invoice: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to create invoice: {str(e)}")


def get_invoice_details(invoice_id: str) -> Dict[str, Any]:
    """
    Get detailed invoice information
    """
    try:
        worksheet = get_crm_admission_sheet()
        all_records = worksheet.get_all_records()
        
        # Find invoice
        invoice_rows = []
        for record in all_records:
            if str(record.get("Invoice Ref", "")) == invoice_id:
                invoice_rows.append(record)
        
        if not invoice_rows:
            raise HTTPException(status_code=404, detail="Invoice not found")
        
        # First row contains main invoice data
        main_row = invoice_rows[0]
        
        # Collect all service items
        services = []
        for row in invoice_rows:
            if row.get("Service Name"):
                services.append({
                    "service_name": row.get("Service Name", ""),
                    "provider": row.get("Provider", ""),
                    "perform_date": row.get("Perform Date", ""),
                    "price": row.get("Price", 0),
                    "quantity": row.get("Quantity", 1),
                    "discount": row.get("Discount", 0),
                    "tax_type": row.get("Tax Type", ""),
                    "tax_amount": row.get("Tax Amount", 0),
                    "amount": row.get("Amount", 0),
                    "notes": row.get("Notes", ""),
                })
        
        return {
            "invoice_id": invoice_id,
            "invoice_date": main_row.get("Invoice Date", ""),
            "patient_id": main_row.get("Patient ID", ""),
            "patient_name": main_row.get("Patient Name", ""),
            "visit_id": main_row.get("Visit ID", ""),
            "care_center": main_row.get("Care Center", ""),
            "corporate_customer": main_row.get("Corporate Customer", "") == "Yes",
            "status": main_row.get("Status", ""),
            "total_amount": main_row.get("Total Amount", 0),
            "services": services,
        }
        
    except HTTPException:
        raise
    except Exception as e:
        print(f"Error fetching invoice details: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to fetch invoice details: {str(e)}")


def calculate_invoice_totals(services: List[Dict[str, Any]]) -> Dict[str, float]:
    """
    Calculate invoice totals from service items
    """
    subtotal = 0
    total_discount = 0
    total_tax = 0
    
    for service in services:
        price = float(service.get("price", 0))
        qty = int(service.get("quantity", 1))
        discount = float(service.get("discount", 0))
        tax_amount = float(service.get("tax_amount", 0))
        
        item_subtotal = price * qty
        subtotal += item_subtotal
        total_discount += discount
        total_tax += tax_amount
    
    grand_total = subtotal - total_discount + total_tax
    rounded_total = round(grand_total)
    
    return {
        "subtotal": subtotal,
        "discount": total_discount,
        "tax": total_tax,
        "grand_total": grand_total,
        "rounded_total": rounded_total,
    }
