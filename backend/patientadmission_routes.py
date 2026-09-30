"""
Patient Admission Routes Module
FastAPI routes for patient admission billing management
"""

from fastapi import APIRouter, HTTPException, Query
from typing import List, Dict, Any, Optional
from pydantic import BaseModel
from datetime import datetime

from patientadmission_service import (
    get_all_patientadmission_clients,
    get_patientadmission_client_by_id,
    get_billing_history,
    generate_patientadmission_invoice,
    calculate_next_billing_date,
    parse_date,
    format_date,
    process_daily_billing,
)

router = APIRouter()


class PatientAdmissionClientCreate(BaseModel):
    patient_name: str
    gender: Optional[str] = ""
    age: Optional[str] = ""
    pain_point: Optional[str] = ""
    care_center: Optional[str] = ""
    service_started_on: str
    service_type: Optional[str] = "Patient Admission"
    home_care_revenue: float
    additional_nursing_charges: float = 0
    discount: float = 0
    shift: Optional[str] = "Regular"
    active_inactive: str = "ACTIVE"


class PatientAdmissionClientUpdate(BaseModel):
    gender: Optional[str] = None
    age: Optional[str] = None
    pain_point: Optional[str] = None
    care_center: Optional[str] = None
    service_type: Optional[str] = None
    home_care_revenue: Optional[float] = None
    additional_nursing_charges: Optional[float] = None
    discount: Optional[float] = None
    shift: Optional[str] = None
    active_inactive: Optional[str] = None
    service_stopped_on: Optional[str] = None


@router.get("/patientadmission/clients")
async def list_homecare_clients(
    status: Optional[str] = Query(None, description="Filter by Twin or Single room type")
):
    """
    Get list of all patient admission clients.
    
    Query Parameters:
        status: Filter by ACTIVE or INACTIVE (optional)
    
    Returns:
        List of patient admission clients with billing information
    """
    try:
        clients = get_all_patientadmission_clients()
        
        # OPTIMIZATION: Fetch ALL billing history in ONE API call instead of 67 separate calls
        # This prevents Google Sheets API rate limit errors
        try:
            from patientadmission_service import get_accounts_receivable_sheet
            ar_sheet = get_accounts_receivable_sheet()
            all_invoices = ar_sheet.get_all_values()
            
            # Build a map of patient_name -> list of invoices
            billing_history_map = {}
            if len(all_invoices) > 1:
                headers = all_invoices[0]
                header_map = {h.strip().lower(): i for i, h in enumerate(headers)}
                
                patient_col = header_map.get('patient name')
                invoice_date_col = header_map.get('invoice date')
                
                if patient_col is not None and invoice_date_col is not None:
                    for row in all_invoices[1:]:
                        if len(row) > max(patient_col, invoice_date_col):
                            patient = row[patient_col].strip()
                            invoice_date = row[invoice_date_col].strip()
                            
                            if patient and invoice_date:
                                if patient not in billing_history_map:
                                    billing_history_map[patient] = []
                                billing_history_map[patient].append({
                                    "invoice_date": invoice_date
                                })
        except Exception as e:
            print(f"[Patient Admission] Warning: Could not fetch billing history in batch: {e}")
            billing_history_map = {}
        
        # Add next billing date to each client
        enriched_clients = []
        for client in clients:
            # Get admission date (from SNF sheet column "Admission Date")
            admission_date_str = client.get("Admission Date", "") or client.get("SERVICE STARTED ON", "")
            admission_date = parse_date(admission_date_str)
            
            # Get discharge date
            discharge_date_str = client.get("Discharge Date", "") or client.get("SERVICE STOPPED ON", "")
            discharge_date = parse_date(discharge_date_str) if discharge_date_str else None
            
            # Get patient name
            patient_name = client.get("Patient Name", "") or client.get("PATIENT NAME", "")
            
            # Calculate occupied bed days (from admission to today or discharge)
            occupied_bed_days = 0
            if admission_date:
                from datetime import datetime
                # Convert both dates to date objects for comparison
                admission_date_only = admission_date.date() if isinstance(admission_date, datetime) else admission_date
                
                if discharge_date:
                    discharge_date_only = discharge_date.date() if isinstance(discharge_date, datetime) else discharge_date
                    end_date = discharge_date_only
                else:
                    end_date = datetime.now().date()
                
                days_diff = (end_date - admission_date_only).days
                occupied_bed_days = max(0, days_diff)
            
            # Calculate next billing date
            next_billing = None
            billing_count = 0
            
            if admission_date:
                # Get billing history from the pre-fetched map
                billing_history = billing_history_map.get(patient_name, [])
                billing_count = len(billing_history)
                
                last_billed_date = None
                if billing_history:
                    last_invoice_date_str = billing_history[0].get("invoice_date", "")
                    last_billed_date = parse_date(last_invoice_date_str)
                
                # If discharge date exists, show discharge date as next billing
                if discharge_date:
                    next_billing = format_date(discharge_date)
                else:
                    # Calculate next billing (same date next month from admission)
                    from patientadmission_service import calculate_next_future_billing_date
                    next_billing_dt = calculate_next_future_billing_date(admission_date, last_billed_date)
                    next_billing = format_date(next_billing_dt)
            
            enriched_clients.append({
                "patient_name": patient_name,
                "gender": client.get("Gender", "") or client.get("GENDER", ""),
                "age": client.get("AGE", ""),
                "care_center": client.get("Location", "") or client.get("LOCATION", "") or client.get("CARE CENTER", ""),
                "pain_point": client.get("Pain point", "") or client.get("PAIN POINT", ""),
                "service_started_on": admission_date_str,
                "service_type": client.get("SERVICE TYPE", "Patient Admission"),
                "home_care_revenue": client.get("Expected Revenue", 0) or client.get("Patient Admission Revenue", 0),
                "additional_nursing_charges": client.get("Additional Nursing Charges", 0),
                "discount": client.get("Discount", 0),
                "revenue": client.get("Revenue", 0) or client.get("REVENUE", 0),
                "shift": occupied_bed_days,  # Now calculated from admission to today
                "status": client.get("ACTIVE / INACTIVE", "ACTIVE"),
                "room_type": client.get("Room Type", "") or client.get("ROOM TYPE", ""),
                "next_billing_date": next_billing,
                "billing_count": billing_count,
            })
        
        # Apply room type filter if provided (but not for "ALL", "ACTIVE", or "INACTIVE")
        if status and status not in ["ALL", "ACTIVE", "INACTIVE"]:
            enriched_clients = [c for c in enriched_clients if c.get("room_type", "") == status]
        
        return {
            "status": "success",
            "count": len(enriched_clients),
            "clients": enriched_clients
        }
        
    except Exception as e:
        print(f"Error listing patient admission clients: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to list clients: {str(e)}")


@router.get("/patientadmission/clients/{patient_name}")
async def get_homecare_client(patient_name: str):
    """
    Get specific patient admission client details.
    
    Path Parameters:
        patient_name: Patient name
    
    Returns:
        Client details with billing history
    """
    try:
        client = get_patientadmission_client_by_id(patient_name)
        
        if not client:
            raise HTTPException(status_code=404, detail="Client not found")
        
        # Get billing history
        billing_history = get_billing_history(patient_name)
        
        # Calculate next billing date
        service_start_str = client.get("Admission Date", "") or client.get("SERVICE STARTED ON", "")
        service_start_date = parse_date(service_start_str)
        
        next_billing = None
        if service_start_date:
            last_billed_date = None
            if billing_history:
                last_invoice_date_str = billing_history[0].get("invoice_date", "")
                last_billed_date = parse_date(last_invoice_date_str)
            
            if last_billed_date:
                next_billing_dt = calculate_next_billing_date(service_start_date, last_billed_date)
            else:
                next_billing_dt = calculate_next_billing_date(service_start_date, service_start_date)
            
            next_billing = format_date(next_billing_dt)
        
        return {
            "status": "success",
            "client": {
                "patient_name": client.get("Patient Name", "") or client.get("PATIENT NAME", ""),
                "gender": client.get("Gender", "") or client.get("GENDER", ""),
                "age": client.get("AGE", ""),
                "care_center": client.get("Location", "") or client.get("LOCATION", "") or client.get("CARE CENTER", ""),
                "pain_point": client.get("Pain point", "") or client.get("PAIN POINT", ""),
                "service_started_on": service_start_str,
                "service_stopped_on": client.get("Discharge Date", "") or client.get("SERVICE STOPPED ON", ""),
                "service_type": client.get("SERVICE TYPE", "Patient Admission"),
                "home_care_revenue": client.get("Expected Revenue", 0) or client.get("Patient Admission Revenue", 0),
                "additional_nursing_charges": client.get("Additional Nursing Charges", 0),
                "discount": client.get("Discount", 0),
                "revenue": client.get("Revenue", 0) or client.get("REVENUE", 0),
                "shift": client.get("SHIFT", ""),
                "status": client.get("ACTIVE / INACTIVE", "ACTIVE"),
                "room_type": client.get("Room Type", "") or client.get("ROOM TYPE", "Twin"),
                "room_no": client.get("Room No", ""),
                "bed_count": client.get("Bed Count", ""),
                "occupants_status": client.get("Occupants as on last day of last month", ""),
                "other_charges": client.get("Other Charges (Amenities)", 0),
                "expected_revenue": client.get("Expected Revenue", 0),
                "last_month_revenue": client.get("Last Month Revenue", 0),
                "next_billing_date": next_billing,
            },
            "billing_history": billing_history
        }
        
    except HTTPException:
        raise
    except Exception as e:
        print(f"Error fetching patient admission client: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to fetch client: {str(e)}")


@router.get("/patientadmission/billing-history/{patient_name}")
async def get_client_billing_history(patient_name: str):
    """
    Get billing history for a patient admission client.
    
    Path Parameters:
        patient_name: Patient name
    
    Returns:
        List of invoices for the client
    """
    try:
        billing_history = get_billing_history(patient_name)
        
        return {
            "status": "success",
            "patient_name": patient_name,
            "count": len(billing_history),
            "history": billing_history
        }
        
    except Exception as e:
        print(f"Error fetching billing history: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to fetch billing history: {str(e)}")


@router.post("/patientadmission/trigger-billing/{patient_name}")
async def trigger_manual_billing(patient_name: str):
    """
    Manually trigger billing for a specific patient admission client.
    
    Path Parameters:
        patient_name: Patient name
    
    Returns:
        Invoice details
    """
    try:
        client = get_patientadmission_client_by_id(patient_name)
        
        if not client:
            raise HTTPException(status_code=404, detail="Client not found")
        
        # Check if client is active
        status = str(client.get("ACTIVE / INACTIVE", "")).strip().upper()
        service_stopped = str(client.get("SERVICE STOPPED ON", "")).strip()
        
        if status != "ACTIVE" or service_stopped:
            raise HTTPException(
                status_code=400, 
                detail="Cannot bill inactive client or client with stopped service"
            )
        
        # Generate invoice
        invoice = generate_patientadmission_invoice(client)
        
        return {
            "status": "success",
            "message": "Invoice generated successfully",
            "invoice": invoice
        }
        
    except HTTPException:
        raise
    except Exception as e:
        print(f"Error triggering manual billing: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to trigger billing: {str(e)}")




@router.post("/patientadmission/run-daily-billing")
async def run_daily_billing():
    """
    Manually trigger the daily billing process.
    This endpoint is useful for testing and manual execution.
    
    Returns:
        Summary of billing operations
    """
    try:
        summary = process_daily_billing()
        return {
            "status": "success",
            "summary": summary
        }
    except Exception as e:
        print(f"Error running daily billing: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/patientadmission/clients")
async def create_client(client_data: dict):
    """
    Create a new patient admission client.
    
    Request body should contain:
    - patient_name (required)
    - gender
    - age
    - pain_point
    - location
    - service_started_on (YYYY-MM-DD format)
    - home_care_revenue (required)
    - additional_nursing_charges
    - discount
    - shift
    - active_inactive
    """
    try:
        from patientadmission_service import create_patientadmission_client
        
        # Validate required fields
        if not client_data.get('patient_name'):
            raise HTTPException(status_code=400, detail="Patient name is required")
        
        if not client_data.get('home_care_revenue'):
            raise HTTPException(status_code=400, detail="Home care revenue is required")
        
        result = create_patientadmission_client(client_data)
        return result
        
    except HTTPException:
        raise
    except Exception as e:
        print(f"Error creating patient admission client: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.put("/patientadmission/clients/{patient_name}")
async def update_client(patient_name: str, client_data: dict):
    """
    Update an existing patient admission client.
    
    Path parameter:
    - patient_name: Name of the patient to update
    
    Request body can contain:
    - gender
    - age
    - pain_point
    - location
    - service_started_on (YYYY-MM-DD format)
    - home_care_revenue
    - additional_nursing_charges
    - discount
    - shift
    - active_inactive
    - service_stopped_on (YYYY-MM-DD format, only if inactive)
    """
    try:
        from patientadmission_service import update_patientadmission_client
        
        result = update_patientadmission_client(patient_name, client_data)
        return result
        
    except HTTPException:
        raise
    except Exception as e:
        print(f"Error updating patient admission client: {e}")
        raise HTTPException(status_code=500, detail=str(e))


# Import timedelta for billing preview
from datetime import timedelta


@router.post("/patientadmission/generate-invoice/{patient_name}")
async def generate_invoice_manual(patient_name: str, billing_date: Optional[str] = None):
    """
    Manually generate invoice for a patient admission client.
    
    Path Parameters:
        patient_name: Name of the patient
    
    Query Parameters:
        billing_date: Optional billing date in YYYY-MM-DD format (defaults to today)
    
    Returns:
        Invoice details
    """
    try:
        # Get client data
        client = get_patientadmission_client_by_id(patient_name)
        
        if not client:
            raise HTTPException(status_code=404, detail="Client not found")
        
        # Generate invoice
        invoice = generate_patientadmission_invoice(client)
        
        # Update LAST BILLED DATE in SNF sheet
        from patientadmission_service import get_patientadmission_sheet
        worksheet = get_patientadmission_sheet()
        all_values = worksheet.get_all_values()
        
        if all_values and len(all_values) >= 2:
            headers = all_values[0]
            header_map = {header: idx for idx, header in enumerate(headers)}
            
            # Find client row
            for row_idx, row in enumerate(all_values[1:], start=2):
                if not any(row):
                    continue
                
                current_name = ""
                if "Patient Name" in header_map:
                    current_name = row[header_map["Patient Name"]]
                elif "PATIENT NAME" in header_map:
                    current_name = row[header_map["PATIENT NAME"]]
                
                if current_name.strip().lower() == patient_name.lower():
                    # Update LAST BILLED DATE
                    if "LAST BILLED DATE" in header_map:
                        col_idx = header_map["LAST BILLED DATE"] + 1
                        billed_date = datetime.now().strftime("%d/%m/%Y")
                        worksheet.update_cell(row_idx, col_idx, billed_date)
                        print(f"[Patient Admission] Updated LAST BILLED DATE for {patient_name}")
                    break
        
        return {
            "status": "success",
            "message": "Invoice generated successfully",
            "invoice": invoice
        }
        
    except HTTPException:
        raise
    except Exception as e:
        print(f"[Patient Admission] Error generating invoice: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to generate invoice: {str(e)}")


@router.get("/patientadmission/billing-preview")
async def get_billing_preview(days: int = Query(default=30, ge=1, le=90)):
    """
    Get upcoming billing preview for patient admission clients.
    
    Query Parameters:
        days: Number of days to preview (default: 30, max: 90)
    
    Returns:
        List of upcoming bills with patient name, billing date, amount, and days until
    """
    print("="*80)
    print(f"[BILLING PREVIEW ENDPOINT CALLED] Days: {days}")
    print("="*80)
    try:
        from datetime import timedelta
        
        # Get all clients
        all_clients = get_all_patientadmission_clients()
        
        upcoming_bills = []
        total_forecast = 0
        today = datetime.now().date()
        end_date = today + timedelta(days=days)
        
        print(f"[Billing Preview] Processing {len(all_clients)} clients for next {days} days")
        print(f"[Billing Preview] Today: {today}, End date: {end_date}")
        print(f"[Billing Preview] About to start loop with {len(all_clients)} clients")
        
        for client in all_clients:
            try:
                patient_name = client.get("Patient Name", "") or client.get("PATIENT NAME", "")
                if not patient_name:
                    continue
                
                print(f"[Billing Preview DEBUG] Processing patient: {patient_name}")
                
                # Check for discharge date - skip if discharged
                discharge_date_str = client.get("Discharge Date", "") or client.get("SERVICE STOPPED ON", "")
                if discharge_date_str:
                    discharge_date = parse_date(discharge_date_str)
                    if discharge_date:
                        discharge_date_only = discharge_date.date() if isinstance(discharge_date, datetime) else discharge_date
                        print(f"[Billing Preview DEBUG] {patient_name} - Discharge date: {discharge_date_only}")
                        if discharge_date_only <= today:
                            print(f"[Billing Preview DEBUG] {patient_name} - SKIPPED (discharged)")
                            continue
                
                # Get admission date
                admission_date_str = client.get("Admission Date", "") or client.get("SERVICE STARTED ON", "")
                if not admission_date_str:
                    print(f"[Billing Preview DEBUG] {patient_name} - SKIPPED (no admission date)")
                    continue
                
                admission_date = parse_date(admission_date_str)
                if not admission_date:
                    print(f"[Billing Preview DEBUG] {patient_name} - SKIPPED (invalid admission date)")
                    continue
                
                print(f"[Billing Preview DEBUG] {patient_name} - Admission date: {admission_date}")
                
                # Get billing history
                billing_history = get_billing_history(patient_name)
                last_billed_date = None
                if billing_history:
                    last_invoice_date_str = billing_history[0].get("invoice_date", "")
                    last_billed_date = parse_date(last_invoice_date_str)
                    print(f"[Billing Preview DEBUG] {patient_name} - Last billed: {last_billed_date}")
                else:
                    print(f"[Billing Preview DEBUG] {patient_name} - No billing history")
                
                # Calculate next billing date
                if last_billed_date:
                    next_billing_dt = calculate_next_billing_date(admission_date, last_billed_date)
                else:
                    next_billing_dt = calculate_next_billing_date(admission_date, admission_date)
                
                print(f"[Billing Preview DEBUG] {patient_name} - Initial next billing: {next_billing_dt.date()}")
                
                # Make sure next billing is in the future
                loop_count = 0
                while next_billing_dt.date() < today and loop_count < 12:
                    next_billing_dt = calculate_next_billing_date(admission_date, next_billing_dt)
                    loop_count += 1
                
                next_billing_date = next_billing_dt.date()
                print(f"[Billing Preview DEBUG] {patient_name} - Final next billing: {next_billing_date}, Days until: {(next_billing_date - today).days}")
                
                # Check if billing date is within preview period
                if today <= next_billing_date <= end_date:
                    # Helper function to safely convert values to float (handles commas)
                    def safe_float(value):
                        if value is None:
                            return 0.0
                        if isinstance(value, (int, float)):
                            return float(value)
                        # Handle string values - remove commas
                        str_value = str(value).strip().replace(',', '')
                        if not str_value or str_value == '':
                            return 0.0
                        try:
                            return float(str_value)
                        except (ValueError, TypeError):
                            return 0.0
                    
                    # Calculate amount using SNF sheet columns
                    patient_admission_revenue = safe_float(client.get("Expected Revenue", 0))
                    additional_nursing = safe_float(client.get("Additional Nursing Charges", 0))
                    discount = safe_float(client.get("Discount", 0))
                    
                    # Calculate total amount: (Expected Revenue + Additional Nursing) - Discount
                    amount = (patient_admission_revenue + additional_nursing) - discount
                    
                    print(f"[Billing Preview DEBUG] {patient_name} - Revenue: Rs.{patient_admission_revenue}, Nursing: Rs.{additional_nursing}, Discount: Rs.{discount}, Total: Rs.{amount}")
                    print(f"[Billing Preview DEBUG] {patient_name} - INCLUDED! Amount: Rs.{amount}")
                    
                    days_until = (next_billing_date - today).days
                    
                    upcoming_bills.append({
                        "patient_name": patient_name,
                        "billing_date": format_date(next_billing_dt),
                        "amount": amount,
                        "days_until": days_until
                    })
                    
                    total_forecast += amount
                else:
                    print(f"[Billing Preview DEBUG] {patient_name} - EXCLUDED (outside preview period: {today} to {end_date})")
                    
            except Exception as e:
                print(f"[Billing Preview] Error processing client {client.get('Patient Name', 'Unknown')}: {e}")
                import traceback
                traceback.print_exc()
                continue
        
        # Sort by billing date
        upcoming_bills.sort(key=lambda x: x["days_until"])
        
        print(f"[Billing Preview] Found {len(upcoming_bills)} upcoming bills, total forecast: Rs.{total_forecast}")
        
        return {
            "status": "success",
            "upcoming_bills": upcoming_bills,
            "total_forecast": total_forecast,
            "preview_days": days
        }
        
    except Exception as e:
        print(f"[Billing Preview] Error in billing preview: {e}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Failed to generate billing preview: {str(e)}")
