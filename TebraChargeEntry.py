# TebraChargeEntry.py
# -*- coding: utf-8 -*-
"""
Tebra Charge Entry Streamlit Application
By Panacea Smart Solutions
Developed by Saqib Sherwani

Handles charge entry by grouping charges per case found for existing patients.
NOTE: Patients MUST have at least one case created in Tebra beforehand.
This script will fail rows for patients where no existing case is found via the API.
Includes flexible provider name matching.
MODIFIED: To handle multiple Dates of Service for the same patient as separate encounters.
"""

import streamlit as st
import pandas as pd
import zeep
import zeep.helpers
from zeep.exceptions import Fault as SoapFault, TransportError, LookupError as ZeepLookupError
from requests.exceptions import ConnectionError as RequestsConnectionError
import datetime
import time
import re
from collections import defaultdict
import io

# --- Application Configuration ---
APP_TITLE = "🤖 Tebra Charge Entry"
APP_SUBTITLE = "By Panacea Smart Solutions"
APP_FOOTER = "Tebra Charge Entry Tool © 2025 | Panacea Smart Solutions | Developed by Saqib Sherwani"
TEBRA_PRACTICE_NAME = "Pediatrics West" # Hardcoded Practice Name
TEBRA_WSDL_URL = "https://webservice.kareo.com/services/soap/2.1/KareoServices.svc?singleWsdl"

# --- SET PAGE CONFIG MUST BE THE FIRST STREAMLIT COMMAND ---
st.set_page_config(page_title=APP_TITLE, page_icon="🤖", layout="wide", initial_sidebar_state="expanded")

# --- Expected Excel Columns ---
EXPECTED_COLUMNS = [
    'Patient ID', 'From Date', 'Through Date', 'Rendering Provider', 'Scheduling Provider',
    'Location', 'Place of Service', 'Encounter Mode', 'Procedures', 'Mod 1', 'Mod 2',
    'Units', 'Diag 1', 'Diag 2', 'Diag 3', 'Diag 4', 'Batch Number'
]
# Define column name constants for consistency
COL_PATIENT_ID = 'Patient ID'; COL_FROM_DATE = 'From Date'; COL_THROUGH_DATE = 'Through Date'
COL_RENDERING_PROVIDER = 'Rendering Provider'; COL_SCHEDULING_PROVIDER = 'Scheduling Provider'
COL_LOCATION = 'Location'; COL_PLACE_OF_SERVICE_EXCEL = 'Place of Service'
COL_ENCOUNTER_MODE = 'Encounter Mode'; COL_PROCEDURES = 'Procedures'; COL_MOD1 = 'Mod 1'; COL_MOD2 = 'Mod 2'
COL_UNITS = 'Units'; COL_DIAG1 = 'Diag 1'; COL_DIAG2 = 'Diag 2'
COL_DIAG3 = 'Diag 3'; COL_DIAG4 = 'Diag 4'; COL_BATCH_NUMBER = 'Batch Number'

# --- Place of Service Mapping ---
POS_CODE_MAP = {
    "OFFICE": {"code": "11", "name": "Office"},
    "IN OFFICE": {"code": "11", "name": "Office"},
    "TELEHEALTH": {"code": "10", "name": "Telehealth Provided in Patient’s Home"},
    "TELEHEALTH OFFICE": {"code": "02", "name": "Telehealth Provided Other than in Patient’s Home"},
}

def apply_custom_styling():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap');
        html, body, [class*="st-"] { font-family: 'Open Sans', sans-serif; }
        .footer { position: fixed; left: 0; bottom: 0; width: 100%; background-color: #004A7C; color: #E0F2F7; text-align: center; padding: 10px; font-size: 13px; z-index: 1000; }
        .stButton>button { border: none; border-radius: 8px; padding: 10px 20px; font-weight: 600; background-color: #007bff; color: white; transition: background-color 0.3s ease; cursor: pointer; }
        .stButton>button:hover { background-color: #0056b3; }
        .stTextInput input, .stFileUploader label, .stTextArea textarea { border-radius: 8px; border: 1px solid #ced4da; }
        .stFileUploader label { border: 2px dashed #007bff; padding: 20px; background-color: var(--secondary-background-color); }
        .stFileUploader>div>div>button { background-color: #6c757d; color:white; }
        .stFileUploader>div>div>button:hover { background-color: #5a6268; color:white; }
        .message-box { border-left-width: 5px; padding: 12px; margin-bottom: 15px; border-radius: 6px; font-size: 14px; box-shadow: 1px 1px 3px rgba(0,0,0,0.1); color: var(--text-color); }
        .info-message { border-left-color: #17a2b8; background-color: var(--secondary-background-color); }
        .success-message { border-left-color: #28a745; background-color: var(--secondary-background-color); }
        .error-message { border-left-color: #dc3545; background-color: var(--secondary-background-color); }
        .warning-message { border-left-color: #ffc107; background-color: var(--secondary-background-color); }
        [data-testid="stSidebar"] { background-color: #f0f2f6; border-right: 1px solid #dee2e6; }
        [data-theme="dark"] [data-testid="stSidebar"] { background-color: #262730 !important; border-right: 1px solid #31333F !important; }
        [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: #004A7C; font-weight: 600; }
        [data-theme="dark"] [data-testid="stSidebar"] h1, [data-theme="dark"] [data-testid="stSidebar"] h2, [data-theme="dark"] [data-testid="stSidebar"] h3 { color: #e1e1e1 !important; }
        [data-theme="dark"] [data-testid="stSidebar"] .stTextInput input { background-color: #FFFFFF !important; color: #000000 !important; border: 1px solid #B0B0B0 !important; }
        [data-theme="dark"] [data-testid="stSidebar"] .stTextInput input::placeholder { color: #555555 !important; }
        [data-theme="dark"] [data-testid="stSidebar"] .stTextInput label { color: #e1e1e1 !important; }
        [data-theme="dark"] [data-testid="stSidebar"] div[data-testid="stFileUploaderFile"] { background-color: #FFFFFF !important; border-radius: 4px !important; padding: 5px 8px !important; margin-bottom: 5px !important; }
        [data-theme="dark"] [data-testid="stSidebar"] div[data-testid="stFileUploaderFile"] *, [data-theme="dark"] [data-testid="stSidebar"] div[data-testid="stFileUploaderFile"] button svg { color: #000000 !important; fill: #000000 !important; }
        [data-theme="dark"] [data-testid="stSidebar"] .stFileUploader label { background-color: var(--secondary-background-color) !important; border: 2px dashed #007bff !important; color: #e1e1e1 !important; }
        [data-theme="dark"] [data-testid="stSidebar"] .stFileUploader>div>div>button { background-color: #6c757d !important; color: white !important; }
        [data-theme="dark"] [data-testid="stSidebar"] .stFileUploader>div>div>button:hover { background-color: #5a6268 !important; color: white !important; }
        </style>
    """, unsafe_allow_html=True)

def display_message(type, message):
    st.markdown(f'<div class="message-box {type}-message">{message}</div>', unsafe_allow_html=True)

@st.cache_resource(ttl=3600)
def create_api_client(wsdl_url):
    try:
        from requests import Session
        from zeep.transports import Transport
        session = Session(); session.timeout = 60
        transport = Transport(session=session, timeout=60)
        client = zeep.Client(wsdl=wsdl_url, transport=transport)
        return client
    except Exception as e:
        st.error(f"Fatal Error: Could not initialize Zeep SOAP client: {e}")
        return None

def build_request_header(credentials, client):
    if not client: return None
    try:
        header_type = client.get_type('ns0:RequestHeader')
        pw = credentials['Password'].replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;').replace("'", '&apos;')
        return header_type(CustomerKey=credentials['CustomerKey'], User=credentials['User'], Password=pw)
    except ZeepLookupError as le: display_message("error", f"Zeep LookupError building header: {le}. WSDL issue?"); return None
    except Exception as e: display_message("error", f"Error building API request header: {e}"); return None

def format_datetime_for_api(date_value):
    if pd.isna(date_value) or date_value is None: return None
    try: return pd.to_datetime(date_value).strftime('%Y-%m-%dT%H:%M:%S')
    except Exception as e: st.warning(f"Date parse warning: '{date_value}'. Error: {e}. Using None."); return None

# --- ID Lookup Functions ---
def get_practice_id_from_name(client_obj, header_obj, practice_name_to_find):
    if not client_obj or not header_obj: return None
    cache_key = f"practice_id_{practice_name_to_find}"
    if cache_key in st.session_state: return st.session_state[cache_key]
    with st.spinner(f"Verifying Practice '{practice_name_to_find}'..."):
        practice_id = None
        try:
            req_type = client_obj.get_type('ns0:GetPracticesReq')
            filter_type = client_obj.get_type('ns0:PracticeFilter')
            fields_type = client_obj.get_type('ns0:PracticeFieldsToReturn')
            fields = fields_type(ID=True, PracticeName=True)
            p_filter = filter_type(PracticeName=practice_name_to_find)
            req = req_type(RequestHeader=header_obj, Filter=p_filter, Fields=fields)
            resp = client_obj.service.GetPractices(request=req)
            if hasattr(resp, 'ErrorResponse') and resp.ErrorResponse and resp.ErrorResponse.IsError: raise Exception(f"API Error: {resp.ErrorResponse.ErrorMessage}")
            if hasattr(resp, 'SecurityResponse') and resp.SecurityResponse and not resp.SecurityResponse.Authorized: raise Exception(f"Auth Error: {resp.SecurityResponse.SecurityResult}")
            if hasattr(resp, 'Practices') and resp.Practices and hasattr(resp.Practices, 'PracticeData') and resp.Practices.PracticeData:
                practice_obj = next((p for p in resp.Practices.PracticeData if p.PracticeName and p.PracticeName.strip().lower() == practice_name_to_find.strip().lower() and p.ID), None)
                if practice_obj: practice_id = int(practice_obj.ID)
                else: display_message("warning", f"Practice '{practice_name_to_find}' name not found.")
            else: display_message("warning", f"No practices returned matching '{practice_name_to_find}'.")
        except SoapFault as sf:
            msg = f"SOAP Fault GetPractices: {sf.message} (Code: {sf.code})"
            if "Unknown fault occured" in str(sf.message): msg += ". Often due to invalid Credentials or Permissions."
            display_message("error", msg)
        except Exception as e: display_message("error", f"Error Verifying Practice: {e}")
        st.session_state[cache_key] = practice_id
        return practice_id

def get_provider_id_by_name(client_obj, header_obj, practice_id, provider_name_from_excel):
    if not all([client_obj, header_obj, practice_id, provider_name_from_excel]):
        display_message("error", "Missing parameters for Provider lookup.")
        return None
    provider_name_search = str(provider_name_from_excel).strip()
    if not provider_name_search:
        display_message("warning", "Provider name to search is empty.")
        return None
    cache_key = f"provider_id_{practice_id}_{provider_name_search}"
    if cache_key in st.session_state: return st.session_state[cache_key]
    provider_id_found = None
    with st.spinner(f"Finding ProviderID for '{provider_name_search}'..."):
        try:
            get_providers_req_type = client_obj.get_type('ns0:GetProvidersReq')
            provider_filter_type = client_obj.get_type('ns0:ProviderFilter')
            provider_fields_type = client_obj.get_type('ns0:ProviderFieldsToReturn')
            fields = provider_fields_type(ID=True, FullName=True, FirstName=True, LastName=True, Active=True, PracticeID=True, Type=True)
            exact_filter = provider_filter_type(FullName=provider_name_search, PracticeID=str(practice_id))
            resp_exact = client_obj.service.GetProviders(request=get_providers_req_type(RequestHeader=header_obj, Filter=exact_filter, Fields=fields))
            if not (hasattr(resp_exact, 'ErrorResponse') and resp_exact.ErrorResponse and resp_exact.ErrorResponse.IsError) and \
               not (hasattr(resp_exact, 'SecurityResponse') and resp_exact.SecurityResponse and not resp_exact.SecurityResponse.Authorized) and \
               hasattr(resp_exact, 'Providers') and resp_exact.Providers and hasattr(resp_exact.Providers, 'ProviderData') and resp_exact.Providers.ProviderData:
                for p_data in resp_exact.Providers.ProviderData:
                    is_active = (hasattr(p_data, 'Active') and p_data.Active is not None and ((isinstance(p_data.Active, bool) and p_data.Active) or (isinstance(p_data.Active, str) and p_data.Active.lower() == 'true')))
                    if is_active and p_data.FullName and p_data.FullName.strip().lower() == provider_name_search.lower() and p_data.ID:
                        provider_id_found = int(p_data.ID)
                        st.session_state[cache_key] = provider_id_found
                        return provider_id_found
            if provider_id_found is None:
                broad_filter = provider_filter_type(PracticeID=str(practice_id))
                resp_all = client_obj.service.GetProviders(request=get_providers_req_type(RequestHeader=header_obj, Filter=broad_filter, Fields=fields))
                if hasattr(resp_all, 'ErrorResponse') and resp_all.ErrorResponse and resp_all.ErrorResponse.IsError: raise Exception(f"API Error Broad Provider Search: {resp_all.ErrorResponse.ErrorMessage}")
                if hasattr(resp_all, 'SecurityResponse') and resp_all.SecurityResponse and not resp_all.SecurityResponse.Authorized: raise Exception(f"Auth Error Broad Provider Search: {resp_all.SecurityResponse.SecurityResult}")
                found_providers_flex = []
                if hasattr(resp_all, 'Providers') and resp_all.Providers and hasattr(resp_all.Providers, 'ProviderData') and resp_all.Providers.ProviderData:
                    terms = [t.lower() for t in provider_name_search.replace(',', '').replace('.', '').split() if t.lower() not in ['md', 'do', 'pa', 'np'] and t]
                    if not terms: terms = [provider_name_search.lower()]
                    for p_item in resp_all.Providers.ProviderData:
                        is_active = (hasattr(p_item, 'Active') and p_item.Active is not None and ((isinstance(p_item.Active, bool) and p_item.Active) or (isinstance(p_item.Active, str) and p_item.Active.lower() == 'true')))
                        if not is_active: continue
                        name_api = p_item.FullName.strip().lower() if p_item.FullName else ""
                        if not name_api: continue
                        score = 0
                        if name_api == provider_name_search.lower(): score = 100
                        elif terms: score = (sum(1 for t_term in terms if t_term in name_api) / len(terms)) * 90 if len(terms) > 0 else 0 # MODIFIED: Added len(terms) > 0 check
                        if score > 70 and p_item.ID is not None: found_providers_flex.append({"ID": int(p_item.ID), "FullName": p_item.FullName or "", "score": score})
                if found_providers_flex:
                    best = sorted(found_providers_flex, key=lambda x: x['score'], reverse=True)[0]
                    provider_id_found = best['ID']
            if provider_id_found is None: display_message("warning", f"Could not find suitable ACTIVE provider matching '{provider_name_search}'.")
        except SoapFault as sf: display_message("error", f"SOAP Fault GetProviders for '{provider_name_search}': {sf.message}")
        except ZeepLookupError as le: display_message("error", f"Zeep Type Error GetProviders: {le}")
        except Exception as e: display_message("error", f"Unexpected Error finding ProviderID for '{provider_name_search}': {e}")
        st.session_state[cache_key] = provider_id_found
        return provider_id_found

def get_location_id_by_name(client_obj, header_obj, practice_id, location_name_to_find):
    if not all([client_obj, header_obj, practice_id, location_name_to_find]): return None
    location_name = str(location_name_to_find).strip()
    if not location_name: return None
    cache_key = f"location_id_{practice_id}_{location_name}"
    if cache_key in st.session_state: return st.session_state[cache_key]
    with st.spinner(f"Finding LocationID for '{location_name}'..."):
        location_id = None
        try:
            req_type = client_obj.get_type('ns6:GetServiceLocationsReq') # Assuming ns6, adjust if needed
            filter_type = client_obj.get_type('ns6:ServiceLocationFilter')
            fields_type = client_obj.get_type('ns6:ServiceLocationFieldsToReturn')
            fields = fields_type(ID=True, Name=True, PracticeID=True) # Active field removed
            loc_filter = filter_type(PracticeID=str(practice_id))
            req = req_type(RequestHeader=header_obj, Filter=loc_filter, Fields=fields)
            resp = client_obj.service.GetServiceLocations(request=req)
            if hasattr(resp, 'ErrorResponse') and resp.ErrorResponse and resp.ErrorResponse.IsError: raise Exception(f"API Error: {resp.ErrorResponse.ErrorMessage}")
            if hasattr(resp, 'SecurityResponse') and resp.SecurityResponse and not resp.SecurityResponse.Authorized: raise Exception(f"Auth Error: {resp.SecurityResponse.SecurityResult}")
            if hasattr(resp, 'ServiceLocations') and resp.ServiceLocations and hasattr(resp.ServiceLocations, 'ServiceLocationData') and resp.ServiceLocations.ServiceLocationData:
                location_obj = next((loc for loc in resp.ServiceLocations.ServiceLocationData if loc.Name and loc.Name.strip().lower() == location_name.lower() and loc.ID), None)
                if location_obj: location_id = int(location_obj.ID)
                else: display_message("warning", f"Location '{location_name}' not found (case-insensitive name match).")
            else: display_message("warning", f"No service locations returned for PracticeID {practice_id}.")
        except Exception as e: display_message("error", f"Error GetLocationID for '{location_name}': {e}")
        st.session_state[cache_key] = location_id
        return location_id

def get_primary_case_for_patient(client_obj, header_obj, patient_id_to_fetch):
    cache_key = f"patient_case_{patient_id_to_fetch}"
    if cache_key in st.session_state: return st.session_state[cache_key]
    with st.spinner(f"Fetching Case info for Pt ID: {patient_id_to_fetch}..."):
        case_id_found = None
        try:
            get_patient_req_type = client_obj.get_type('ns0:GetPatientReq')
            filter_type = client_obj.get_type('ns0:SinglePatientFilter')
            p_filter = filter_type(PatientID=int(patient_id_to_fetch))
            request_data = get_patient_req_type(RequestHeader=header_obj, Filter=p_filter) # Default fields
            api_response = client_obj.service.GetPatient(request=request_data)
            if hasattr(api_response, 'ErrorResponse') and api_response.ErrorResponse and api_response.ErrorResponse.IsError: raise Exception(f"API Error: {api_response.ErrorResponse.ErrorMessage}")
            if hasattr(api_response, 'SecurityResponse') and api_response.SecurityResponse and not api_response.SecurityResponse.Authorized: raise Exception(f"Auth Error: {api_response.SecurityResponse.SecurityResult}")
            if hasattr(api_response, 'Patient') and api_response.Patient and hasattr(api_response.Patient, 'Cases') and \
               api_response.Patient.Cases and hasattr(api_response.Patient.Cases, 'PatientCaseData') and api_response.Patient.Cases.PatientCaseData:
                cases = api_response.Patient.Cases.PatientCaseData
                if not isinstance(cases, list): cases = [cases] # Ensure 'cases' is a list
                primary = next((c for c in cases if hasattr(c, 'IsPrimaryCase') and ((isinstance(c.IsPrimaryCase, bool) and c.IsPrimaryCase) or (isinstance(c.IsPrimaryCase, str) and c.IsPrimaryCase.lower() == 'true')) and hasattr(c, 'PatientCaseID') and c.PatientCaseID), None) # MODIFIED: Added hasattr checks
                if primary: case_id_found = int(primary.PatientCaseID)
                elif cases and hasattr(cases[0], 'PatientCaseID') and cases[0].PatientCaseID : case_id_found = int(cases[0].PatientCaseID)
        except Exception as e: display_message("error", f"Error fetching CaseID for Pt {patient_id_to_fetch}: {e}")
        if case_id_found is None: display_message("warning", f"No usable case found via API for Patient ID {patient_id_to_fetch}. Charge entry will fail.")
        st.session_state[cache_key] = case_id_found
        return case_id_found

# --- Payload Creation Functions ---
def create_patient_identifier_payload(c, p): return c.get_type('ns0:PatientIdentifierReq')(PatientID=int(p))
def create_provider_identifier_payload(c, p): return c.get_type('ns0:ProviderIdentifierDetailedReq')(ProviderID=int(p))
def create_service_location_payload(c, l): return c.get_type('ns0:EncounterServiceLocation')(LocationID=int(l))
def create_practice_identifier_payload(c, p): return c.get_type('ns0:PracticeIdentifierReq')(PracticeID=int(p))

def create_place_of_service_payload(client_obj, pos_val_excel):
    payload_type = client_obj.get_type('ns0:EncounterPlaceOfService')
    code, name = None, None
    if pd.isna(pos_val_excel) or not str(pos_val_excel).strip(): return None
    norm_pos = str(pos_val_excel).strip()
    if norm_pos.upper() in POS_CODE_MAP: code, name = POS_CODE_MAP[norm_pos.upper()]["code"], POS_CODE_MAP[norm_pos.upper()]["name"]
    elif norm_pos.isdigit() and len(norm_pos) <=2 : #MODIFIED: ensure it's a 2 digit code only
        code, name = norm_pos, next((d["name"] for _, d in POS_CODE_MAP.items() if d["code"] == norm_pos), norm_pos)
    else: # If not in map or not a simple 2-digit code, try to use it as is if it's short, else error
        if len(norm_pos) <= 2: # Accept as code if short
             code, name = norm_pos, norm_pos
             display_message("info", f"POS value '{norm_pos}' not in standard map, using directly as code and name.")
        else: # Too long to be a valid code if not in map
            display_message("error", f"POS value '{norm_pos}' is not in standard map and is not a valid 2-digit code format.")
            return None
    if not code: return None # Should be caught by earlier checks
    try: return payload_type(PlaceOfServiceCode=str(code), PlaceOfServiceName=str(name)) # Ensure strings
    except Exception:
        try: return payload_type(PlaceOfServiceCode=str(code))
        except Exception as e2: display_message("error", f"POS Payload Error: {e2}"); return None


def create_service_line_payload(client_obj, sld, start_dt, end_dt):
    slt = client_obj.get_type('ns0:ServiceLineReq')
    pc, u, d1 = str(sld.get(COL_PROCEDURES, "")).strip(), sld.get(COL_UNITS), str(sld.get(COL_DIAG1, "")).strip()
    if not pc or u is None or pd.isna(u) or not d1: return None
    try: uf = float(u)
    except ValueError: return None
    
    def clean_val(val, is_modifier=False): # MODIFIED: Combined cleaner
        s_val = str(val if pd.notna(val) else "").strip()
        if is_modifier and s_val.endswith(".0"): # Specific cleaning for modifiers like "59.0"
            s_val = s_val[:-2]
        return s_val if s_val and s_val.lower() != 'nan' else None

    m1 = clean_val(sld.get(COL_MOD1), is_modifier=True)
    m2 = clean_val(sld.get(COL_MOD2), is_modifier=True)
    
    args = {
        'ProcedureCode': pc, 'Units': uf, 
        'ServiceStartDate': format_datetime_for_api(start_dt), # Ensure consistent API format
        'ServiceEndDate': format_datetime_for_api(end_dt),     # Ensure consistent API format
        'DiagnosisCode1': clean_val(d1) # Use clean_val for diags too (for stripping and nan check)
    }
    if m1: args['ProcedureModifier1'] = m1
    if m2: args['ProcedureModifier2'] = m2
    diag2 = clean_val(sld.get(COL_DIAG2)); diag3 = clean_val(sld.get(COL_DIAG3)); diag4 = clean_val(sld.get(COL_DIAG4))
    if diag2: args['DiagnosisCode2'] = diag2
    if diag3: args['DiagnosisCode3'] = diag3
    if diag4: args['DiagnosisCode4'] = diag4
    
    try: return slt(**args)
    except Exception as e: display_message("error", f"SvcLine Payload Error for Proc {pc}: {e} with args {args}"); return None

# --- Main Processing Logic ---
def process_excel_data(client_obj, header_obj, current_practice_id, df_excel_data):
    processed_rows_data = df_excel_data.to_dict(orient='records')
    for r_idx, r_val in enumerate(processed_rows_data): # Use enumerate for index
        r_val['Charge Entry Status'], r_val['Reason for Failure'] = "Pending", ""
        processed_rows_data[r_idx] = r_val # Ensure update takes effect

    try:
        enc_type = client_obj.get_type('ns0:EncounterCreate')
        create_req_type = client_obj.get_type('ns0:CreateEncounterReq')
        case_id_type = client_obj.get_type('ns0:PatientCaseIdentifierReq')
        arr_sl_req_type = client_obj.get_type('ns0:ArrayOfServiceLineReq')
    except Exception as e:
        display_message("error", f"Fatal WSDL Type Error: {e}. Cannot process.")
        for r_idx in range(len(processed_rows_data)):
            processed_rows_data[r_idx]['Charge Entry Status'] = "Failed"
            processed_rows_data[r_idx]['Reason for Failure'] = f"WSDL Type Error: {e}"
        return processed_rows_data

    keys_to_clear = [k for k in st.session_state if k.startswith("provider_id_") or k.startswith("location_id_") or k.startswith("patient_case_")]
    for k in keys_to_clear: del st.session_state[k]

    # MODIFIED GROUPING LOGIC
    grouped_charges = defaultdict(lambda: {'encounter_details_source_row_dict': None, 
                                           'service_lines_data_list': [], 
                                           'original_df_indices': []})
    display_message("info", "Grouping Excel Rows by Patient, Date of Service, and Case...")
    pb_group = st.progress(0)
    grouping_warnings = []

    for df_idx, row_series in df_excel_data.iterrows(): # Use df_idx from iterrows()
        pb_group.progress((df_idx + 1) / len(df_excel_data))
        pid_val = row_series.get(COL_PATIENT_ID)
        from_date_val = row_series.get(COL_FROM_DATE) # Key for grouping

        current_row_status, current_row_reason = "Failed", ""

        if pd.isna(pid_val): current_row_reason = f"'{COL_PATIENT_ID}' missing."
        elif pd.isna(from_date_val): current_row_reason = f"'{COL_FROM_DATE}' missing (required for grouping)."
        else:
            try:
                pid_grp = int(pid_val)
                # Normalize from_date_val to string 'YYYY-MM-DD' for consistent grouping key
                from_date_str_key = pd.to_datetime(from_date_val).strftime('%Y-%m-%d')
            except ValueError:
                current_row_reason = f"'{COL_PATIENT_ID}' ('{pid_val}') or '{COL_FROM_DATE}' ('{from_date_val}') invalid format."
            else:
                cid_for_group = get_primary_case_for_patient(client_obj, header_obj, pid_grp)
                if not cid_for_group:
                    current_row_reason = f"No existing case found in Tebra for Patient {pid_grp}."
                else:
                    current_row_status = "Grouped"
                    # MODIFIED grp_key to include normalized From Date string
                    grp_key = (pid_grp, from_date_str_key, cid_for_group)
                    
                    row_dict = row_series.to_dict() # Convert current row to dict for storing
                    if not grouped_charges[grp_key]['encounter_details_source_row_dict']:
                        grouped_charges[grp_key]['encounter_details_source_row_dict'] = row_dict
                    grouped_charges[grp_key]['service_lines_data_list'].append(row_dict)
                    grouped_charges[grp_key]['original_df_indices'].append(df_idx) # Store original DataFrame index

        if current_row_status == "Failed":
             processed_rows_data[df_idx]['Charge Entry Status'] = current_row_status # Update specific row by index
             processed_rows_data[df_idx]['Reason for Failure'] = current_row_reason
             grouping_warnings.append(f"Row {df_idx+2}: {current_row_reason}") # df_idx is 0-based, Excel is 1-based + header
    
    pb_group.empty()
    if grouping_warnings: display_message("warning", "Issues during grouping:<br>" + "<br>".join(grouping_warnings))

    display_message("info", f"Processing {len(grouped_charges)} Grouped Encounters...")
    if not grouped_charges:
         if len(df_excel_data) > 0: display_message("warning", "No valid groups formed.")
         else: display_message("info", "No data rows found in file.")
         return processed_rows_data

    pb_proc = st.progress(0); proc_grp_cnt = 0; success_groups = 0; fail_groups = 0

    for grp_key, data_dict in grouped_charges.items(): # MODIFIED: Use consistent naming
        proc_grp_cnt += 1; pb_proc.progress(proc_grp_cnt / len(grouped_charges))
        # pid_for_api, from_date_key, case_id_for_api = grp_key # Unpack the new group key
        pid_for_api, _, case_id_for_api = grp_key # from_date_key is in enc_src for dates

        enc_src_dict = data_dict['encounter_details_source_row_dict'] # Renamed for clarity
        sl_to_proc_list = data_dict['service_lines_data_list']     # Renamed for clarity
        orig_indices_list = data_dict['original_df_indices']       # Renamed for clarity
        
        grp_status, grp_fail_reason = "Failed", "Group processing did not complete."
        log_ph = st.empty()
        # Use original Excel row numbers from the list for logging
        excel_row_nums_str = ", ".join([str(processed_rows_data[i].get('original_excel_row_num', i+2)) for i in orig_indices_list])
        log_ph.info(f"Processing Grp: Pt {pid_for_api}, Case {case_id_for_api}, DOS {enc_src_dict[COL_FROM_DATE]} (Orig. Excel Rows: {excel_row_nums_str})")


        try:
            rp_name = str(enc_src_dict.get(COL_RENDERING_PROVIDER, "")).strip()
            if not rp_name: raise ValueError(f"'{COL_RENDERING_PROVIDER}' missing.")
            rp_id = get_provider_id_by_name(client_obj, header_obj, current_practice_id, rp_name)
            if not rp_id: raise ValueError(f"Active Provider ID not found for '{rp_name}'.")

            loc_name = str(enc_src_dict.get(COL_LOCATION, "")).strip()
            if not loc_name: raise ValueError(f"'{COL_LOCATION}' missing.")
            loc_id = get_location_id_by_name(client_obj, header_obj, current_practice_id, loc_name)
            if not loc_id: raise ValueError(f"Location ID not found for '{loc_name}'.")

            sch_p_name = str(enc_src_dict.get(COL_SCHEDULING_PROVIDER, "")).strip()
            sch_p_pyld = None
            if sch_p_name:
                # log_ph.info(f"Grp {grp_key}: Looking up Scheduling Provider '{sch_p_name}'...") # Redundant with main log
                sch_p_id = get_provider_id_by_name(client_obj, header_obj, current_practice_id, sch_p_name)
                if not sch_p_id: display_message("warning", f"Grp (Pt {pid_for_api}, DOS {enc_src_dict[COL_FROM_DATE]}): Active Scheduling Provider ID NOT FOUND for '{sch_p_name}'. Encounter will be created without it.")
                else: sch_p_pyld = create_provider_identifier_payload(client_obj, sch_p_id)
            
            batch_num_val = str(enc_src_dict.get(COL_BATCH_NUMBER, "")).strip()
                
            enc_start_dt_api = format_datetime_for_api(enc_src_dict.get(COL_FROM_DATE))
            enc_end_dt_api = format_datetime_for_api(enc_src_dict.get(COL_THROUGH_DATE)) # ThroughDate still used for header
            if not enc_start_dt_api : raise ValueError(f"Encounter '{COL_FROM_DATE}' invalid for group.")
            if not enc_end_dt_api : enc_end_dt_api = enc_start_dt_api # Fallback if Through Date is missing/invalid

            pt_pyld = create_patient_identifier_payload(client_obj, pid_for_api)
            rp_pyld = create_provider_identifier_payload(client_obj, rp_id)
            sloc_pyld = create_service_location_payload(client_obj, loc_id)
            case_pyld_obj = case_id_type(CaseID=case_id_for_api)
            prac_pyld = create_practice_identifier_payload(client_obj, current_practice_id)

            pos_excel_val = str(enc_src_dict.get(COL_PLACE_OF_SERVICE_EXCEL, "")).strip()
            if not pos_excel_val:
                enc_mode_val = str(enc_src_dict.get(COL_ENCOUNTER_MODE, "")).strip()
                if enc_mode_val: pos_excel_val = enc_mode_val
                else: raise ValueError(f"'{COL_PLACE_OF_SERVICE_EXCEL}' & '{COL_ENCOUNTER_MODE}' missing.")
            pos_pyld = create_place_of_service_payload(client_obj, pos_excel_val)
            if not pos_pyld: raise ValueError(f"POS payload creation failed for '{pos_excel_val}'.")

            all_sl_objs = []
            line_errors_grp = [] # Renamed to avoid conflict
            for line_idx, sld_item_dict in enumerate(sl_to_proc_list):
                # Service line dates should be consistent with the encounter header for this group
                sl_start_dt = format_datetime_for_api(sld_item_dict.get(COL_FROM_DATE))
                sl_end_dt = format_datetime_for_api(sld_item_dict.get(COL_THROUGH_DATE))
                if not sl_start_dt: sl_start_dt = enc_start_dt_api # Default to encounter header date
                if not sl_end_dt: sl_end_dt = enc_end_dt_api     # Default to encounter header date

                # Crucial check: ensure service line dates match the group's DOS if strict adherence is needed.
                # For now, we use the dates from the line, defaulting to encounter header dates.
                # The grouping by FromDate should ensure consistency for enc_start_dt_api.

                sl_obj = create_service_line_payload(client_obj, sld_item_dict, sl_start_dt, sl_end_dt)
                if not sl_obj: 
                    line_errors_grp.append(f"L{line_idx+1}(Proc:{sld_item_dict.get(COL_PROCEDURES, 'N/A')}, ExcelRow:{processed_rows_data[orig_indices_list[line_idx]].get('original_excel_row_num', orig_indices_list[line_idx]+2)}): Creation Failed")
                    continue
                all_sl_objs.append(sl_obj)
            
            if line_errors_grp: raise ValueError("Service Line Errors: " + "; ".join(line_errors_grp))
            if not all_sl_objs: raise ValueError("No valid service lines created for this group.")

            sl_arr_pyld = arr_sl_req_type(ServiceLineReq=all_sl_objs)
            enc_args = {"Patient": pt_pyld, "RenderingProvider": rp_pyld, "ServiceLocation": sloc_pyld,
                    "PlaceOfService": pos_pyld, "ServiceStartDate": enc_start_dt_api, "ServiceEndDate": enc_end_dt_api,
                    "ServiceLines": sl_arr_pyld, "Practice": prac_pyld, "EncounterStatus": "Draft", "Case": case_pyld_obj}
            if sch_p_pyld: enc_args["SchedulingProvider"] = sch_p_pyld
            if batch_num_val: enc_args["BatchNumber"] = batch_num_val
            enc_pyld = enc_type(**enc_args)
            final_req = create_req_type(RequestHeader=header_obj, Encounter=enc_pyld)
            
            log_ph.info(f"Grp (Pt {pid_for_api}, DOS {enc_src_dict[COL_FROM_DATE]}): Calling CreateEncounter API...")
            api_resp = client_obj.service.CreateEncounter(request=final_req)

            if hasattr(api_resp, 'ErrorResponse') and api_resp.ErrorResponse and api_resp.ErrorResponse.IsError: grp_fail_reason = f"API Error: {api_resp.ErrorResponse.ErrorMessage}"
            elif hasattr(api_resp, 'SecurityResponse') and api_resp.SecurityResponse and not api_resp.SecurityResponse.Authorized: grp_fail_reason = f"Auth Error: {api_resp.SecurityResponse.SecurityResult}"
            elif hasattr(api_resp, 'EncounterID') and api_resp.EncounterID:
                grp_status, grp_fail_reason = "Done", f"{api_resp.EncounterID}"
                log_ph.success(f"Grp (Pt {pid_for_api}, DOS {enc_src_dict[COL_FROM_DATE]}): SUCCESS! EncounterID: {api_resp.EncounterID}")
                success_groups += 1
            else: grp_fail_reason = f"Unknown API resp: {zeep.helpers.serialize_object(api_resp,dict) if api_resp else 'None'}"
        
        except ValueError as ve: grp_fail_reason = str(ve)
        except SoapFault as sf: grp_fail_reason = f"SOAP FAULT: {sf.message} (Code: {sf.code})"
        except ZeepLookupError as le: grp_fail_reason = f"Zeep Type Lookup Error: {le}"
        except Exception as e: grp_fail_reason = f"UNEXPECTED SCRIPT ERROR: {type(e).__name__} - {e}"

        if grp_status == "Failed":
             log_ph.error(f"Grp (Pt {pid_for_api}, DOS {enc_src_dict[COL_FROM_DATE]}): FAILED. {grp_fail_reason}")
             fail_groups += 1
        
        for orig_df_idx in orig_indices_list: # Update original DataFrame rows
            processed_rows_data[orig_df_idx]['Charge Entry Status'] = grp_status
            processed_rows_data[orig_df_idx]['Reason for Failure'] = grp_fail_reason
        time.sleep(0.05)

    pb_proc.empty()
    summary_msg = f"Encounter processing finished. Groups Processed: {proc_grp_cnt}, Successful Groups: {success_groups}, Failed Groups: {fail_groups}."
    if fail_groups > 0 : display_message("warning", summary_msg + " Check 'Reason for Failure' column in results.")
    else: display_message("success", summary_msg)
    return processed_rows_data

# --- Streamlit Application UI ---
def main():
    apply_custom_styling()
    st.title(APP_TITLE) 
    st.subheader(APP_SUBTITLE)

    st.sidebar.header("Tebra Credentials")
    customer_key_val = st.sidebar.text_input("Customer Key", type="password", key="sb_customer_key")
    user_email_val = st.sidebar.text_input("Username (email)", key="sb_user_email")
    user_password_val = st.sidebar.text_input("Password", type="password", key="sb_user_password")

    st.sidebar.header("Upload Charge Data")
    uploaded_file_val = st.sidebar.file_uploader("Upload Excel File (.xlsx)", type="xlsx", key="sb_uploaded_file")
    
    # Initialize original_excel_row_num in session state if it doesn't exist
    if 'original_excel_row_num' not in st.session_state:
        st.session_state.original_excel_row_num = 2 # Start from row 2 (1-based index + header)

    process_button_val = st.sidebar.button("Process Charges", key="sb_process_button")
    results_placeholder = st.container()

    if process_button_val:
        results_placeholder.empty()
        with results_placeholder:
            if not customer_key_val or not user_email_val or not user_password_val: display_message("error", "❌ Please enter all Tebra credentials."); st.stop()
            if uploaded_file_val is None: display_message("error", "❌ Please upload an Excel file."); st.stop()
            credentials = {"CustomerKey": customer_key_val, "User": user_email_val, "Password": user_password_val}
            keys_to_clear = [k for k in st.session_state if k.startswith("practice_id_") or k.startswith("provider_id_") or k.startswith("location_id_") or k.startswith("patient_case_")]
            for k in keys_to_clear: del st.session_state[k]
            with st.spinner("Connecting to Tebra API and verifying practice..."):
                client = create_api_client(TEBRA_WSDL_URL)
                if not client: st.stop()
                header = build_request_header(credentials, client)
                if not header: st.stop()
                practice_id_check = get_practice_id_from_name(client, header, TEBRA_PRACTICE_NAME)
                if not practice_id_check: st.stop()
            display_message("success", f"✅ Connected to Tebra. Practice '{TEBRA_PRACTICE_NAME}' ID: {practice_id_check}.")
            try:
                df_excel = pd.read_excel(uploaded_file_val)
                df_excel.columns = df_excel.columns.str.strip()
                # Add original_excel_row_num for better error reporting context
                df_excel['original_excel_row_num'] = range(st.session_state.original_excel_row_num, st.session_state.original_excel_row_num + len(df_excel))

                missing_cols = [c for c in EXPECTED_COLUMNS if c not in df_excel.columns]
                if missing_cols: display_message("error", f"❌ Excel missing columns: {', '.join(missing_cols)}."); st.stop()
                display_message("info", "Excel loaded. Processing charges...")
                output_data = process_excel_data(client, header, practice_id_check, df_excel)
                if output_data:
                    df_results = pd.DataFrame(output_data)
                    st.subheader("Processing Results Summary")
                    total_rows = len(df_results)
                    success_rows = len(df_results[df_results['Charge Entry Status'] == 'Done'])
                    failed_rows = total_rows - success_rows
                    st.metric("Total Rows Processed", total_rows)
                    col1, col2 = st.columns(2)
                    col1.metric("Rows Successful", success_rows)
                    col2.metric("Rows Failed", failed_rows)
                    st.subheader("Detailed Results"); st.dataframe(df_results)
                    output_excel_io = io.BytesIO()
                    with pd.ExcelWriter(output_excel_io, engine='xlsxwriter') as writer:
                        df_results.to_excel(writer, index=False, sheet_name='ChargeEntryResults')
                    excel_bytes = output_excel_io.getvalue()
                    st.download_button(label="📥 Download Results Excel", data=excel_bytes,
                                       file_name=f"Tebra_Results_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       key="sb_download_excel_button")
                else: display_message("info", "ℹ️ No results data generated.")
            except Exception as e:
                display_message("error", f"❌ An unexpected error occurred: {e}")
                st.exception(e)
    st.markdown(f'<div class="footer">{APP_FOOTER}</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
