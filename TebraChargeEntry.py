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
MODIFIED: To change result column header and simplify API error messages.
MODIFIED: To control output column order and remove 'original_excel_row_num' from final Excel.
"""

import streamlit as st
import pandas as pd
import zeep
import zeep.helpers
from zeep.exceptions import Fault as SoapFault, TransportError, LookupError as ZeepLookupError
from requests.exceptions import ConnectionError as RequestsConnectionError
import datetime
import time
import re # For parsing XML errors
from collections import defaultdict
import io
from xml.etree import ElementTree as ET # For more robust XML parsing

# --- Application Configuration ---
APP_TITLE = "🤖 Tebra Charge Entry"
APP_SUBTITLE = "By Panacea Smart Solutions"
APP_FOOTER = "Tebra Charge Entry Tool © 2025 | Panacea Smart Solutions | Developed by Saqib Sherwani"
TEBRA_PRACTICE_NAME = "Pediatrics West" # Hardcoded Practice Name
TEBRA_WSDL_URL = "https://webservice.kareo.com/services/soap/2.1/KareoServices.svc?singleWsdl"

# --- SET PAGE CONFIG MUST BE THE FIRST STREAMLIT COMMAND ---
st.set_page_config(page_title=APP_TITLE, page_icon="🤖", layout="wide", initial_sidebar_state="expanded")

# --- Expected Excel Columns (from the original app structure) ---
# This list defines what columns the script expects to find in the input Excel.
# The actual processing logic will determine if a row has enough data for an encounter.
EXPECTED_COLUMNS = [
    'Patient ID', 'From Date', 'Through Date', 'Rendering Provider', 'Scheduling Provider',
    'Location', 'Place of Service', 'Encounter Mode', 'Procedures', 'Mod 1', 'Mod 2',
    'Units', 'Diag 1', 'Diag 2', 'Diag 3', 'Diag 4', 'Batch Number'
]
# Define column name constants for consistency (matching the input Excel)
COL_PATIENT_ID = 'Patient ID'; COL_FROM_DATE = 'From Date'; COL_THROUGH_DATE = 'Through Date'
COL_RENDERING_PROVIDER = 'Rendering Provider'; COL_SCHEDULING_PROVIDER = 'Scheduling Provider'
COL_LOCATION = 'Location'; COL_PLACE_OF_SERVICE_EXCEL = 'Place of Service'
COL_ENCOUNTER_MODE = 'Encounter Mode'; COL_PROCEDURES = 'Procedures'; COL_MOD1 = 'Mod 1'; COL_MOD2 = 'Mod 2'
COL_UNITS = 'Units'; COL_DIAG1 = 'Diag 1'; COL_DIAG2 = 'Diag 2'
COL_DIAG3 = 'Diag 3'; COL_DIAG4 = 'Diag 4'; COL_BATCH_NUMBER = 'Batch Number'

# NEW Column name for results - MODIFIED as per user request
COL_RESULT_MESSAGE = "Encounter ID or Reasons for Failure"


# --- Place of Service Mapping ---
POS_CODE_MAP = {
    "OFFICE": {"code": "11", "name": "Office"},
    "IN OFFICE": {"code": "11", "name": "Office"},
    "INOFFICE": {"code": "11", "name": "Office"}, 
    "TELEHEALTH": {"code": "10", "name": "Telehealth Provided in Patient’s Home"},
    "TELEHEALTH OFFICE": {"code": "02", "name": "Telehealth Provided Other than in Patient’s Home"},
}

def apply_custom_styling():
    # Styles remain the same as in TebraChargeEntry_v1.txt
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
    except Exception as e: 
        display_message("warning", f"Date parse warning: '{date_value}'. Error: {e}. Using None.")
        return None

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
                practices_data = resp.Practices.PracticeData
                if not isinstance(practices_data, list): practices_data = [practices_data]
                practice_obj = next((p for p in practices_data if hasattr(p,'PracticeName') and p.PracticeName and p.PracticeName.strip().lower() == practice_name_to_find.strip().lower() and hasattr(p,'ID') and p.ID), None)
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
            
            exact_providers_data = []
            if hasattr(resp_exact, 'Providers') and resp_exact.Providers and hasattr(resp_exact.Providers, 'ProviderData') and resp_exact.Providers.ProviderData:
                exact_providers_data = resp_exact.Providers.ProviderData
                if not isinstance(exact_providers_data, list): exact_providers_data = [exact_providers_data]

            if not (hasattr(resp_exact, 'ErrorResponse') and resp_exact.ErrorResponse and resp_exact.ErrorResponse.IsError) and \
               not (hasattr(resp_exact, 'SecurityResponse') and resp_exact.SecurityResponse and not resp_exact.SecurityResponse.Authorized) and \
               exact_providers_data:
                for p_data in exact_providers_data:
                    is_active = (hasattr(p_data, 'Active') and p_data.Active is not None and ((isinstance(p_data.Active, bool) and p_data.Active) or (isinstance(p_data.Active, str) and p_data.Active.lower() == 'true')))
                    if is_active and hasattr(p_data,'FullName') and p_data.FullName and p_data.FullName.strip().lower() == provider_name_search.lower() and hasattr(p_data,'ID') and p_data.ID:
                        provider_id_found = int(p_data.ID)
                        st.session_state[cache_key] = provider_id_found
                        return provider_id_found
            
            if provider_id_found is None:
                broad_filter = provider_filter_type(PracticeID=str(practice_id))
                resp_all = client_obj.service.GetProviders(request=get_providers_req_type(RequestHeader=header_obj, Filter=broad_filter, Fields=fields))
                if hasattr(resp_all, 'ErrorResponse') and resp_all.ErrorResponse and resp_all.ErrorResponse.IsError: raise Exception(f"API Error Broad Provider Search: {resp_all.ErrorResponse.ErrorMessage}")
                if hasattr(resp_all, 'SecurityResponse') and resp_all.SecurityResponse and not resp_all.SecurityResponse.Authorized: raise Exception(f"Auth Error Broad Provider Search: {resp_all.SecurityResponse.SecurityResult}")
                
                found_providers_flex = []
                all_providers_data_broad = []
                if hasattr(resp_all, 'Providers') and resp_all.Providers and hasattr(resp_all.Providers, 'ProviderData') and resp_all.Providers.ProviderData:
                    all_providers_data_broad = resp_all.Providers.ProviderData
                    if not isinstance(all_providers_data_broad, list): all_providers_data_broad = [all_providers_data_broad]

                if all_providers_data_broad:
                    terms = [t.lower() for t in provider_name_search.replace(',', '').replace('.', '').split() if t.lower() not in ['md', 'do', 'pa', 'np'] and t]
                    if not terms: terms = [provider_name_search.lower()]
                    for p_item in all_providers_data_broad:
                        is_active = (hasattr(p_item, 'Active') and p_item.Active is not None and ((isinstance(p_item.Active, bool) and p_item.Active) or (isinstance(p_item.Active, str) and p_item.Active.lower() == 'true')))
                        if not is_active: continue
                        name_api = p_item.FullName.strip().lower() if hasattr(p_item,'FullName') and p_item.FullName else ""
                        if not name_api: continue
                        score = 0
                        if name_api == provider_name_search.lower(): score = 100
                        elif terms: score = (sum(1 for t_term in terms if t_term in name_api) / len(terms)) * 90 if len(terms) > 0 else 0
                        if score > 70 and hasattr(p_item,'ID') and p_item.ID is not None: found_providers_flex.append({"ID": int(p_item.ID), "FullName": p_item.FullName or "", "score": score})
                
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
            req_type = client_obj.get_type('ns6:GetServiceLocationsReq')
            filter_type = client_obj.get_type('ns6:ServiceLocationFilter')
            fields_type = client_obj.get_type('ns6:ServiceLocationFieldsToReturn')
            fields = fields_type(ID=True, Name=True, PracticeID=True)
            loc_filter = filter_type(PracticeID=str(practice_id))
            req = req_type(RequestHeader=header_obj, Filter=loc_filter, Fields=fields)
            resp = client_obj.service.GetServiceLocations(request=req)
            if hasattr(resp, 'ErrorResponse') and resp.ErrorResponse and resp.ErrorResponse.IsError: raise Exception(f"API Error: {resp.ErrorResponse.ErrorMessage}")
            if hasattr(resp, 'SecurityResponse') and resp.SecurityResponse and not resp.SecurityResponse.Authorized: raise Exception(f"Auth Error: {resp.SecurityResponse.SecurityResult}")
            
            if hasattr(resp, 'ServiceLocations') and resp.ServiceLocations and hasattr(resp.ServiceLocations, 'ServiceLocationData') and resp.ServiceLocations.ServiceLocationData:
                locations_data = resp.ServiceLocations.ServiceLocationData
                if not isinstance(locations_data, list): locations_data = [locations_data]
                location_obj = next((loc for loc in locations_data if hasattr(loc, 'Name') and loc.Name and loc.Name.strip().lower() == location_name.lower() and hasattr(loc, 'ID') and loc.ID), None)
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
            request_data = get_patient_req_type(RequestHeader=header_obj, Filter=p_filter)
            api_response = client_obj.service.GetPatient(request=request_data)
            if hasattr(api_response, 'ErrorResponse') and api_response.ErrorResponse and api_response.ErrorResponse.IsError: raise Exception(f"API Error: {api_response.ErrorResponse.ErrorMessage}")
            if hasattr(api_response, 'SecurityResponse') and api_response.SecurityResponse and not api_response.SecurityResponse.Authorized: raise Exception(f"Auth Error: {api_response.SecurityResponse.SecurityResult}")
            if hasattr(api_response, 'Patient') and api_response.Patient and hasattr(api_response.Patient, 'Cases') and \
               api_response.Patient.Cases and hasattr(api_response.Patient.Cases, 'PatientCaseData') and api_response.Patient.Cases.PatientCaseData:
                cases = api_response.Patient.Cases.PatientCaseData
                if not isinstance(cases, list): cases = [cases]
                primary = next((c for c in cases if hasattr(c, 'IsPrimaryCase') and ((isinstance(c.IsPrimaryCase, bool) and c.IsPrimaryCase) or (isinstance(c.IsPrimaryCase, str) and c.IsPrimaryCase.lower() == 'true')) and hasattr(c, 'PatientCaseID') and c.PatientCaseID), None)
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
    if pd.isna(pos_val_excel) or not str(pos_val_excel).strip(): 
        display_message("warning", "Place of Service value is blank in Excel. Cannot create POS payload.")
        return None
    norm_pos_input = str(pos_val_excel).strip()
    norm_pos_upper = norm_pos_input.upper()

    if norm_pos_upper in POS_CODE_MAP:
        code = POS_CODE_MAP[norm_pos_upper]["code"]
        name = POS_CODE_MAP[norm_pos_upper]["name"]
    elif norm_pos_input.isdigit() and len(norm_pos_input) <= 2:
        code = norm_pos_input
        name = next((d["name"] for _, d in POS_CODE_MAP.items() if d["code"] == code), norm_pos_input) 
    else:
        display_message("error", f"POS value '{norm_pos_input}' is not in standard map and not a valid 2-digit code format.")
        return None
        
    if not code: return None 
    try: return payload_type(PlaceOfServiceCode=str(code), PlaceOfServiceName=str(name))
    except Exception:
        try: return payload_type(PlaceOfServiceCode=str(code))
        except Exception as e2: display_message("error", f"POS Payload Error: {e2}"); return None

def create_service_line_payload(client_obj, sld, start_dt_api_str, end_dt_api_str):
    slt = client_obj.get_type('ns0:ServiceLineReq')
    pc = str(sld.get(COL_PROCEDURES, "")).strip()
    u = sld.get(COL_UNITS)
    d1_val = sld.get(COL_DIAG1) 

    row_num_for_log = sld.get('original_excel_row_num', 'N/A')

    if not pc: 
        display_message("warning", f"Row {row_num_for_log}: Procedure code missing. SvcLine not created.")
        return None
    if u is None or pd.isna(u) or str(u).strip() == "":
        display_message("warning", f"Row {row_num_for_log} (Proc {pc}): Units missing or blank. SvcLine not created.")
        return None
    if d1_val is None or pd.isna(d1_val) or not str(d1_val).strip():
        display_message("warning", f"Row {row_num_for_log} (Proc {pc}): Diag1 missing or blank. SvcLine not created.")
        return None
    
    try: 
        uf = float(u)
        if uf <= 0: 
            display_message("warning", f"Row {row_num_for_log} (Proc {pc}): Units must be > 0. Value: {uf}. SvcLine not created.")
            return None
    except ValueError: 
        display_message("warning", f"Row {row_num_for_log} (Proc {pc}): Units '{u}' not valid number. SvcLine not created.")
        return None
    
    def clean_val(val, is_modifier=False, is_diag=False):
        s_val = str(val if pd.notna(val) else "").strip()
        if is_modifier:
            if s_val.endswith(".0"): s_val = s_val[:-2]
            if s_val and len(s_val) > 2 :
                 display_message("warning", f"Row {row_num_for_log} (Proc {pc}): Modifier '{s_val}' is longer than 2 characters. Using first 2: '{s_val[:2]}'.")
                 s_val = s_val[:2]
        return s_val if s_val and s_val.lower() != 'nan' else None

    m1 = clean_val(sld.get(COL_MOD1), is_modifier=True)
    m2 = clean_val(sld.get(COL_MOD2), is_modifier=True)
    d1_cleaned = clean_val(d1_val, is_diag=True) 

    if not d1_cleaned: 
        display_message("warning", f"Row {row_num_for_log} (Proc {pc}): Diag1 became invalid after cleaning. SvcLine not created.")
        return None

    args = {
        'ProcedureCode': pc, 'Units': uf, 
        'ServiceStartDate': start_dt_api_str, 
        'ServiceEndDate': end_dt_api_str,    
        'DiagnosisCode1': d1_cleaned
    }
    if m1: args['ProcedureModifier1'] = m1
    if m2: args['ProcedureModifier2'] = m2
    
    diag2 = clean_val(sld.get(COL_DIAG2), is_diag=True); 
    diag3 = clean_val(sld.get(COL_DIAG3), is_diag=True); 
    diag4 = clean_val(sld.get(COL_DIAG4), is_diag=True)
    
    if diag2: args['DiagnosisCode2'] = diag2
    if diag3: args['DiagnosisCode3'] = diag3
    if diag4: args['DiagnosisCode4'] = diag4
    
    try: return slt(**args)
    except Exception as e: 
        display_message("error", f"SvcLine Payload Error for Proc {pc} (Row {row_num_for_log}): {e} with args {args}")
        return None

# --- XML Error Parsing Function ---
def parse_and_simplify_tebra_xml_error(xml_string, patient_id_context="N/A", dos_context="N/A"):
    if not xml_string or not isinstance(xml_string, str) or "<Encounter" not in xml_string :
        return xml_string 

    simplified_errors = []
    try:
        if xml_string.startswith("API Error: "):
            xml_string = xml_string[len("API Error: "):].strip()
        
        service_line_pattern = r"<ServiceLine>(.*?)</ServiceLine>"
        diag_error_pattern = r"<DiagnosisCode(?P<diag_num>\d)>(?P<diag_code_val>[^<]+?)<err id=\"\d+\">(?P<err_msg>[^<]+)</err>"
        mod_error_pattern = r"<ProcedureModifier(?P<mod_num>\d)>(?P<mod_code_val>[^<]+?)<err id=\"\d+\">(?P<err_msg>[^<]+)</err>"
        proc_code_pattern = r"<ProcedureCode>(?P<proc_code_val>[^<]+)</ProcedureCode>"

        service_lines_xml = re.findall(service_line_pattern, xml_string, re.DOTALL)
        
        sl_counter_for_log = 0
        for sl_xml_content in service_lines_xml:
            sl_counter_for_log += 1 
            proc_code = "N/A"
            proc_match = re.search(proc_code_pattern, sl_xml_content)
            if proc_match:
                proc_code = proc_match.group("proc_code_val")

            for match in re.finditer(diag_error_pattern, sl_xml_content):
                group_dict = match.groupdict()
                simple_msg = group_dict['err_msg'].split(',')[0].strip() 
                simplified_errors.append(f"L{sl_counter_for_log} (Proc {proc_code}): Diag {group_dict['diag_num']} ('{group_dict['diag_code_val']}') - {simple_msg}.")
            
            for match in re.finditer(mod_error_pattern, sl_xml_content):
                group_dict = match.groupdict()
                simple_msg = group_dict['err_msg'].split(',')[0].strip()
                mod_val = group_dict['mod_code_val']
                error_text = f"L{sl_counter_for_log} (Proc {proc_code}): Mod {group_dict['mod_num']} ('{mod_val}') - {simple_msg}."
                if ".0" in mod_val:
                    error_text += " (Note: Modifiers should be 2 chars, e.g., '59' not '59.0')."
                simplified_errors.append(error_text)

        overall_encounter_error_match = re.search(r'<err id="6100">(.*?)</err>', xml_string)
        if overall_encounter_error_match :
            if not simplified_errors: # Only add general if no specific line errors found
                 simplified_errors.append(f"Encounter Creation Failed: {overall_encounter_error_match.group(1).strip()}.")

        if not simplified_errors and xml_string.startswith("<Encounter"): 
            simplified_errors.append("Encounter creation failed. No specific line errors parsed. Review raw API response.")
        elif not simplified_errors: 
            return xml_string 

        return "; ".join(list(set(simplified_errors))) if simplified_errors else "Encounter processing failed with unspecified service line errors."
        
    except Exception as e_parse:
        return "Error simplifying API message. Original: " + xml_string[:300] + "..."


# --- Main Processing Logic (MODIFIED for DOS Grouping and Error Simplification) ---
def process_excel_data(client_obj, header_obj, current_practice_id, df_excel_data):
    df_results = df_excel_data.copy()
    df_results['Charge Entry Status'] = "Pending" 
    df_results[COL_RESULT_MESSAGE] = "" 

    try:
        enc_type = client_obj.get_type('ns0:EncounterCreate')
        create_req_type = client_obj.get_type('ns0:CreateEncounterReq')
        case_id_type = client_obj.get_type('ns0:PatientCaseIdentifierReq') 
        arr_sl_req_type = client_obj.get_type('ns0:ArrayOfServiceLineReq')
    except Exception as e:
        display_message("error", f"Fatal WSDL Type Error: {e}. Cannot process.")
        df_results['Charge Entry Status'] = "Failed"
        df_results[COL_RESULT_MESSAGE] = f"WSDL Type Error: {e}"
        # Define expected output columns for consistent error output
        output_columns = df_excel_data.columns.tolist() + ['original_excel_row_num', 'Charge Entry Status', COL_RESULT_MESSAGE]
        # Remove duplicates while preserving order
        output_columns = sorted(list(set(output_columns)), key=lambda x: output_columns.index(x) if x in output_columns else float('inf'))
        for col in output_columns:
            if col not in df_results.columns: df_results[col] = None # Add if missing
        return df_results.reindex(columns=output_columns).fillna(''), 0, 0


    keys_to_clear = [k for k in st.session_state if k.startswith("provider_id_") or k.startswith("location_id_") or k.startswith("patient_case_")]
    for k in keys_to_clear: 
        if k in st.session_state: del st.session_state[k]

    grouped_charges = defaultdict(lambda: {'encounter_details_source_row_dict': None, 
                                           'service_lines_data_list': [], 
                                           'original_df_indices': []})
    
    display_message("info", "Grouping Excel Rows by Patient, From Date (DOS), and Case...")
    pb_group = st.progress(0)
    grouping_warnings = []

    for df_idx, row_series in df_excel_data.iterrows():
        pb_group.progress((df_idx + 1) / len(df_excel_data))
        
        pid_val = row_series.get(COL_PATIENT_ID)
        from_date_val = row_series.get(COL_FROM_DATE) 

        current_row_grouping_status, current_row_grouping_reason = "Failed to Group", ""

        if pd.isna(pid_val) or str(pid_val).strip() == "":
            current_row_grouping_reason = f"'{COL_PATIENT_ID}' is missing."
        elif pd.isna(from_date_val) or str(from_date_val).strip() == "": 
            current_row_grouping_reason = f"'{COL_FROM_DATE}' (Date of Service) is missing for grouping."
        else:
            try:
                pid_grp = int(pid_val)
                from_date_str_key = pd.to_datetime(from_date_val).strftime('%Y-%m-%d')
            except ValueError:
                current_row_grouping_reason = f"Invalid format for '{COL_PATIENT_ID}' ('{pid_val}') or '{COL_FROM_DATE}' ('{from_date_val}')."
            except Exception as e_date_parse: 
                 current_row_grouping_reason = f"Error parsing '{COL_FROM_DATE}' ('{from_date_val}') for grouping key: {e_date_parse}."
            else:
                cid_for_group = get_primary_case_for_patient(client_obj, header_obj, pid_grp)
                if not cid_for_group:
                    current_row_grouping_reason = f"No existing Tebra case found for Patient ID {pid_grp}."
                else:
                    current_row_grouping_status = "Grouped Successfully"
                    grp_key = (pid_grp, from_date_str_key, cid_for_group) 
                    
                    row_dict_for_group = row_series.to_dict()
                    row_dict_for_group['original_excel_row_num'] = df_excel_data.loc[df_idx, 'original_excel_row_num']
                    
                    if not grouped_charges[grp_key]['encounter_details_source_row_dict']:
                        grouped_charges[grp_key]['encounter_details_source_row_dict'] = row_dict_for_group
                    grouped_charges[grp_key]['service_lines_data_list'].append(row_dict_for_group) 
                    grouped_charges[grp_key]['original_df_indices'].append(df_idx) 
        
        if current_row_grouping_status == "Failed to Group":
             df_results.loc[df_idx, 'Charge Entry Status'] = "Failed"
             df_results.loc[df_idx, COL_RESULT_MESSAGE] = current_row_grouping_reason
             grouping_warnings.append(f"Row {row_series.get('original_excel_row_num', df_idx+2)}: {current_row_grouping_reason}")
    
    pb_group.empty()
    if grouping_warnings: display_message("warning", "Issues during grouping stage:<br>" + "<br>".join(grouping_warnings))

    valid_groups_to_process = {k: v for k, v in grouped_charges.items() if v['encounter_details_source_row_dict'] is not None}

    display_message("info", f"Processing {len(valid_groups_to_process)} unique encounter groups...")
    
    proc_grp_cnt = 0
    success_groups = 0 
    fail_groups = 0    

    if not valid_groups_to_process and not grouping_warnings:
         if len(df_excel_data) > 0: display_message("warning", "No valid encounter groups were formed for processing.")
         else: display_message("info", "No data rows found in the uploaded file to process.")
         # Ensure consistent output columns even if no groups processed
         final_cols_on_no_groups = df_excel_data.columns.tolist() + ['original_excel_row_num', 'Charge Entry Status', COL_RESULT_MESSAGE]
         final_cols_on_no_groups = sorted(list(set(final_cols_on_no_groups)), key=lambda x: final_cols_on_no_groups.index(x) if x in final_cols_on_no_groups else float('inf'))
         for col in final_cols_on_no_groups:
            if col not in df_results.columns: df_results[col] = None
         return df_results.reindex(columns=final_cols_on_no_groups).fillna(''), success_groups, fail_groups


    pb_proc = st.progress(0)

    for grp_key, data_dict in valid_groups_to_process.items():
        proc_grp_cnt += 1; pb_proc.progress(proc_grp_cnt / len(valid_groups_to_process))
        pid_for_api, dos_key_str, case_id_for_api = grp_key

        enc_src_dict = data_dict['encounter_details_source_row_dict']
        sl_to_proc_list = data_dict['service_lines_data_list']
        orig_indices_list = data_dict['original_df_indices']
        
        grp_api_status, grp_api_message = "Failed", "Group processing did not complete."
        log_ph = st.empty()
        
        first_row_excel_num_for_log = df_results.loc[orig_indices_list[0], 'original_excel_row_num']
        log_ph.info(f"Processing Group: Pt {pid_for_api}, DOS {dos_key_str}, Case {case_id_for_api} (Excel Rows ~{first_row_excel_num_for_log})")

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
                sch_p_id = get_provider_id_by_name(client_obj, header_obj, current_practice_id, sch_p_name)
                if not sch_p_id: display_message("warning", f"Group (Pt {pid_for_api}, DOS {dos_key_str}): Active Scheduling Provider ID NOT FOUND for '{sch_p_name}'. Encounter will omit it.")
                else: sch_p_pyld = create_provider_identifier_payload(client_obj, sch_p_id)
            
            batch_num_val = str(enc_src_dict.get(COL_BATCH_NUMBER, "")).strip() if pd.notna(enc_src_dict.get(COL_BATCH_NUMBER)) else None
                
            enc_start_dt_api = format_datetime_for_api(enc_src_dict.get(COL_FROM_DATE))
            enc_end_dt_api = format_datetime_for_api(enc_src_dict.get(COL_THROUGH_DATE))
            if not enc_start_dt_api : raise ValueError(f"Encounter '{COL_FROM_DATE}' invalid for group (key: {dos_key_str}).")
            if not enc_end_dt_api : enc_end_dt_api = enc_start_dt_api

            pt_pyld = create_patient_identifier_payload(client_obj, pid_for_api)
            rp_pyld = create_provider_identifier_payload(client_obj, rp_id)
            sloc_pyld = create_service_location_payload(client_obj, loc_id)
            case_pyld_obj = case_id_type(CaseID=case_id_for_api)
            prac_pyld = create_practice_identifier_payload(client_obj, current_practice_id)

            pos_excel_val = str(enc_src_dict.get(COL_PLACE_OF_SERVICE_EXCEL, "")).strip()
            if not pos_excel_val:
                enc_mode_val = str(enc_src_dict.get(COL_ENCOUNTER_MODE, "")).strip()
                if enc_mode_val: pos_excel_val = enc_mode_val
                else: raise ValueError(f"Both '{COL_PLACE_OF_SERVICE_EXCEL}' & '{COL_ENCOUNTER_MODE}' are missing.")
            pos_pyld = create_place_of_service_payload(client_obj, pos_excel_val)
            if not pos_pyld: raise ValueError(f"POS payload creation failed for '{pos_excel_val}'.")

            all_sl_objs = []
            line_errors_grp = []
            for line_idx, sld_item_dict in enumerate(sl_to_proc_list):
                if 'original_excel_row_num' not in sld_item_dict:
                     sld_item_dict['original_excel_row_num'] = df_results.loc[orig_indices_list[line_idx], 'original_excel_row_num']

                sl_obj = create_service_line_payload(client_obj, sld_item_dict, enc_start_dt_api, enc_end_dt_api)
                if not sl_obj: 
                    line_errors_grp.append(f"SvcLine for Proc '{sld_item_dict.get(COL_PROCEDURES, 'N/A')}' (orig Excel row {sld_item_dict.get('original_excel_row_num','N/A')}) failed creation.")
                    continue
                all_sl_objs.append(sl_obj)
            
            if line_errors_grp: raise ValueError("Service Line Payload Errors: " + "; ".join(line_errors_grp))
            if not all_sl_objs: raise ValueError("No valid service lines were created for this encounter.")

            sl_arr_pyld = arr_sl_req_type(ServiceLineReq=all_sl_objs)
            enc_args = {
                "Patient": pt_pyld, "RenderingProvider": rp_pyld, "ServiceLocation": sloc_pyld,
                "PlaceOfService": pos_pyld, "ServiceStartDate": enc_start_dt_api, 
                "ServiceEndDate": enc_end_dt_api, "ServiceLines": sl_arr_pyld, 
                "Practice": prac_pyld, "EncounterStatus": "Draft", "Case": case_pyld_obj
            }
            if sch_p_pyld: enc_args["SchedulingProvider"] = sch_p_pyld
            if batch_num_val: enc_args["BatchNumber"] = batch_num_val
            
            enc_pyld_obj = enc_type(**enc_args) 
            final_req = create_req_type(RequestHeader=header_obj, Encounter=enc_pyld_obj) 
            
            log_ph.info(f"Group (Pt {pid_for_api}, DOS {dos_key_str}): Calling CreateEncounter API...")
            api_resp = client_obj.service.CreateEncounter(request=final_req)

            if hasattr(api_resp, 'ErrorResponse') and api_resp.ErrorResponse and api_resp.ErrorResponse.IsError:
                grp_api_message = parse_and_simplify_tebra_xml_error(api_resp.ErrorResponse.ErrorMessage, pid_for_api, dos_key_str)
            elif hasattr(api_resp, 'SecurityResponse') and api_resp.SecurityResponse and not api_resp.SecurityResponse.Authorized:
                grp_api_message = f"API Auth Error: {api_resp.SecurityResponse.SecurityResult}"
            elif hasattr(api_resp, 'EncounterID') and api_resp.EncounterID is not None:
                grp_api_status = "Done" 
                grp_api_message = f"{api_resp.EncounterID}" 
                log_ph.success(f"Group (Pt {pid_for_api}, DOS {dos_key_str}): SUCCESS! EncounterID: {api_resp.EncounterID}")
                success_groups += 1
            else:
                raw_resp_str = str(zeep.helpers.serialize_object(api_resp, dict) if api_resp else 'None')
                grp_api_message = f"Unknown API response: {raw_resp_str[:250]}..."
        
        except ValueError as ve: grp_api_message = str(ve)
        except SoapFault as sf: grp_api_message = f"SOAP FAULT: {sf.message} (Code: {sf.code})"
        except ZeepLookupError as le: grp_api_message = f"Zeep Type Lookup Error (WSDL issue?): {le}"
        except Exception as e: 
            grp_api_message = f"UNEXPECTED SCRIPT ERROR: {type(e).__name__} - {str(e)[:150]}"
        
        if grp_api_status == "Failed":
             log_ph.error(f"Group (Pt {pid_for_api}, DOS {dos_key_str}): FAILED. {grp_api_message}")
             fail_groups += 1
        
        for orig_df_idx in orig_indices_list:
            df_results.loc[orig_df_idx, 'Charge Entry Status'] = grp_api_status
            df_results.loc[orig_df_idx, COL_RESULT_MESSAGE] = grp_api_message
        
        time.sleep(0.05)

    pb_proc.empty()
    summary_msg = f"Encounter processing finished. Groups Processed: {proc_grp_cnt}, Successful Groups: {success_groups}, Failed Groups: {fail_groups}."
    if fail_groups > 0 : display_message("warning", summary_msg + f" Check '{COL_RESULT_MESSAGE}' column in results.")
    else: display_message("success", summary_msg)
    
    # --- Define the final output column order ---
    # Start with the original columns from the input Excel, in their original order
    output_column_order = df_excel_data.columns.tolist()
    # Append the two new status columns at the end
    if 'Charge Entry Status' not in output_column_order: # Should be added by df_results init
        output_column_order.append('Charge Entry Status')
    if COL_RESULT_MESSAGE not in output_column_order: # New result column
        output_column_order.append(COL_RESULT_MESSAGE)
    
    # Remove 'original_excel_row_num' if it exists, as it's not for final output
    if 'original_excel_row_num' in output_column_order:
        output_column_order.remove('original_excel_row_num')
    
    # Ensure all columns in df_results are included, even if not in original_input_columns
    # This handles any unexpected columns that might have been added.
    for col in df_results.columns:
        if col not in output_column_order and col != 'original_excel_row_num':
            output_column_order.append(col)
            
    return df_results.reindex(columns=output_column_order).fillna(''), success_groups, fail_groups


# --- Streamlit Application UI ---
def main():
    apply_custom_styling()
    st.title(APP_TITLE) 
    st.subheader(APP_SUBTITLE)

    st.sidebar.header("Tebra Credentials")
    customer_key_val = st.sidebar.text_input("Customer Key", type="password", key="sb_customer_key_v5") 
    user_email_val = st.sidebar.text_input("Username (email)", key="sb_user_email_v5")
    user_password_val = st.sidebar.text_input("Password", type="password", key="sb_user_password_v5")

    st.sidebar.header("Upload Charge Data")
    uploaded_file_val = st.sidebar.file_uploader("Upload Excel File (.xlsx)", type="xlsx", key="sb_uploaded_file_v5")
    
    if 'original_excel_row_num_start' not in st.session_state: 
        st.session_state.original_excel_row_num_start = 2 

    process_button_val = st.sidebar.button("Process Charges", key="sb_process_button_v5")
    results_placeholder = st.container()

    if process_button_val:
        results_placeholder.empty()
        with results_placeholder:
            if not customer_key_val or not user_email_val or not user_password_val: display_message("error", "❌ Please enter all Tebra credentials."); st.stop()
            if uploaded_file_val is None: display_message("error", "❌ Please upload an Excel file."); st.stop()
            
            credentials = {"CustomerKey": customer_key_val, "User": user_email_val, "Password": user_password_val}
            
            keys_to_clear_main = [k for k in st.session_state if k.startswith("practice_id_") or k.startswith("provider_id_") or k.startswith("location_id_") or k.startswith("patient_case_")]
            for k_main in keys_to_clear_main: 
                if k_main in st.session_state: del st.session_state[k_main]

            with st.spinner("Connecting to Tebra API and verifying practice..."):
                client = create_api_client(TEBRA_WSDL_URL)
                if not client: st.stop()
                header = build_request_header(credentials, client)
                if not header: st.stop()
                practice_id_check = get_practice_id_from_name(client, header, TEBRA_PRACTICE_NAME)
                if not practice_id_check: st.stop()
            display_message("success", f"✅ Connected to Tebra. Practice '{TEBRA_PRACTICE_NAME}' ID: {practice_id_check}.")
            
            try:
                df_excel_input = pd.read_excel(uploaded_file_val, dtype=str) 
                df_excel_input.columns = df_excel_input.columns.str.strip()
                # Add 'original_excel_row_num' for internal use, will be dropped before final Excel output
                df_excel_input['original_excel_row_num'] = range(st.session_state.original_excel_row_num_start, st.session_state.original_excel_row_num_start + len(df_excel_input))

                # Validate input columns based on EXPECTED_COLUMNS (original app's list)
                actual_cols = df_excel_input.columns.tolist()
                missing_cols = [c for c in EXPECTED_COLUMNS if c not in actual_cols]
                if missing_cols: 
                    display_message("error", f"❌ Excel file is missing the following required columns from the expected template: {', '.join(missing_cols)}.")
                    st.stop()
                
                display_message("info", f"Excel loaded ({len(df_excel_input)} rows). Processing charges...")
                output_df_results, success_groups_count, fail_groups_count = process_excel_data(client, header, practice_id_check, df_excel_input)
                
                if output_df_results is not None and not output_df_results.empty:
                    st.subheader("Processing Results Summary")
                    total_input_rows = len(df_excel_input) 
                    
                    st.metric("Total Input Excel Rows", total_input_rows) 
                    col1, col2 = st.columns(2)
                    col1.metric("Encounter Groups Successfully Processed", success_groups_count) 
                    col2.metric("Encounter Groups Failed", fail_groups_count)

                    st.subheader("Detailed Results")
                    # Display the DataFrame with the new column name and simplified errors
                    st.dataframe(output_df_results) 
                    
                    # Prepare DataFrame for download (it should already have the correct columns from process_excel_data)
                    df_for_download = output_df_results.copy()
                    # Explicitly drop 'original_excel_row_num' if it's still there before download
                    if 'original_excel_row_num' in df_for_download.columns:
                        df_for_download = df_for_download.drop(columns=['original_excel_row_num'])


                    output_excel_io = io.BytesIO()
                    with pd.ExcelWriter(output_excel_io, engine='xlsxwriter') as writer:
                        df_for_download.to_excel(writer, index=False, sheet_name='ChargeEntryResults')
                    excel_bytes = output_excel_io.getvalue()
                    st.download_button(label="📥 Download Results Excel", data=excel_bytes,
                                       file_name=f"Tebra_Results_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       key="sb_download_excel_v5") 
                else: display_message("info", "ℹ️ No results data generated from processing.")
            except Exception as e:
                display_message("error", f"❌ An unexpected error occurred during Excel processing or API interaction: {e}")
                st.exception(e) 
    
    st.markdown(f'<div class="footer">{APP_FOOTER}</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
