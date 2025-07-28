import streamlit as st
import pandas as pd
from attend_web import get_available_dates, get_panchayath_and_work_codes, get_attendance_reports

st.set_page_config(page_title="NMMS Attendance Downloader", layout="wide")

st.title("NMMS Attendance Downloader")

# --- Constants ---
PANCHAYATH_LIST = [
    "BAGEWADI", "BAGGURU", "BALAKUNDHI", "BEERAHALLI", "B.M. SUGURU", 
    "BYRAPURA", "DESANOORU", "HACHCHOLI", "HALEKOTE", "H. HOSAHALLI", 
    "KARURU", "K. BELAGALLU", "KENCHANAGUDDA", "KONCHIGERI", "K. SUGURU", 
    "KUDUDHURAHAL", "KURUVALLI", "M. SUGURU", "MUDDATANURU", "NADAVI", 
    "RARAVI", "SANAVASAPURA", "SIRIGERI", "TALURU", "UPPARA HOSAHALLI", 
    "UTHTHANURU"
]

# --- State Management ---
def init_session_state():
    st.session_state.dates = []
    st.session_state.work_codes = []
    st.session_state.selected_panchayath = None
    st.session_state.reports = None
    st.session_state.loading = False
    st.session_state.page_key = 0 # Used to force re-render on reset

if 'page_key' not in st.session_state:
    init_session_state()

def reset_app():
    st.session_state.page_key += 1
    keys_to_clear = ['dates', 'work_codes', 'selected_panchayath', 'reports', 'loading']
    for key in keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]
    init_session_state()
    st.rerun()

# --- Helper Functions ---
def get_friendly_filename(filename):
    if filename.startswith("nmr_images_"):
        return "NMR Images Report"
    elif filename.startswith("verification_format_"):
        return "Verification Format"
    elif filename.startswith("nmr_raw_"):
        return "Raw Attendance Data"
    elif filename.startswith("nmr_"):
        return "Attendance with Images report"
    return filename

# --- Data Fetching Functions ---
def fetch_dates():
    with st.spinner("Fetching available dates..."):
        st.session_state.dates = get_available_dates()
    if not st.session_state.dates:
        st.warning("Could not fetch available dates. Please check the backend or network.")

def fetch_panchayaths_and_workcodes(attendance_date, panchayath_name):
    st.session_state.loading = True
    st.session_state.work_codes = [] 
    st.session_state.reports = None
    with st.spinner(f"Fetching work codes for {panchayath_name} on {attendance_date}..."):
        work_codes = get_panchayath_and_work_codes(attendance_date, panchayath_name)
        if work_codes:
            st.session_state.work_codes = work_codes
            st.session_state.selected_panchayath = panchayath_name
            st.success(f"Found {len(work_codes)} work codes for {panchayath_name}.")
        else:
            st.error(f"Could not find any work codes for '{panchayath_name}'. Please check the name and date and try again.")
    st.session_state.loading = False

def generate_reports(attendance_date, panchayath_name, choice, workcodes):
    st.session_state.loading = True
    st.session_state.reports = None
    
    progress_bar = st.progress(0)
    progress_text = st.empty()

    def update_progress(progress, message):
        progress_bar.progress(progress)
        progress_text.text(message)

    reports = get_attendance_reports(
        attendance_date, panchayath_name, choice, workcodes, progress_callback=update_progress
    )
    
    progress_bar.empty()
    progress_text.empty()

    if reports:
        st.session_state.reports = reports
        st.success("Reports generated successfully!")
    else:
        st.error("Failed to generate reports.")
    
    st.session_state.loading = False

# --- UI Layout ---
main_container = st.container(key=f"main_container_{st.session_state.page_key}")

with main_container:
    if not st.session_state.dates:
        fetch_dates()

    if st.session_state.dates:
        col1, col2 = st.columns(2)
        with col1:
            selected_date = st.selectbox("Select Attendance Date", st.session_state.dates)
        with col2:
            panchayath_name_input = st.selectbox("Select Panchayath Name", PANCHAYATH_LIST)

        if st.button("Fetch Work Codes", disabled=st.session_state.loading or not panchayath_name_input):
            fetch_panchayaths_and_workcodes(selected_date, panchayath_name_input)

        if st.session_state.work_codes and st.session_state.selected_panchayath == panchayath_name_input:
            st.markdown("---")
            st.subheader(f"Work Codes for {st.session_state.selected_panchayath}")

            choice = st.radio("Select Muster Rolls to Download", ('All', 'Specific Work Codes'), key='work_choice')

            selected_work_codes = []
            if choice == 'Specific Work Codes':
                with st.expander("Select Work Codes", expanded=True):
                    for code in st.session_state.work_codes:
                        if st.checkbox(code, key=f"cb_{code}"):
                            selected_work_codes.append(code)
            
            if st.button("Generate and Download Reports", disabled=st.session_state.loading):
                if choice == 'Specific Work Codes' and not selected_work_codes:
                    st.warning("Please select at least one work code.")
                else:
                    report_choice = 'work' if choice == 'Specific Work Codes' else 'all'
                    generate_reports(selected_date, panchayath_name_input, report_choice, selected_work_codes)

        if st.session_state.reports:
            st.markdown("---")
            st.subheader("Downloads")
            st.info("Click the buttons below to download your generated Excel reports.")
            
            num_reports = len(st.session_state.reports)
            cols = st.columns(num_reports if num_reports > 0 else 1)
            
            for idx, (filename, data) in enumerate(st.session_state.reports.items()):
                with cols[idx]:
                    st.download_button(
                        label=f"Download {get_friendly_filename(filename)}",
                        data=data,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{filename}"
                    )
    else:
        st.error("Application could not start. Please ensure the backend is running and configured correctly.")

    st.markdown("---")
    if st.button("Reset Application"):
        reset_app()
