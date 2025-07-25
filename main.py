import streamlit as st
import pandas as pd
from datetime import datetime
from fpdf import FPDF
from io import BytesIO
from PIL import Image
from supabase import create_client, Client
import re
import time
import base64
import math
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

st.set_page_config(page_title="CloudIA Quota Verifier", layout="centered")
QUOTA_PER_HA = 800
LOGO_PATH = "cloudia_logo.png"
LOGO_COCOA = "cocoasourcelogo.jpg"

# --- Language Switcher ---
lang = st.sidebar.radio("🌐 Language / Langue", ["English", "Français"])

def t(key):
    translations = {
        "upload_title": {
            "English": "Upload Delivery Template",
            "Français": "Téléverser le Modèle de Livraison"
        },
        "file_format_caption": {
            "English": "✅ Format: .xlsx | Max size: 200MB",
            "Français": "✅ Format : .xlsx | Taille max : 200 Mo"
        },
        "title": {
            "English": "☁️ CloudIA – Farmer Quota Verification System – Coop Level",
            "Français": "☁️ CloudIA – Système de Vérification des Quotas – Niveau Coopérative"
        },
        "missing_exporter_column": {
            "English": "Missing 'exporter' column in the Excel file.",
            "Français": "La colonne 'exporter' est manquante dans le fichier Excel."
        },
        "missing_columns": {
            "English": "Missing columns: {}",
            "Français": "Colonnes manquantes : {}"
        },
        "unknown_farmers_error": {
            "English": "The following farmers are NOT in the database:",
            "Français": "Les producteurs suivants ne sont PAS dans la base de données :"
        },
        "quota_overview_title": {
            "English": "### Quota Overview (Only Warnings and Exceeded)",
            "Français": "### Aperçu des Quotas (Avertissements et Dépassements)"
        },
        "quota_warning_count": {
            "English": "⚠️ {} farmers in the uploaded file have quota warnings or exceeded limits.",
            "Français": "⚠️ {} producteurs ont des avertissements ou des dépassements de quota."
        },
        "quota_ok": {
            "English": "✅ All farmers in the uploaded file are within their assigned quotas.",
            "Français": "✅ Tous les producteurs respectent leurs quotas."
        },
        "lot_status_out_of_range": {
            "English": "### Lot Status Overview - Out of Range",
            "Français": "### Aperçu des Lots - Hors Plage Autorisée"
        },
        "rollback_error": {
            "English": "❌ Uploaded delivery has been rolled back due to validation errors. PDF cannot be generated.",
            "Français": "❌ Livraison annulée en raison d'erreurs de validation. PDF non généré."
        },
        "file_approved": {
            "English": "✅ File approved. All farmers valid, quotas OK, and delivered kg per lot within allowed range.",
            "Français": "✅ Fichier approuvé. Tous les producteurs sont valides et les quotas respectés."
        },
        "generate_pdf": {
            "English": "Generate Approval PDF",
            "Français": "Générer le PDF d'Approbation"
        },
        "download_pdf": {
            "English": "Download Approval PDF",
            "Français": "Télécharger le PDF"
        },
        "lot_too_low": {
            "English": "Too low",
            "Français": "Trop faible"
        },
        "lot_within_range": {
            "English": "Within range",
            "Français": "Dans la plage autorisée"
        },
         "insert_success": {
            "English": "✅ Data successfully inserted! {0} new records added.",
            "Français": "✅ Données insérées avec succès ! {0} nouveaux enregistrements ajoutés."
        },
        "approval_save_error": {
            "English": "❌ Error saving approval to the database",
            "Français": "❌ Erreur lors de l'enregistrement de l'approbation dans la base de données"
        },
         "missing_farmer_id_column": {
            "English": "❌ quota_view does not contain 'farmer_id'. Columns returned: {}",
            "Français": "❌ La vue quota_view ne contient pas 'farmer_id'. Colonnes retournées : {}"
        },
        "saving": {
            "English": "💾 Saving data...",
            "Français": "💾 Sauvegarde des données..."
        }


    }
    return translations.get(key, {}).get(lang, key)

# Load SharePoint configuration from secrets
sharepoint_config = st.secrets.get("sharepoint", {})

@st.cache_resource
def get_supabase() -> Client:
    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["key"]
    return create_client(url, key)

supabase = get_supabase()

@st.cache_data
def load_all_farmers():
    all_rows = []
    page_size = 1000
    last_farmer_id = None
    while True:
        query = supabase.table("farmers").select("*").limit(page_size).order("farmer_id")
        if last_farmer_id:
            query = query.gt("farmer_id", last_farmer_id)
        result = query.execute()
        rows = result.data
        if not rows:
            break
        all_rows.extend(rows)
        last_farmer_id = rows[-1]["farmer_id"]
    farmers_df = pd.DataFrame(all_rows)
    farmers_df.columns = farmers_df.columns.str.lower()
    farmers_df['farmer_id'] = farmers_df['farmer_id'].astype(str).str.strip().str.lower()
    return farmers_df

def delete_existing_delivery_rpc(export_lot, exporter_name, farmer_ids):
    export_lot = str(export_lot)
    exporter_name = str(exporter_name)
    if hasattr(farmer_ids, 'tolist'):
        farmer_ids = farmer_ids.tolist()
    farmer_ids = [str(farmer_id) for farmer_id in farmer_ids]
    try:
        supabase.rpc('delete_traceability_records', {
            'lot': export_lot,
            'exporter_param': exporter_name,
            'farmer_ids': farmer_ids
        }).execute()
    except Exception as e:
        st.error(f"❌ RPC Delete Error: {e}")


def save_delivery_to_supabase(df):
    column_mapping = {
        'cooperative name': 'cooperative_name',
        'export lot n°/connaissement': 'export_lot',
        'date of purchase from cooperative': 'purchase_date',
        'certification': 'certification',
        'farmer_id': 'farmer_id',
        'net weight (kg)': 'net_weight_kg',
        'exporter': 'exporter'
    }
    df = df.rename(columns=column_mapping)
    required_columns = ['export_lot', 'exporter', 'farmer_id', 'net_weight_kg']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"Missing required columns: {', '.join(missing_columns)}")
        return False

    df_cleaned = df.copy()
    df_cleaned['farmer_id'] = df_cleaned['farmer_id'].str.strip().str.lower()
    df_cleaned['purchase_date'] = df_cleaned['purchase_date'].fillna(datetime.today().strftime('%Y-%m-%d'))
    # Ensure certification is string before replacing
    df_cleaned['certification'] = df_cleaned['certification'].astype(str)
    df_cleaned['certification'] = df_cleaned['certification'].replace(
        ['N/A', 'n/a', 'na', 'NA', 'NaN', 'nan', '', 'None'], None
    )


    def excel_date_to_date(excel_date):
        if isinstance(excel_date, (int, float)):
            return (pd.to_datetime('1899-12-30') + pd.to_timedelta(excel_date, unit='D')).strftime('%Y-%m-%d')
        return excel_date

    df_cleaned['purchase_date'] = df_cleaned['purchase_date'].apply(excel_date_to_date)
    df_cleaned['purchase_date'] = df_cleaned['purchase_date'].astype(str)
    data = df_cleaned.to_dict(orient="records")

    # Check for missing values in required fields before inserting
    required_fields = ['export_lot', 'exporter', 'farmer_id', 'net_weight_kg']
    missing_values = df_cleaned[required_fields].isnull().any(axis=1)

    if missing_values.any():
        st.error("❌ Some rows have missing values in required fields:")
        st.dataframe(df_cleaned[missing_values])
        return False

    try:
        with st.spinner(t("saving")):
            supabase.table("traceability").insert(data).execute()
        st.success(t("insert_success").format(len(data)))
        return True
    except Exception as e:
        st.error(f"{t('insert_error')}: {e}")
        return False

# Modified generate_pdf_confirmation to return buffer and filename
def generate_pdf_confirmation(lot_numbers, exporter_name, farmer_count, total_kg, lot_kg_summary, logo_path, logo_cocoa, cooperative_names):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(200, 10, "Delivery Approval Certificate", ln=True, align="C")

    if logo_path:
        try:
            pdf.image(logo_path, x=10, y=20, w=40)
        except Exception as e:
            st.warning(f"Could not embed logo from {logo_path}: {e}")
    if logo_cocoa:
        try:
            pdf.image(logo_cocoa, x=(210 - 110) / 2, y=20, w=110)
        except Exception as e:
            st.warning(f"Could not embed logo from {LOGO_COCOA}: {e}")


    pdf.set_y(70)
    pdf.set_font("Arial", "", 12)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    pdf.multi_cell(0, 10, f"Generated on: {now}")
    pdf.multi_cell(0, 10, f"Exporter: {exporter_name}")
    pdf.multi_cell(0, 10, f"Cooperatives: {', '.join(sorted(set(cooperative_names)))}")
    pdf.multi_cell(0, 10, f"Lots: {', '.join(str(l) for l in lot_numbers)}")
    pdf.multi_cell(0, 10, f"Total Farmers: {farmer_count}")
    pdf.multi_cell(0, 10, f"Total Net Weight: {round(total_kg / 1000, 2)} MT")

    pdf.ln(5)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Lot Summary", ln=True)
    pdf.set_font("Arial", "", 12)
    for lot, kg in lot_kg_summary.items():
        pdf.cell(0, 10, f"{lot}: {round(kg / 1000, 2)} MT", ln=True)

    pdf.ln(5)
    pdf.cell(0, 10, "Approved by CloudIA", ln=True)

    reference_number = lot_numbers[0] if len(lot_numbers) == 1 else "MULTI"
    reference_number = re.sub(r"[^\w\-]", "_", str(reference_number))
    today_str = datetime.now().strftime('%Y%m%d')
    exporter_clean = exporter_name.replace(" ", "_").replace("/", "_")[:20]
    total_volume_mt = round(total_kg / 1000, 2)

    filename = f"Approval_{reference_number}_{today_str}_{exporter_clean}_{total_volume_mt}MT.pdf"
    # Save approval record to database
    data = {
        "created_at": now,
        "lot_number": ", ".join(str(l) for l in lot_numbers),
        "exporter_name": exporter_name,
        "approved_by": "CloudIA",
        "file_name": filename
    }
    try:
        supabase.table("approvals").insert(data).execute()
    except Exception as e:
        st.error(f"{t('approval_save_error')}: {e}")


    # Return PDF as BytesIO buffer
    pdf_bytes = pdf.output(dest='S').encode('latin1')
    pdf_buffer = BytesIO(pdf_bytes)
    pdf_buffer.seek(0)

    return filename, pdf_buffer

# SharePoint upload function based on the working code
def upload_to_sharepoint(file_buffer, filename, sharepoint_config):
    """Uploads a file to a specified SharePoint folder."""
    try:
        site_url = sharepoint_config.get("site_url")
        client_id = sharepoint_config.get("client_id")
        client_secret = sharepoint_config.get("client_secret")
        library_name = sharepoint_config.get("library_name")

        if not all([site_url, client_id, client_secret, library_name]):
            st.error("❌ SharePoint credentials or library name are missing in secrets.toml")
            return False

        credentials = ClientCredential(client_id, client_secret)
        ctx = ClientContext(site_url).with_credentials(credentials)

        # Construct the folder URL - assuming it's under 'sites/TRACAFILES/' as seen in working code
        # Adjust if your site structure is different
        folder_url = f"/sites/TRACAFILES/{library_name}"

        target_folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        ctx.load(target_folder)
        ctx.execute_query() # Ensure folder exists and is accessible

        target_folder.upload_file(filename, file_buffer.getvalue()).execute_query()
        st.success(f"✅ File uploaded to SharePoint: {filename}")
        return True

    except Exception as e:
        st.error(f"❌ SharePoint upload failed for file '{filename}'. Error: {e}")
        # You might want more detailed logging here
        return False


def load_quota_view():
    result = supabase.table("quota_view").select("*").execute()
    return pd.DataFrame(result.data)

# --- UI Layout ---
st.markdown("---")
logo_col1, logo_col2 = st.columns([1, 1])
with logo_col1:
    # Need to handle image loading robustly, especially in different environments
    try:
        st.image(Image.open(LOGO_PATH), width=120)
    except FileNotFoundError:
        st.warning(f"Logo file not found at {LOGO_PATH}")
        st.markdown(f"**CloudIA Logo Placeholder**") # Placeholder text
with logo_col2:
    try:
        st.image(Image.open(LOGO_COCOA), width=200)
    except FileNotFoundError:
         st.warning(f"Logo file not found at {LOGO_COCOA}")
         st.markdown(f"**Cocoa Source Logo Placeholder**") # Placeholder text


st.markdown("""
    <div style='
        text-align: center;
        font-size: 60px;
        font-weight: bold;
        margin-top: 30px;
        margin-bottom: 20px;
        color: #2c3e50;
        letter-spacing: 5px;
    '>
        COOP
    </div>
""", unsafe_allow_html=True)

st.markdown(f"### {t('title')}")



# --- Główna logika ---
st.subheader("📥 Step 1: Upload Excel for Validation")
delivery_file = st.file_uploader(t("upload_title"), type=["xlsx"])
st.caption(t("file_format_caption")) # Add caption back
farmers_df = load_all_farmers()

# Initialize session state for PDF buffer and filename
if 'pdf_buffer' not in st.session_state:
    st.session_state['pdf_buffer'] = None
    st.session_state['pdf_filename'] = None
if 'uploaded_excel_file_obj' not in st.session_state:
     st.session_state['uploaded_excel_file_obj'] = None


if delivery_file:
    # Store the file object in session state for later access
    st.session_state['uploaded_excel_file_obj'] = delivery_file

    uploaded_df = pd.read_excel(delivery_file)
    uploaded_df.columns = uploaded_df.columns.str.strip().str.lower()
    uploaded_df['farmer_id'] = uploaded_df['farmer_id'].astype(str).str.strip().str.lower()

    if 'exporter' not in uploaded_df.columns:
        st.error(t("missing_exporter_column"))
        st.stop()

    exporter_names = uploaded_df['exporter'].dropna().astype(str).str.strip().unique()
    # Assuming a single exporter per file for simplicity based on previous code
    exporter_name = ", ".join(exporter_names) # Keep this as it was

    expected_columns = ['cooperative name', 'export lot n°/connaissement', 'date of purchase from cooperative',
                        'certification', 'farmer_id', 'farm_id', 'net weight (kg)', 'exporter']
    missing_columns = [col for col in expected_columns if col not in uploaded_df.columns]
    if missing_columns:
        st.error(t("missing_columns").format(', '.join(missing_columns)))
        st.stop()

    uploaded_df.rename(columns={
        'export lot n°/connaissement': 'export_lot',
        'net weight (kg)': 'net_weight_kg',
        'date of purchase from cooperative': 'purchase_date'
    }, inplace=True)

    uploaded_df['purchase_date'] = uploaded_df['purchase_date'].fillna(datetime.today().strftime('%Y-%m-%d'))
    uploaded_df['exporter'] = exporter_name # Ensure exporter name is consistent

    uploaded_df = uploaded_df.drop_duplicates(subset=['export_lot', 'exporter', 'farmer_id', 'net_weight_kg'], keep='last')

     # ZABEZPIECZENIE: blokuj puste pliki
    if uploaded_df.empty:
        st.error("❌ The uploaded file is empty or contains no valid delivery records.")
        st.stop()


    unknown_farmers = uploaded_df[
        ~uploaded_df['farmer_id'].str.lower().isin(farmers_df['farmer_id'].str.lower())
    ]['farmer_id'].unique()

    if unknown_farmers.size > 0:
        st.error(t("unknown_farmers_error"))
        st.write(list(unknown_farmers))
        st.stop()

    # --- Process each exporter separately for RPC deletion ---
    # This logic seems intended to clean up previous uploads for the same lot/exporter/farmers
    # It might be better placed *before* reading the file if you want to prevent processing
    # duplicates entirely, but keeping it here matches the original flow.
    # Note: The original code iterated through exporter_names again here, which is redundant
    # if you assume one exporter per file as implied by `exporter_name = ", ".join(exporter_names)`.
    # Keeping the loop structure but using the single `exporter_name` for clarity.
    # If multiple exporters per file are expected, this deletion logic needs refinement.
    lot_numbers_in_file = uploaded_df['export_lot'].unique()
    for lot in lot_numbers_in_file:
        farmer_ids_for_lot = uploaded_df[uploaded_df['export_lot'] == lot]['farmer_id'].unique().tolist()
        # Assuming exporter_name is representative for the file
        delete_existing_delivery_rpc(lot, exporter_name, farmer_ids_for_lot)


    # Save to Supabase
    inserted_ok = save_delivery_to_supabase(uploaded_df)
    if not inserted_ok:
        st.stop()

    time.sleep(1)  # Give time for data propagation
    quota_df = load_quota_view()

    # Diagnoza – check if farmer_id column exists
    if 'farmer_id' not in quota_df.columns:
        st.error(t("missing_farmer_id_column").format(list(quota_df.columns)))
        st.stop()

    # Cleaning and Filtering Quota Data
    uploaded_ids = pd.Series(uploaded_df['farmer_id']).astype(str).str.strip().str.lower()
    quota_df['farmer_id'] = quota_df['farmer_id'].astype(str).str.strip().str.lower()
    quota_df = quota_df[quota_df['farmer_id'].isin(uploaded_ids)]

    quota_filtered = quota_df[quota_df['quota_status'].isin(['EXCEEDED', 'WARNING'])]


    if not quota_filtered.empty:
        st.write(t("quota_overview_title"))

        def highlight_status(val):
            if val == 'EXCEEDED':
                return 'background-color: #ffcccc'
            elif val == 'WARNING':
                return 'background-color: #fff3cd'
            return ''

        styled_quota = quota_filtered[[
            'farmer_id', 'max_quota_kg', 'total_net_weight_kg', 'quota_used_pct', 'quota_status'
        ]].style.applymap(highlight_status, subset=['quota_status']).format({
            'max_quota_kg': '{:.0f}',
            'total_net_weight_kg': '{:.0f}',
            'quota_used_pct': '{:.2f}'
        })

        st.dataframe(styled_quota, use_container_width=True)
        st.warning(t("quota_warning_count").format(len(quota_filtered)))
    else:
        st.success(t("quota_ok"))

    # Validation Checks
    all_ids_valid = len(unknown_farmers) == 0
    any_quota_exceeded = 'EXCEEDED' in quota_filtered['quota_status'].values
    lot_totals = uploaded_df.groupby('export_lot')['net_weight_kg'].sum()

    def check_lot_status(weight_in_kg):
        weight_in_mt = weight_in_kg / 1000
        # Correcting the lot weight check based on the working code
        if weight_in_mt < 21:
            return t("lot_too_low")
        else:
            return t("lot_within_range")

    lot_status = lot_totals.apply(check_lot_status)
    # Correcting the comparison based on the working code
    lot_status_ok = lot_status == "Within range"

    lot_status_info = pd.DataFrame({
        'export_lot': lot_totals.index,
        'total_net_weight_kg': lot_totals.values,
        'lot_status': lot_status
    })

    if not lot_status_ok.all():
        st.write(t("lot_status_out_of_range"))
        st.dataframe(lot_status_info[~lot_status_ok])

    def rollback_delivery(uploaded_df):
        lot_numbers = uploaded_df['export_lot'].unique()
        # Ensure exporter_name is correctly retrieved for rollback
        exporter_name_for_rollback = uploaded_df['exporter'].iloc[0] if not uploaded_df.empty else "Unknown"
        for lot in lot_numbers:
            farmer_ids_for_lot = uploaded_df[uploaded_df['export_lot'] == lot]['farmer_id'].unique().tolist()
            delete_existing_delivery_rpc(lot, exporter_name_for_rollback, farmer_ids_for_lot)
        st.error(t("rollback_error"))


    # --- Display Approval Status and Buttons ---
    if all_ids_valid and not any_quota_exceeded and lot_status_ok.all():
        st.success(t("file_approved"))
        st.subheader("📄 Step 2: Generate and Download PDF")

        # Buttons for PDF generation and SharePoint upload
        col1, col2 = st.columns([1, 1])

        with col1:
            # Generate PDF button
            if st.button(t("generate_pdf")):
                total_kg = int(lot_totals.sum())
                # Call generate_pdf_confirmation and store results in session state
                filename, pdf_buffer = generate_pdf_confirmation(
                    lot_numbers=lot_totals.index.tolist(),
                    exporter_name=exporter_name,
                    farmer_count=uploaded_df['farmer_id'].nunique(),
                    total_kg=total_kg,
                    lot_kg_summary=lot_totals.to_dict(),
                    cooperative_names=uploaded_df['cooperative name'].dropna().unique().tolist(),
                    logo_path=LOGO_PATH,
                    logo_cocoa=LOGO_COCOA
                )
                st.session_state['pdf_buffer'] = pdf_buffer
                st.session_state['pdf_filename'] = filename
                st.success("✅ PDF generated. You can now download or upload it.")


        # Download PDF button (appears after PDF is generated)
        if st.session_state['pdf_buffer'] and st.session_state['pdf_filename']:
             with col1: # Place download button under generate
                st.download_button(
                    label=t("download_pdf"),
                    data=st.session_state['pdf_buffer'],
                    file_name=st.session_state['pdf_filename'],
                    mime="application/pdf"
                )

        with col2:
            st.subheader("📤 Step 3: Upload to SharePoint")
            # SharePoint Upload button (appears after PDF is generated)
            if st.session_state['pdf_buffer'] and st.session_state['pdf_filename']:
                if st.button("📤 Upload to SharePoint"):
                    success_pdf = success_excel = False

                    # Check if SharePoint configuration is available
                    if not sharepoint_config:
                         st.error("❌ SharePoint configuration not found in Streamlit secrets.")
                    else:
                        st.info("📤 Uploading PDF to SharePoint...")
                        # Upload the generated PDF
                        success_pdf = upload_to_sharepoint(
                            st.session_state['pdf_buffer'],
                            st.session_state['pdf_filename'],
                            sharepoint_config
                        )

                        st.info("📤 Uploading Excel to SharePoint...")
                        # Upload the original Excel file
                        uploaded_file_obj = st.session_state.get('uploaded_excel_file_obj')
                        if uploaded_file_obj:
                             # Seek to the beginning of the file buffer before reading
                             uploaded_file_obj.seek(0)
                             success_excel = upload_to_sharepoint(
                                 uploaded_file_obj, # Pass the file object/buffer
                                 uploaded_file_obj.name, # Pass the original file name
                                 sharepoint_config
                             )
                        else:
                            st.error("❌ Original Excel file object not found in session state.")
                            success_excel = False


                        if success_pdf and success_excel:
                            st.success("✅ Both PDF and Excel uploaded to SharePoint.")
                        elif success_pdf or success_excel:
                             st.warning("⚠️ One of the files failed to upload. See error messages above.")
                        else:
                             st.error("❌ Both files failed to upload. See error messages above.")

            else:
                 st.info("Generate the PDF first to enable SharePoint upload.")


    else:
        # If validation fails, rollback delivery and show error
        rollback_delivery(uploaded_df)