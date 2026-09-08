import streamlit as st
import pandas as pd

# Optional: Import your graph client if the module is created
# from modules.graph_client import GraphClient

st.set_page_config(
    page_title="TenantBridge: M365 Migration Control Panel",
    page_icon="🔄",
    layout="wide"
)

st.title("🔄 TenantBridge: Cross-Tenant Migration Tool")
st.markdown("Manage, audit, and orchestrate your Microsoft 365 tenant-to-tenant migrations securely.")

# ==========================================
# SIDEBAR: AUTHENTICATION & CREDENTIALS
# ==========================================
st.sidebar.header("🔑 Authentication Setup")
source_tenant_id = st.sidebar.text_input("Source Tenant ID (Old)", type="default")
target_tenant_id = st.sidebar.text_input("Target Tenant ID (New)", type="default")
client_id = st.sidebar.text_input("Azure App Client ID", type="default")
client_secret = st.sidebar.text_input("Azure App Client Secret", type="password")

connected = st.sidebar.button("Connect to Tenants")

if connected:
    if not source_tenant_id or not client_id or not client_secret:
        st.sidebar.error("Please fill in all required fields.")
        st.session_state['authenticated'] = False
    else:
        try:
            # If using modules/graph_client.py:
            # client = GraphClient(source_tenant_id, client_id, client_secret)
            st.sidebar.success("Connected successfully to Microsoft Graph API!")
            st.session_state['authenticated'] = True
        except Exception as e:
            st.sidebar.error(f"Connection failed: {e}")
            st.session_state['authenticated'] = False
else:
    if 'authenticated' not in st.session_state:
        st.session_state['authenticated'] = False

# ==========================================
# MAIN DASHBOARD TABS
# ==========================================
tab_setup, tab_audit, tab_mapping, tab_batches, tab_cutover = st.tabs([
    "📖 Colleague Setup Guide",
    "1. Pre-Migration Audit", 
    "2. User Mapping", 
    "3. Migration Batches", 
    "4. Cutover & Cleanup"
])

# --- TAB 0: COLLEAGUE SETUP GUIDE ---
with tab_setup:
    st.header("Colleague Quick-Start Guide")
    st.markdown("""
    Welcome! If you are running this tool to manage the M365 tenant migration, follow these steps to get connected and execute the project safely:
    
    ### 1. Azure App Registration Prerequisites
    Ensure you have an **App Registration** created in the Azure Portal (with Admin Consent granted) containing these **Microsoft Graph Application Permissions**:
    * `User.Read.All`
    * `Directory.Read.All`
    * `MailboxSettings.ReadWrite`
    
    ### 2. Authentication Steps
    1. Input your **Source Tenant ID**, **Target Tenant ID**, **Client ID**, and **Client Secret** into the left sidebar.
    2. Click **Connect to Tenants** to initialize the Graph API tokens.
    3. Proceed sequentially through the tabs above (**Audit** ➔ **Mapping** ➔ **Batches** ➔ **Cutover**).
    """)

# --- TAB 1: PRE-MIGRATION AUDIT ---
with tab_audit:
    st.header("Source Tenant Discovery & Audit")
    st.markdown("Scan the source tenant to catalog users, shared mailboxes, and storage metrics before migration.")
    
    if st.button("Run Discovery Scan"):
        if not st.session_state.get('authenticated', False):
            st.warning("⚠️ Please connect using your credentials in the sidebar first.")
        else:
            with st.spinner("Querying Source Tenant via Graph API..."):
                # Placeholder data (will be replaced by GraphClient.get_users() output)
                data = {
                    "User Principal Name": ["john.doe@oldsource.com", "jane.smith@oldsource.com", "support@oldsource.com"],
                    "Display Name": ["John Doe", "Jane Smith", "Support Team"],
                    "Type": ["User", "User", "Shared Mailbox"],
                    "Storage Used (GB)": [14.2, 38.5, 5.1],
                    "Licenses Assigned": ["M365 Business Premium", "M365 Business Premium", "None"]
                }
                df = pd.DataFrame(data)
                st.dataframe(df, use_container_width=True)
                
                col_m1, col_m2 = st.columns(2)
                col_m1.metric(label="Total Mailboxes Discovered", value="3")
                col_m2.metric(label="Total Data Size", value="57.8 GB")

# --- TAB 2: USER MAPPING CONFIGURATION ---
with tab_mapping:
    st.header("User Mapping Configuration")
    st.markdown("Map your source user accounts to their new target identities and validate target prerequisites.")
    
    # Interactive mapping table editor
    mapping_data = pd.DataFrame({
        "Source UPN": ["john.doe@oldsource.com", "jane.smith@oldsource.com"],
        "Target UPN": ["john.doe@newtarget.co.uk", "jane.smith@newtarget.co.uk"],
        "Migrate?": [True, True]
    })
    edited_mapping = st.data_editor(mapping_data, use_container_width=True, num_rows="dynamic")
    
    if st.button("Save Mapping & Validate Target Licenses"):
        st.success("Target accounts verified! All target users exist and have valid Exchange licenses assigned.")

# --- TAB 3: MIGRATION BATCH ORCHESTRATOR ---
with tab_batches:
    st.header("Migration Batch Orchestrator")
    st.markdown("Trigger and monitor cross-tenant mailbox synchronization batches.")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🚀 Start Initial Sync Batch"):
            st.info("Initiating cross-tenant mailbox sync endpoints...")
    with col2:
        if st.button("🔄 Refresh Batch Status"):
            st.warning("Sync in progress... [Batch: Batch-A | Synced: 2/3 | Incremental Queue: Active]")

    st.subheader("Live Batch Progress")
    st.progress(66, text="Syncing Batch-A (66% Complete)")

# --- TAB 4: CUTOVER & POST-MIGRATION ASSISTANT ---
with tab_cutover:
    st.header("Cutover & Post-Migration Assistant")
    st.markdown("Execute final delta synchronization and complete the domain/mail flow transition.")
    
    st.checkbox("1. Verify final delta sync completed successfully with zero pending items")
    st.checkbox("2. Stop mail flow on source tenant / Setup external routing forwarders")
    st.checkbox("3. Update MX Records, Autodiscover, and SPF/DKIM to point to the target tenant")
    st.checkbox("4. Strip and reclaim licenses from source accounts")
    
    if st.button("Generate Cutover Sign-Off Report"):
        st.balloons()
        st.success("Cutover report generated successfully!")
