import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="TenantBridge: M365 Migration Control Panel",
    page_icon="🔄",
    layout="wide"
)

st.title("🔄 TenantBridge: Cross-Tenant Migration Tool")
st.markdown("Manage, audit, and orchestrate your Microsoft 365 tenant-to-tenant migrations.")

# Sidebar for Authentication / Tenant Selection
st.sidebar.header("🔑 Authentication")
source_tenant_id = st.sidebar.text_input("Source Tenant ID (Old)", type="default")
target_tenant_id = st.sidebar.text_input("Target Tenant ID (New)", type="default")
client_id = st.sidebar.text_input("Azure App Client ID", type="default")
client_secret = st.sidebar.text_input("Azure App Client Secret", type="password")

connected = st.sidebar.button("Connect to Tenants")

if connected:
    st.sidebar.success("Connected successfully to Graph API!")
    st.session_state['authenticated'] = True
else:
    if 'authenticated' not in st.session_state:
        st.session_state['authenticated'] = False

if not st.session_state['authenticated']:
    st.warning("👈 Please enter your Azure App Registration credentials in the sidebar to begin.")
else:
    # Main Dashboard Tabs
    tab1, tab2, tab3, tab4 = st.tabs([
        "1. Pre-Migration Audit", 
        "2. User Mapping & Plan", 
        "3. Migration Batches", 
        "4. Cutover & Cleanup"
    ])

    with tab1:
        st.header("Source Tenant Discovery & Audit")
        st.markdown("Scan the source tenant to catalog mailboxes, sizes, and shared resources.")
        if st.button("Run Discovery Scan"):
            with st.spinner("Querying Source Tenant via Graph API..."):
                # Mock data for initial layout build
                data = {
                    "User Principal Name": ["john.doe@oldsource.com", "jane.smith@oldsource.com", "support@oldsource.com"],
                    "Display Name": ["John Doe", "Jane Smith", "Support Team"],
                    "Type": ["User", "User", "Shared Mailbox"],
                    "Storage Used (GB)": [14.2, 38.5, 5.1],
                    "Licenses Assigned": ["M365 Business Premium", "M365 Business Premium", "None"]
                }
                df = pd.DataFrame(data)
                st.dataframe(df, use_container_width=True)
                st.metric(label="Total Mailboxes Discovered", value="3")
                st.metric(label="Total Data Size", value="57.8 GB")

    with tab2:
        st.header("User Mapping Configuration")
        st.markdown("Map your source user accounts to their new target identities.")
        
        # Interactive mapping table editor
        mapping_data = pd.DataFrame({
            "Source UPN": ["john.doe@oldsource.com", "jane.smith@oldsource.com"],
            "Target UPN": ["john.doe@newtarget.co.uk", "jane.smith@newtarget.co.uk"],
            "Migrate?": [True, True]
        })
        edited_mapping = st.data_editor(mapping_data, use_container_width=True)
        
        if st.button("Save Mapping & Validate Target Licenses"):
            st.success("Target accounts verified! All target users exist and have valid Exchange licenses assigned.")

    with tab3:
        st.header("Migration Batch Orchestrator")
        st.markdown("Trigger and monitor cross-tenant mailbox sync batches.")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🚀 Start Initial Sync Batch"):
                st.info("Initiating cross-tenant mailbox sync endpoints...")
        with col2:
            if st.button("🔄 Refresh Batch Status"):
                st.warning("Sync in progress... [Batch: Batch-A | Synced: 2/3 | Incremental Queue: Active]")

        # Progress visualization mock
        st.subheader("Live Batch Status")
        st.progress(66, text="Syncing Batch-A (66% Complete)")

    with tab4:
        st.header("Cutover & Post-Migration Assistant")
        st.markdown("Finalize the cutover once initial syncs are caught up.")
        st.checkbox("1. Verify final delta sync completed successfully")
        st.checkbox("2. Stop mail flow on source tenant / Setup external forwarding")
        st.checkbox("3. Update MX Records & Autodiscover to point to target tenant")
        st.checkbox("4. Strip licenses from source accounts")
        
        if st.button("Generate Cutover Checklist Report"):
            st.balloons()
            st.success("Cutover report generated!")
