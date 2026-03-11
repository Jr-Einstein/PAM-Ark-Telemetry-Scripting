import streamlit as st
import pandas as pd
import plotly.express as px

# ---------------------------------------------------------
# 1. PAGE CONFIGURATION & SESSION STATE (Data Persistence)
# ---------------------------------------------------------
st.set_page_config(page_title="Proactive PAM Compliance", layout="wide")

# Initialize session state to store data temporarily while the app runs
if 'df' not in st.session_state:
    # Creating an empty dataframe with the required Phase I fields
    st.session_state.df = pd.DataFrame(columns=[
        'Account Name', 'Address', 'Platform ID', 'Last Error Message', 'Status'
    ])
if 'total_managed' not in st.session_state:
    st.session_state.total_managed = 56771 # Placeholder based on previous context

# ---------------------------------------------------------
# 2. SIDEBAR NAVIGATION
# ---------------------------------------------------------
st.sidebar.title("CyberArk PAM Dashboard")
page = st.sidebar.radio(
    "Navigation", 
    ["Home: Remediation Matrix", "Upload CSV Data", "Failed Accounts (Action)", "Analytics Dashboard", "FAQ & Tutorial"]
)

# ---------------------------------------------------------
# 3. PAGE: HOME / REMEDIATION MATRIX
# ---------------------------------------------------------
if page == "Home: Remediation Matrix":
    st.title("CPM Failure Remediation Matrix")
    st.markdown("Use this runbook to map failure trends to immediate operational fixes.")
    
    # Building the matrix based directly on your Phase III requirement doc
    remediation_data = {
        "Failure Reason (Pattern)": [
            "Access Denied (0x80070005)", 
            "Network Path Not Found", 
            "Account Locked Out", 
            "Timeout (Prompt not found)"
        ],
        "Likely Root Cause": [
            "CPM User lacks permissions on the target.",
            "Firewall blocking SMB/WMI ports.",
            "Target system lockout policy triggered.",
            "Slow target response or custom SSH prompt."
        ],
        "Remediation Steps": [
            "Verify CPM user is in the local 'Administrators' group.",
            "Request firewall opening for ports 445/135-139.",
            "Unlock account on Domain Controller; check for cached credentials.",
            "Increase platform timeout or update Process/Prompts file."
        ]
    }
    st.table(pd.DataFrame(remediation_data))

# ---------------------------------------------------------
# 4. PAGE: UPLOAD CSV DATA
# ---------------------------------------------------------
elif page == "Upload CSV Data":
    st.title("Phase I: Data Extraction & Upload")
    st.markdown("Upload your 'Accounts Inventory Report' containing Fully Managed accounts in an Error state.")
    
    uploaded_file = st.file_uploader("Upload Daily CyberArk CSV", type=['csv'])
    
    if uploaded_file is not None:
        new_data = pd.read_csv(uploaded_file)
        
        # Ensure the 'Status' column exists, default to 'Failed'
        if 'Status' not in new_data.columns:
            new_data['Status'] = 'Failed'
            
        # Append to our persistent session memory
        st.session_state.df = pd.concat([st.session_state.df, new_data], ignore_index=True)
        st.success(f"Successfully loaded {len(new_data)} records into the database.")
        st.dataframe(new_data.head())

# ---------------------------------------------------------
# 5. PAGE: FAILED ACCOUNTS & RECONCILIATION ACTION
# ---------------------------------------------------------
elif page == "Failed Accounts (Action)":
    st.title("Active CPM Failures Workflow")
    st.markdown("Review failed accounts and mark them as reconciled once the engineering fix is applied.")
    
    df = st.session_state.df
    failed_df = df[df['Status'] == 'Failed'].copy()
    
    if failed_df.empty:
        st.success("Zero failed accounts! The vault is healthy.")
    else:
        # Create a checkbox column for the user to mark as reconciled
        failed_df.insert(0, 'Mark Reconciled', False)
        
        # st.data_editor allows interactive editing in the browser
        edited_df = st.data_editor(
            failed_df,
            column_config={"Mark Reconciled": st.column_config.CheckboxColumn("Reconciled?", default=False)},
            disabled=["Account Name", "Address", "Platform ID", "Last Error Message", "Status"],
            hide_index=True,
            use_container_width=True
        )
        
        # When the user clicks the "Update Vault" button, save changes
        if st.button("Update Vault Health"):
            # Find rows where the checkbox was ticked
            reconciled_indices = edited_df[edited_df['Mark Reconciled'] == True].index
            
            # Update the main session state dataframe
            st.session_state.df.loc[reconciled_indices, 'Status'] = 'Reconciled'
            st.success(f"Successfully reconciled {len(reconciled_indices)} accounts!")
            st.rerun() # Refresh the page to remove reconciled accounts from this view

# ---------------------------------------------------------
# 6. PAGE: ANALYTICS DASHBOARD
# ---------------------------------------------------------
elif page == "Analytics Dashboard":
    st.title("Phase IV: Executive Impact Dashboard")
    
    df = st.session_state.df
    total_managed = st.session_state.total_managed
    
    # Math for the metrics
    total_failures_ever = len(df)
    current_active_failures = len(df[df['Status'] == 'Failed'])
    reconciled_count = len(df[df['Status'] == 'Reconciled'])
    
    # Success Rate = (Total Managed - Current Active Failures) / Total Managed
    initial_success_rate = ((total_managed - total_failures_ever) / total_managed) * 100
    current_success_rate = ((total_managed - current_active_failures) / total_managed) * 100
    improvement = current_success_rate - initial_success_rate
    
    # 1. Top Level Metrics (Before / After Health Check)
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Managed Accounts", f"{total_managed:,}")
    col2.metric("Active CPM Failures", current_active_failures, delta=f"-{reconciled_count} Reconciled", delta_color="inverse")
    col3.metric("Overall Success Rate %", f"{current_success_rate:.2f}%", delta=f"{improvement:.2f}% Improvement")
    
    st.divider()
    
    # 2. Primary Visualization: Bar Graph
    if not df.empty:
        st.subheader("Count of Failures vs. Failure Reason")
        # Grouping by Last Error Message to find trends
        trend_df = df[df['Status'] == 'Failed'].groupby('Last Error Message').size().reset_index(name='Count')
        
        fig = px.bar(trend_df, x='Last Error Message', y='Count', color='Last Error Message',
                     title="Active Failure Trends", text_auto=True)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Upload data to view trend analysis.")

# ---------------------------------------------------------
# 7. PAGE: FAQ & TUTORIAL
# ---------------------------------------------------------
elif page == "FAQ & Tutorial":
    st.title("Dashboard Documentation & Tutorial")
    st.markdown("""
    ### How to use this platform:
    1. **Upload CSV Data:** Start your day by uploading the *Accounts Inventory Report* extracted from CyberArk PVWA. 
    2. **Failed Accounts Workflow:** Operations engineers should review the 'Failed Accounts' tab. Once an issue (like an ODBC error) is fixed via the runbook, check the box and click 'Update Vault'.
    3. **Analytics Dashboard:** Use this tab during executive reviews to show the Before/After health check and how much the Success Rate % has improved.
    
    *Note: Data is stored in the browser's session state. In a production environment, this would be connected directly to the CyberArk API or an SQL database.*
    """)







# ---------------------------------------------------------
# 6. PAGE: ANALYTICS DASHBOARD (Updated with Auto-Trend Detection)
# ---------------------------------------------------------
elif page == "Analytics Dashboard":
    st.title("Phase IV: Executive Impact & Trend Analysis")
    
    df = st.session_state.df
    total_managed = st.session_state.total_managed
    
    # Filter only the accounts currently in a "Failed" state
    failed_df = df[df['Status'] == 'Failed']
    total_failures_ever = len(df)
    current_active_failures = len(failed_df)
    
    # Success Rate Math
    initial_success_rate = ((total_managed - total_failures_ever) / total_managed) * 100
    current_success_rate = ((total_managed - current_active_failures) / total_managed) * 100
    improvement = current_success_rate - initial_success_rate
    
    # 1. Top Level Health Check Metrics
    st.subheader("Before / After Health Check")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Managed Accounts", f"{total_managed:,}")
    col2.metric("Overall Success Rate %", f"{current_success_rate:.2f}%", delta=f"{improvement:.2f}% Improvement")
    col3.metric("Total Accounts Reconciled", len(df[df['Status'] == 'Reconciled']))
    
    st.divider()
    
    # 2. Automated Trend Analysis & Visualization
    st.subheader("CPM Failure Trends")
    
    if not failed_df.empty:
        # PANDAS MAGIC: Automatically find the maximum trends and sort them
        # (Assuming your CSV column is named 'Failure Reason'. If it's 'Last Error Message', change it here)
        trend_counts = failed_df['Failure Reason'].value_counts()
        
        # Create a layout: Chart on the Left (70% width), Metrics on the Right (30% width)
        col_chart, col_metrics = st.columns([7, 3])
        
        with col_chart:
            # Create the interactive Pie Chart
            fig = px.pie(
                failed_df, 
                names='Failure Reason', 
                title="Distribution of Failure Reasons",
                hole=0.4 # Makes it a modern donut chart
            )
            fig.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig, use_container_width=True)
            
        with col_metrics:
            # Display the right-side metrics dynamically based on the top 3 trends found
            st.markdown("### Top Failure Trends")
            st.metric("Total Active Failures", current_active_failures)
            
            # Safely check if we have at least 1 reason, 2 reasons, etc., before displaying
            if len(trend_counts) > 0:
                reason_1_name = trend_counts.index[0]
                reason_1_count = trend_counts.iloc[0]
                st.metric(f"🥇 Top 1: {reason_1_name}", reason_1_count)
                
            if len(trend_counts) > 1:
                reason_2_name = trend_counts.index[1]
                reason_2_count = trend_counts.iloc[1]
                st.metric(f"🥈 Top 2: {reason_2_name}", reason_2_count)
                
            if len(trend_counts) > 2:
                reason_3_name = trend_counts.index[2]
                reason_3_count = trend_counts.iloc[2]
                st.metric(f"🥉 Top 3: {reason_3_name}", reason_3_count)
    else:
        st.success("No active failures to analyze! Upload data or check the vault.")
        
