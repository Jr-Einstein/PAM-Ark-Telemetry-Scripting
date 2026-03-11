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
