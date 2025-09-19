# Updated filter section with synchronized ID/Name selection
if st.session_state.admin_logged_in:
    st.subheader("Admin Dashboard - Employee Results")
    if st.button("🔄 Refresh Data"):
        st.cache_data.clear()
        st.rerun()
    
    results_df = load_all_results()
    if not results_df.empty:
        st.markdown("---")
        st.subheader("🔍 Filters")
        if "filter_reset_counter" not in st.session_state:
            st.session_state.filter_reset_counter = 0
        
        # Create ID-Name mapping for synchronization
        id_name_mapping = dict(zip(results_df["ID"].astype(str), results_df["Name"]))
        name_id_mapping = dict(zip(results_df["Name"], results_df["ID"].astype(str)))
        
        filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)
        
        with filter_col1:
            # Employee ID filter
            employee_ids = ["All"] + sorted(results_df["ID"].astype(str).unique().tolist())
            
            # Check if name was changed and sync ID
            selected_name_key = f"emp_name_filter_{st.session_state.filter_reset_counter}"
            if selected_name_key in st.session_state:
                selected_name = st.session_state[selected_name_key]
                if selected_name != "All" and selected_name in name_id_mapping:
                    corresponding_id = name_id_mapping[selected_name]
                    if corresponding_id in employee_ids:
                        id_index = employee_ids.index(corresponding_id)
                    else:
                        id_index = 0
                else:
                    id_index = 0
            else:
                id_index = 0
            
            selected_emp_id = st.selectbox(
                "Filter by Employee ID", 
                employee_ids, 
                index=id_index,
                key=f"emp_id_filter_{st.session_state.filter_reset_counter}"
            )
        
        with filter_col2:
            # Employee Name filter
            employee_names = ["All"] + sorted(results_df["Name"].unique().tolist())
            
            # Check if ID was changed and sync name
            selected_id_key = f"emp_id_filter_{st.session_state.filter_reset_counter}"
            if selected_id_key in st.session_state:
                selected_id = st.session_state[selected_id_key]
                if selected_id != "All" and selected_id in id_name_mapping:
                    corresponding_name = id_name_mapping[selected_id]
                    if corresponding_name in employee_names:
                        name_index = employee_names.index(corresponding_name)
                    else:
                        name_index = 0
                else:
                    name_index = 0
            else:
                name_index = 0
            
            selected_emp_name = st.selectbox(
                "Filter by Employee Name", 
                employee_names, 
                index=name_index,
                key=f"emp_name_filter_{st.session_state.filter_reset_counter}"
            )
        
        with filter_col3:
            # Status filter
            statuses = ["All"] + sorted(results_df["Status"].unique().tolist())
            selected_status = st.selectbox(
                "Filter by Status", 
                statuses, 
                index=0,
                key=f"status_filter_{st.session_state.filter_reset_counter}"
            )
        
        with filter_col4:
            # Test Type/Standard filter
            test_types = ["All"] + sorted(results_df["Test Type"].unique().tolist())
            selected_test_type = st.selectbox(
                "Filter by Test Type", 
                test_types, 
                index=0,
                key=f"test_type_filter_{st.session_state.filter_reset_counter}"
            )
        
        filter_col5, filter_col6, filter_col7, filter_col8 = st.columns(4)
        with filter_col5:
            st.write("")
            if st.button("🗑️ Clear All Filters"):
                st.session_state.filter_reset_counter += 1
                keys_to_remove = [key for key in st.session_state.keys() if key.startswith(('emp_id_filter_', 'emp_name_filter_', 'status_filter_', 'test_type_filter_'))]
                for key in keys_to_remove:
                    del st.session_state[key]
                st.rerun()
        
        # Apply filters - use whichever filter is not "All"
        filtered_df = results_df.copy()
        
        # If either ID or Name is selected (not "All"), filter by that employee
        if selected_emp_id != "All":
            filtered_df = filtered_df[filtered_df["ID"].astype(str) == selected_emp_id]
        elif selected_emp_name != "All":
            filtered_df = filtered_df[filtered_df["Name"] == selected_emp_name]
        
        if selected_status != "All":
            filtered_df = filtered_df[filtered_df["Status"] == selected_status]
        if selected_test_type != "All":
            filtered_df = filtered_df[filtered_df["Test Type"] == selected_test_type]

        # Display sync status
        if selected_emp_id != "All" or selected_emp_name != "All":
            if selected_emp_id != "All":
                display_name = id_name_mapping.get(selected_emp_id, "Unknown")
                st.info(f"🔗 **Employee Selected**: ID: {selected_emp_id} | Name: {display_name}")
            else:
                display_id = name_id_mapping.get(selected_emp_name, "Unknown")
                st.info(f"🔗 **Employee Selected**: Name: {selected_emp_name} | ID: {display_id}")

        st.markdown("---")
        st.subheader("📥 Individual Test Download")
        
        # Use the synchronized selection for individual test download
        if selected_emp_id != "All" or selected_emp_name != "All":
            if selected_emp_id != "All":
                emp_filtered = filtered_df[filtered_df["ID"].astype(str) == selected_emp_id]
                emp_name_display = id_name_mapping.get(selected_emp_id, selected_emp_id)
                emp_id_display = selected_emp_id
            else:
                emp_filtered = filtered_df[filtered_df["Name"] == selected_emp_name]
                emp_name_display = selected_emp_name
                emp_id_display = name_id_mapping.get(selected_emp_name, "Unknown")
            
            if not emp_filtered.empty:
                st.info(f"Showing {len(emp_filtered)} test(s) for employee: **{emp_name_display}** (ID: {emp_id_display})")
                emp_filtered = emp_filtered.sort_values("Date / Time", ascending=False).reset_index(drop=True)
                
                for idx, test_row in emp_filtered.iterrows():
                    with st.expander(f"Test {idx+1}: {test_row['Test Type']} - {test_row['Date / Time']} ({test_row['Status']})", expanded=False):
                        col1, col2, col3 = st.columns([2, 1, 1])
                        with col1:
                            st.metric("Score", f"{test_row['Right']}/{test_row['Total']}")
                            st.metric("Percentage", f"{test_row['Percentage']:.1f}%")
                        with col2:
                            st.metric("Status", test_row['Status'])
                        with col3:
                            csv_data, filename = download_individual_test(
                                test_row['ID'], 
                                test_row['Name'], 
                                test_row
                            )
                            st.download_button(
                                label=f"📄 Download Test Report",
                                data=csv_data,
                                file_name=filename,
                                mime="text/csv",
                                use_container_width=True
                            )
                        st.write("**Test Details:**")
                        st.json({
                            "Employee ID": test_row['ID'],
                            "Employee Name": test_row['Name'],
                            "Standard": test_row['Test Type'],
                            "Total Questions": test_row['Total'],
                            "Correct": test_row['Right'],
                            "Wrong": test_row['Wrong'],
                            "Passing Criteria": f"{test_row['Criteria']}%",
                            "Completed": test_row['Date / Time']
                        })
            else:
                st.warning("No test results found for the selected employee.")
        else:
            st.info("👆 **Select an Employee ID or Name** to view and download individual test reports")
        
        st.markdown("---")
        st.subheader("📊 Test Summary")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Tests", len(filtered_df))
        with col2:
            pass_count = len(filtered_df[filtered_df["Status"] == "Pass"]) if "Status" in filtered_df.columns else 0
            st.metric("Passed", pass_count)
        with col3:
            fail_count = len(filtered_df[filtered_df["Status"] == "Fail"]) if "Status" in filtered_df.columns else 0
            st.metric("Failed", fail_count)
        with col4:
            if "Percentage" in filtered_df.columns and len(filtered_df) > 0:
                avg_score = filtered_df["Percentage"].mean()
                st.metric("Avg Score", f"{avg_score:.1f}%")
            else:
                st.metric("Avg Score", "N/A")
        
        if len(filtered_df) != len(results_df):
            st.info(f"Showing {len(filtered_df)} of {len(results_df)} total records")
        
        st.markdown("---")
        if not filtered_df.empty:
            display_df = filtered_df.copy()
            display_df.insert(0, 'S.No.', range(1, len(display_df) + 1))
            export_col1, export_col2, export_col3 = st.columns([1, 1, 2])
            with export_col1:
                csv = display_df.to_csv(index=False)
                st.download_button(
                    label="📄 Download CSV",
                    data=csv,
                    file_name=f"all_test_results_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            with export_col2:
                if st.button("⚙️ Column Settings"):
                    st.session_state.show_column_settings = not st.session_state.get("show_column_settings", False)
            
            if st.session_state.get("show_column_settings", False):
                st.subheader("Column Visibility")
                cols_to_show = []
                col_settings = st.columns(5)
                for i, col in enumerate(filtered_df.columns):
                    with col_settings[i % 5]:
                        if st.checkbox(col, value=True, key=f"show_{col}"):
                            cols_to_show.append(col)
                filtered_df = filtered_df[cols_to_show] if cols_to_show else filtered_df
                display_df = filtered_df.copy()
                display_df.insert(0, 'S.No.', range(1, len(display_df) + 1))
            
            st.dataframe(
                display_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "S.No.": st.column_config.NumberColumn("S.No.", help="Serial Number", format="%d", width="small"),
                    "Percentage": st.column_config.ProgressColumn("Percentage", help="Test Score Percentage", format="%.1f%%", min_value=0, max_value=100),
                    "Status": st.column_config.TextColumn("Status", help="Pass/Fail Status"),
                    "Total": st.column_config.NumberColumn("Total Questions", help="Total number of questions in the test", format="%d"),
                    "Right": st.column_config.NumberColumn("Correct Answers", help="Number of correct answers", format="%d"),
                    "Wrong": st.column_config.NumberColumn("Wrong Answers", help="Number of wrong answers", format="%d"),
                    "Date / Time": st.column_config.TextColumn("Date / Time", help="Test completion date and time"),
                }
            )
        else:
            st.warning("No results found matching the current filters")
        
        st.markdown("---")
        if st.button("Logout"):
            st.session_state.admin_logged_in = False
            st.session_state.pop("quiz", None)
            st.rerun()
    else:
        st.info("No results available yet in the Result 2 sheet.")
        if st.button("Logout"):
            st.session_state.admin_logged_in = False
            st.session_state.pop("quiz", None)
            st.rerun()
