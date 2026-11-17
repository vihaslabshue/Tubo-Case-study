st.markdown("<h2 style='color: #1F4E78; border-bottom: 2px solid #1F4E78; padding-bottom: 10px;'>4. Export Master List</h2>", unsafe_allow_html=True)
    output = io.BytesIO()
    
    # We must import openpyxl tools
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Export the main sheets
        master.to_excel(writer, sheet_name='Master_Data', index=False)
        
        # Also create a separate sheet just for the Push items
        push_only = master[master['class'] == 'Push'].sort_values('aged_value', ascending=False)
        push_only.to_excel(writer, sheet_name='Actionable_Push_Items', index=False)
        
        # Now, apply professional styling
        workbook = writer.book
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        border = Border(left=Side(style='thin', color='D9D9D9'),
                        right=Side(style='thin', color='D9D9D9'),
                        top=Side(style='thin', color='D9D9D9'),
                        bottom=Side(style='thin', color='D9D9D9'))
        
        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions
            
            for col_idx, col in enumerate(ws.columns, 1):
                max_length = 0
                col_letter = get_column_letter(col_idx)
                header_name = str(ws.cell(row=1, column=col_idx).value).lower()
                
                for row_idx, cell in enumerate(col, 1):
                    cell.border = border
                    
                    if row_idx == 1:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        
                    elif cell.value is not None:
                        if isinstance(cell.value, (int, float)):
                            if 'value' in header_name or 'cost' in header_name:
                                cell.number_format = '₹#,##0.00'
                            elif 'months' in header_name or 'avg' in header_name:
                                cell.number_format = '0.0'
                            else:
                                cell.number_format = '#,##0'
                    
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                ws.column_dimensions[col_letter].width = min(max_length + 2, 35)

    st.download_button(
        label="Download Professional Results to Excel", 
        data=output.getvalue(), 
        file_name="Professional_Slow_Stock_Action_List.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )