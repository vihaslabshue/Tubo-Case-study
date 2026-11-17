import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Slow-Stock Action Tool", layout="wide")
st.title("Company X: Slow-Stock Action Tool")
st.markdown("Upload the monthly combined Excel file to generate targeted sales lists.")

uploaded_file = st.file_uploader("Upload 'Case_Data_Technology_and_Data_Management_All_Sheets.xlsx'", type=['xlsx'])

if uploaded_file:
    # 1. Load Data
    items = pd.read_excel(uploaded_file, sheet_name='Item_Master')
    stock = pd.read_excel(uploaded_file, sheet_name='Stock_Aging')
    sales = pd.read_excel(uploaded_file, sheet_name='Sales_Lines')
    customers = pd.read_excel(uploaded_file, sheet_name='Customers')

    # Clean headers (lowercase and strip spaces)
    for df in [items, stock, sales, customers]:
        df.columns = df.columns.str.strip().str.lower()
    
    # [FIX APPLIED HERE: Exact column mapping for Kothari Excel file]
    items.rename(columns={
        'item name': 'name', 
        'master pack qty': 'master_pack', 
        'unit cost (rs)': 'unit_cost', 
        'alternate item (name)': 'alt_name', 
        'alternate flag': 'alt_flag'
    }, inplace=True)
    
    sales.rename(columns={
        'customerid': 'customer',
        'qty': 'quantity',
        'value (rs)': 'value'
    }, inplace=True)
    
    customers.rename(columns={
        'customerid': 'customer',
        'customername': 'name'
    }, inplace=True)

    # Clean strings of trailing spaces
    for df in [items, stock, sales, customers]:
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip()

    # 2. Data Problems: Check for Inactive holding stock
    inactive_skus = items[items['status'] == 'Inactive']['sku'].tolist()
    inactive_stock = stock[stock['sku'].isin(inactive_skus)].copy()
    
    stock_cols = ['qty 0-30 days', 'qty 31-60 days', 'qty 61-90 days', 'qty 91-180 days', 'qty 180+ days']
    
    for c in stock_cols: 
        inactive_stock[c] = pd.to_numeric(inactive_stock[c], errors='coerce').fillna(0)
    
    inactive_holding = inactive_stock[inactive_stock[stock_cols].sum(axis=1) > 0]
    if not inactive_holding.empty:
        st.warning(f"⚠️ Data Problem Found: {len(inactive_holding)} inactive items hold warehouse stock. They have been removed from the action list per Rule 2, but notify Supply Chain.")

    # 3. Rule 1: Pack-Size Switches
    active_items = items[items['status'] == 'Active'].copy()
    switch_items = active_items[active_items['alt_flag'] == 'SWITCH']
    name_to_sku = items.set_index('name')['sku'].to_dict()
    
    sku_map = {}
    for _, row in switch_items.iterrows():
        if row['alt_name'] in name_to_sku:
            sku_map[row['sku']] = name_to_sku[row['alt_name']]
    
    stock['mapped_sku'] = stock['sku'].map(lambda x: sku_map.get(x, x))
    sales['mapped_sku'] = sales['sku'].map(lambda x: sku_map.get(x, x))
    
    final_items = active_items[~active_items['sku'].isin(sku_map.keys())].copy()

    # 4. Aggregations (Rule 3 & 4)
    for c in stock_cols: 
        stock[c] = pd.to_numeric(stock[c], errors='coerce').fillna(0)
    agg_stock = stock.groupby('mapped_sku')[stock_cols].sum().reset_index().rename(columns={'mapped_sku': 'sku'})
    
    sales['quantity'] = pd.to_numeric(sales['quantity'], errors='coerce').fillna(0)
    agg_sales = sales.groupby('mapped_sku')['quantity'].sum().reset_index().rename(columns={'mapped_sku': 'sku', 'quantity': 'sales_6m'})

    master = pd.merge(final_items, agg_stock, on='sku', how='left').fillna(0)
    master = pd.merge(master, agg_sales, on='sku', how='left').fillna(0)

    master['aged_stock'] = master['qty 91-180 days'] + master['qty 180+ days']
    master['total_stock'] = master[stock_cols].sum(axis=1)
    
    master['unit_cost'] = pd.to_numeric(master['unit_cost'], errors='coerce')
    master['master_pack'] = pd.to_numeric(master['master_pack'], errors='coerce')
    master['aged_value'] = master['aged_stock'] * master['unit_cost']
    master['avg_mo_sales'] = master['sales_6m'] / 6
    master['months_to_clear'] = np.where(master['avg_mo_sales'] > 0, master['total_stock'] / master['avg_mo_sales'], float('inf'))

    # 5. Classification (Rule 5)
    def classify_item(r):
        if r['sales_6m'] < r['master_pack']: return 'Red List'
        elif r['aged_stock'] <= 0 or r['months_to_clear'] < 3: return 'No Action'
        else: return 'Push'
    
    master['class'] = master.apply(classify_item, axis=1)

    # UI Rendering
    st.header("1. Executive Summary")
    col1, col2, col3, col4 = st.columns(4)
    counts = master['class'].value_counts()
    col1.metric("Push Items", counts.get('Push', 0))
    col2.metric("No Action Items", counts.get('No Action', 0))
    col3.metric("Red List Items", counts.get('Red List', 0))
    
    push_items = master[master['class'] == 'Push']
    col4.metric("Total Aged Value (Push)", f"₹{push_items['aged_value'].sum():,.0f}")

    st.header("2. Top 10 Push Items by Aged Value")
    top_10 = push_items.sort_values('aged_value', ascending=False).head(10)
    st.dataframe(top_10[['sku', 'name', 'category', 'aged_stock', 'aged_value', 'months_to_clear']], use_container_width=True)

    st.header("3. Customer Targets for Sales (Mobile View)")
    selected_sku = st.selectbox("Select a Push Item to generate the call list:", push_items['sku'].tolist())
    
    if selected_sku:
        # Rules 6 & 7: Customer Targeting
        sales['value'] = pd.to_numeric(sales['value'], errors='coerce')
        cust_totals = sales.groupby('customer')['value'].sum().reset_index()
        elig_custs = cust_totals[cust_totals['value'] > 100000]['customer'].tolist()
        
        sku_sales = sales[sales['mapped_sku'] == selected_sku]
        b1 = set(sku_sales['customer'].unique()).intersection(set(elig_custs))
        
        buyers = set(sku_sales['customer'].unique())
        buyer_inds = set(customers[customers['customer'].isin(buyers)]['industry'].dropna().astype(str).str.strip().unique())
        for v in ['nan', 'None', '', 'NaN']: buyer_inds.discard(v)
        
        b2 = set(customers[customers['industry'].isin(buyer_inds)]['customer'].unique()).intersection(set(elig_custs)) - b1
        
        cat = master[master['sku'] == selected_sku]['category'].iloc[0]
        cat_skus = master[master['category'] == cat]['sku'].tolist()
        b3 = set(sales[sales['mapped_sku'].isin(cat_skus)]['customer'].unique()).intersection(set(elig_custs)) - b1 - b2

        targets = []
        for c in elig_custs:
            if c in b1: b = 1
            elif c in b2: b = 2
            elif c in b3: b = 3
            else: continue
            
            v = cust_totals[cust_totals['customer'] == c]['value'].iloc[0]
            cname_row = customers[customers['customer'] == c]
            cname = cname_row['name'].iloc[0] if not cname_row.empty else 'Unknown'
            targets.append({'Bucket': b, 'Customer Name': cname, 'Total 6-Mo Spend (₹)': v, 'CustomerID': c})
        
        if targets:
            tdf = pd.DataFrame(targets).sort_values(by=['Bucket', 'Total 6-Mo Spend (₹)'], ascending=[True, False])
            st.dataframe(tdf, use_container_width=True, hide_index=True)
        else:
            st.info("No eligible target customers found for this item.")

    st.header("4. Export Master List")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        master.to_excel(writer, sheet_name='Master_Classifications', index=False)
    st.download_button(label="Download Full Results to Excel", data=output.getvalue(), file_name="Slow_Stock_Action_List.xlsx", mime="application/vnd.ms-excel")