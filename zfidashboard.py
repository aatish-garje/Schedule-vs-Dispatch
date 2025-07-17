import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(layout="wide")
st.title('Dispatch Data Dashboard 📊')

page = st.sidebar.radio("Select Page", ['Overview', 'SPD', 'OEM', 'Daywise Dispatch', 'Invoice Value', 'Dispatch Details'])

uploaded_file = st.file_uploader("Upload your Dispatch Data Excel file", type=['xlsx', 'csv'])

if uploaded_file is not None:
    if uploaded_file.name.endswith('.xlsx'):
        dispatch_data = pd.read_excel(uploaded_file)
    else:
        dispatch_data = pd.read_csv(uploaded_file)

    dispatch_data.columns = dispatch_data.columns.str.strip()

    dispatch_data.insert(
        dispatch_data.columns.get_loc('Customer Name') + 1,
        'Updated Customer Name',
        np.select(
            [
                dispatch_data['Customer Name'].str.lower().str.startswith('ashok'),
                dispatch_data['Customer Name'].str.lower().str.startswith('tata') & ~dispatch_data['Customer Name'].str.lower().str.startswith('tata advanced'),
                dispatch_data['Customer Name'].str.lower().str.startswith('blue energy'),
                dispatch_data['Customer Name'].str.lower().str.startswith('force motors'),
                dispatch_data['Customer Name'].str.lower().str.startswith('cnh'),
                dispatch_data['Customer Name'].str.lower().str.startswith('bajaj auto'),
                dispatch_data['Sold-to Party'].str.upper().isin(['M0163', 'M0164', 'M0231']),
                dispatch_data['Sold-to Party'].str.upper().isin(['M0009', 'M0010', 'M0221']),
            ],
            [
                'Ashok Leyland',
                'Tata Motors',
                'Blue Energy',
                'Force Motors',
                'CNH',
                'Bajaj Auto',
                'Mahindra Swaraj',
                'M&M'
            ],
            default=dispatch_data['Customer Name']
        )
    )

    def categorize_material(material):
        material_str = str(material)
        if material_str == '7632975501':
            return 'Oil Tank'
        elif material_str.startswith(('80339', '80349', '80379', '80439', '80469', '80489', '80499', 'M0339', 'M0439', '88439')):
            return 'Power STG'
        elif material_str.startswith(('76139', '76729', '76739', '76749', '76919')):
            return 'Vane Pump'
        elif material_str.startswith(('78209', '73409')):
            return 'Mechanical Stg'
        elif material_str.startswith('78609'):
            return 'Bevel Gear'
        elif len(material_str) >= 7 and material_str[4:7] == '012':
            return 'Drop Arm'
        elif len(material_str) >= 7 and material_str[4:7] == '472':
            return 'Oil Tank'
        else:
            return 'Child Parts'

    dispatch_data.insert(
        dispatch_data.columns.get_loc('Material') + 1,
        'Material Category',
        dispatch_data['Material'].apply(categorize_material)
    )

    dispatch_data['Billing Date'] = pd.to_datetime(dispatch_data['Billing Date'], errors='coerce').dt.strftime('%d-%m-%Y')
    dispatch_data['Cust PO Date'] = pd.to_datetime(dispatch_data['Cust PO Date'], errors='coerce').dt.strftime('%d-%m-%Y')
    month_year = pd.to_datetime(dispatch_data['Billing Date'], dayfirst=True, errors='coerce').dt.strftime('%B-%y')
    dispatch_data.insert(
        dispatch_data.columns.get_loc('Billing Date') + 1,
        'Month-Year',
        month_year
    )

    dispatch_data.insert(
        dispatch_data.columns.get_loc('Material'),
        'Model New',
        dispatch_data['Material'].astype(str).str[:5]
    )

    specific_customers = ['C0003', 'F0006', 'G1044', 'I0047', 'M0163', 'M0231', 'T0138']
    dispatch_data.loc[
        (dispatch_data['Sold-to Party'].str.upper().isin(specific_customers)) &
        (dispatch_data['Model New'].str.lower() == 'm0339'),
        'Model New'
    ] = 'M0339 H-Pas'

    dispatch_data['Customer Group'] = dispatch_data['Customer Group'].astype(str).str.strip().str.replace('.0', '', regex=False)
    dispatch_data.insert(
        dispatch_data.columns.get_loc('Customer Group') + 1,
        'Customer Category',
        dispatch_data['Customer Group'].apply(
            lambda x: 'OEM' if x == '10' else
                      'SPD' if x.isdigit() and 11 <= int(x) <= 15 else
                      'Internal'
        )
    )

    def assign_financial_year(date_str):
        try:
            date = pd.to_datetime(date_str, dayfirst=True)
            if date.month >= 4:
                fy_start = date.year
                fy_end = date.year + 1
            else:
                fy_start = date.year - 1
                fy_end = date.year
            return f"FY {fy_start}-{str(fy_end)[-2:]}"
        except:
            return None

    dispatch_data['Financial Year'] = dispatch_data['Billing Date'].apply(assign_financial_year)

    # ------------------------- Pages -------------------------
    if page == 'Overview':
        st.header('Overview Page')
        st.dataframe(dispatch_data)

    elif page == 'SPD':
        st.header('SPD Page')
        spd_data = dispatch_data[dispatch_data['Customer Category'] == 'SPD']
        st.dataframe(spd_data)

    elif page == 'OEM':
        st.header('OEM Page')
        oem_data = dispatch_data[dispatch_data['Customer Category'] == 'OEM']
        st.dataframe(oem_data)

    elif page == 'Daywise Dispatch':
        st.header('Daywise Dispatch Page')
        st.dataframe(dispatch_data)

    elif page == 'Invoice Value':
        st.header('Invoice Value Page')
        st.dataframe(dispatch_data)

    elif page == 'Dispatch Details':
        st.header('Dispatch Details Page')

        category_options = ['All', 'OEM', 'SPD']
        selected_category = st.sidebar.radio('Select Customer Category', category_options)

        filtered_for_customer_list = dispatch_data.copy()
        if selected_category != 'All':
            filtered_for_customer_list = filtered_for_customer_list[filtered_for_customer_list['Customer Category'] == selected_category]

        month_list = sorted(dispatch_data['Month-Year'].dropna().unique().tolist())
        month_list.insert(0, 'All')
        selected_month = st.sidebar.selectbox('Select Month-Year', month_list)

        fy_list = sorted(dispatch_data['Financial Year'].dropna().unique().tolist())
        fy_list.insert(0, 'All')
        selected_fy = st.sidebar.selectbox('Select Financial Year', fy_list)

        updated_customer_list = sorted(filtered_for_customer_list['Updated Customer Name'].dropna().unique().tolist())
        updated_customer_list.insert(0, 'All')
        selected_updated_customer = st.sidebar.selectbox('Select Updated Customer Name', updated_customer_list)

        filtered_for_original = filtered_for_customer_list.copy()
        if selected_updated_customer != 'All':
            filtered_for_original = filtered_for_original[filtered_for_original['Updated Customer Name'] == selected_updated_customer]

        customer_list = sorted(filtered_for_original['Customer Name'].dropna().unique().tolist())
        customer_list.insert(0, 'All')
        selected_customer = st.sidebar.selectbox('Select Customer Name', customer_list)

        # Date Range
        st.sidebar.markdown('---')
        st.sidebar.subheader('Select Date Range (Billing Date)')
        billing_dates = pd.to_datetime(dispatch_data['Billing Date'], dayfirst=True, errors='coerce')
        min_date = billing_dates.min()
        max_date = billing_dates.max()

        start_date, end_date = st.sidebar.date_input(
            "Billing Date Range:",
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )

        # Apply Filters
        filtered_data = dispatch_data.copy()

        if selected_category != 'All':
            filtered_data = filtered_data[filtered_data['Customer Category'] == selected_category]

        if selected_month != 'All':
            filtered_data = filtered_data[filtered_data['Month-Year'] == selected_month]

        if selected_fy != 'All':
            filtered_data = filtered_data[filtered_data['Financial Year'] == selected_fy]

        if selected_updated_customer != 'All':
            filtered_data = filtered_data[filtered_data['Updated Customer Name'] == selected_updated_customer]

        if selected_customer != 'All':
            filtered_data = filtered_data[filtered_data['Customer Name'] == selected_customer]

        filtered_data['Billing Date'] = pd.to_datetime(filtered_data['Billing Date'], dayfirst=True, errors='coerce')
        filtered_data = filtered_data[
            (filtered_data['Billing Date'] >= pd.to_datetime(start_date)) &
            (filtered_data['Billing Date'] <= pd.to_datetime(end_date))
        ]

        st.dataframe(filtered_data)

        def convert_df_to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Filtered Data')
            processed_data = output.getvalue()
            return processed_data

        excel_file = convert_df_to_excel(filtered_data)
        st.download_button(
            label="📥 Download Filtered Data as Excel",
            data=excel_file,
            file_name="filtered_dispatch_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
