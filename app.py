import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns


st.set_page_config(layout="wide")
st.title('Dispatch Data Dashboard 📊 (Raw File Upload)')


uploaded_file = st.file_uploader("Upload your OLD Sales Register Excel file", type=['xlsx'])

if uploaded_file:
    dispatch_df = pd.read_excel(uploaded_file)
    dispatch_df = dispatch_df[dispatch_df['Customer Group'] == 10]
    st.success('File Uploaded Successfully!')

    # ---------------- DATA CLEANING ----------------
    # Customer Name Mapping
    def update_customer_name(row):
        name = str(row['Customer Name']).lower()
        sold_to = str(row['Sold-to Party']).upper()
        if name.startswith('ashok'):
            return 'Ashok Leyland'
        elif name.startswith('tata') and 'advanced' not in name:
            return 'Tata Motors'
        elif name.startswith('blue energy'):
            return 'Blue Energy'
        elif name.startswith('force'):
            return 'Force Motors'
        elif name.startswith('cnh'):
            return 'CNH'
        elif name.startswith('bajaj auto'):
            return 'Bajaj Auto'
        elif sold_to in ['M0163', 'M0164', 'M0231']:
            return 'Mahindra Swaraj'
        elif sold_to in ['M0009', 'M0010', 'M0221']:
            return 'M&M'
        else:
            return row['Customer Name']

    dispatch_df['Updated Customer Name'] = dispatch_df.apply(update_customer_name, axis=1)


    # Material Category Mapping
    def categorize_material(material):
        material = str(material)
        if material.startswith(('80339', '80349', '80379', '80439', '80469', '80489', '80499', 'M0339', 'M0439', '88439')):
            return 'Power STG'
        elif material.startswith(('76139', '76729', '76739', '76749', '76919')):
            return 'Vane Pump'
        elif material.startswith(('78209', '73409')):
            return 'Mechanical Stg'
        elif material.startswith('78609'):
            return 'Bevel gear'
        elif material[4:7] == '012':
            return 'Drop Arm'
        elif material[4:7] == '472':
            return 'Oil Tank'
        else:
            return 'Child Parts'

    dispatch_df['Material Category'] = dispatch_df['Material'].apply(categorize_material)


    # Month-Year Column
    dispatch_df['Billing Date'] = pd.to_datetime(dispatch_df['Billing Date'])
    dispatch_df['Month-Year'] = dispatch_df['Billing Date'].dt.strftime('%b-%y')


    # Model New (First 5 Digits)
    dispatch_df['Model New'] = dispatch_df['Material'].astype(str).str[:5]


    # ---------------- CHARTS SAME AS BEFORE ----------------
    month_selected = st.sidebar.selectbox('Select Month:', sorted(dispatch_df['Month-Year'].dropna().unique()))

    month_data = dispatch_df[dispatch_df['Month-Year'] == month_selected]


    # ---------------- POWER STG Customer-wise Qty ----------------
    st.header('Power STG - Customer-wise Quantity')
    power_stg = month_data[month_data['Material Category'] == 'Power STG']
    power_cust_qty = power_stg.groupby('Updated Customer Name')['Inv Qty'].sum().sort_values(ascending=False)

    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(y=power_cust_qty.index, x=power_cust_qty.values, palette='Blues_r', ax=ax)
    for i, (name, value) in enumerate(zip(power_cust_qty.index, power_cust_qty.values)):
        ax.text(value, i, f'{value:,.0f}', va='center')
    st.pyplot(fig)


    # ---------------- MECHANICAL STG Customer-wise Qty ----------------
    st.header('Mechanical Stg - Customer-wise Quantity')
    mech_stg = month_data[month_data['Material Category'] == 'Mechanical Stg']
    mech_cust_qty = mech_stg.groupby('Updated Customer Name')['Inv Qty'].sum().sort_values(ascending=False)

    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(y=mech_cust_qty.index, x=mech_cust_qty.values, palette='Greens_r', ax=ax)
    for i, (name, value) in enumerate(zip(mech_cust_qty.index, mech_cust_qty.values)):
        ax.text(value, i, f'{value:,.0f}', va='center')
    st.pyplot(fig)


    # ---------------- Customer-wise Total Value ----------------
    st.header('Updated Customer Name - Total Value (₹)')
    cust_value = month_data.groupby('Updated Customer Name')['Basic Amt.LocCur'].sum().sort_values(ascending=False)

    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(y=cust_value.index, x=cust_value.values, palette='Oranges_r', ax=ax)
    for i, (name, value) in enumerate(zip(cust_value.index, cust_value.values)):
        ax.text(value, i, f'{value:,.0f}', va='center')
    st.pyplot(fig)


    # ---------------- MODEL-WISE Power STG ----------------
    st.header('Model-wise Quantity - Power STG')
    power_model_qty = power_stg.groupby('Model New')['Inv Qty'].sum().sort_values(ascending=False)

    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(y=power_model_qty.index, x=power_model_qty.values, palette='Blues', ax=ax)
    for i, (name, value) in enumerate(zip(power_model_qty.index, power_model_qty.values)):
        ax.text(value, i, f'{value:,.0f}', va='center')
    st.pyplot(fig)


    # ---------------- MODEL-WISE Vane Pump ----------------
    st.header('Model-wise Quantity - Vane Pump')
    vane_pump = month_data[month_data['Material Category'] == 'Vane Pump']
    vane_model_qty = vane_pump.groupby('Model New')['Inv Qty'].sum().sort_values(ascending=False)

    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(y=vane_model_qty.index, x=vane_model_qty.values, palette='Purples', ax=ax)
    for i, (name, value) in enumerate(zip(vane_model_qty.index, vane_model_qty.values)):
        ax.text(value, i, f'{value:,.0f}', va='center')
    st.pyplot(fig)


    # ---------------- MODEL-WISE Mechanical Stg ----------------
    st.header('Model-wise Quantity - Mechanical Stg')
    mech_model_qty = mech_stg.groupby('Model New')['Inv Qty'].sum().sort_values(ascending=False)

    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(y=mech_model_qty.index, x=mech_model_qty.values, palette='Greens', ax=ax)
    for i, (name, value) in enumerate(zip(mech_model_qty.index, mech_model_qty.values)):
        ax.text(value, i, f'{value:,.0f}', va='center')
    st.pyplot(fig)


    # ---------------- CUSTOMER INPUT (PARTIAL) - Power STG Model-wise Qty ----------------
    st.header('Customer-wise (Partial Input) - Model-wise Quantity - Power STG')
    customer_input = st.text_input('Type Customer (Partial allowed):').lower()

    matching_customers = power_stg['Updated Customer Name'].dropna().unique()
    matching_customer = [c for c in matching_customers if customer_input in c.lower()]

    if matching_customer:
        customer_selected = matching_customer[0]
        st.write(f'Auto-selected: **{customer_selected}**')

        cust_power_stg = power_stg[power_stg['Updated Customer Name'] == customer_selected]
        cust_model_qty = cust_power_stg.groupby('Model New')['Inv Qty'].sum().sort_values(ascending=False)

        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(y=cust_model_qty.index, x=cust_model_qty.values, palette='Blues', ax=ax)
        for i, (name, value) in enumerate(zip(cust_model_qty.index, cust_model_qty.values)):
            ax.text(value, i, f'{value:,.0f}', va='center')
        st.pyplot(fig)
    else:
        if customer_input:
            st.warning('No matching customer found.')

    # ---------------- DOWNLOAD CLEANED DATA ----------------
    st.header('Download Cleaned Data')
    csv = dispatch_df.to_csv(index=False).encode('utf-8')
    st.download_button("Download Processed Data as CSV", data=csv, file_name='Processed_Dispatch_Data.csv', mime='text/csv')

else:
    st.info('Please upload your raw Excel file to continue.')
