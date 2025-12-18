import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px
import seaborn as sns
import matplotlib.pyplot as plt

def to_cr(value):
    return value / 1e7
    
st.set_page_config(layout="wide")
st.title('Dispatch Data Dashboard 📊')

page = st.sidebar.radio("Select Page", ['Overview', 'SPD', 'OEM', 'Daywise Dispatch', 'Invoice Value', 'Dispatch Details'])

uploaded_file = st.file_uploader("Upload your Dispatch Data Excel file", type=['xlsx', 'csv'])

if uploaded_file is not None:
    if uploaded_file.name.lower().endswith('.xlsx'):
        dispatch_data = pd.read_excel(uploaded_file)
    else:
        dispatch_data = pd.read_csv(uploaded_file, encoding='latin1')

    dispatch_data.columns = dispatch_data.columns.str.strip()

    dispatch_data.insert(
        dispatch_data.columns.get_loc('Customer Name') + 1,
        'Updated Customer Name',
        np.select(
            [
                dispatch_data['Customer Name'].str.lower().str.startswith('ashok', na=False),
                dispatch_data['Customer Name'].str.lower().str.startswith('tata', na=False) & ~dispatch_data['Customer Name'].str.lower().str.startswith('tata advanced', na=False),
                dispatch_data['Customer Name'].str.lower().str.startswith('blue energy', na=False),
                dispatch_data['Customer Name'].str.lower().str.startswith('force motors', na=False),
                dispatch_data['Customer Name'].str.lower().str.startswith('cnh', na=False),
                dispatch_data['Customer Name'].str.lower().str.startswith('bajaj auto', na=False),
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
        elif material_str.startswith(('78209', '73409','73408')):
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

    dispatch_data['Billing Date'] = pd.to_datetime(dispatch_data['Billing Date'], dayfirst=True, errors='coerce')
    dispatch_data['Cust PO Date'] = pd.to_datetime(dispatch_data['Cust PO Date'], dayfirst=True, errors='coerce')

    dispatch_data.insert(
        dispatch_data.columns.get_loc('Billing Date') + 1,
        'Month-Year',
        dispatch_data['Billing Date'].dt.strftime('%B-%y')
    )

    dispatch_data['Month Start Date'] = pd.to_datetime(
        '01 ' + dispatch_data['Month-Year'], 
        format='%d %B-%y', 
        errors='coerce'
    )

    dispatch_data['Billing Date'] = dispatch_data['Billing Date'].dt.strftime('%d-%m-%Y')
    dispatch_data['Cust PO Date'] = dispatch_data['Cust PO Date'].dt.strftime('%d-%m-%Y')

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

    dispatch_data.loc[dispatch_data['Model New'] == 'M0339 H-Pas', 'Material Category'] = 'Power STG H-Pas'

    dispatch_data.loc[
        dispatch_data['Material'].astype(str).str.endswith('/RF', na=False),
        ['Model New', 'Material Category']
    ] = ['M0339 H-Pas', 'Power STG H-Pas']

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
    if page == 'Overview':
        st.header('Overview Page')
        month_list = sorted(dispatch_data['Month-Year'].dropna().unique().tolist())
        month_list.insert(0, 'All')
        selected_month = st.sidebar.selectbox('Select Month-Year (Overview)', month_list)
        
        overview_data = dispatch_data.copy()
        if selected_month != 'All':
            overview_data = overview_data[overview_data['Month-Year'] == selected_month]

        overview_data = overview_data.sort_values('Month Start Date')
            
        monthly_sales = overview_data.groupby(['Month-Year', 'Month Start Date'])['Basic Amt.LocCur'].sum().reset_index()
        monthly_sales = monthly_sales.sort_values('Month Start Date')
        y_max1 = monthly_sales['Basic Amt.LocCur'].max() * 1.15

        fig_total_sales = px.bar(
            monthly_sales,
            x='Month-Year',
            y='Basic Amt.LocCur',
            title='Total Monthly Sales (Basic Amt.LocCur)',
            labels={'Basic Amt.LocCur': 'Basic Amount (₹)', 'Month-Year': 'Month-Year'},
            text='Basic Amt.LocCur'
        )
        
        fig_total_sales.update_layout(
            yaxis_tickprefix='₹ ',
            xaxis_title='Month-Year',
            uniformtext_minsize=8,
            uniformtext_mode='hide',
            bargap=0.3,
            yaxis=dict(range=[0, y_max1])
        )
            
        fig_total_sales.update_traces(
            texttemplate='₹ %{text:,.0f}',
            textposition='outside',
            marker_line_width=0.5,
            width=0.3
        )
        fig_total_sales.update_traces(texttemplate='₹ %{text:,.0f}', textposition='outside')
        
        split_sales = overview_data[overview_data['Customer Category'].isin(['OEM', 'SPD'])].groupby(
            ['Month-Year', 'Month Start Date', 'Customer Category']
        )['Basic Amt.LocCur'].sum().reset_index()

        split_sales = split_sales.sort_values('Month Start Date')
        y_max2 = split_sales['Basic Amt.LocCur'].max() * 1.15

        fig_oem_spd = px.bar(
            split_sales,
            x='Month-Year',
            y='Basic Amt.LocCur',
            color='Customer Category',
            barmode='group',
            title='OEM & SPD Sales (Basic Amt.LocCur) Month-wise',
            labels={'Basic Amt.LocCur': 'Basic Amount (₹)', 'Month-Year': 'Month-Year'},
            text='Basic Amt.LocCur'
        )
        
        fig_oem_spd.update_layout(
            yaxis_tickprefix='₹ ',
            xaxis_title='Month-Year',
            uniformtext_minsize=8,
            uniformtext_mode='hide',
            bargap=0.3,
            yaxis=dict(range=[0, y_max2])
        )
        
        fig_oem_spd.update_traces(
            texttemplate='₹ %{text:,.0f}',
            textposition='outside',
            marker_line_width=0.5,
            width=0.2
        )

        fig_oem_spd.update_traces(texttemplate='₹ %{text:,.0f}', textposition='outside')
        
        plant_sales = overview_data.groupby('Plant')['Basic Amt.LocCur'].sum().reset_index()
        plant_sales['Plant'] = plant_sales['Plant'].astype(str)

        y_max3 = plant_sales['Basic Amt.LocCur'].max() * 1.15
        
        fig_plant_sales = px.bar(
            plant_sales,
            x='Plant',
            y='Basic Amt.LocCur',
            title='Plant-wise Sales (Basic Amt.LocCur)',
            labels={'Basic Amt.LocCur': 'Basic Amount (₹)', 'Plant': 'Plant'},
            text='Basic Amt.LocCur'
        )
        
        fig_plant_sales.update_layout(
            xaxis=dict(type='category'),
            yaxis_tickprefix='₹ ',
            xaxis_title='Plant',
            uniformtext_minsize=8,
            uniformtext_mode='hide',
            bargap=0.3,
            yaxis=dict(range=[0, y_max3])
        )
        
        fig_plant_sales.update_traces(
            texttemplate='₹ %{text:,.0f}',
            textposition='outside',
            marker_line_width=0.5,
            width=0.3
        )
        fig_plant_sales.update_traces(texttemplate='₹ %{text:,.0f}', textposition='outside')

        st.header("Material Category vs Customer Category (OEM & SPD) – Qty Wise")
        categories_to_include = ['OEM', 'SPD']
        material_categories_to_include = ['Power STG', 'Mechanical Stg', 'Power STG H-Pas']
        
        overview_data = dispatch_data.copy()
        
        overview_data['Inv Qty'] = pd.to_numeric(overview_data['Inv Qty'], errors='coerce').fillna(0)
        overview_data['Kit Qty'] = pd.to_numeric(overview_data['Kit Qty'], errors='coerce').fillna(0)
        
        if selected_month != 'All':
            overview_data = overview_data[overview_data['Month-Year'] == selected_month]
            
        overview_data = overview_data[
            (overview_data['Customer Category'].isin(['OEM', 'SPD'])) &
            (overview_data['Material Category'].isin(['Power STG', 'Mechanical Stg', 'Power STG H-Pas']))
        ]
        
        overview_data['Effective Qty'] = overview_data.apply(
            lambda row: row['Inv Qty']
            if (row['Customer Category'] == 'OEM' or row['Inv Qty'] > 0)
            else row['Kit Qty'],
            axis=1
        )
        
        grouped = overview_data.groupby(['Material Category', 'Customer Category'])['Effective Qty'].sum().reset_index()
        
        fig = px.bar(
            grouped,
            x='Material Category',
            y='Effective Qty',
            color='Customer Category',
            barmode='group',
            text='Effective Qty',
            title='Qty by Material & Customer Category'
        )

        y_max = grouped['Effective Qty'].max() * 1.2

        fig.update_layout(
            xaxis_title='Material Category',
            yaxis_title='Quantity',
            uniformtext_minsize=8,
            uniformtext_mode='hide',
            bargap=0.3,
            yaxis=dict(range=[0, y_max])
        )

        fig.update_traces(texttemplate='%{text:.0f}', textposition='outside', cliponaxis=False)

        st.plotly_chart(fig_total_sales, use_container_width=True)
        st.plotly_chart(fig_oem_spd, use_container_width=True)
        st.plotly_chart(fig_plant_sales, use_container_width=True)
        st.plotly_chart(fig, use_container_width=True)

    elif page == 'SPD':
        st.header('SPD Page')
        spd_data = dispatch_data[dispatch_data['Customer Category'] == 'SPD']
        st.dataframe(spd_data)

    elif page == 'OEM':
        st.header('OEM Dashboard')
        
        oem_df = dispatch_data[dispatch_data['Customer Category'] == 'OEM']
        oem_df['Material Category'] = oem_df['Material Category'].replace('Power STG H-Pas', 'Power STG')
        oem_months = sorted(oem_df['Month-Year'].dropna().unique())
        oem_months_with_all = ['All'] + list(oem_months)

        selected_month = st.sidebar.selectbox('Select Month (OEM):', oem_months_with_all)

        filtered_df = oem_df.copy()
        
        if selected_month != 'All':
            filtered_df = filtered_df[filtered_df['Month-Year'] == selected_month]
            
        updated_customers = sorted(filtered_df['Updated Customer Name'].dropna().unique())
        updated_customers.insert(0, 'All')
        selected_updated_customer = st.sidebar.selectbox("Select Updated Customer Name (OEM):", updated_customers)
        
        if selected_updated_customer != 'All':
            filtered_df = filtered_df[filtered_df['Updated Customer Name'] == selected_updated_customer]
        
        customer_names = sorted(filtered_df['Customer Name'].dropna().unique())
        customer_names.insert(0, 'All')
        selected_customer_name = st.sidebar.selectbox("Select Customer Name (OEM):", customer_names)
        
        if selected_customer_name != 'All':
            filtered_df = filtered_df[filtered_df['Customer Name'] == selected_customer_name]
            
        st.subheader('OEM - Power STG - Customer-wise Quantity')
        oem_power_stg = filtered_df[filtered_df['Material Category'] == 'Power STG']
        oem_power_cust_qty = oem_power_stg.groupby('Updated Customer Name')['Inv Qty'].sum().sort_values(ascending=False)
        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(y=oem_power_cust_qty.index, x=oem_power_cust_qty.values, palette='Blues_r', ax=ax)
        for i, (name, value) in enumerate(zip(oem_power_cust_qty.index, oem_power_cust_qty.values)):
            ax.text(value, i, f'{value:,.0f}', va='center')
        st.pyplot(fig)
        
        st.subheader('OEM - Mechanical Stg - Customer-wise Quantity')
        oem_mech_stg = filtered_df[filtered_df['Material Category'] == 'Mechanical Stg']
        oem_mech_cust_qty = oem_mech_stg.groupby('Updated Customer Name')['Inv Qty'].sum().sort_values(ascending=False)
        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(y=oem_mech_cust_qty.index, x=oem_mech_cust_qty.values, palette='Greens_r', ax=ax)
        for i, (name, value) in enumerate(zip(oem_mech_cust_qty.index, oem_mech_cust_qty.values)):
            ax.text(value, i, f'{value:,.0f}', va='center')
        st.pyplot(fig)
        
        st.subheader('OEM - Customer-wise Total Value (₹)')
        oem_cust_value = filtered_df.groupby('Updated Customer Name')['Basic Amt.LocCur'].sum().sort_values(ascending=False)
        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(y=oem_cust_value.index, x=oem_cust_value.values, palette='Oranges_r', ax=ax)
        for i, (name, value) in enumerate(zip(oem_cust_value.index, oem_cust_value.values)):
            ax.text(value, i, f'₹{value:,.0f}', va='center')
        st.pyplot(fig)
        
        st.subheader('OEM - Model-wise Quantity - Power STG')
        oem_power_model_qty = oem_power_stg.groupby('Model New')['Inv Qty'].sum().sort_values(ascending=False)
        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(y=oem_power_model_qty.index, x=oem_power_model_qty.values, palette='Blues', ax=ax)
        for i, (name, value) in enumerate(zip(oem_power_model_qty.index, oem_power_model_qty.values)):
            ax.text(value, i, f'{value:,.0f}', va='center')
        st.pyplot(fig)
        
        st.subheader('OEM - Model-wise Quantity - Vane Pump')
        oem_vane_pump = filtered_df[filtered_df['Material Category'] == 'Vane Pump']
        oem_vane_model_qty = oem_vane_pump.groupby('Model New')['Inv Qty'].sum().sort_values(ascending=False)
        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(y=oem_vane_model_qty.index, x=oem_vane_model_qty.values, palette='Purples', ax=ax)
        for i, (name, value) in enumerate(zip(oem_vane_model_qty.index, oem_vane_model_qty.values)):
            ax.text(value, i, f'{value:,.0f}', va='center')
        st.pyplot(fig)
            
        st.subheader('OEM - Model-wise Quantity - Mechanical Stg')
        oem_mech_model_qty = oem_mech_stg.groupby('Model New')['Inv Qty'].sum().sort_values(ascending=False)
        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(y=oem_mech_model_qty.index, x=oem_mech_model_qty.values, palette='Greens', ax=ax)
        for i, (name, value) in enumerate(zip(oem_mech_model_qty.index, oem_mech_model_qty.values)):
            ax.text(value, i, f'{value:,.0f}', va='center')
        st.pyplot(fig)
        
        st.subheader('OEM - Top 20 Material + Customer combinations by Basic Amount (₹) with Quantity')
        
        top_mat_cust = (
            filtered_df.groupby(['Material', 'Updated Customer Name'])[['Basic Amt.LocCur', 'Inv Qty']]
            .sum()
            .sort_values(by='Basic Amt.LocCur', ascending=False)
            .head(20)
            .reset_index()
        )
        
        top_mat_cust['Label'] = (
            top_mat_cust['Material'].astype(str) + ' - ' + top_mat_cust['Updated Customer Name'] +
            ' | Qty: ' + top_mat_cust['Inv Qty'].astype(int).astype(str)
        )
        
        fig, ax = plt.subplots(figsize=(12, 8))
        sns.barplot(y=top_mat_cust['Label'], x=top_mat_cust['Basic Amt.LocCur'], palette='rocket', ax=ax)
        ax.set_xlabel("Basic Amount (₹)")
        ax.set_ylabel("Material - Customer")
        ax.set_title("Top 20 Material + Customer combinations by Basic Amt.LocCur")
        
        for i, (val, qty) in enumerate(zip(top_mat_cust['Basic Amt.LocCur'], top_mat_cust['Inv Qty'])):
            ax.text(val, i, f'₹{int(val):,}', va='center')
            
        st.pyplot(fig)
        
        filtered_df = filtered_df.sort_values('Month Start Date')
        
        st.subheader("OEM – Month-wise Revenue Trend (₹ Cr)")
        
        filtered_df = filtered_df.sort_values('Month Start Date')
        
        revenue_monthly = (
            filtered_df
            .groupby(
                ['Month-Year', 'Month Start Date', 'Updated Customer Name']
            )['Basic Amt.LocCur']
            .sum()
            .reset_index()
            .sort_values('Month Start Date')
        )
        
        revenue_monthly['Revenue (Cr)'] = revenue_monthly['Basic Amt.LocCur'] / 1e7
        month_order = (
            revenue_monthly
            .sort_values('Month Start Date')['Month-Year']
            .unique()
            .tolist()
        )

        if selected_updated_customer == 'All':
            fig_revenue = px.line(
                revenue_monthly,
                x='Month-Year',
                y='Revenue (Cr)',
                color='Updated Customer Name',
                markers=True,
                text=revenue_monthly['Revenue (Cr)'],
                category_orders={'Month-Year': month_order},
                title='Month-wise Revenue Comparison – All OEM Customers (₹ Cr)',
            )
        
        else:
            single_cust_df = revenue_monthly[
            revenue_monthly['Updated Customer Name'] == selected_updated_customer
            ]
            fig_revenue = px.line(
                single_cust_df,
                x='Month-Year',
                y='Revenue (Cr)',
                markers=True,
                text=revenue_monthly['Revenue (Cr)'],
                category_orders={'Month-Year': month_order},
                title=f'Month-wise Revenue Trend – {selected_updated_customer} (₹ Cr)',
            )
            fig_revenue.update_traces(
                textposition='top center',
                texttemplate='₹ %{text:.2f} Cr',
                cliponaxis=False
            )
            fig_revenue.update_layout(
                xaxis_title='Month',
                yaxis_title='Revenue (₹ Cr)',
                yaxis_tickformat='.2f',
                legend_title='Customer',
                hovermode='x unified',
                margin=dict(t=80),
                uniformtext_minsize=9,
                uniformtext_mode='hide'
            )
            fig_revenue.update_yaxes(range=[
                0,
                revenue_monthly['Revenue (Cr)'].max() * 1.25
            ])
            
        st.plotly_chart(fig_revenue, use_container_width=True)
        
        st.subheader("OEM – Power STG Quantity Trend")
        filtered_df = filtered_df.sort_values('Month Start Date')
        power_qty_monthly = (
            filtered_df[filtered_df['Material Category'] == 'Power STG']
            .groupby(['Month-Year', 'Month Start Date', 'Updated Customer Name'])['Inv Qty']
            .sum()
            .reset_index()
            .sort_values('Month Start Date')
        
        )
        month_order = (
            power_qty_monthly
            .sort_values('Month Start Date')['Month-Year']
            .unique()
            .tolist()
        )
        if selected_updated_customer == 'All':
            fig_power = px.line(
                power_qty_monthly,
                x='Month-Year',
                y='Inv Qty',
                color='Updated Customer Name',
                markers=True,
                text=power_qty_monthly['Inv Qty'],
                category_orders={'Month-Year': month_order},
                title='Month-wise Power STG Quantity – All OEM Customers'
            )
        
        else:
            single_cust_power = power_qty_monthly[
            power_qty_monthly['Updated Customer Name'] == selected_updated_customer
            ]
            fig_power = px.line(
                single_cust_power,
                x='Month-Year',
                y='Inv Qty',
                markers=True,
                text=single_cust_power['Inv Qty'],
                category_orders={'Month-Year': month_order},
                title=f'Month-wise Power STG Quantity – {selected_updated_customer}'
            )
            fig_power.update_traces(
                textposition='top center',
                texttemplate='%{text:,.0f}',  # 19,200 format
                cliponaxis=False
            )
            
            fig_power.update_layout(
                xaxis_title='Month',
                yaxis_title='Quantity',
                yaxis_tickformat=',',        # Force comma format
                legend_title='Customer',
                hovermode='x unified',
                margin=dict(t=80),
                uniformtext_minsize=9,
                uniformtext_mode='hide'
            )
            
            fig_power.update_yaxes(
                range=[0, power_qty_monthly['Inv Qty'].max() * 1.25],
                separatethousands=True      # Explicitly prevent K/M
            )
            
        st.plotly_chart(fig_power, use_container_width=True)

    elif page == 'Invoice Value':
        st.header('Invoice Value Page')
        
        category_options = ['All', 'OEM', 'SPD', 'OEM + SPD']
        selected_category = st.sidebar.radio('Select Customer Category', category_options)

        filtered_for_customer_list = dispatch_data.copy()
        if selected_category == 'OEM + SPD':
            filtered_data = filtered_for_customer_list[filtered_for_customer_list['Customer Category'].isin(['OEM', 'SPD'])]
        elif selected_category != 'All':
            filtered_data = filtered_for_customer_list[filtered_for_customer_list['Customer Category'] == selected_category]
            
        month_list = sorted(dispatch_data['Month-Year'].dropna().unique().tolist())
        month_list.insert(0, 'All')
        selected_month = st.sidebar.selectbox('Select Month-Year', month_list)
            
        fy_list = sorted(dispatch_data['Financial Year'].dropna().unique().tolist())
        fy_list.insert(0, 'All')
        selected_fy = st.sidebar.selectbox('Select Financial Year', fy_list)
            
        updated_customer_list = sorted(filtered_for_customer_list['Updated Customer Name'].dropna().unique().tolist())
        updated_customer_list.insert(0, 'All')
        selected_updated_customer = st.sidebar.selectbox('Select Updated Customer Name', updated_customer_list)

        model_list = sorted(dispatch_data['Model New'].dropna().unique().tolist())
        model_list.insert(0, 'All')
        selected_model = st.sidebar.selectbox('Select Model New', model_list)

        filtered_for_original = filtered_for_customer_list.copy()
        if selected_updated_customer != 'All':
            filtered_for_original = filtered_for_original[filtered_for_original['Updated Customer Name'] == selected_updated_customer]
                
        customer_list = sorted(filtered_for_original['Customer Name'].dropna().unique().tolist())
        customer_list.insert(0, 'All')
        selected_customer = st.sidebar.selectbox('Select Customer Name', customer_list)
        
        plant_list = sorted(dispatch_data['Plant'].dropna().unique().astype(str).tolist())
        plant_list.insert(0, 'All')
        selected_plant = st.sidebar.selectbox('Select Plant', plant_list)

        material_category_list = sorted(dispatch_data['Material Category'].dropna().unique().tolist())
        material_category_list.insert(0, 'All')
        selected_material_category = st.sidebar.selectbox('Select Material Category', material_category_list)


        st.sidebar.markdown('---')
        st.sidebar.subheader('Invoice No. Filter (Type to Search)')
        
        invoice_numbers = sorted(dispatch_data['Billing Doc No.'].dropna().unique().astype(str).tolist())
        typed_invoice = st.sidebar.text_input('Type Invoice No.')
        suggested_invoices = [inv for inv in invoice_numbers if typed_invoice in inv] if typed_invoice else []

        selected_invoice = st.sidebar.selectbox(
            'Select from Suggestions', 
            ['All'] + suggested_invoices, 
            index=0, 
            key='invoice_value_invoice_filter'
        )

        clear_invoice_filter = st.sidebar.button("Clear Invoice Filter")

        billing_dates = pd.to_datetime(dispatch_data['Billing Date'], dayfirst=True, errors='coerce')

        if selected_month != 'All':
            month_year_date = pd.to_datetime('01 ' + selected_month, format='%d %B-%y', errors='coerce')
            min_date = month_year_date
            max_date = month_year_date + pd.offsets.MonthEnd(0)
        else:
            min_date = billing_dates.min()
            max_date = billing_dates.max()
            
        st.sidebar.markdown('---')
        st.sidebar.subheader('Select Date Range (Billing Date)')
        
        date_range = st.sidebar.date_input(
            "Billing Date Range:",
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )
        
        clear_date_filter = st.sidebar.button("Clear Date Filter")
        
        st.sidebar.markdown('---')
        st.sidebar.subheader('Material Filter (Type to Search)')
        material_numbers = sorted(dispatch_data['Material'].dropna().unique().astype(str).tolist())
        typed_material = st.sidebar.text_input('Type Material')

        suggested_materials = [p for p in material_numbers if typed_material.lower() in p.lower()] if typed_material else []

        selected_material = st.sidebar.selectbox(
            'Select from Suggestions', 
            ['All'] + suggested_materials, 
            index=0, 
            key='invoice_value_material_filter'
        )

        clear_material_filter = st.sidebar.button("Clear Material Filter")
        
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

        if selected_invoice != 'All':
            filtered_data = filtered_data[filtered_data['Billing Doc No.'].astype(str) == selected_invoice]
            
        if not clear_invoice_filter:
            if typed_invoice:
                filtered_data = filtered_data[filtered_data['Billing Doc No.'].astype(str).str.contains(typed_invoice, na=False)] 

        if selected_plant != 'All':
            filtered_data = filtered_data[filtered_data['Plant'].astype(str) == selected_plant]

        if selected_material_category != 'All':
            filtered_data = filtered_data[filtered_data['Material Category'] == selected_material_category]
        
        if selected_model != 'All':
            filtered_data = filtered_data[filtered_data['Model New'] == selected_model]


        filtered_data['Billing Date'] = pd.to_datetime(filtered_data['Billing Date'], dayfirst=True, errors='coerce')
        
        if not clear_date_filter:
            start_date, end_date = date_range
            filtered_data = filtered_data[
                (filtered_data['Billing Date'] >= pd.to_datetime(start_date)) &
                (filtered_data['Billing Date'] <= pd.to_datetime(end_date))
            ]
            
        if not clear_material_filter:
            if typed_material:
                filtered_data = filtered_data[filtered_data['Material'].astype(str).str.lower().str.contains(typed_material.lower(), na=False)]
            elif selected_material != 'All':
                filtered_data = filtered_data[filtered_data['Material'].astype(str) == selected_material]
                
        filtered_data['Billing Date'] = filtered_data['Billing Date'].dt.strftime('%d-%m-%Y')
        
        filtered_data['Qty'] = filtered_data['Inv Qty'] + filtered_data['Kit Qty']
        
        filtered_data['Basic Value Per Item'] = np.where(
            filtered_data['Qty'] > 0,
            filtered_data['Basic Amt.LocCur'] / filtered_data['Qty'],
            0
        )

        # Deduplicate Logic:
        filtered_data['Basic Amt.LocCur'] = pd.to_numeric(filtered_data['Basic Amt.LocCur'], errors='coerce').fillna(0)
        filtered_data['Tax Amount'] = pd.to_numeric(filtered_data['Tax Amount'], errors='coerce').fillna(0)
        filtered_data['Amt.Locl Currency'] = pd.to_numeric(filtered_data['Amt.Locl Currency'], errors='coerce').fillna(0)
        filtered_data['Inv Qty'] = pd.to_numeric(filtered_data['Inv Qty'], errors='coerce').fillna(0)
        filtered_data['Kit Qty'] = pd.to_numeric(filtered_data['Kit Qty'], errors='coerce').fillna(0)

        def invoice_filter(group):
            if (group['Billing Doc No.'].nunique() > 1) or group['Sales Order No'].astype(str).str.startswith('10').any():
                return group
            return group[group['Item'] == 10]
        
        invoice_totals = (
            filtered_data.groupby('Billing Doc No.')[['Basic Amt.LocCur', 'Tax Amount', 'Amt.Locl Currency']]
            .sum()
            .reset_index()
        )
        
        filtered_data = filtered_data.merge(invoice_totals, on='Billing Doc No.', suffixes=('', '_Total'))
        
        mask_item_10 = filtered_data['Item'] == 10
        filtered_data.loc[mask_item_10, 'Basic Amt.LocCur'] = filtered_data.loc[mask_item_10, 'Basic Amt.LocCur_Total']
        filtered_data.loc[mask_item_10, 'Tax Amount'] = filtered_data.loc[mask_item_10, 'Tax Amount_Total']
        filtered_data.loc[mask_item_10, 'Amt.Locl Currency'] = filtered_data.loc[mask_item_10, 'Amt.Locl Currency_Total']
        
        filtered_data = filtered_data.drop(columns=['Basic Amt.LocCur_Total', 'Tax Amount_Total', 'Amt.Locl Currency_Total'])

        filtered_data = filtered_data.groupby('Billing Doc No.').apply(invoice_filter).reset_index(drop=True)

        filtered_data['Qty'] = filtered_data['Inv Qty'] + filtered_data['Kit Qty']
        filtered_data['Basic Value Per Item'] = np.where(
            filtered_data['Qty'] > 0,
            filtered_data['Basic Amt.LocCur'] / filtered_data['Qty'],
            0
        )
        
        filtered_data = filtered_data.drop(columns=['Inv Qty', 'Kit Qty'], errors='ignore')
        
        cols = filtered_data.columns.tolist()
        if 'Material' in cols and 'Qty' in cols and 'Basic Value Per Item' in cols:
            material_idx = cols.index('Material')
            cols.remove('Qty')
            cols.remove('Basic Value Per Item')
            cols.insert(material_idx + 1, 'Qty')
            cols.insert(material_idx + 2, 'Basic Value Per Item')
            filtered_data = filtered_data[cols]
        
        subtotal_data = filtered_data.copy()
        
        basic_amt_sum = subtotal_data['Basic Amt.LocCur'].sum()
        tax_amt_sum = subtotal_data['Tax Amount'].sum()
        amt_loc_sum = subtotal_data['Amt.Locl Currency'].sum()
        
        if 'Month Start Date' in filtered_data.columns:
            filtered_data = filtered_data.drop(columns=['Month Start Date'])
            
        st.dataframe(filtered_data)


    elif page == 'Dispatch Details':
        st.header('Dispatch Details Page')

        category_options = ['All', 'OEM', 'SPD', 'OEM + SPD']
        selected_category = st.sidebar.radio('Select Customer Category', category_options)

        filtered_for_customer_list = dispatch_data.copy()

        if selected_category == 'OEM + SPD':
            filtered_data = filtered_for_customer_list[filtered_for_customer_list['Customer Category'].isin(['OEM', 'SPD'])]
        elif selected_category != 'All':
            filtered_data = filtered_for_customer_list[filtered_for_customer_list['Customer Category'] == selected_category]

        month_list = sorted(dispatch_data['Month-Year'].dropna().unique().tolist())
        month_list.insert(0, 'All')
        selected_month = st.sidebar.selectbox('Select Month-Year', month_list)

        fy_list = sorted(dispatch_data['Financial Year'].dropna().unique().tolist())
        fy_list.insert(0, 'All')
        selected_fy = st.sidebar.selectbox('Select Financial Year', fy_list)

        updated_customer_list = sorted(filtered_for_customer_list['Updated Customer Name'].dropna().unique().tolist())
        updated_customer_list.insert(0, 'All')
        selected_updated_customer = st.sidebar.selectbox('Select Updated Customer Name', updated_customer_list)

        model_list = sorted(dispatch_data['Model New'].dropna().unique().tolist())
        model_list.insert(0, 'All')
        selected_model = st.sidebar.selectbox('Select Model New', model_list)

        filtered_for_original = filtered_for_customer_list.copy()
        if selected_updated_customer != 'All':
            filtered_for_original = filtered_for_original[filtered_for_original['Updated Customer Name'] == selected_updated_customer]

        customer_list = sorted(filtered_for_original['Customer Name'].dropna().unique().tolist())
        customer_list.insert(0, 'All')
        selected_customer = st.sidebar.selectbox('Select Customer Name', customer_list)

        plant_list = sorted(dispatch_data['Plant'].dropna().unique().astype(str).tolist())
        plant_list.insert(0, 'All')
        selected_plant = st.sidebar.selectbox('Select Plant', plant_list)

        material_category_list = sorted(dispatch_data['Material Category'].dropna().unique().tolist())
        material_category_list.insert(0, 'All')
        selected_material_category = st.sidebar.selectbox('Select Material Category', material_category_list)

        billing_dates = pd.to_datetime(dispatch_data['Billing Date'], dayfirst=True, errors='coerce')

        if selected_month != 'All':
            month_year_date = pd.to_datetime('01 ' + selected_month, format='%d %B-%y', errors='coerce')
            min_date = month_year_date
            max_date = month_year_date + pd.offsets.MonthEnd(0)
        else:
            min_date = billing_dates.min()
            max_date = billing_dates.max()

        st.sidebar.markdown('---')
        st.sidebar.subheader('Select Date Range (Billing Date)')

        date_range = st.sidebar.date_input(
            "Billing Date Range:",
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )

        clear_date_filter = st.sidebar.button("Clear Date Filter")

        st.sidebar.markdown('---')
        st.sidebar.subheader('Material Filter (Type to Search)')

        material_numbers = sorted(dispatch_data['Material'].dropna().unique().astype(str).tolist())
        typed_material = st.sidebar.text_input('Type Material')
        suggested_materials = [p for p in material_numbers if typed_material.lower() in p.lower()] if typed_material else []

        selected_material = st.sidebar.selectbox('Select from Suggestions', ['All'] + suggested_materials, index=0)
        clear_material_filter = st.sidebar.button("Clear Material Filter")

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

        if selected_plant != 'All':
            filtered_data = filtered_data[filtered_data['Plant'].astype(str) == selected_plant]

        if selected_material_category != 'All':
            filtered_data = filtered_data[filtered_data['Material Category'] == selected_material_category]
        
        if selected_model != 'All':
            filtered_data = filtered_data[filtered_data['Model New'] == selected_model]

        filtered_data['Billing Date'] = pd.to_datetime(filtered_data['Billing Date'], dayfirst=True, errors='coerce')

        if not clear_date_filter:
            start_date, end_date = date_range
            filtered_data = filtered_data[
                (filtered_data['Billing Date'] >= pd.to_datetime(start_date)) &
                (filtered_data['Billing Date'] <= pd.to_datetime(end_date))
            ]

        filtered_data['Billing Date'] = filtered_data['Billing Date'].dt.strftime('%d-%m-%Y')

        if not clear_material_filter:
            if typed_material:
                filtered_data = filtered_data[filtered_data['Material'].astype(str).str.lower().str.contains(typed_material.lower(), na=False)]
            elif selected_material != 'All':
                filtered_data = filtered_data[filtered_data['Material'].astype(str) == selected_material]

        filtered_data['Inv Qty'] = pd.to_numeric(filtered_data['Inv Qty'], errors='coerce').fillna(0)
        filtered_data['Kit Qty'] = pd.to_numeric(filtered_data['Kit Qty'], errors='coerce').fillna(0)

        inv_qty_sum = filtered_data['Inv Qty'].sum()
        kit_qty_sum = filtered_data['Kit Qty'].sum()
        basic_amt_sum = filtered_data['Basic Amt.LocCur'].sum()

        st.markdown(
            """
            <style>
            .subtotal-box {
                padding: 10px;
                border-radius: 5px;
                border: 1px solid;
                font-weight: bold;
            }
            .subtotal-box-light {
                background-color: #f0f0f0;
                color: #000;
                border-color: #ccc;
            }
            .subtotal-box-dark {
                background-color: #222;
                color: #fff;
                border-color: #555;
            }
            </style>
            """,
            unsafe_allow_html=True
        )

        theme = st.get_option("theme.base")
        box_class = "subtotal-box-light" if theme == "light" else "subtotal-box-dark"

        st.markdown(
            f"""
            <div class="subtotal-box {box_class}">
            Subtotal (Filtered Data):<br>
            Inv Qty: {inv_qty_sum:,.0f} &nbsp;&nbsp;&nbsp;
            Kit Qty: {kit_qty_sum:,.0f} &nbsp;&nbsp;&nbsp;
            Basic Amt.LocCur: ₹ {basic_amt_sum:,.2f}
            </div>
            """,
            unsafe_allow_html=True
        )

        if 'Month Start Date' in filtered_data.columns:
            filtered_data = filtered_data.drop(columns=['Month Start Date'])

        st.dataframe(filtered_data)

    elif page == 'Daywise Dispatch':
        st.header('Daywise Dispatch Page')

        dispatch_data['Inv Qty'] = pd.to_numeric(dispatch_data['Inv Qty'], errors='coerce').fillna(0)
        dispatch_data['Kit Qty'] = pd.to_numeric(dispatch_data['Kit Qty'], errors='coerce').fillna(0)

        filtered_daywise = dispatch_data[
            ~dispatch_data['Material'].astype(str).str.upper().str.startswith('C') &
            (dispatch_data['Material'].astype(str) != '8043975905')
        ].copy()

        def should_keep(row, billing_counts):
            if billing_counts[row['Billing Doc No.']] == 1:
                return True
            if str(row['Sales Order No']).startswith('10'):
                return True
            return row['Item'] == 10

        billing_counts = filtered_daywise['Billing Doc No.'].value_counts()
        filtered_daywise = filtered_daywise[filtered_daywise.apply(lambda row: should_keep(row, billing_counts), axis=1)].copy()

        if 'Total Dispatch' not in filtered_daywise.columns:
            kit_qty_index = filtered_daywise.columns.get_loc('Kit Qty')
            filtered_daywise.insert(kit_qty_index + 1, 'Total Dispatch', filtered_daywise['Inv Qty'] + filtered_daywise['Kit Qty'])

        category_options = ['All', 'OEM', 'SPD', 'OEM + SPD']
        selected_category = st.sidebar.radio('Select Customer Category', category_options)

        filtered_for_customer_list = filtered_daywise.copy()

        if selected_category == 'OEM + SPD':
            filtered_data = filtered_for_customer_list[filtered_for_customer_list['Customer Category'].isin(['OEM', 'SPD'])]
        elif selected_category != 'All':
            filtered_data = filtered_for_customer_list[filtered_for_customer_list['Customer Category'] == selected_category]

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

        plant_list = sorted(dispatch_data['Plant'].dropna().unique().astype(str).tolist())
        plant_list.insert(0, 'All')
        selected_plant = st.sidebar.selectbox('Select Plant', plant_list)

        material_category_list = sorted(dispatch_data['Material Category'].dropna().unique().tolist())
        material_category_list.insert(0, 'All')
        selected_material_category = st.sidebar.selectbox('Select Material Category', material_category_list)

        model_list = sorted(dispatch_data['Model New'].dropna().unique().tolist())
        model_list.insert(0, 'All')
        selected_model = st.sidebar.selectbox('Select Model New', model_list)

        st.sidebar.markdown('---')
        st.sidebar.subheader('Material Filter (Type to Search)')
        material_numbers = sorted(dispatch_data['Material'].dropna().unique().astype(str).tolist())
        typed_material = st.sidebar.text_input('Type Material')
        suggested_materials = [p for p in material_numbers if typed_material.lower() in p.lower()] if typed_material else []

        selected_material = st.sidebar.selectbox('Select from Suggestions', ['All'] + suggested_materials, index=0)
        clear_material_filter = st.sidebar.button("Clear Material Filter")

        billing_dates = pd.to_datetime(filtered_daywise['Billing Date'], dayfirst=True, errors='coerce')

        if selected_month != 'All':
            month_year_date = pd.to_datetime('01 ' + selected_month, format='%d %B-%y', errors='coerce')
            min_date = month_year_date
            max_date = month_year_date + pd.offsets.MonthEnd(0)
        else:
            min_date = billing_dates.min()
            max_date = billing_dates.max()

        st.sidebar.markdown('---')
        st.sidebar.subheader('Select Date Range (Billing Date)')

        date_range = st.sidebar.date_input(
            "Billing Date Range:",
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )

        clear_date_filter = st.sidebar.button("Clear Date Filter")

        final_daywise = filtered_daywise.copy()

        if selected_category != 'All':
            final_daywise = final_daywise[final_daywise['Customer Category'] == selected_category]

        if selected_month != 'All':
            final_daywise = final_daywise[final_daywise['Month-Year'] == selected_month]

        if selected_fy != 'All':
            final_daywise = final_daywise[final_daywise['Financial Year'] == selected_fy]

        if selected_updated_customer != 'All':
            final_daywise = final_daywise[final_daywise['Updated Customer Name'] == selected_updated_customer]

        if selected_customer != 'All':
            final_daywise = final_daywise[final_daywise['Customer Name'] == selected_customer]

        if selected_plant != 'All':
            final_daywise = final_daywise[final_daywise['Plant'].astype(str) == selected_plant]

        if selected_material_category != 'All':
            final_daywise = final_daywise[final_daywise['Material Category'] == selected_material_category]

        if selected_model != 'All':
            final_daywise = final_daywise[final_daywise['Model New'] == selected_model]

        final_daywise['Billing Date'] = pd.to_datetime(final_daywise['Billing Date'], dayfirst=True, errors='coerce')

        if not clear_date_filter:
            start_date, end_date = date_range
            final_daywise = final_daywise[
                (final_daywise['Billing Date'] >= pd.to_datetime(start_date)) &
                (final_daywise['Billing Date'] <= pd.to_datetime(end_date))
            ]

        if not clear_material_filter:
            if typed_material:
                final_daywise = final_daywise[final_daywise['Material'].astype(str).str.lower().str.contains(typed_material.lower(), na=False)]
            elif selected_material != 'All':
                final_daywise = final_daywise[final_daywise['Material'].astype(str) == selected_material]

        final_daywise['Billing Date'] = pd.to_datetime(final_daywise['Billing Date'], dayfirst=True, errors='coerce')

        pivot_table = final_daywise.pivot_table(
            index=['Sold-to Party', 'Customer Name', 'Material', 'Plant'],
            columns='Billing Date',
            values='Total Dispatch',
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        pivot_table.columns = [
            col.strftime('%d-%m-%Y') if isinstance(col, pd.Timestamp) else col
            for col in pivot_table.columns
        ]

        pivot_table.columns.name = None

        st.dataframe(pivot_table)










