import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side

st.set_page_config(layout="wide")
st.title("Schedule vs Dispatch Report (Manual Upload, Auto Kit File)")

# --- Manual Uploads ---
sales_register_file = st.file_uploader("Upload Sales Register (xlsx)", type="xlsx")
schedule_file = st.file_uploader("Upload Schedule (xlsx)", type="xlsx")
fg_file = st.file_uploader("Upload FG File (xlsx)", type="xlsx")

# --- Auto Kit File from Google Drive ---
kit_part_url = "https://drive.google.com/uc?id=18YkiGvirKsrrwg8IZq2H3Aje5HAw-Djp"
kit_file = BytesIO(requests.get(kit_part_url).content)

if not (sales_register_file and schedule_file and fg_file):
    st.stop()

dispatch_df = pd.read_excel(sales_register_file)
schedule_power = pd.read_excel(schedule_file, sheet_name="POWER", header=3)
schedule_mech = pd.read_excel(schedule_file, sheet_name="MECH", header=3)
fg_df = pd.read_excel(fg_file)

kit_psg = pd.read_excel(kit_file, sheet_name="PSG", usecols="K:M")
kit_psg.columns = ["Part Number", "Desc", "Kit Part Number"]
kit_psg["Part Number"] = kit_psg["Part Number"].astype(str)
lookup_power_stg = dict(zip(kit_psg["Part Number"], kit_psg["Kit Part Number"]))

kit_psg_mech = pd.read_excel(kit_file, sheet_name="PSG", usecols="S:T")
kit_psg_mech.columns = ["Part Number", "Kit Part Number"]
kit_psg_mech["Part Number"] = kit_psg_mech["Part Number"].astype(str)
lookup_mech = dict(zip(kit_psg_mech["Part Number"], kit_psg_mech["Kit Part Number"]))

kit_vp = pd.read_excel(kit_file, sheet_name="VP", usecols="B:D")
kit_vp.columns = ["Part Number", "Desc", "Kit Part Number"]
kit_vp["Part Number"] = kit_vp["Part Number"].astype(str)
lookup_power_vp = dict(zip(kit_vp["Part Number"], kit_vp["Kit Part Number"]))

# --- Dispatch Logic ---
dispatch_df['Sold-to Party'] = dispatch_df.apply(lambda x: str(x['Sold-to Party']) + '.' if (x['Plant'] == 2000 and not str(x['Sold-to Party']).upper().startswith('V')) else x['Sold-to Party'], axis=1)
dispatch_df = dispatch_df[~dispatch_df['Material'].astype(str).str.startswith('C')]
dispatch_df.loc[dispatch_df['Inv Qty'] == 0, 'Inv Qty'] = dispatch_df['Kit Qty']
dispatch_df = dispatch_df[dispatch_df['Material'] != 8043975905]
dispatch_df = dispatch_df[dispatch_df['Customer Group'] == 10]

keep_sales_order_10 = dispatch_df[dispatch_df['Sales Order No'].astype(str).str.startswith('10')]
remaining = dispatch_df[~dispatch_df['Sales Order No'].astype(str).str.startswith('10')]
duplicates = remaining['Billing Doc No.'].value_counts()[lambda x: x > 1].index.tolist()
duplicates_df = remaining[remaining['Billing Doc No.'].isin(duplicates)]
unique_df = remaining[~remaining['Billing Doc No.'].isin(duplicates)]
duplicates_df = duplicates_df[duplicates_df['Item'] == 10]
dispatch_df = pd.concat([keep_sales_order_10, unique_df, duplicates_df], ignore_index=True)

dispatch_summary = dispatch_df.groupby(['Sold-to Party', 'Material'], as_index=False)['Inv Qty'].sum()
dispatch_summary.rename(columns={'Inv Qty': 'Dispatch Qty'}, inplace=True)

schedule_power['Code'] = schedule_power['Code'].astype(str)
schedule_power['Part Number'] = schedule_power['Part Number'].astype(str)
schedule_mech['Code'] = schedule_mech['Code'].astype(str)
schedule_mech['Part Number'] = schedule_mech['Part Number'].astype(str)

schedule_power = pd.merge(schedule_power, dispatch_summary, left_on=['Code', 'Part Number'], right_on=['Sold-to Party', 'Material'], how='left')
schedule_mech = pd.merge(schedule_mech, dispatch_summary, left_on=['Code', 'Part Number'], right_on=['Sold-to Party', 'Material'], how='left')

schedule_power['Dispatch Qty'] = schedule_power['Dispatch Qty'].fillna(0)
schedule_mech['Dispatch Qty'] = schedule_mech['Dispatch Qty'].fillna(0)

def get_power_kit(row):
    if row['Description'] in ['STG GEAR KIT', 'STG GEAR KIT H-Pas']:
        return lookup_power_stg.get(row['Part Number'], '')
    if 'VANE PUMP KIT' in str(row['Description']):
        return lookup_power_vp.get(row['Part Number'], '')
    return ''

schedule_power.insert(schedule_power.columns.get_loc('Part Number') + 1, 'Kit Part Number', schedule_power.apply(get_power_kit, axis=1))
schedule_mech.insert(schedule_mech.columns.get_loc('Part Number') + 1, 'Kit Part Number', schedule_mech.apply(lambda row: lookup_mech.get(row['Part Number'], '') if row['Part Number'].startswith(('7820975', '734097')) else '', axis=1))

fg_df['Material'] = fg_df['Material'].astype(str)
fg_df['Plant'] = fg_df['Plant'].astype(str)

def fg_sum(df, part_col, plant_col):
    return df.apply(lambda row: fg_df.loc[(fg_df['Material'] == str(row[part_col])) & (fg_df['Plant'] == str(row[plant_col])), 'Unrestricted'].sum(), axis=1)

schedule_power['FG'] = fg_sum(schedule_power, 'Part Number', 'BILLING PLANT') + fg_sum(schedule_power, 'Kit Part Number', 'BILLING PLANT')
schedule_mech['FG'] = fg_sum(schedule_mech, 'Part Number', 'Billing Plant') + fg_sum(schedule_mech, 'Kit Part Number', 'Billing Plant')

marketing_columns_power = [col for col in schedule_power.columns if str(col).startswith('Marketing Requirement')]
marketing_columns_mech = [col for col in schedule_mech.columns if str(col).startswith('Marketing Requirement')]

schedule_power['Balance Dispatch'] = (schedule_power[marketing_columns_power].sum(axis=1) - schedule_power['Dispatch Qty']).clip(lower=0)
schedule_power['Excess Dispatch'] = schedule_power.apply(lambda x: x['Dispatch Qty'] - x[marketing_columns_power].sum() if x['Dispatch Qty'] > x[marketing_columns_power].sum() else 0, axis=1)

schedule_mech['Balance Dispatch'] = (schedule_mech[marketing_columns_mech].sum(axis=1) - schedule_mech['Dispatch Qty']).clip(lower=0)
schedule_mech['Excess Dispatch'] = schedule_mech.apply(lambda x: x['Dispatch Qty'] - x[marketing_columns_mech].sum() if x['Dispatch Qty'] > x[marketing_columns_mech].sum() else 0, axis=1)

schedule_power = schedule_power[['Code', 'Customer', 'MODEL', 'BILLING PLANT', 'Part Number', 'Kit Part Number',
                                 'Customer Part', 'Description', 'Initial Schedule', 'REV-1', 'REV-2'] +
                                 marketing_columns_power +
                                 ['Dispatch Qty', 'Balance Dispatch', 'FG', 'Excess Dispatch', 'ZFI SCOPE']]

schedule_mech = schedule_mech[['Code', 'Customer', 'Model', 'Billing Plant', 'Part Number', 'Kit Part Number',
                               'Customer Part', 'Description', 'Initial Schedule', 'REV-1', 'REV-2'] +
                               marketing_columns_mech +
                               ['Dispatch Qty', 'Balance Dispatch', 'FG', 'Excess Dispatch']]

view_option = st.sidebar.radio("Select View", ["All", "Power Schedule", "Mech Schedule"])

def apply_filters(df, code, customer, billing_plant, model, part_number_search, sheet_type):
    if code:
        df = df[df['Code'].isin(code)]
    if customer:
        df = df[df['Customer'].isin(customer)]
    if billing_plant:
        plant_col = 'BILLING PLANT' if sheet_type == 'Power' else 'Billing Plant'
        df = df[df[plant_col].isin(billing_plant)]
    if model:
        model_col = 'MODEL' if sheet_type == 'Power' else 'Model'
        df = df[df[model_col].isin(model)]
    if part_number_search:
        df = df[df['Part Number'].str.contains(part_number_search, case=False, na=False)]
    return df

def display_subtotals(df):
    if st.get_option("theme.base") == "dark":
        bg_color = "#262730"
        text_color = "#FFFFFF"
    else:
        bg_color = "#FFFFFF"
        text_color = "#000000"

    marketing_cols = [col for col in df.columns if str(col).startswith('Marketing Requirement')]
    subtotal_data = {col: df[col].sum() for col in marketing_cols}
    subtotal_data['Dispatch Qty'] = df['Dispatch Qty'].sum()
    subtotal_data['Balance Dispatch'] = df['Balance Dispatch'].sum()
    subtotal_data['Excess Dispatch'] = df['Excess Dispatch'].sum()

    st.markdown(
        f"""
        <div style="background-color:{bg_color}; padding:10px; border-radius:5px;">
            {"".join([f"<p style='color:{text_color};'>{k}: {v:.0f}</p>" for k, v in subtotal_data.items()])}
        </div>
        """,
        unsafe_allow_html=True
    )

if view_option == "Power Schedule":
    code = st.sidebar.multiselect('Code', schedule_power['Code'].unique())
    customer = st.sidebar.multiselect('Customer', schedule_power['Customer'].unique())
    billing_plant = st.sidebar.multiselect('Billing Plant', schedule_power['BILLING PLANT'].unique())
    model = st.sidebar.multiselect('Model', schedule_power['MODEL'].unique())
    part_number_search = st.sidebar.text_input('Part Number (Type & Press Enter)')
    filtered_power = apply_filters(schedule_power, code, customer, billing_plant, model, part_number_search, 'Power')

    display_subtotals(filtered_power)
    st.dataframe(filtered_power, use_container_width=True)
    power_to_download = filtered_power
    mech_to_download = pd.DataFrame()

elif view_option == "Mech Schedule":
    code = st.sidebar.multiselect('Code', schedule_mech['Code'].unique())
    customer = st.sidebar.multiselect('Customer', schedule_mech['Customer'].unique())
    billing_plant = st.sidebar.multiselect('Billing Plant', schedule_mech['Billing Plant'].unique())
    model = st.sidebar.multiselect('Model', schedule_mech['Model'].unique())
    part_number_search = st.sidebar.text_input('Part Number (Type & Press Enter)')
    filtered_mech = apply_filters(schedule_mech, code, customer, billing_plant, model, part_number_search, 'Mech')

    display_subtotals(filtered_mech)
    st.dataframe(filtered_mech, use_container_width=True)
    power_to_download = pd.DataFrame()
    mech_to_download = filtered_mech

else:
    power_to_download = schedule_power.copy()
    mech_to_download = schedule_mech.copy()
    st.write("### Power Schedule")
    st.dataframe(schedule_power, use_container_width=True)
    st.write("### Mech Schedule")
    st.dataframe(schedule_mech, use_container_width=True)

if not power_to_download.empty or not mech_to_download.empty:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not power_to_download.empty:
            power_to_download.to_excel(writer, sheet_name='Power', index=False)
        if not mech_to_download.empty:
            mech_to_download.to_excel(writer, sheet_name='Mech', index=False)
        workbook = writer.book
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for sheet_name in writer.sheets:
            worksheet = workbook[sheet_name]
            for col in worksheet.columns:
                max_len = max(len(str(cell.value)) for cell in col if cell.value) + 2
                worksheet.column_dimensions[get_column_letter(col[0].column)].width = max_len
            for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                for cell in row:
                    cell.border = thin_border
    output.seek(0)
    st.download_button("Download Excel", output, "Schedule_with_Dispatch.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
