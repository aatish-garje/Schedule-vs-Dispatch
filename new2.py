import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px
import seaborn as sns
import matplotlib.pyplot as plt

def tocr(value):
    return value / 1e7

st.set_page_config(layout="wide")
st.title("Dispatch Data Dashboard")

page = st.sidebar.radio(
    "Select Page",
    ["Overview", "SPD", "OEM", "Daywise Dispatch", "Invoice Value", "Dispatch Details"]
)

uploaded_file = st.file_uploader("Upload your Dispatch Data Excel file", type=["xlsx", "csv"])

if uploaded_file is not None:
    # -------------------- Load Data --------------------
    if uploaded_file.name.lower().endswith(".xlsx"):
        dispatchdata = pd.read_excel(uploaded_file)
    else:
        dispatchdata = pd.read_csv(uploaded_file, encoding="latin1")

    dispatchdata.columns = dispatchdata.columns.str.strip()

    # -------------------- Customer Mapping --------------------
    dispatchdata.insert(
        dispatchdata.columns.get_loc("Customer Name") + 1,
        "Updated Customer Name",
        np.select(
            [
                dispatchdata["Customer Name"].str.lower().str.startswith("ashok", na=False),
                dispatchdata["Customer Name"].str.lower().str.startswith("tata advanced", na=False),
                dispatchdata["Customer Name"].str.lower().str.startswith("tata", na=False),
                dispatchdata["Customer Name"].str.lower().str.startswith("blue energy", na=False),
                dispatchdata["Customer Name"].str.lower().str.startswith("force motors", na=False),
                dispatchdata["Customer Name"].str.lower().str.startswith("cnh", na=False),
                dispatchdata["Customer Name"].str.lower().str.startswith("bajaj auto", na=False),
                dispatchdata["Sold-to Party"].astype(str).str.upper().isin(["M0163", "M0164", "M0231"]),
                dispatchdata["Sold-to Party"].astype(str).str.upper().isin(["M0009", "M0010", "M0221"]),
            ],
            [
                "Ashok Leyland",
                "Tata Motors",
                "Tata Motors",
                "Blue Energy",
                "Force Motors",
                "CNH",
                "Bajaj Auto",
                "Mahindra Swaraj",
                "MM",
            ],
            default=dispatchdata["Customer Name"]
        )
    )

    # -------------------- Material Category --------------------
    def categorize_material(material):
        materialstr = str(material)
        if materialstr == "7632975501":
            return "Oil Tank"
        elif materialstr.startswith(("80339", "80349", "80379", "80439", "80469", "80489", "80499", "M0339", "M0439", "88439")):
            return "Power STG"
        elif materialstr.startswith(("76139", "76729", "76739", "76749", "76919")):
            return "Vane Pump"
        elif materialstr.startswith(("78209", "73409", "73408")):
            return "Mechanical Stg"
        elif materialstr.startswith(("78609",)):
            return "Bevel Gear"
        elif len(materialstr) == 7 and materialstr[4:7] == "012":
            return "Drop Arm"
        elif len(materialstr) == 7 and materialstr[4:7] == "472":
            return "Oil Tank"
        else:
            return "Child Parts"

    dispatchdata.insert(
        dispatchdata.columns.get_loc("Material") + 1,
        "Material Category",
        dispatchdata["Material"].apply(categorize_material)
    )

    # -------------------- Dates --------------------
    dispatchdata["Billing Date"] = pd.to_datetime(dispatchdata["Billing Date"], dayfirst=True, errors="coerce")
    dispatchdata["Cust PO Date"] = pd.to_datetime(dispatchdata["Cust PO Date"], dayfirst=True, errors="coerce")

    dispatchdata.insert(
        dispatchdata.columns.get_loc("Billing Date") + 1,
        "Month-Year",
        dispatchdata["Billing Date"].dt.strftime("%B-%y")
    )
    dispatchdata["Month Start Date"] = pd.to_datetime("01 " + dispatchdata["Month-Year"], format="%d %B-%y", errors="coerce")
    dispatchdata["Billing Date"] = dispatchdata["Billing Date"].dt.strftime("%d-%m-%Y")
    dispatchdata["Cust PO Date"] = dispatchdata["Cust PO Date"].dt.strftime("%d-%m-%Y")

    # -------------------- Model New --------------------
    dispatchdata.insert(
        dispatchdata.columns.get_loc("Material"),
        "Model New",
        dispatchdata["Material"].astype(str).str[:5]
    )

    specificcustomers = ["C0003", "F0006", "G1044", "I0047", "M0163", "M0231", "T0138"]
    dispatchdata.loc[
        dispatchdata["Sold-to Party"].astype(str).str.upper().isin(specificcustomers)
        & (dispatchdata["Model New"].str.lower() == "m0339"),
        "Model New"
    ] = "M0339 H-Pas"

    dispatchdata.loc[
        dispatchdata["Model New"] == "M0339 H-Pas",
        "Material Category"
    ] = "Power STG H-Pas"

    dispatchdata.loc[
        dispatchdata["Material"].astype(str).str.endswith("RF", na=False),
        ["Model New", "Material Category"]
    ] = ["M0339 H-Pas", "Power STG H-Pas"]

    # -------------------- Customer Category --------------------
    dispatchdata["Customer Group"] = dispatchdata["Customer Group"].astype(str).str.strip().str.replace(".0", "", regex=False)

    dispatchdata.insert(
        dispatchdata.columns.get_loc("Customer Group") + 1,
        "Customer Category",
        dispatchdata["Customer Group"].apply(
            lambda x: "OEM" if x == "10" else ("SPD" if (x.isdigit() and 11 <= int(x) <= 15) else "Internal")
        )
    )

    # -------------------- Financial Year --------------------
    def assign_financial_year(date_str):
        try:
            date = pd.to_datetime(date_str, dayfirst=True)
            if date.month >= 4:
                fystart = date.year
                fyend = date.year + 1
            else:
                fystart = date.year - 1
                fyend = date.year
            return f"FY {fystart}-{str(fyend)[-2:]}"
        except Exception:
            return None

    dispatchdata["Financial Year"] = pd.to_datetime(dispatchdata["Billing Date"], dayfirst=True, errors="coerce").apply(assign_financial_year)

    # ==========================
    # Helper: multi-filter apply
    # ==========================
    def apply_multifilter(df, col, selected_values, cast_str=False):
        """
        If selected_values is empty => no filter.
        Else filter df where df[col] is in selected_values.
        """
        if not selected_values:
            return df
        if cast_str:
            return df[df[col].astype(str).isin([str(x) for x in selected_values])]
        return df[df[col].isin(selected_values)]

    # ==========================
    # Overview Page (unchanged)
    # ==========================
    if page == "Overview":
        st.header("Overview Page")

        monthlist = sorted(dispatchdata["Month-Year"].dropna().unique().tolist())
        selectedmonth = st.sidebar.selectbox("Select Month-Year (Overview)", ["All"] + monthlist)

        overviewdata = dispatchdata.copy()
        if selectedmonth != "All":
            overviewdata = overviewdata[overviewdata["Month-Year"] == selectedmonth]

        overviewdata = overviewdata.sort_values("Month Start Date")

        monthlysales = (overviewdata.groupby(["Month-Year", "Month Start Date"])["Basic Amt.LocCur"]
                        .sum().reset_index().sort_values("Month Start Date"))
        ymax1 = monthlysales["Basic Amt.LocCur"].max() * 1.15 if len(monthlysales) else 0

        figtotalsales = px.bar(
            monthlysales, x="Month-Year", y="Basic Amt.LocCur",
            title="Total Monthly Sales (Basic Amt.LocCur)",
            labels={"Basic Amt.LocCur": "Basic Amount", "Month-Year": "Month-Year"},
            text="Basic Amt.LocCur"
        )
        figtotalsales.update_layout(yaxis_tickprefix="", xaxis_title="Month-Year",
                                    uniformtext_minsize=8, uniformtext_mode="hide",
                                    bargap=0.3, yaxis=dict(range=[0, ymax1]))
        figtotalsales.update_traces(
            texttemplate="%{text:,.0f}",
            textposition="outside",
            marker_line_width=0.5
        )

        splitsales = (overviewdata[overviewdata["Customer Category"].isin(["OEM", "SPD"])]
                      .groupby(["Month-Year", "Month Start Date", "Customer Category"])["Basic Amt.LocCur"]
                      .sum().reset_index().sort_values("Month Start Date"))
        ymax2 = splitsales["Basic Amt.LocCur"].max() * 1.15 if len(splitsales) else 0

        figoemspd = px.bar(
            splitsales, x="Month-Year", y="Basic Amt.LocCur", color="Customer Category",
            barmode="group", title="OEM / SPD Sales (Basic Amt.LocCur) Month-wise",
            labels={"Basic Amt.LocCur": "Basic Amount", "Month-Year": "Month-Year"},
            text="Basic Amt.LocCur"
        )
        figoemspd.update_layout(yaxis_tickprefix="", xaxis_title="Month-Year",
                                uniformtext_minsize=8, uniformtext_mode="hide",
                                bargap=0.3, yaxis=dict(range=[0, ymax2]))
        figoemspd.update_traces(texttemplate="%{text:,.0f}", textposition="outside", marker_linewidth=0.5, width=0.2)

        plantsales = overviewdata.groupby("Plant")["Basic Amt.LocCur"].sum().reset_index()
        plantsales["Plant"] = plantsales["Plant"].astype(str)
        ymax3 = plantsales["Basic Amt.LocCur"].max() * 1.15 if len(plantsales) else 0

        figplantsales = px.bar(
            plantsales, x="Plant", y="Basic Amt.LocCur",
            title="Plant-wise Sales (Basic Amt.LocCur)",
            labels={"Basic Amt.LocCur": "Basic Amount", "Plant": "Plant"},
            text="Basic Amt.LocCur"
        )
        figplantsales.update_layout(xaxis=dict(type="category"), yaxis_tickprefix="",
                                    xaxis_title="Plant", uniformtext_minsize=8,
                                    uniformtext_mode="hide", bargap=0.3,
                                    yaxis=dict(range=[0, ymax3]))
        figplantsales.update_traces(texttemplate="%{text:,.0f}", textposition="outside", marker_linewidth=0.5, width=0.3)

        st.plotly_chart(figtotalsales, use_container_width=True)
        st.plotly_chart(figoemspd, use_container_width=True)
        st.plotly_chart(figplantsales, use_container_width=True)

    # ==========================
    # SPD Page (unchanged)
    # ==========================
    elif page == "SPD":
        st.header("SPD Page")
        spddata = dispatchdata[dispatchdata["Customer Category"] == "SPD"]
        st.dataframe(spddata)

    # ==========================
    # OEM Page (Month, Updated Customer, Customer => MULTISELECT)
    # ==========================
    elif page == "OEM":
        st.header("OEM Dashboard")

        oemdf = dispatchdata[dispatchdata["Customer Category"] == "OEM"].copy()
        oemdf["Material Category"] = oemdf["Material Category"].replace("Power STG H-Pas", "Power STG")

        # multiselect filters
        oemmonths = sorted(oemdf["Month-Year"].dropna().unique().tolist())
        selectedmonths = st.sidebar.multiselect("Select Month (OEM)", oemmonths)

        filtereddf = oemdf.copy()
        filtereddf = apply_multifilter(filtereddf, "Month-Year", selectedmonths)

        updatedcustomers = sorted(filtereddf["Updated Customer Name"].dropna().unique().tolist())
        selectedupdatedcustomers = st.sidebar.multiselect("Select Updated Customer Name (OEM)", updatedcustomers)

        filtereddf = apply_multifilter(filtereddf, "Updated Customer Name", selectedupdatedcustomers)

        customernames = sorted(filtereddf["Customer Name"].dropna().unique().tolist())
        selectedcustomernames = st.sidebar.multiselect("Select Customer Name (OEM)", customernames)

        filtereddf = apply_multifilter(filtereddf, "Customer Name", selectedcustomernames)

        st.subheader("OEM - Power STG - Customer-wise Quantity")
        oempowerstg = filtereddf[filtereddf["Material Category"] == "Power STG"]
        oempowercustqty = oempowerstg.groupby("Updated Customer Name")["Inv Qty"].sum().sort_values(ascending=False)

        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(y=oempowercustqty.index, x=oempowercustqty.values, palette="Blues_r", ax=ax)
        for i, (name, value) in enumerate(zip(oempowercustqty.index, oempowercustqty.values)):
            ax.text(value, i, f"{value:,.0f}", va="center")
        st.pyplot(fig)

    # ==========================================================
    # Invoice Value Page (ALL dropdowns => MULTISELECT)
    # ==========================================================
    elif page == "Invoice Value":
        st.header("Invoice Value Page")

        categoryoptions = ["All", "OEM", "SPD", "OEM SPD"]
        selectedcategory = st.sidebar.radio("Select Customer Category", categoryoptions)

        filteredforcustomerlist = dispatchdata.copy()
        if selectedcategory == "OEM SPD":
            filteredforcustomerlist = filteredforcustomerlist[filteredforcustomerlist["Customer Category"].isin(["OEM", "SPD"])]
        elif selectedcategory != "All":
            filteredforcustomerlist = filteredforcustomerlist[filteredforcustomerlist["Customer Category"] == selectedcategory]

        # MULTISELECT filters
        monthlist = sorted(dispatchdata["Month-Year"].dropna().unique().tolist())
        selectedmonths = st.sidebar.multiselect("Select Month-Year", monthlist)

        fylist = sorted(dispatchdata["Financial Year"].dropna().unique().tolist())
        selectedfys = st.sidebar.multiselect("Select Financial Year", fylist)

        updatedcustomerlist = sorted(filteredforcustomerlist["Updated Customer Name"].dropna().unique().tolist())
        selectedupdatedcustomers = st.sidebar.multiselect("Select Updated Customer Name", updatedcustomerlist)

        modellist = sorted(dispatchdata["Model New"].dropna().unique().tolist())
        selectedmodels = st.sidebar.multiselect("Select Model New", modellist)

        filteredfororiginal = filteredforcustomerlist.copy()
        filteredfororiginal = apply_multifilter(filteredfororiginal, "Updated Customer Name", selectedupdatedcustomers)

        customerlist = sorted(filteredfororiginal["Customer Name"].dropna().unique().tolist())
        selectedcustomers = st.sidebar.multiselect("Select Customer Name", customerlist)

        plantlist = sorted(dispatchdata["Plant"].dropna().astype(str).unique().tolist())
        selectedplants = st.sidebar.multiselect("Select Plant", plantlist)

        materialcategorylist = sorted(dispatchdata["Material Category"].dropna().unique().tolist())
        selectedmaterialcategories = st.sidebar.multiselect("Select Material Category", materialcategorylist)

        # Invoice number type-to-search (kept as before)
        st.sidebar.markdown("---")
        st.sidebar.subheader("Invoice No. Filter (Type to Search)")
        invoicenumbers = sorted(dispatchdata["Billing Doc No."].dropna().unique().astype(str).tolist())
        typedinvoice = st.sidebar.textinput("Type Invoice No.")
        suggestedinvoices = [inv for inv in invoicenumbers if typedinvoice in inv] if typedinvoice else []
        selectedinvoice = st.sidebar.selectbox("Select from Suggestions", ["All"] + suggestedinvoices, index=0, key="invoicevalueinvoicefilter")
        clearinvoicefilter = st.sidebar.button("Clear Invoice Filter")

        # Date range (kept as before)
        billingdates = pd.to_datetime(dispatchdata["Billing Date"], dayfirst=True, errors="coerce")
        if selectedmonths:
            # take min/max month boundaries from selected months
            month_start_dates = pd.to_datetime(["01 " + m for m in selectedmonths], format="%d %B-%y", errors="coerce")
            mindate = month_start_dates.min()
            maxdate = (month_start_dates.max() + pd.offsets.MonthEnd(0))
        else:
            mindate = billingdates.min()
            maxdate = billingdates.max()

        st.sidebar.markdown("---")
        st.sidebar.subheader("Select Date Range (Billing Date)")
        daterange = st.sidebar.dateinput("Billing Date Range", (mindate, maxdate), min_value=mindate, max_value=maxdate)
        cleardatefilter = st.sidebar.button("Clear Date Filter")

        # Material type-to-search (kept as before)
        st.sidebar.markdown("---")
        st.sidebar.subheader("Material Filter (Type to Search)")
        materialnumbers = sorted(dispatchdata["Material"].dropna().unique().astype(str).tolist())
        typedmaterial = st.sidebar.textinput("Type Material")
        suggestedmaterials = [p for p in materialnumbers if typedmaterial.lower() in p.lower()] if typedmaterial else []
        selectedmaterial = st.sidebar.selectbox("Select from Suggestions", ["All"] + suggestedmaterials, index=0, key="invoicevaluematerialfilter")
        clearmaterialfilter = st.sidebar.button("Clear Material Filter")

        # Apply filters
        filtereddata = dispatchdata.copy()

        if selectedcategory != "All":
            if selectedcategory == "OEM SPD":
                filtereddata = filtereddata[filtereddata["Customer Category"].isin(["OEM", "SPD"])]
            else:
                filtereddata = filtereddata[filtereddata["Customer Category"] == selectedcategory]

        filtereddata = apply_multifilter(filtereddata, "Month-Year", selectedmonths)
        filtereddata = apply_multifilter(filtereddata, "Financial Year", selectedfys)
        filtereddata = apply_multifilter(filtereddata, "Updated Customer Name", selectedupdatedcustomers)
        filtereddata = apply_multifilter(filtereddata, "Customer Name", selectedcustomers)
        filtereddata = apply_multifilter(filtereddata, "Plant", selectedplants, cast_str=True)
        filtereddata = apply_multifilter(filtereddata, "Material Category", selectedmaterialcategories)
        filtereddata = apply_multifilter(filtereddata, "Model New", selectedmodels)

        if selectedinvoice != "All":
            filtereddata = filtereddata[filtereddata["Billing Doc No."].astype(str) == str(selectedinvoice)]

        if not clearinvoicefilter and typedinvoice:
            filtereddata = filtereddata[filtereddata["Billing Doc No."].astype(str).str.contains(typedinvoice, na=False)]

        filtereddata["Billing Date"] = pd.to_datetime(filtereddata["Billing Date"], dayfirst=True, errors="coerce")

        if not cleardatefilter:
            startdate, enddate = daterange
            filtereddata = filtereddata[
                (filtereddata["Billing Date"] >= pd.to_datetime(startdate)) &
                (filtereddata["Billing Date"] <= pd.to_datetime(enddate))
            ]

        if not clearmaterialfilter:
            if typedmaterial:
                filtereddata = filtereddata[filtereddata["Material"].astype(str).str.lower().str.contains(typedmaterial.lower(), na=False)]
            elif selectedmaterial != "All":
                filtereddata = filtereddata[filtereddata["Material"].astype(str) == str(selectedmaterial)]

        filtereddata["Billing Date"] = filtereddata["Billing Date"].dt.strftime("%d-%m-%Y")

        # Existing computations (kept)
        filtereddata["Qty"] = pd.to_numeric(filtereddata["Inv Qty"], errors="coerce").fillna(0) + pd.to_numeric(filtereddata["Kit Qty"], errors="coerce").fillna(0)
        filtereddata["Basic Amt.LocCur"] = pd.to_numeric(filtereddata["Basic Amt.LocCur"], errors="coerce").fillna(0)
        filtereddata["Tax Amount"] = pd.to_numeric(filtereddata["Tax Amount"], errors="coerce").fillna(0)
        filtereddata["Amt.Locl Currency"] = pd.to_numeric(filtereddata["Amt.Locl Currency"], errors="coerce").fillna(0)

        filtereddata["Basic Value Per Item"] = np.where(filtereddata["Qty"] != 0, filtereddata["Basic Amt.LocCur"] / filtereddata["Qty"], 0)

        def invoicefilter(group):
            if group["Billing Doc No."].nunique() == 1 or group["Sales Order No"].astype(str).str.startswith("10").any():
                return group
            return group[group["Item"] == 10]

        invoicetotals = (filtereddata.groupby("Billing Doc No.")[["Basic Amt.LocCur", "Tax Amount", "Amt.Locl Currency"]]
                         .sum().reset_index())
        filtereddata = filtereddata.merge(invoicetotals, on="Billing Doc No.", suffixes=("", "Total"))

        maskitem10 = filtereddata["Item"] == 10
        filtereddata.loc[maskitem10, "Basic Amt.LocCur"] = filtereddata.loc[maskitem10, "Basic Amt.LocCurTotal"]
        filtereddata.loc[maskitem10, "Tax Amount"] = filtereddata.loc[maskitem10, "Tax AmountTotal"]
        filtereddata.loc[maskitem10, "Amt.Locl Currency"] = filtereddata.loc[maskitem10, "Amt.Locl CurrencyTotal"]

        filtereddata = filtereddata.drop(columns=["Basic Amt.LocCurTotal", "Tax AmountTotal", "Amt.Locl CurrencyTotal"])
        filtereddata = filtereddata.groupby("Billing Doc No.").apply(invoicefilter).reset_index(drop=True)

        filtereddata["Qty"] = pd.to_numeric(filtereddata["Inv Qty"], errors="coerce").fillna(0) + pd.to_numeric(filtereddata["Kit Qty"], errors="coerce").fillna(0)
        filtereddata["Basic Value Per Item"] = np.where(filtereddata["Qty"] != 0, filtereddata["Basic Amt.LocCur"] / filtereddata["Qty"], 0)

        filtereddata = filtereddata.drop(columns=["Inv Qty", "Kit Qty"], errors="ignore")

        cols = filtereddata.columns.tolist()
        if "Material" in cols and "Qty" in cols and "Basic Value Per Item" in cols:
            materialidx = cols.index("Material")
            cols.remove("Qty")
            cols.remove("Basic Value Per Item")
            cols.insert(materialidx + 1, "Qty")
            cols.insert(materialidx + 2, "Basic Value Per Item")
            filtereddata = filtereddata[cols]

        if "Month Start Date" in filtereddata.columns:
            filtereddata = filtereddata.drop(columns=["Month Start Date"])

        st.dataframe(filtereddata)

    # ==========================================================
    # Dispatch Details Page (ALL dropdowns => MULTISELECT)
    # ==========================================================
    elif page == "Dispatch Details":
        st.header("Dispatch Details Page")

        categoryoptions = ["All", "OEM", "SPD", "OEM SPD"]
        selectedcategory = st.sidebar.radio("Select Customer Category", categoryoptions)

        filteredforcustomerlist = dispatchdata.copy()
        if selectedcategory == "OEM SPD":
            filteredforcustomerlist = filteredforcustomerlist[filteredforcustomerlist["Customer Category"].isin(["OEM", "SPD"])]
        elif selectedcategory != "All":
            filteredforcustomerlist = filteredforcustomerlist[filteredforcustomerlist["Customer Category"] == selectedcategory]

        # MULTISELECT filters
        monthlist = sorted(dispatchdata["Month-Year"].dropna().unique().tolist())
        selectedmonths = st.sidebar.multiselect("Select Month-Year", monthlist)

        fylist = sorted(dispatchdata["Financial Year"].dropna().unique().tolist())
        selectedfys = st.sidebar.multiselect("Select Financial Year", fylist)

        updatedcustomerlist = sorted(filteredforcustomerlist["Updated Customer Name"].dropna().unique().tolist())
        selectedupdatedcustomers = st.sidebar.multiselect("Select Updated Customer Name", updatedcustomerlist)

        modellist = sorted(dispatchdata["Model New"].dropna().unique().tolist())
        selectedmodels = st.sidebar.multiselect("Select Model New", modellist)

        filteredfororiginal = filteredforcustomerlist.copy()
        filteredfororiginal = apply_multifilter(filteredfororiginal, "Updated Customer Name", selectedupdatedcustomers)

        customerlist = sorted(filteredfororiginal["Customer Name"].dropna().unique().tolist())
        selectedcustomers = st.sidebar.multiselect("Select Customer Name", customerlist)

        plantlist = sorted(dispatchdata["Plant"].dropna().astype(str).unique().tolist())
        selectedplants = st.sidebar.multiselect("Select Plant", plantlist)

        materialcategorylist = sorted(dispatchdata["Material Category"].dropna().unique().tolist())
        selectedmaterialcategories = st.sidebar.multiselect("Select Material Category", materialcategorylist)

        # Date filter (kept)
        billingdates = pd.to_datetime(dispatchdata["Billing Date"], dayfirst=True, errors="coerce")
        if selectedmonths:
            month_start_dates = pd.to_datetime(["01 " + m for m in selectedmonths], format="%d %B-%y", errors="coerce")
            mindate = month_start_dates.min()
            maxdate = (month_start_dates.max() + pd.offsets.MonthEnd(0))
        else:
            mindate = billingdates.min()
            maxdate = billingdates.max()

        st.sidebar.markdown("---")
        st.sidebar.subheader("Select Date Range (Billing Date)")
        daterange = st.sidebar.dateinput("Billing Date Range", (mindate, maxdate), min_value=mindate, max_value=maxdate)
        cleardatefilter = st.sidebar.button("Clear Date Filter")

        # Material type-to-search (kept)
        st.sidebar.markdown("---")
        st.sidebar.subheader("Material Filter (Type to Search)")
        materialnumbers = sorted(dispatchdata["Material"].dropna().unique().astype(str).tolist())
        typedmaterial = st.sidebar.textinput("Type Material")
        suggestedmaterials = [p for p in materialnumbers if typedmaterial.lower() in p.lower()] if typedmaterial else []
        selectedmaterial = st.sidebar.selectbox("Select from Suggestions", ["All"] + suggestedmaterials, index=0)
        clearmaterialfilter = st.sidebar.button("Clear Material Filter")

        # Apply filters
        filtereddata = dispatchdata.copy()

        if selectedcategory != "All":
            if selectedcategory == "OEM SPD":
                filtereddata = filtereddata[filtereddata["Customer Category"].isin(["OEM", "SPD"])]
            else:
                filtereddata = filtereddata[filtereddata["Customer Category"] == selectedcategory]

        filtereddata = apply_multifilter(filtereddata, "Month-Year", selectedmonths)
        filtereddata = apply_multifilter(filtereddata, "Financial Year", selectedfys)
        filtereddata = apply_multifilter(filtereddata, "Updated Customer Name", selectedupdatedcustomers)
        filtereddata = apply_multifilter(filtereddata, "Customer Name", selectedcustomers)
        filtereddata = apply_multifilter(filtereddata, "Plant", selectedplants, cast_str=True)
        filtereddata = apply_multifilter(filtereddata, "Material Category", selectedmaterialcategories)
        filtereddata = apply_multifilter(filtereddata, "Model New", selectedmodels)

        filtereddata["Billing Date"] = pd.to_datetime(filtereddata["Billing Date"], dayfirst=True, errors="coerce")

        if not cleardatefilter:
            startdate, enddate = daterange
            filtereddata = filtereddata[
                (filtereddata["Billing Date"] >= pd.to_datetime(startdate)) &
                (filtereddata["Billing Date"] <= pd.to_datetime(enddate))
            ]

        filtereddata["Billing Date"] = filtereddata["Billing Date"].dt.strftime("%d-%m-%Y")

        if not clearmaterialfilter:
            if typedmaterial:
                filtereddata = filtereddata[filtereddata["Material"].astype(str).str.lower().str.contains(typedmaterial.lower(), na=False)]
            elif selectedmaterial != "All":
                filtereddata = filtereddata[filtereddata["Material"].astype(str) == str(selectedmaterial)]

        filtereddata["Inv Qty"] = pd.to_numeric(filtereddata["Inv Qty"], errors="coerce").fillna(0)
        filtereddata["Kit Qty"] = pd.to_numeric(filtereddata["Kit Qty"], errors="coerce").fillna(0)

        invqtysum = filtereddata["Inv Qty"].sum()
        kitqtysum = filtereddata["Kit Qty"].sum()
        basicamtsum = pd.to_numeric(filtereddata["Basic Amt.LocCur"], errors="coerce").fillna(0).sum()

        st.markdown(
            """
            <style>
            .subtotal-box { padding: 10px; border-radius: 5px; border: 1px solid; font-weight: bold; }
            .subtotal-box-light { background-color: #f0f0f0; color: #000; border-color: #ccc; }
            .subtotal-box-dark { background-color: #222; color: #fff; border-color: #555; }
            </style>
            """,
            unsafe_allow_html=True
        )
        theme = st.get_option("theme.base")
        boxclass = "subtotal-box-light" if theme == "light" else "subtotal-box-dark"
        st.markdown(
            f"<div class='subtotal-box {boxclass}'>"
            f"Subtotal Filtered Data<br>"
            f"Inv Qty: {invqtysum:,.0f} &nbsp;&nbsp;&nbsp; "
            f"Kit Qty: {kitqtysum:,.0f} &nbsp;&nbsp;&nbsp; "
            f"Basic Amt.LocCur: {basicamtsum:,.2f}"
            f"</div>",
            unsafe_allow_html=True
        )

        if "Month Start Date" in filtereddata.columns:
            filtereddata = filtereddata.drop(columns=["Month Start Date"])

        st.dataframe(filtereddata)

    # ==========================================================
    # Daywise Dispatch Page (ALL dropdowns => MULTISELECT)
    # ==========================================================
    elif page == "Daywise Dispatch":
        st.header("Daywise Dispatch Page")

        dispatchdata["Inv Qty"] = pd.to_numeric(dispatchdata["Inv Qty"], errors="coerce").fillna(0)
        dispatchdata["Kit Qty"] = pd.to_numeric(dispatchdata["Kit Qty"], errors="coerce").fillna(0)

        filtereddaywise = dispatchdata[
            dispatchdata["Material"].astype(str).str.upper().str.startswith("C")
            & (dispatchdata["Material"].astype(str) != "8043975905")
        ].copy()

        def should_keep(row, billingcounts):
            if billingcounts.get(row["Billing Doc No."], 0) == 1:
                return True
            if str(row["Sales Order No"]).startswith("10"):
                return True
            return row["Item"] == 10

        billingcounts = filtereddaywise["Billing Doc No."].value_counts().to_dict()
        filtereddaywise = filtereddaywise[filtereddaywise.apply(lambda r: should_keep(r, billingcounts), axis=1)].copy()

        if "Total Dispatch" not in filtereddaywise.columns:
            kitqtyindex = filtereddaywise.columns.get_loc("Kit Qty")
            filtereddaywise.insert(kitqtyindex + 1, "Total Dispatch", filtereddaywise["Inv Qty"] + filtereddaywise["Kit Qty"])

        categoryoptions = ["All", "OEM", "SPD", "OEM SPD"]
        selectedcategory = st.sidebar.radio("Select Customer Category", categoryoptions)

        filteredforcustomerlist = filtereddaywise.copy()
        if selectedcategory == "OEM SPD":
            filteredforcustomerlist = filteredforcustomerlist[filteredforcustomerlist["Customer Category"].isin(["OEM", "SPD"])]
        elif selectedcategory != "All":
            filteredforcustomerlist = filteredforcustomerlist[filteredforcustomerlist["Customer Category"] == selectedcategory]

        # MULTISELECT filters
        monthlist = sorted(dispatchdata["Month-Year"].dropna().unique().tolist())
        selectedmonths = st.sidebar.multiselect("Select Month-Year", monthlist)

        fylist = sorted(dispatchdata["Financial Year"].dropna().unique().tolist())
        selectedfys = st.sidebar.multiselect("Select Financial Year", fylist)

        updatedcustomerlist = sorted(filteredforcustomerlist["Updated Customer Name"].dropna().unique().tolist())
        selectedupdatedcustomers = st.sidebar.multiselect("Select Updated Customer Name", updatedcustomerlist)

        filteredfororiginal = filteredforcustomerlist.copy()
        filteredfororiginal = apply_multifilter(filteredfororiginal, "Updated Customer Name", selectedupdatedcustomers)

        customerlist = sorted(filteredfororiginal["Customer Name"].dropna().unique().tolist())
        selectedcustomers = st.sidebar.multiselect("Select Customer Name", customerlist)

        plantlist = sorted(dispatchdata["Plant"].dropna().astype(str).unique().tolist())
        selectedplants = st.sidebar.multiselect("Select Plant", plantlist)

        materialcategorylist = sorted(dispatchdata["Material Category"].dropna().unique().tolist())
        selectedmaterialcategories = st.sidebar.multiselect("Select Material Category", materialcategorylist)

        modellist = sorted(dispatchdata["Model New"].dropna().unique().tolist())
        selectedmodels = st.sidebar.multiselect("Select Model New", modellist)

        # Material type-to-search (kept)
        st.sidebar.markdown("---")
        st.sidebar.subheader("Material Filter (Type to Search)")
        materialnumbers = sorted(dispatchdata["Material"].dropna().unique().astype(str).tolist())
        typedmaterial = st.sidebar.textinput("Type Material")
        suggestedmaterials = [p for p in materialnumbers if typedmaterial.lower() in p.lower()] if typedmaterial else []
        selectedmaterial = st.sidebar.selectbox("Select from Suggestions", ["All"] + suggestedmaterials, index=0)
        clearmaterialfilter = st.sidebar.button("Clear Material Filter")

        # Date range (kept)
        billingdates = pd.to_datetime(filtereddaywise["Billing Date"], dayfirst=True, errors="coerce")
        if selectedmonths:
            month_start_dates = pd.to_datetime(["01 " + m for m in selectedmonths], format="%d %B-%y", errors="coerce")
            mindate = month_start_dates.min()
            maxdate = (month_start_dates.max() + pd.offsets.MonthEnd(0))
        else:
            mindate = billingdates.min()
            maxdate = billingdates.max()

        st.sidebar.markdown("---")
        st.sidebar.subheader("Select Date Range (Billing Date)")
        daterange = st.sidebar.dateinput("Billing Date Range", (mindate, maxdate), min_value=mindate, max_value=maxdate)
        cleardatefilter = st.sidebar.button("Clear Date Filter")

        # Apply filters
        finaldaywise = filtereddaywise.copy()

        if selectedcategory != "All":
            if selectedcategory == "OEM SPD":
                finaldaywise = finaldaywise[finaldaywise["Customer Category"].isin(["OEM", "SPD"])]
            else:
                finaldaywise = finaldaywise[finaldaywise["Customer Category"] == selectedcategory]

        finaldaywise = apply_multifilter(finaldaywise, "Month-Year", selectedmonths)
        finaldaywise = apply_multifilter(finaldaywise, "Financial Year", selectedfys)
        finaldaywise = apply_multifilter(finaldaywise, "Updated Customer Name", selectedupdatedcustomers)
        finaldaywise = apply_multifilter(finaldaywise, "Customer Name", selectedcustomers)
        finaldaywise = apply_multifilter(finaldaywise, "Plant", selectedplants, cast_str=True)
        finaldaywise = apply_multifilter(finaldaywise, "Material Category", selectedmaterialcategories)
        finaldaywise = apply_multifilter(finaldaywise, "Model New", selectedmodels)

        finaldaywise["Billing Date"] = pd.to_datetime(finaldaywise["Billing Date"], dayfirst=True, errors="coerce")

        if not cleardatefilter:
            startdate, enddate = daterange
            finaldaywise = finaldaywise[
                (finaldaywise["Billing Date"] >= pd.to_datetime(startdate)) &
                (finaldaywise["Billing Date"] <= pd.to_datetime(enddate))
            ]

        if not clearmaterialfilter:
            if typedmaterial:
                finaldaywise = finaldaywise[finaldaywise["Material"].astype(str).str.lower().str.contains(typedmaterial.lower(), na=False)]
            elif selectedmaterial != "All":
                finaldaywise = finaldaywise[finaldaywise["Material"].astype(str) == str(selectedmaterial)]

        # Pivot output
        pivottable = (finaldaywise.pivot_table(
            index=["Sold-to Party", "Customer Name", "Material", "Plant"],
            columns="Billing Date",
            values="Total Dispatch",
            aggfunc="sum",
            fill_value=0
        ).reset_index())

        pivottable.columns = [
            col.strftime("%d-%m-%Y") if isinstance(col, pd.Timestamp) else col
            for col in pivottable.columns
        ]
        pivottable.columns.name = None

        st.dataframe(pivottable)

