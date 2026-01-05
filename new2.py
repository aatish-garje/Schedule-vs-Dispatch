import streamlit as st
import pandas as pd
import numpy as np
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

    # -------------------- Updated Customer Name --------------------
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

    # Keep formatted dates (matches your original behavior in later pages)
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

    # --------------------
    # Helper: multiselect filter
    # --------------------
    def apply_multifilter(df, col, selected_values):
        if not selected_values:
            return df
        return df[df[col].isin(selected_values)]

    # ==========================
    # OVERVIEW
    # ==========================
    if page == "Overview":
        st.header("Overview Page")

        monthlist = sorted(dispatchdata["Month-Year"].dropna().unique().tolist())
        monthlist = ["All"] + monthlist
        selectedmonth = st.sidebar.selectbox("Select Month-Year (Overview)", monthlist)

        overviewdata = dispatchdata.copy()
        if selectedmonth != "All":
            overviewdata = overviewdata[overviewdata["Month-Year"] == selectedmonth]

        overviewdata = overviewdata.sort_values("Month Start Date")

        monthlysales = (overviewdata.groupby(["Month-Year", "Month Start Date"])["Basic Amt.LocCur"]
                        .sum().reset_index().sort_values("Month Start Date"))

        ymax1 = monthlysales["Basic Amt.LocCur"].max() * 1.15 if len(monthlysales) else 0

        figtotalsales = px.bar(
            monthlysales,
            x="Month-Year",
            y="Basic Amt.LocCur",
            title="Total Monthly Sales (Basic Amt.LocCur)",
            labels={"Basic Amt.LocCur": "Basic Amount", "Month-Year": "Month-Year"},
            text="Basic Amt.LocCur"
        )
        figtotalsales.update_layout(
            yaxis_tickprefix="",
            xaxis_title="Month-Year",
            uniformtext_minsize=8,
            uniformtext_mode="hide",
            bargap=0.3,
            yaxis=dict(range=[0, ymax1])
        )
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
            splitsales,
            x="Month-Year",
            y="Basic Amt.LocCur",
            color="Customer Category",
            barmode="group",
            title="OEM / SPD Sales (Basic Amt.LocCur) Month-wise",
            labels={"Basic Amt.LocCur": "Basic Amount", "Month-Year": "Month-Year"},
            text="Basic Amt.LocCur"
        )
        figoemspd.update_layout(
            yaxis_tickprefix="",
            xaxis_title="Month-Year",
            uniformtext_minsize=8,
            uniformtext_mode="hide",
            bargap=0.3,
            yaxis=dict(range=[0, ymax2])
        )
        figoemspd.update_traces(
            texttemplate="%{text:,.0f}",
            textposition="outside",
            marker_line_width=0.5
        )

        plantsales = overviewdata.groupby("Plant")["Basic Amt.LocCur"].sum().reset_index()
        plantsales["Plant"] = plantsales["Plant"].astype(str)
        ymax3 = plantsales["Basic Amt.LocCur"].max() * 1.15 if len(plantsales) else 0

        figplantsales = px.bar(
            plantsales,
            x="Plant",
            y="Basic Amt.LocCur",
            title="Plant-wise Sales (Basic Amt.LocCur)",
            labels={"Basic Amt.LocCur": "Basic Amount", "Plant": "Plant"},
            text="Basic Amt.LocCur"
        )
        figplantsales.update_layout(
            xaxis=dict(type="category"),
            yaxis_tickprefix="",
            xaxis_title="Plant",
            uniformtext_minsize=8,
            uniformtext_mode="hide",
            bargap=0.3,
            yaxis=dict(range=[0, ymax3])
        )
        figplantsales.update_traces(
            texttemplate="%{text:,.0f}",
            textposition="outside",
            marker_line_width=0.5
        )

        st.plotly_chart(figtotalsales, use_container_width=True)
        st.plotly_chart(figoemspd, use_container_width=True)
        st.plotly_chart(figplantsales, use_container_width=True)

        st.header("Material Category vs Customer Category (OEM / SPD) Qty Wise")
        overviewdata2 = dispatchdata.copy()
        overviewdata2["Inv Qty"] = pd.to_numeric(overviewdata2["Inv Qty"], errors="coerce").fillna(0)
        overviewdata2["Kit Qty"] = pd.to_numeric(overviewdata2["Kit Qty"], errors="coerce").fillna(0)

        if selectedmonth != "All":
            overviewdata2 = overviewdata2[overviewdata2["Month-Year"] == selectedmonth]

        overviewdata2 = overviewdata2[overviewdata2["Customer Category"].isin(["OEM", "SPD"])]
        overviewdata2 = overviewdata2[overviewdata2["Material Category"].isin(["Power STG", "Mechanical Stg", "Power STG H-Pas"])]

        overviewdata2["Effective Qty"] = overviewdata2.apply(
            lambda row: row["Inv Qty"] if (row["Customer Category"] == "OEM" or row["Inv Qty"] != 0) else row["Kit Qty"],
            axis=1
        )

        grouped = overviewdata2.groupby(["Material Category", "Customer Category"])["Effective Qty"].sum().reset_index()
        ymax = grouped["Effective Qty"].max() * 1.2 if len(grouped) else 0

        fig = px.bar(
            grouped,
            x="Material Category",
            y="Effective Qty",
            color="Customer Category",
            barmode="group",
            text="Effective Qty",
            title="Qty by Material Category & Customer Category"
        )
        fig.update_layout(
            xaxis_title="Material Category",
            yaxis_title="Quantity",
            uniformtext_minsize=8,
            uniformtext_mode="hide",
            bargap=0.3,
            yaxis=dict(range=[0, ymax])
        )
        fig.update_traces(
            texttemplate="%{text:,.0f}",
            textposition="outside",
            cliponaxis=False
        )
        st.plotly_chart(fig, use_container_width=True)

    # ==========================
    # SPD
    # ==========================
    elif page == "SPD":
        st.header("SPD Page")
        spddata = dispatchdata[dispatchdata["Customer Category"] == "SPD"]
        st.dataframe(spddata)

    # ==========================
    # OEM (ALL CHARTS RESTORED)
    # ==========================
    elif page == "OEM":
        st.header("OEM Dashboard")

        oemdf = dispatchdata[dispatchdata["Customer Category"] == "OEM"].copy()
        oemdf["Material Category"] = oemdf["Material Category"].replace("Power STG H-Pas", "Power STG")

        # --------- Multiselect Filters (instead of selectbox All) ----------
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

        # Ensure numeric
        filtereddf["Inv Qty"] = pd.to_numeric(filtereddf["Inv Qty"], errors="coerce").fillna(0)
        filtereddf["Basic Amt.LocCur"] = pd.to_numeric(filtereddf["Basic Amt.LocCur"], errors="coerce").fillna(0)

        # --- Chart 1: Power STG customer-wise qty
        st.subheader("OEM - Power STG - Customer-wise Quantity")
        oempowerstg = filtereddf[filtereddf["Material Category"] == "Power STG"]
        oempowercustqty = oempowerstg.groupby("Updated Customer Name")["Inv Qty"].sum().sort_values(ascending=False)

        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(y=oempowercustqty.index, x=oempowercustqty.values, palette="Blues_r", ax=ax)
        for i, (name, value) in enumerate(zip(oempowercustqty.index, oempowercustqty.values)):
            ax.text(value, i, f"{value:,.0f}", va="center")
        st.pyplot(fig)

        # --- Chart 2: Mechanical Stg customer-wise qty
        st.subheader("OEM - Mechanical Stg - Customer-wise Quantity")
        oemmechstg = filtereddf[filtereddf["Material Category"] == "Mechanical Stg"]
        oemmechcustqty = oemmechstg.groupby("Updated Customer Name")["Inv Qty"].sum().sort_values(ascending=False)

        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(y=oemmechcustqty.index, x=oemmechcustqty.values, palette="Greens_r", ax=ax)
        for i, (name, value) in enumerate(zip(oemmechcustqty.index, oemmechcustqty.values)):
            ax.text(value, i, f"{value:,.0f}", va="center")
        st.pyplot(fig)

        # --- Chart 3: Customer-wise total value
        st.subheader("OEM - Customer-wise Total Value")
        oemcustvalue = filtereddf.groupby("Updated Customer Name")["Basic Amt.LocCur"].sum().sort_values(ascending=False)

        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(y=oemcustvalue.index, x=oemcustvalue.values, palette="Oranges_r", ax=ax)
        for i, (name, value) in enumerate(zip(oemcustvalue.index, oemcustvalue.values)):
            ax.text(value, i, f"{value:,.0f}", va="center")
        st.pyplot(fig)

        # --- Chart 4: Model-wise qty - Power STG
        st.subheader("OEM - Model-wise Quantity - Power STG")
        oempowermodelqty = oempowerstg.groupby("Model New")["Inv Qty"].sum().sort_values(ascending=False)

        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(y=oempowermodelqty.index, x=oempowermodelqty.values, palette="Blues", ax=ax)
        for i, (name, value) in enumerate(zip(oempowermodelqty.index, oempowermodelqty.values)):
            ax.text(value, i, f"{value:,.0f}", va="center")
        st.pyplot(fig)

        # --- Chart 5: Model-wise qty - Vane Pump
        st.subheader("OEM - Model-wise Quantity - Vane Pump")
        oemvanepump = filtereddf[filtereddf["Material Category"] == "Vane Pump"]
        oemvanemodelqty = oemvanepump.groupby("Model New")["Inv Qty"].sum().sort_values(ascending=False)

        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(y=oemvanemodelqty.index, x=oemvanemodelqty.values, palette="Purples", ax=ax)
        for i, (name, value) in enumerate(zip(oemvanemodelqty.index, oemvanemodelqty.values)):
            ax.text(value, i, f"{value:,.0f}", va="center")
        st.pyplot(fig)

        # --- Chart 6: Model-wise qty - Mechanical Stg
        st.subheader("OEM - Model-wise Quantity - Mechanical Stg")
        oemmechmodelqty = oemmechstg.groupby("Model New")["Inv Qty"].sum().sort_values(ascending=False)

        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(y=oemmechmodelqty.index, x=oemmechmodelqty.values, palette="Greens", ax=ax)
        for i, (name, value) in enumerate(zip(oemmechmodelqty.index, oemmechmodelqty.values)):
            ax.text(value, i, f"{value:,.0f}", va="center")
        st.pyplot(fig)

        # --- Chart 7: Top 20 Material-Customer by Basic Amt with Qty
        st.subheader("OEM - Top 20 Material-Customer combinations by Basic Amount with Quantity")
        topmatcust = (filtereddf.groupby(["Material", "Updated Customer Name"])[["Basic Amt.LocCur", "Inv Qty"]]
                      .sum()
                      .sort_values(by="Basic Amt.LocCur", ascending=False)
                      .head(20)
                      .reset_index())

        topmatcust["Label"] = (
            topmatcust["Material"].astype(str)
            + " - "
            + topmatcust["Updated Customer Name"].astype(str)
            + " Qty "
            + topmatcust["Inv Qty"].fillna(0).astype(int).astype(str)
        )

        fig, ax = plt.subplots(figsize=(12, 8))
        sns.barplot(y=topmatcust["Label"], x=topmatcust["Basic Amt.LocCur"], palette="rocket", ax=ax)
        ax.set_xlabel("Basic Amount")
        ax.set_ylabel("Material - Customer")
        ax.set_title("Top 20 Material-Customer combinations by Basic Amt.LocCur")
        for i, (val, qty) in enumerate(zip(topmatcust["Basic Amt.LocCur"], topmatcust["Inv Qty"])):
            ax.text(val, i, f"{int(val):,}", va="center")
        st.pyplot(fig)

        # --- Trend charts need Month Start Date ordering
        filtereddf = filtereddf.sort_values("Month Start Date")

        # --- Chart 8: OEM Month-wise Revenue Trend (Cr)
        st.subheader("OEM Month-wise Revenue Trend (Cr)")
        revenuemonthly = (filtereddf.groupby(["Month-Year", "Month Start Date", "Updated Customer Name"])["Basic Amt.LocCur"]
                          .sum()
                          .reset_index()
                          .sort_values("Month Start Date"))
        revenuemonthly["Revenue Cr"] = revenuemonthly["Basic Amt.LocCur"] / 1e7

        monthorder = revenuemonthly.sort_values("Month Start Date")["Month-Year"].unique().tolist()

        if not selectedupdatedcustomers:  # means "All" customers effectively
            figrevenue = px.line(
                revenuemonthly,
                x="Month-Year",
                y="Revenue Cr",
                color="Updated Customer Name",
                markers=True,
                text="Revenue Cr",
                category_orders={"Month-Year": monthorder},
                title="Month-wise Revenue Comparison (All OEM Customers) (Cr)"
            )
        else:
            # if multiple selected, plot them; if one selected also works
            singlecustdf = revenuemonthly[revenuemonthly["Updated Customer Name"].isin(selectedupdatedcustomers)]
            figrevenue = px.line(
                singlecustdf,
                x="Month-Year",
                y="Revenue Cr",
                color="Updated Customer Name" if len(selectedupdatedcustomers) > 1 else None,
                markers=True,
                text="Revenue Cr",
                category_orders={"Month-Year": monthorder},
                title="Month-wise Revenue Trend (Selected OEM Customer(s)) (Cr)"
            )

        figrevenue.update_traces(textposition="top center", texttemplate="%{text:.2f} Cr", cliponaxis=False)
        figrevenue.update_layout(
            xaxis_title="Month",
            yaxis_title="Revenue (Cr)",
            legend_title="Customer",
            hovermode="x unified",
            margin=dict(t=80),
            uniformtext_minsize=9,
            uniformtext_mode="hide"
        )
        if len(revenuemonthly):
            figrevenue.update_yaxes(range=[0, revenuemonthly["Revenue Cr"].max() * 1.25])
        st.plotly_chart(figrevenue, use_container_width=True)

        # --- Chart 9: OEM Power STG Quantity Trend
        st.subheader("OEM Power STG Quantity Trend")
        powerqtymonthly = (filtereddf[filtereddf["Material Category"] == "Power STG"]
                           .groupby(["Month-Year", "Month Start Date", "Updated Customer Name"])["Inv Qty"]
                           .sum()
                           .reset_index()
                           .sort_values("Month Start Date"))

        monthorder2 = powerqtymonthly.sort_values("Month Start Date")["Month-Year"].unique().tolist()

        if not selectedupdatedcustomers:
            figpower = px.line(
                powerqtymonthly,
                x="Month-Year",
                y="Inv Qty",
                color="Updated Customer Name",
                markers=True,
                text="Inv Qty",
                category_orders={"Month-Year": monthorder2},
                title="Month-wise Power STG Quantity (All OEM Customers)"
            )
        else:
            singlecustpower = powerqtymonthly[powerqtymonthly["Updated Customer Name"].isin(selectedupdatedcustomers)]
            figpower = px.line(
                singlecustpower,
                x="Month-Year",
                y="Inv Qty",
                color="Updated Customer Name" if len(selectedupdatedcustomers) > 1 else None,
                markers=True,
                text="Inv Qty",
                category_orders={"Month-Year": monthorder2},
                title="Month-wise Power STG Quantity (Selected OEM Customer(s))"
            )

        figpower.update_traces(textposition="top center", texttemplate="%{text:,.0f}", cliponaxis=False)
        figpower.update_layout(
            xaxis_title="Month",
            yaxis_title="Quantity",
            legend_title="Customer",
            hovermode="x unified",
            margin=dict(t=80),
            uniformtext_minsize=9,
            uniformtext_mode="hide"
        )
        if len(powerqtymonthly):
            figpower.update_yaxes(range=[0, powerqtymonthly["Inv Qty"].max() * 1.25])
        st.plotly_chart(figpower, use_container_width=True)

    # ==========================
    # The remaining pages (Invoice Value / Dispatch Details / Daywise Dispatch)
    # are unchanged here because your request was specifically OEM charts.
    # ==========================
    else:
        st.info("Please select Overview/SPD/OEM in this version.")

else:
    st.info("Upload a file to start.")
