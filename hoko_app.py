import math
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="HOKO Mobility Financial Copilot v3", layout="wide")


# --------------------------------------------------
# HELPERS
# --------------------------------------------------

def try_get_param(params_df: pd.DataFrame, key: str, default=None):
    if params_df is None or params_df.empty:
        return default
    cols = [c.lower().strip() for c in params_df.columns]
    if "parameter" not in cols or "value" not in cols:
        return default
    pcol = params_df.columns[cols.index("parameter")]
    vcol = params_df.columns[cols.index("value")]
    hit = params_df[params_df[pcol].astype(str).str.strip().str.lower() == key.strip().lower()]
    if hit.empty:
        return default
    return hit.iloc[0][vcol]


def fmt_try(x):
    return f"₺{x:,.0f}"


def fmt_num(x):
    return f"{x:,.0f}"


def fmt_pct(x):
    return f"{x*100:,.1f}%"


def make_template_excel():
    main_parameters = pd.DataFrame({
        "Parameter": [
            "Motor Count",
            "Km per Day",
            "Wh per km",
            "Sell Price",
            "Buy Price",
            "Company Share",
            "Cycle per CS24",
            "Battery kWh Each",
            "Battery Slots",
            "Delivery SoC",
            "CS24 Hardware USD",
            "Battery USD per kWh",
            "USDTRY",
            "Grace Period Months",
            "CS24 Down Payment %",
            "Battery Life Years",
        ],
        "Value": [
            6000, 150, 40, 15, 3.5, 0.80, 2.0, 3.24, 24, 0.85, 5000, 200, 44.7795, 5, 0.25, 8
        ]
    })

    personnel = pd.DataFrame({
        "Role": ["CEO", "CFO", "Operations Manager", "Technician", "Customer Support"],
        "Count": [1, 1, 2, 8, 5],
        "Monthly Salary": [150000, 120000, 90000, 45000, 35000]
    })

    bom = pd.DataFrame({
        "Item": ["Frame", "Motor", "Controller", "Battery Pack", "Charger", "Telematics"],
        "Unit Cost": [5000, 12000, 4000, 22000, 1500, 1000]
    })

    cashflow = pd.DataFrame({
        "Month": list(range(1, 49)),
        "Notes": ["" for _ in range(48)]
    })

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        main_parameters.to_excel(writer, sheet_name="Main Parameters", index=False)
        personnel.to_excel(writer, sheet_name="Personnel", index=False)
        bom.to_excel(writer, sheet_name="BoM List", index=False)
        cashflow.to_excel(writer, sheet_name="Cash Flow Statement", index=False)
    bio.seek(0)
    return bio.getvalue()


# --------------------------------------------------
# HEADER
# --------------------------------------------------

st.title("HOKO Mobility Financial Copilot v3")
st.caption("Excel-driven Business Plan model with override logic, 4-year cash flow, and investor-style outputs.")

with st.expander("Excel Upload Instructions", expanded=True):
    st.markdown(
        """
**Please upload your Business Plan Excel file.**

Required sheets:
1. `Main Parameters`
2. `Personnel`
3. `BoM List`
4. `Cash Flow Statement`

Rules:
1. Excel is used as a **structure + source file**
2. **Conflicting parameters are NOT taken from Excel** if overridden in app
3. The app values remain dominant for locked assumptions
4. Cash flow horizon is fixed at **48 months**
5. Grace Period is fixed at **5 months** unless you intentionally change it
        """
    )

template_bytes = make_template_excel()
st.download_button(
    "Download Template Excel",
    data=template_bytes,
    file_name="HOKO_BP_Template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

uploaded_file = st.file_uploader("Upload Business Plan Excel", type=["xlsx"])


# --------------------------------------------------
# READ EXCEL
# --------------------------------------------------

excel = None
params_df = None
personnel_df = None
bom_df = None
cashflow_df = None
sheet_check_ok = False

if uploaded_file is not None:
    try:
        excel = pd.ExcelFile(uploaded_file)
        required_sheets = ["Main Parameters", "Personnel", "BoM List", "Cash Flow Statement"]
        missing = [s for s in required_sheets if s not in excel.sheet_names]

        if missing:
            st.error(f"Missing required sheet(s): {', '.join(missing)}")
        else:
            params_df = pd.read_excel(excel, "Main Parameters")
            personnel_df = pd.read_excel(excel, "Personnel")
            bom_df = pd.read_excel(excel, "BoM List")
            cashflow_df = pd.read_excel(excel, "Cash Flow Statement")
            sheet_check_ok = True
            st.success("Excel uploaded successfully.")
    except Exception as e:
        st.error(f"Excel could not be read: {e}")


# --------------------------------------------------
# LOCKED / OVERRIDE PARAMETERS
# --------------------------------------------------

st.sidebar.header("Locked / Override Parameters")

# Locked model assumptions from current workstream
fleet_base = st.sidebar.number_input("Base Fleet (for reference)", value=6000, step=1000)
daily_km = st.sidebar.number_input("Km per Day", value=150.0, step=5.0)
wh_per_km = st.sidebar.number_input("Wh per km", value=40.0, step=1.0)

battery_slots = st.sidebar.number_input("Battery Slots per CS24", value=24, step=1)
battery_kwh_each = st.sidebar.number_input("Battery kWh Each", value=3.24, step=0.01)
delivery_soc = st.sidebar.number_input("Delivery SoC", value=0.85, step=0.01)

cycle_per_cs24 = st.sidebar.number_input("Charging Cycle per CS24", value=2.0, step=0.1)

buy_price = st.sidebar.number_input("Electricity Buy Price (TL/kWh)", value=3.5, step=0.1)
sell_price = st.sidebar.number_input("Service Sell Price (TL/kWh)", value=15.0, step=0.5)
company_share = st.sidebar.number_input("NewCo Share", value=0.80, step=0.01)

cs24_hardware_usd = st.sidebar.number_input("CS24 Hardware Cost (USD)", value=5000.0, step=100.0)
battery_usd_per_kwh = st.sidebar.number_input("Battery Cost (USD/kWh)", value=200.0, step=5.0)
usdtry = st.sidebar.number_input("USDTRY", value=44.7795, step=0.1)

grace_period_months = st.sidebar.number_input("Grace Period (Months)", value=5, step=1)
battery_life_years = st.sidebar.number_input("Battery Life (Years)", value=8, step=1)
down_payment_pct = st.sidebar.number_input("CS24 Down Payment %", value=0.25, step=0.01)

# Additional model items requested
franchise_share = 1 - company_share
monthly_other_opex = st.sidebar.number_input("Monthly Other OPEX (TL)", value=500000.0, step=50000.0)
monthly_rent = st.sidebar.number_input("Monthly Rent / Admin (TL)", value=250000.0, step=50000.0)
monthly_marketing = st.sidebar.number_input("Monthly Marketing (TL)", value=150000.0, step=50000.0)
monthly_gna = st.sidebar.number_input("Monthly G&A (TL)", value=200000.0, step=50000.0)
monthly_interest = st.sidebar.number_input("Monthly Financing Cost (TL)", value=0.0, step=50000.0)
tax_rate = st.sidebar.number_input("Tax Rate", value=0.20, step=0.01)

motor_gross_profit = st.sidebar.number_input("Motor Gross Profit per Unit (TL)", value=80000.0, step=5000.0)
cs24_other_gross_profit_usd = st.sidebar.number_input("CS24 Other Gross Profit per Year (USD)", value=1200.0, step=100.0)

st.sidebar.markdown("---")
st.sidebar.caption("Excel values are read for structure, but the sidebar overrides are dominant for conflicts.")


# --------------------------------------------------
# SALES SCHEDULE (LOCKED)
# --------------------------------------------------

sales_schedule = {}
for m in range(1, 6):
    sales_schedule[m] = 0

sales_schedule[6] = 2000
sales_schedule[7] = 2500
sales_schedule[8] = 3000
sales_schedule[9] = 3000
sales_schedule[10] = 3250
sales_schedule[11] = 3250
sales_schedule[12] = 3250

for m in range(13, 19):
    sales_schedule[m] = 3500

for m in range(19, 49):
    sales_schedule[m] = 4000


# --------------------------------------------------
# PERSONNEL
# --------------------------------------------------

personnel_monthly_total = 0.0
if personnel_df is not None and not personnel_df.empty:
    lower_cols = [c.lower().strip() for c in personnel_df.columns]
    if "count" in lower_cols and "monthly salary" in lower_cols:
        ccol = personnel_df.columns[lower_cols.index("count")]
        scol = personnel_df.columns[lower_cols.index("monthly salary")]
        personnel_df["Monthly Personnel Cost"] = personnel_df[ccol].fillna(0) * personnel_df[scol].fillna(0)
        personnel_monthly_total = personnel_df["Monthly Personnel Cost"].sum()
else:
    personnel_monthly_total = 0.0


# --------------------------------------------------
# CORE CALCULATION
# --------------------------------------------------

daily_kwh_per_motor = daily_km * wh_per_km / 1000
annual_kwh_per_motor = daily_kwh_per_motor * 365

cs24_total_battery_kwh = battery_slots * battery_kwh_each
sellable_kwh_per_cycle = cs24_total_battery_kwh * delivery_soc
daily_kwh_per_cs24 = sellable_kwh_per_cycle * cycle_per_cs24
monthly_kwh_per_cs24 = daily_kwh_per_cs24 * 30

gross_margin_per_kwh = sell_price - buy_price
company_margin_per_kwh = gross_margin_per_kwh * company_share
franchise_margin_per_kwh = gross_margin_per_kwh * franchise_share

battery_total_usd = cs24_total_battery_kwh * battery_usd_per_kwh
cs24_total_capex_try = (cs24_hardware_usd + battery_total_usd) * usdtry
station_down_payment_try = cs24_total_capex_try * down_payment_pct
company_financed_per_cs24_try = cs24_total_capex_try - station_down_payment_try

active = 0
cum_cash = 0.0
cash_rows = []

installed_cs24 = 0

for m in range(1, 49):
    sold = sales_schedule[m]
    active += sold

    daily_kwh = active * daily_kwh_per_motor
    monthly_kwh = daily_kwh * 30

    required_cs24 = math.ceil(daily_kwh / daily_kwh_per_cs24) if daily_kwh_per_cs24 > 0 else 0
    new_cs24 = max(0, required_cs24 - installed_cs24)

    installed_cs24 = required_cs24

    # Revenue starts after grace period
    if m <= grace_period_months:
        monthly_revenue = 0.0
        energy_cost = 0.0
        energy_gross_profit = 0.0
        motor_gross = 0.0
        cs24_other_gross = 0.0
    else:
        monthly_revenue = monthly_kwh * sell_price
        energy_cost = monthly_kwh * buy_price
        energy_gross_profit = monthly_kwh * company_margin_per_kwh
        motor_gross = sold * motor_gross_profit
        cs24_other_gross = installed_cs24 * ((cs24_other_gross_profit_usd * usdtry) / 12)

    total_gross_profit = energy_gross_profit + motor_gross + cs24_other_gross

    monthly_personnel = personnel_monthly_total
    monthly_opex_total = monthly_personnel + monthly_other_opex + monthly_rent + monthly_marketing + monthly_gna

    ebitda_proxy = total_gross_profit - monthly_opex_total

    capex_out = new_cs24 * company_financed_per_cs24_try
    ebt = ebitda_proxy - monthly_interest
    tax = max(0.0, ebt) * tax_rate
    net_income = ebt - tax

    net_cash_flow = net_income - capex_out
    cum_cash += net_cash_flow

    inventory_battery_asset = installed_cs24 * cs24_total_capex_try
    debt_like_financing = installed_cs24 * company_financed_per_cs24_try
    equity_like_down_payment = installed_cs24 * station_down_payment_try
    retained_earnings = cum_cash
    cash_balance = cum_cash

    cash_rows.append({
        "Month": m,
        "Sold Motors": sold,
        "Active Motors": active,
        "Daily kWh": daily_kwh,
        "Monthly kWh": monthly_kwh,
        "Required CS24": required_cs24,
        "New CS24": new_cs24,
        "Revenue": monthly_revenue,
        "Energy Cost": energy_cost,
        "Energy GP (NewCo)": energy_gross_profit,
        "Motor Gross Profit": motor_gross,
        "CS24 Other GP": cs24_other_gross,
        "Total Gross Profit": total_gross_profit,
        "Personnel Cost": monthly_personnel,
        "Other OPEX": monthly_other_opex,
        "Rent/Admin": monthly_rent,
        "Marketing": monthly_marketing,
        "G&A": monthly_gna,
        "EBITDA Proxy": ebitda_proxy,
        "Interest": monthly_interest,
        "Tax": tax,
        "Net Income": net_income,
        "CapEx Outflow": capex_out,
        "Net Cash Flow": net_cash_flow,
        "Cumulative Cash": cum_cash,
        "Assets - Installed CS24": inventory_battery_asset,
        "Liabilities - Financing": debt_like_financing,
        "Equity Proxy": equity_like_down_payment + retained_earnings,
        "Cash Balance": cash_balance,
    })

cash_df = pd.DataFrame(cash_rows)

# Annual roll-up
cash_df["Year"] = ((cash_df["Month"] - 1) // 12) + 1

annual_df = cash_df.groupby("Year", as_index=False).agg({
    "Sold Motors": "sum",
    "Active Motors": "last",
    "Monthly kWh": "sum",
    "Revenue": "sum",
    "Energy Cost": "sum",
    "Energy GP (NewCo)": "sum",
    "Motor Gross Profit": "sum",
    "CS24 Other GP": "sum",
    "Total Gross Profit": "sum",
    "Personnel Cost": "sum",
    "Other OPEX": "sum",
    "Rent/Admin": "sum",
    "Marketing": "sum",
    "G&A": "sum",
    "EBITDA Proxy": "sum",
    "Interest": "sum",
    "Tax": "sum",
    "Net Income": "sum",
    "CapEx Outflow": "sum",
    "Net Cash Flow": "sum",
    "Cumulative Cash": "last",
    "Required CS24": "last",
    "Assets - Installed CS24": "last",
    "Liabilities - Financing": "last",
    "Equity Proxy": "last",
    "Cash Balance": "last",
})

# IRR approximation over monthly cash flows converted from annual series
irr_cashflows = [-0.0] + annual_df["Net Cash Flow"].tolist()


def npv(rate, cfs):
    return sum(cf / ((1 + rate) ** i) for i, cf in enumerate(cfs))


def irr_bisect(cfs, low=-0.99, high=10.0, iterations=300):
    try:
        f_low = npv(low, cfs)
        f_high = npv(high, cfs)
        if f_low * f_high > 0:
            return None
        for _ in range(iterations):
            mid = (low + high) / 2
            f_mid = npv(mid, cfs)
            if abs(f_mid) < 1e-8:
                return mid
            if f_low * f_mid <= 0:
                high = mid
                f_high = f_mid
            else:
                low = mid
                f_low = f_mid
        return mid
    except Exception:
        return None


annual_irr = irr_bisect(irr_cashflows)
min_cash = cash_df["Cumulative Cash"].min()
breakeven_month = None
positive_months = cash_df[cash_df["Cumulative Cash"] >= 0]
if not positive_months.empty:
    breakeven_month = int(positive_months.iloc[0]["Month"])

roi_cs24 = None
if company_financed_per_cs24_try > 0:
    annual_energy_gp_per_cs24 = monthly_kwh_per_cs24 * company_margin_per_kwh * 12
    annual_other_gp_per_cs24 = cs24_other_gross_profit_usd * usdtry
    annual_net_per_cs24 = annual_energy_gp_per_cs24 + annual_other_gp_per_cs24 - (20000 + 0)
    roi_cs24 = annual_net_per_cs24 / company_financed_per_cs24_try


# --------------------------------------------------
# TABS
# --------------------------------------------------

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "Investor Dashboard",
    "Assumptions",
    "Monthly Cash Flow",
    "Income Statement",
    "Balance Sheet",
    "Excel Preview",
])

# --------------------------------------------------
# TAB 1 - DASHBOARD
# --------------------------------------------------

with tab1:
    st.subheader("Investor Dashboard")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Y4 Total Revenue", fmt_try(annual_df.iloc[-1]["Revenue"]))
    c2.metric("Y4 EBITDA Proxy", fmt_try(annual_df.iloc[-1]["EBITDA Proxy"]))
    c3.metric("Lowest Cash Balance", fmt_try(min_cash))
    c4.metric("Break-even Month", "-" if breakeven_month is None else breakeven_month)

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Y4 Active Motors", fmt_num(annual_df.iloc[-1]["Active Motors"]))
    c6.metric("Y4 Required CS24", fmt_num(annual_df.iloc[-1]["Required CS24"]))
    c7.metric("CS24 ROI", "-" if roi_cs24 is None else fmt_pct(roi_cs24))
    c8.metric("Project IRR", "-" if annual_irr is None else fmt_pct(annual_irr))

    st.markdown("### 4-Year Summary")
    view_cols = [
        "Year", "Sold Motors", "Active Motors", "Monthly kWh", "Revenue",
        "Total Gross Profit", "EBITDA Proxy", "CapEx Outflow", "Net Cash Flow", "Cumulative Cash"
    ]
    annual_display = annual_df[view_cols].copy()
    annual_display.rename(columns={"Monthly kWh": "Annual kWh"}, inplace=True)
    st.dataframe(annual_display, use_container_width=True)

    st.markdown("### Key Trend Charts")
    st.line_chart(annual_df.set_index("Year")[["Revenue", "EBITDA Proxy", "Net Cash Flow"]])
    st.line_chart(annual_df.set_index("Year")[["Active Motors", "Required CS24"]])

# --------------------------------------------------
# TAB 2 - ASSUMPTIONS
# --------------------------------------------------

with tab2:
    st.subheader("Assumptions")

    assumptions = pd.DataFrame({
        "Parameter": [
            "Base Fleet",
            "Km per Day",
            "Wh per km",
            "Daily kWh per Motor",
            "Battery Slots per CS24",
            "Battery kWh Each",
            "Total Battery kWh per CS24",
            "Delivery SoC",
            "Sellable kWh per Cycle",
            "Charging Cycle per CS24",
            "Daily kWh per CS24",
            "Buy Price",
            "Sell Price",
            "Gross Margin per kWh",
            "NewCo Share",
            "Franchise Share",
            "NewCo Margin per kWh",
            "CS24 Hardware USD",
            "Battery USD per kWh",
            "USDTRY",
            "CS24 Total CapEx TRY",
            "Down Payment %",
            "Company Financed per CS24",
            "Grace Period Months",
            "Battery Life Years",
            "Tax Rate",
        ],
        "Value": [
            fleet_base,
            daily_km,
            wh_per_km,
            daily_kwh_per_motor,
            battery_slots,
            battery_kwh_each,
            cs24_total_battery_kwh,
            delivery_soc,
            sellable_kwh_per_cycle,
            cycle_per_cs24,
            daily_kwh_per_cs24,
            buy_price,
            sell_price,
            gross_margin_per_kwh,
            company_share,
            franchise_share,
            company_margin_per_kwh,
            cs24_hardware_usd,
            battery_usd_per_kwh,
            usdtry,
            cs24_total_capex_try,
            down_payment_pct,
            company_financed_per_cs24_try,
            grace_period_months,
            battery_life_years,
            tax_rate,
        ]
    })
    st.dataframe(assumptions, use_container_width=True)

    st.markdown("### Excel Sheets Found")
    if sheet_check_ok:
        st.write(excel.sheet_names)
    else:
        st.info("No valid Excel loaded yet.")

# --------------------------------------------------
# TAB 3 - MONTHLY CASH FLOW
# --------------------------------------------------

with tab3:
    st.subheader("Monthly Cash Flow")
    st.dataframe(cash_df, use_container_width=True)
    st.line_chart(cash_df.set_index("Month")[["Net Cash Flow", "Cumulative Cash"]])

# --------------------------------------------------
# TAB 4 - INCOME STATEMENT
# --------------------------------------------------

with tab4:
    st.subheader("Income Statement (Annual)")

    income_statement = annual_df[[
        "Year", "Revenue", "Energy Cost", "Energy GP (NewCo)", "Motor Gross Profit",
        "CS24 Other GP", "Total Gross Profit", "Personnel Cost", "Other OPEX",
        "Rent/Admin", "Marketing", "G&A", "EBITDA Proxy", "Interest", "Tax", "Net Income"
    ]].copy()

    st.dataframe(income_statement, use_container_width=True)
    st.bar_chart(income_statement.set_index("Year")[["Revenue", "EBITDA Proxy", "Net Income"]])

# --------------------------------------------------
# TAB 5 - BALANCE SHEET
# --------------------------------------------------

with tab5:
    st.subheader("Balance Sheet (Annual Ending)")

    balance_sheet = annual_df[[
        "Year", "Cash Balance", "Assets - Installed CS24", "Liabilities - Financing", "Equity Proxy"
    ]].copy()

    balance_sheet.rename(columns={
        "Cash Balance": "Cash",
        "Assets - Installed CS24": "Fixed Assets Proxy",
        "Liabilities - Financing": "Financing Liability",
        "Equity Proxy": "Equity Proxy"
    }, inplace=True)

    st.dataframe(balance_sheet, use_container_width=True)
    st.line_chart(balance_sheet.set_index("Year")[["Cash", "Fixed Assets Proxy", "Financing Liability", "Equity Proxy"]])

# --------------------------------------------------
# TAB 6 - EXCEL PREVIEW
# --------------------------------------------------

with tab6:
    st.subheader("Excel Preview")

    if params_df is not None:
        st.markdown("### Main Parameters")
        st.dataframe(params_df, use_container_width=True)

    if personnel_df is not None:
        st.markdown("### Personnel")
        st.dataframe(personnel_df, use_container_width=True)

    if bom_df is not None:
        st.markdown("### BoM List")
        st.dataframe(bom_df, use_container_width=True)

    if cashflow_df is not None:
        st.markdown("### Cash Flow Statement")
        st.dataframe(cashflow_df, use_container_width=True)
