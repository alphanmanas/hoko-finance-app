import math
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="HOKO Mobility Financial Copilot v04.3", layout="wide")


# =========================================================
# FORMAT HELPERS
# =========================================================

def up_int(x):
    if pd.isna(x):
        return 0
    return int(math.ceil(float(x)))


def fmt_num(x):
    return f"{up_int(x):,}"


def fmt_try(x):
    return f"₺{up_int(x):,}"


def fmt_pct(x):
    return f"%{up_int(x)}"


def fmt_plain_int(x):
    return str(up_int(x))


def safe_float(x, default=0.0):
    try:
        if pd.isna(x):
            return default
        return float(x)
    except Exception:
        return default


def safe_int(x, default=0):
    try:
        if pd.isna(x):
            return default
        return int(math.ceil(float(x)))
    except Exception:
        return default


# =========================================================
# TEMPLATE EXCEL
# =========================================================

def make_template_excel():
    assumptions = pd.DataFrame({
        "Parameter": [
            "Grace Period (Months)",
            "Raw Material Required Before (Months)",
            "Stock-In Rate %",
            "Energy Buy Price (2025)",
            "Energy Sell Price (2025)",
            "NewCo Share %",
            "Battery Slots Per CS24",
            "Battery kWh Each",
            "Delivery SoC %",
            "Charging Cycle Per CS24",
            "CS24 Hardware & Installation USD",
            "Battery USD Per kWh",
            "Down Payment %",
            "Battery Life (Years)",
            "Tax Rate %",
            "Wh Per Km",
            "Active User % (Courier E-Scooter)",
            "Active User % (Commuter E-Scooter)",
            "Heavy Duty User % (Courier Scooter)",
            "Standard Duty User % (Courier Scooter)",
            "Light Duty User % (Courier Scooter)",
            "Heavy Duty Km Per Day",
            "Standard Duty Km Per Day",
            "Light Duty Km Per Day",
        ],
        "Value": [
            5, 3, 10, 3.5, 15, 40, 24, 3.24, 85, 2, 5000, 200, 25, 8, 20, 40,
            70, 30, 40, 40, 20, 180, 140, 100
        ]
    })

    monthly_inputs = pd.DataFrame({
        "Month": list(range(1, 49)),
        "Courier E-Scooter (Facelift) Units": [0]*48,
        "Courier E-Scooter (New) Units": [0]*48,
        "Commuter E-Scooter Units": [0]*48,
        "Active User % (Courier E-Scooter)": [70]*48,
        "Active User % (Commuter E-Scooter)": [30]*48,
    })

    macro = pd.DataFrame({
        "Year": [2025, 2026, 2027, 2028],
        "Average USD/TRY (H1)": [45, 52, 60, 68],
        "Average USD/TRY (H2)": [48, 56, 64, 72],
        "Average Inflation Rate (H1)": [20, 16, 12, 10],
        "Average Inflation Rate (H2)": [18, 14, 10, 8],
    })

    opex = pd.DataFrame({
        "Cost Item": ["Rent", "Marketing", "Software & IT", "HQ Expenses"],
        "Monthly Cost (2025)": [250000, 150000, 80000, 200000],
        "Type": ["Fixed", "Semi-Variable", "Fixed", "Fixed"],
        "Category": ["Admin", "Growth", "Tech", "Admin"],
        "Escalation Type": ["Inflation", "Inflation", "Inflation", "Inflation"],
        "Escalation Rate %": [0, 0, 0, 0],
        "Notes": ["", "", "", ""],
    })

    personnel_base = pd.DataFrame({
        "Role": [
            "Plant Manager (Fabrika Müdürü)",
            "Assembly Operator (Montaj Operatörü)",
            "Line Leader (Hat Lideri)",
            "Quality Control Operator (Kalite Kontrol Operatörü)",
            "Test Operator (Test Operatörü)",
            "Line Feeder / Internal Logistics Operator (Hat Besleme / İç Lojistik Operatörü)",
            "CEO",
            "Business Development Manager",
        ],
        "Base Salary (2025)": [120000, 35000, 45000, 40000, 38000, 36000, 250000, 120000],
        "Notes": ["", "", "", "", "", "", "", ""],
    })

    personnel_rules = pd.DataFrame({
        "Role": [
            "Plant Manager (Fabrika Müdürü)",
            "Assembly Operator (Montaj Operatörü)",
            "Line Leader (Hat Lideri)",
            "Quality Control Operator (Kalite Kontrol Operatörü)",
            "Test Operator (Test Operatörü)",
            "Line Feeder / Internal Logistics Operator (Hat Besleme / İç Lojistik Operatörü)",
            "CEO",
            "Business Development Manager",
        ],
        "First Hire Month After Grace": [1, 1, 1, 1, 1, 1, 4, 4],
        "Scaling Factor": [1, 1000, 1000, 1000, 1000, 1000, 1, 3000],
        "Max Cap": [5, 500, 50, 50, 50, 50, 1, 10],
        "Base Headcount": [1, 1, 1, 1, 1, 1, 1, 1],
    })

    bom = pd.DataFrame({
        "Item": ["Frame", "Motor", "Controller", "Plastic Parts"],
        "Product": [
            "Courier E-Scooter (Facelift)",
            "Courier E-Scooter (New)",
            "Commuter E-Scooter",
            "Courier E-Scooter (New)"
        ],
        "Unit Cost USD": [120, 180, 90, 60],
    })

    product_capex = pd.DataFrame({
        "Product": [
            "Courier E-Scooter (Facelift)",
            "Courier E-Scooter (New)",
            "Commuter E-Scooter",
        ],
        "CapEx Per Unit USD": [700, 900, 650],
        "Battery kWh": [6.48, 6.48, 4.86],
        "Notes": ["", "", ""],
    })

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        assumptions.to_excel(writer, sheet_name="Assumptions", index=False)
        monthly_inputs.to_excel(writer, sheet_name="Monthly Inputs", index=False)
        macro.to_excel(writer, sheet_name="Macro Assumptions", index=False)
        opex.to_excel(writer, sheet_name="OPEX", index=False)
        personnel_base.to_excel(writer, sheet_name="Personnel Base", index=False)
        personnel_rules.to_excel(writer, sheet_name="Personnel Monthly Rules", index=False)
        bom.to_excel(writer, sheet_name="BoM List", index=False)
        product_capex.to_excel(writer, sheet_name="Product CapEx", index=False)

    bio.seek(0)
    return bio.getvalue()


# =========================================================
# PARAM READERS
# =========================================================

def get_param_value(df, name, default=0):
    if df is None or df.empty:
        return default
    cols = [str(c).strip().lower() for c in df.columns]
    if "parameter" not in cols or "value" not in cols:
        return default
    pcol = df.columns[cols.index("parameter")]
    vcol = df.columns[cols.index("value")]
    hit = df[df[pcol].astype(str).str.strip().str.lower() == name.strip().lower()]
    if hit.empty:
        return default
    return hit.iloc[0][vcol]


def get_half_for_month(month_no):
    year_index = (month_no - 1) // 12
    month_in_year = ((month_no - 1) % 12) + 1
    year = 2025 + year_index
    half = "H1" if month_in_year <= 6 else "H2"
    return year, half


def get_macro_value(macro_df, year, half, field_base):
    if macro_df is None or macro_df.empty:
        return 0
    year_row = macro_df[macro_df["Year"] == year]
    if year_row.empty:
        return 0
    col = f"{field_base} ({half})"
    if col not in macro_df.columns:
        return 0
    return safe_float(year_row.iloc[0][col], 0)


# =========================================================
# HEADER
# =========================================================

st.title("HOKO Mobility Financial Copilot v04.3")
st.caption("Excel-Driven, Multi-Product, Monthly Financial Model")

with st.expander("Excel Upload Instructions", expanded=True):
    st.markdown("""
**Please upload your Business Plan Excel file.**

Required sheets:
1. Assumptions
2. Monthly Inputs
3. Macro Assumptions
4. OPEX
5. Personnel Base
6. Personnel Monthly Rules
7. BoM List
8. Product CapEx

Cash Flow Statement and Balance Sheet are generated inside the app.
""")

st.download_button(
    "Download Template Excel",
    data=make_template_excel(),
    file_name="HOKO_BP_Template_v04_3.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

uploaded_file = st.file_uploader("Upload Business Plan Excel", type=["xlsx"])


# =========================================================
# READ EXCEL
# =========================================================

assumptions_df = None
monthly_inputs_df = None
macro_df = None
opex_df = None
personnel_base_df = None
personnel_rules_df = None
bom_df = None
product_capex_df = None

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        required = [
            "Assumptions",
            "Monthly Inputs",
            "Macro Assumptions",
            "OPEX",
            "Personnel Base",
            "Personnel Monthly Rules",
            "BoM List",
            "Product CapEx",
        ]
        missing = [s for s in required if s not in xls.sheet_names]
        if missing:
            st.error(f"Missing sheet(s): {', '.join(missing)}")
        else:
            assumptions_df = pd.read_excel(xls, "Assumptions")
            monthly_inputs_df = pd.read_excel(xls, "Monthly Inputs")
            macro_df = pd.read_excel(xls, "Macro Assumptions")
            opex_df = pd.read_excel(xls, "OPEX")
            personnel_base_df = pd.read_excel(xls, "Personnel Base")
            personnel_rules_df = pd.read_excel(xls, "Personnel Monthly Rules")
            bom_df = pd.read_excel(xls, "BoM List")
            product_capex_df = pd.read_excel(xls, "Product CapEx")
            st.success("Excel uploaded successfully.")
    except Exception as e:
        st.error(f"Excel could not be read: {e}")


# =========================================================
# DEFAULTS FROM EXCEL / FALLBACK
# =========================================================

base_grace_period_months = safe_int(get_param_value(assumptions_df, "Grace Period (Months)", 5))
base_raw_material_required_before_months = safe_int(get_param_value(assumptions_df, "Raw Material Required Before (Months)", 3))
base_stock_in_rate_pct = safe_float(get_param_value(assumptions_df, "Stock-In Rate %", 10))
base_energy_buy_price_2025 = safe_float(get_param_value(assumptions_df, "Energy Buy Price (2025)", 3.5))
base_energy_sell_price_2025 = safe_float(get_param_value(assumptions_df, "Energy Sell Price (2025)", 15))
base_newco_share_pct = safe_float(get_param_value(assumptions_df, "NewCo Share %", 40))
base_battery_slots_per_cs24 = safe_int(get_param_value(assumptions_df, "Battery Slots Per CS24", 24))
base_battery_kwh_each = safe_float(get_param_value(assumptions_df, "Battery kWh Each", 3.24))
base_delivery_soc_pct = safe_float(get_param_value(assumptions_df, "Delivery SoC %", 85))
base_charging_cycle_per_cs24 = safe_float(get_param_value(assumptions_df, "Charging Cycle Per CS24", 2))
base_cs24_hardware_installation_usd = safe_float(get_param_value(assumptions_df, "CS24 Hardware & Installation USD", 5000))
base_battery_usd_per_kwh = safe_float(get_param_value(assumptions_df, "Battery USD Per kWh", 200))
base_down_payment_pct = safe_float(get_param_value(assumptions_df, "Down Payment %", 25))
base_battery_life_years = safe_int(get_param_value(assumptions_df, "Battery Life (Years)", 8))
base_tax_rate_pct = safe_float(get_param_value(assumptions_df, "Tax Rate %", 20))
base_wh_per_km = safe_float(get_param_value(assumptions_df, "Wh Per Km", 40))

base_active_user_pct_courier = safe_float(get_param_value(assumptions_df, "Active User % (Courier E-Scooter)", 70))
base_active_user_pct_commuter = safe_float(get_param_value(assumptions_df, "Active User % (Commuter E-Scooter)", 30))
base_heavy_duty_user_pct = safe_float(get_param_value(assumptions_df, "Heavy Duty User % (Courier Scooter)", 40))
base_standard_duty_user_pct = safe_float(get_param_value(assumptions_df, "Standard Duty User % (Courier Scooter)", 40))
base_light_duty_user_pct = safe_float(get_param_value(assumptions_df, "Light Duty User % (Courier Scooter)", 20))
base_heavy_duty_km_per_day = safe_float(get_param_value(assumptions_df, "Heavy Duty Km Per Day", 180))
base_standard_duty_km_per_day = safe_float(get_param_value(assumptions_df, "Standard Duty Km Per Day", 140))
base_light_duty_km_per_day = safe_float(get_param_value(assumptions_df, "Light Duty Km Per Day", 100))


# =========================================================
# ASSUMPTIONS
# =========================================================

st.markdown("## Assumptions")

with st.expander("Editable Assumptions", expanded=True):
    e1, e2, e3 = st.columns(3)

    with e1:
        grace_period_months = st.number_input("Grace Period (Months)", min_value=0, value=base_grace_period_months, step=1)
        raw_material_required_before_months = st.number_input("Raw Material Required Before (Months)", min_value=0, value=base_raw_material_required_before_months, step=1)
        stock_in_rate_pct = st.number_input("Stock-In Rate %", min_value=0, max_value=100, value=up_int(base_stock_in_rate_pct), step=1)
        energy_buy_price_2025 = st.number_input("Energy Buy Price (2025)", min_value=0, value=up_int(base_energy_buy_price_2025), step=1)
        energy_sell_price_2025 = st.number_input("Energy Sell Price (2025)", min_value=0, value=up_int(base_energy_sell_price_2025), step=1)

    with e2:
        newco_share_pct = st.number_input("NewCo Share %", min_value=0, max_value=100, value=up_int(base_newco_share_pct), step=1)
        battery_slots_per_cs24 = st.number_input("Battery Slots Per CS24", min_value=1, value=base_battery_slots_per_cs24, step=1)
        battery_kwh_each = st.number_input("Battery kWh Each", min_value=0, value=up_int(base_battery_kwh_each), step=1)
        delivery_soc_pct = st.number_input("Delivery SoC %", min_value=0, max_value=100, value=up_int(base_delivery_soc_pct), step=1)
        charging_cycle_per_cs24 = st.number_input("Charging Cycle Per CS24", min_value=0, value=up_int(base_charging_cycle_per_cs24), step=1)

    with e3:
        cs24_hardware_installation_usd = st.number_input("CS24 Hardware & Installation USD", min_value=0, value=up_int(base_cs24_hardware_installation_usd), step=100)
        battery_usd_per_kwh = st.number_input("Battery USD Per kWh", min_value=0, value=up_int(base_battery_usd_per_kwh), step=10)
        down_payment_pct = st.number_input("Down Payment %", min_value=0, max_value=100, value=up_int(base_down_payment_pct), step=1)
        battery_life_years = st.number_input("Battery Life (Years)", min_value=1, value=base_battery_life_years, step=1)
        tax_rate_pct = st.number_input("Tax Rate %", min_value=0, max_value=100, value=up_int(base_tax_rate_pct), step=1)
        wh_per_km = st.number_input("Wh Per Km", min_value=0, value=up_int(base_wh_per_km), step=1)

with st.expander("Override Assumptions", expanded=True):
    o1, o2 = st.columns(2)

    with o1:
        active_user_pct_courier_avg = st.text_input("Active User % (Courier E-Scooter)", value=str(up_int(base_active_user_pct_courier)))
        active_user_pct_commuter_avg = st.text_input("Active User % (Commuter E-Scooter)", value=str(up_int(base_active_user_pct_commuter)))

    with o2:
        heavy_duty_user_pct = st.text_input("Heavy Duty User % (Courier Scooter)", value=str(up_int(base_heavy_duty_user_pct)))
        standard_duty_user_pct = st.text_input("Standard Duty User % (Courier Scooter)", value=str(up_int(base_standard_duty_user_pct)))
        light_duty_user_pct = st.text_input("Light Duty User % (Courier Scooter)", value=str(up_int(base_light_duty_user_pct)))
        heavy_duty_km_per_day = st.text_input("Heavy Duty Km Per Day", value=str(up_int(base_heavy_duty_km_per_day)))
        standard_duty_km_per_day = st.text_input("Standard Duty Km Per Day", value=str(up_int(base_standard_duty_km_per_day)))
        light_duty_km_per_day = st.text_input("Light Duty Km Per Day", value=str(up_int(base_light_duty_km_per_day)))

active_user_pct_courier_avg = safe_int(active_user_pct_courier_avg, up_int(base_active_user_pct_courier))
active_user_pct_commuter_avg = safe_int(active_user_pct_commuter_avg, up_int(base_active_user_pct_commuter))
heavy_duty_user_pct = safe_int(heavy_duty_user_pct, up_int(base_heavy_duty_user_pct))
standard_duty_user_pct = safe_int(standard_duty_user_pct, up_int(base_standard_duty_user_pct))
light_duty_user_pct = safe_int(light_duty_user_pct, up_int(base_light_duty_user_pct))
heavy_duty_km_per_day = safe_int(heavy_duty_km_per_day, up_int(base_heavy_duty_km_per_day))
standard_duty_km_per_day = safe_int(standard_duty_km_per_day, up_int(base_standard_duty_km_per_day))
light_duty_km_per_day = safe_int(light_duty_km_per_day, up_int(base_light_duty_km_per_day))

if heavy_duty_user_pct + standard_duty_user_pct + light_duty_user_pct != 100:
    st.warning(
        f"Heavy Duty + Standard Duty + Light Duty must equal %100. Current total: "
        f"%{heavy_duty_user_pct + standard_duty_user_pct + light_duty_user_pct}"
    )


# =========================================================
# CALCULATED ASSUMPTIONS
# =========================================================

weighted_avg_km_per_day_courier = math.ceil(
    heavy_duty_km_per_day * heavy_duty_user_pct / 100
    + standard_duty_km_per_day * standard_duty_user_pct / 100
    + light_duty_km_per_day * light_duty_user_pct / 100
)

commuter_km_per_day = 35

km_per_day_system_avg = math.ceil(
    weighted_avg_km_per_day_courier * active_user_pct_courier_avg / 100
    + commuter_km_per_day * active_user_pct_commuter_avg / 100
)

franchise_share_pct = 100 - newco_share_pct
gross_margin_per_kwh = math.ceil(energy_sell_price_2025 - energy_buy_price_2025)
newco_margin_per_kwh = math.ceil((energy_sell_price_2025 - energy_buy_price_2025) * newco_share_pct / 100)

total_battery_kwh_per_cs24 = math.ceil(battery_slots_per_cs24 * battery_kwh_each)
sellable_kwh_per_cycle = math.ceil(total_battery_kwh_per_cs24 * delivery_soc_pct / 100)
daily_kwh_per_cs24 = math.ceil(sellable_kwh_per_cycle * charging_cycle_per_cs24)
kwh_per_motor_per_day = math.ceil(km_per_day_system_avg * wh_per_km / 1000)


# =========================================================
# PERSONNEL MAPS
# =========================================================

base_salary_map = {}
if personnel_base_df is not None and not personnel_base_df.empty:
    for _, row in personnel_base_df.iterrows():
        base_salary_map[str(row["Role"]).strip()] = safe_float(row["Base Salary (2025)"], 0)

personnel_rule_rows = []
if personnel_rules_df is not None and not personnel_rules_df.empty:
    for _, row in personnel_rules_df.iterrows():
        personnel_rule_rows.append({
            "Role": str(row["Role"]).strip(),
            "First Hire Month After Grace": safe_int(row["First Hire Month After Grace"], 1),
            "Scaling Factor": max(1, safe_int(row["Scaling Factor"], 1)),
            "Max Cap": max(1, safe_int(row["Max Cap"], 1)),
            "Base Headcount": max(1, safe_int(row["Base Headcount"], 1)),
        })


# =========================================================
# PRODUCT MAPS
# =========================================================

bom_monthly_usd_per_product = {}
if bom_df is not None and not bom_df.empty:
    grouped = bom_df.groupby("Product", as_index=False)["Unit Cost USD"].sum()
    for _, row in grouped.iterrows():
        bom_monthly_usd_per_product[str(row["Product"]).strip()] = safe_float(row["Unit Cost USD"], 0)


# =========================================================
# FALLBACK MONTHLY INPUTS
# =========================================================

def build_default_monthly_inputs():
    rows = []
    for m in range(1, 49):
        if 1 <= m <= 5:
            total = 0
        elif m == 6:
            total = 2000
        elif m == 7:
            total = 2500
        elif m in [8, 9]:
            total = 3000
        elif m in [10, 11, 12]:
            total = 3250
        elif 13 <= m <= 18:
            total = 3500
        else:
            total = 4000

        rows.append({
            "Month": m,
            "Courier E-Scooter (Facelift) Units": 0,
            "Courier E-Scooter (New) Units": total,
            "Commuter E-Scooter Units": 0,
            "Active User % (Courier E-Scooter)": 70,
            "Active User % (Commuter E-Scooter)": 30,
        })
    return pd.DataFrame(rows)


if monthly_inputs_df is None or monthly_inputs_df.empty:
    monthly_inputs_df = build_default_monthly_inputs()


# =========================================================
# MONTHLY ENGINE
# =========================================================

results = []
inventory_units = 0
inventory_raw_material_value = 0
active_escooters = 0
installed_cs24 = 0
cumulative_cash = 0
personnel_records = []
raw_material_purchase_schedule = {m: 0 for m in range(1, 49)}

for _, row in monthly_inputs_df.iterrows():
    month_no = safe_int(row["Month"], 1)
    year, half = get_half_for_month(month_no)

    usd_try_month = get_macro_value(macro_df, year, half, "Average USD/TRY")
    inflation_rate_month = get_macro_value(macro_df, year, half, "Average Inflation Rate")

    if usd_try_month == 0:
        usd_try_month = 45
    if inflation_rate_month == 0:
        inflation_rate_month = 15

    courier_facelift_units = safe_int(row.get("Courier E-Scooter (Facelift) Units", 0), 0)
    courier_new_units = safe_int(row.get("Courier E-Scooter (New) Units", 0), 0)
    commuter_units = safe_int(row.get("Commuter E-Scooter Units", 0), 0)

    monthly_production_units = courier_facelift_units + courier_new_units + commuter_units

    stock_in_rate_pct_month = stock_in_rate_pct
    if "Stock-In Rate %" in row.index:
        stock_in_rate_pct_month = safe_float(row["Stock-In Rate %"], stock_in_rate_pct)

    units_added_to_stock = math.ceil(monthly_production_units * stock_in_rate_pct_month / 100)
    units_sold_month = monthly_production_units - units_added_to_stock

    inventory_units += units_added_to_stock
    active_escooters += units_sold_month

    energy_buy_price_month = math.ceil(energy_buy_price_2025 * (1 + inflation_rate_month / 100))
    energy_sell_price_month = math.ceil(energy_sell_price_2025 * (1 + inflation_rate_month / 100))
    newco_margin_per_kwh_month = math.ceil((energy_sell_price_month - energy_buy_price_month) * newco_share_pct / 100)

    monthly_kwh_sold = math.ceil(active_escooters * kwh_per_motor_per_day * 30)
    monthly_revenue = math.ceil(monthly_kwh_sold * energy_sell_price_month)
    monthly_energy_cost = math.ceil(monthly_kwh_sold * energy_buy_price_month)
    monthly_energy_gp_newco = math.ceil(monthly_kwh_sold * newco_margin_per_kwh_month)

    monthly_motor_gross_profit = math.ceil(units_sold_month * 80000)

    monthly_cs24_required = math.ceil((active_escooters * kwh_per_motor_per_day) / max(1, daily_kwh_per_cs24))
    monthly_new_cs24 = max(0, monthly_cs24_required - installed_cs24)
    installed_cs24 = monthly_cs24_required

    cs24_total_capex_try_month = math.ceil(
        (cs24_hardware_installation_usd + battery_slots_per_cs24 * battery_kwh_each * battery_usd_per_kwh) * usd_try_month
    )
    financed_per_cs24_try_month = math.ceil(cs24_total_capex_try_month * (1 - down_payment_pct / 100))
    monthly_cs24_capex_outflow = math.ceil(monthly_new_cs24 * financed_per_cs24_try_month)

    monthly_cs24_other_gp = math.ceil(installed_cs24 * ((1200 * usd_try_month) / 12))

    bom_usd_total = (
        courier_facelift_units * bom_monthly_usd_per_product.get("Courier E-Scooter (Facelift)", 0)
        + courier_new_units * bom_monthly_usd_per_product.get("Courier E-Scooter (New)", 0)
        + commuter_units * bom_monthly_usd_per_product.get("Commuter E-Scooter", 0)
    )
    current_raw_material_need_try = math.ceil(bom_usd_total * usd_try_month)

    slowdown_factor = max(0, 1 - stock_in_rate_pct_month / 100)
    adjusted_raw_material_need_try = math.ceil(current_raw_material_need_try * slowdown_factor)

    purchase_month = month_no - raw_material_required_before_months
    if 1 <= purchase_month <= 48:
        raw_material_purchase_schedule[purchase_month] += adjusted_raw_material_need_try

    raw_material_cost_month = raw_material_purchase_schedule.get(month_no, 0)
    inventory_raw_material_value = max(0, inventory_raw_material_value + raw_material_cost_month - adjusted_raw_material_need_try)

    total_personnel_cost_month = 0

    for pr in personnel_rule_rows:
        role = pr["Role"]
        first_hire_month_after_grace = pr["First Hire Month After Grace"]
        actual_hire_month = grace_period_months + first_hire_month_after_grace - 1

        if month_no < actual_hire_month:
            headcount = 0
        else:
            scaled = math.ceil(monthly_production_units / max(1, pr["Scaling Factor"]))
            headcount = max(pr["Base Headcount"], scaled)
            headcount = min(headcount, pr["Max Cap"])

        salary_base = base_salary_map.get(role, 0)
        monthly_salary_actual = math.ceil(salary_base * (1 + inflation_rate_month / 100))
        monthly_personnel_cost = math.ceil(headcount * monthly_salary_actual)

        total_personnel_cost_month += monthly_personnel_cost

        personnel_records.append({
            "Month": month_no,
            "Role": role,
            "Headcount": headcount,
            "Monthly Salary (Actual)": monthly_salary_actual,
            "Monthly Personnel Cost": monthly_personnel_cost,
        })

    total_opex_month = 0
    if opex_df is not None and not opex_df.empty:
        for _, orow in opex_df.iterrows():
            base_cost = safe_float(orow["Monthly Cost (2025)"], 0)
            esc_type = str(orow["Escalation Type"]).strip().lower()
            custom_rate = safe_float(orow["Escalation Rate %"], 0)

            if esc_type == "custom":
                actual_cost = math.ceil(base_cost * (1 + custom_rate / 100))
            elif esc_type == "fx linked":
                actual_cost = math.ceil(base_cost * (usd_try_month / 45))
            else:
                actual_cost = math.ceil(base_cost * (1 + inflation_rate_month / 100))

            total_opex_month += actual_cost

    monthly_total_gross_profit = math.ceil(
        monthly_energy_gp_newco + monthly_motor_gross_profit + monthly_cs24_other_gp
    )
    monthly_ebitda_proxy = math.ceil(
        monthly_total_gross_profit - total_personnel_cost_month - total_opex_month
    )

    monthly_interest_cost = 0
    monthly_tax = math.ceil(max(0, monthly_ebitda_proxy) * tax_rate_pct / 100)
    monthly_net_income = math.ceil(monthly_ebitda_proxy - monthly_interest_cost - monthly_tax)

    monthly_net_cash_flow = math.ceil(
        monthly_net_income - monthly_cs24_capex_outflow - raw_material_cost_month
    )
    cumulative_cash += monthly_net_cash_flow

    installed_cs24_asset_value = math.ceil(installed_cs24 * cs24_total_capex_try_month)
    financing_liability = math.ceil(installed_cs24 * financed_per_cs24_try_month)
    equity_proxy = math.ceil(max(0, cumulative_cash) + installed_cs24 * (cs24_total_capex_try_month - financed_per_cs24_try_month))

    results.append({
        "Month": month_no,
        "Year": year,
        "Half": half,
        "Courier E-Scooter (Facelift) Units": courier_facelift_units,
        "Courier E-Scooter (New) Units": courier_new_units,
        "Commuter E-Scooter Units": commuter_units,
        "Monthly Production Units": monthly_production_units,
        "Units Sold": units_sold_month,
        "Units Added To Stock": units_added_to_stock,
        "Ending Stock Units": inventory_units,
        "Active E-Scooters": active_escooters,
        "USD/TRY (Month)": usd_try_month,
        "Inflation Rate (Month)": inflation_rate_month,
        "Energy Buy Price (Month)": energy_buy_price_month,
        "Energy Sell Price (Month)": energy_sell_price_month,
        "Monthly kWh Sold": monthly_kwh_sold,
        "Monthly Revenue": monthly_revenue,
        "Monthly Energy Cost": monthly_energy_cost,
        "Monthly Energy Gross Profit (NewCo)": monthly_energy_gp_newco,
        "Monthly Motor Gross Profit": monthly_motor_gross_profit,
        "Monthly CS24 Other Gross Profit": monthly_cs24_other_gp,
        "Monthly Total Gross Profit": monthly_total_gross_profit,
        "Required CS24": monthly_cs24_required,
        "New CS24": monthly_new_cs24,
        "CS24 CapEx Outflow": monthly_cs24_capex_outflow,
        "Raw Material Cost": raw_material_cost_month,
        "Total Personnel Cost": total_personnel_cost_month,
        "Total OPEX": total_opex_month,
        "Monthly EBITDA Proxy": monthly_ebitda_proxy,
        "Monthly Interest Cost": monthly_interest_cost,
        "Monthly Tax": monthly_tax,
        "Monthly Net Income": monthly_net_income,
        "Monthly Net Cash Flow": monthly_net_cash_flow,
        "Cumulative Cash Balance": cumulative_cash,
        "Raw Material Inventory": inventory_raw_material_value,
        "Finished Goods Inventory": inventory_units,
        "Installed CS24 Asset Value": installed_cs24_asset_value,
        "Financing Liability": financing_liability,
        "Equity Proxy": equity_proxy,
    })

monthly_df = pd.DataFrame(results)

annual_df = monthly_df.groupby("Year", as_index=False).agg({
    "Monthly Production Units": "sum",
    "Units Sold": "sum",
    "Active E-Scooters": "last",
    "Monthly kWh Sold": "sum",
    "Monthly Revenue": "sum",
    "Monthly Energy Cost": "sum",
    "Monthly Energy Gross Profit (NewCo)": "sum",
    "Monthly Motor Gross Profit": "sum",
    "Monthly CS24 Other Gross Profit": "sum",
    "Monthly Total Gross Profit": "sum",
    "Total Personnel Cost": "sum",
    "Total OPEX": "sum",
    "Monthly EBITDA Proxy": "sum",
    "CS24 CapEx Outflow": "sum",
    "Raw Material Cost": "sum",
    "Monthly Net Cash Flow": "sum",
    "Cumulative Cash Balance": "last",
    "Required CS24": "last",
    "Installed CS24 Asset Value": "last",
    "Financing Liability": "last",
    "Equity Proxy": "last",
})

personnel_detail_df = pd.DataFrame(personnel_records) if personnel_records else pd.DataFrame()

cash_flow_df = monthly_df[[
    "Month",
    "Monthly Revenue",
    "Monthly Energy Cost",
    "Monthly Energy Gross Profit (NewCo)",
    "Monthly Motor Gross Profit",
    "Monthly CS24 Other Gross Profit",
    "Monthly Total Gross Profit",
    "Total Personnel Cost",
    "Total OPEX",
    "Monthly EBITDA Proxy",
    "CS24 CapEx Outflow",
    "Raw Material Cost",
    "Monthly Interest Cost",
    "Monthly Tax",
    "Monthly Net Income",
    "Monthly Net Cash Flow",
    "Cumulative Cash Balance",
]].copy()

balance_sheet_df = monthly_df[[
    "Month",
    "Cumulative Cash Balance",
    "Raw Material Inventory",
    "Finished Goods Inventory",
    "Installed CS24 Asset Value",
    "Financing Liability",
    "Equity Proxy",
]].copy()

assumption_rows = [
    ["Grace Period (Months)", grace_period_months, "Editable"],
    ["Raw Material Required Before (Months)", raw_material_required_before_months, "Editable"],
    ["Stock-In Rate %", stock_in_rate_pct, "Editable"],
    ["Energy Buy Price (2025)", energy_buy_price_2025, "Editable"],
    ["Energy Sell Price (2025)", energy_sell_price_2025, "Editable"],
    ["Gross Margin Per kWh", gross_margin_per_kwh, "Calculated"],
    ["NewCo Share %", newco_share_pct, "Editable"],
    ["Franchise Share %", franchise_share_pct, "Calculated"],
    ["Battery Slots Per CS24", battery_slots_per_cs24, "Editable"],
    ["Battery kWh Each", battery_kwh_each, "Editable"],
    ["Delivery SoC %", delivery_soc_pct, "Editable"],
    ["Charging Cycle Per CS24", charging_cycle_per_cs24, "Editable"],
    ["Total Battery kWh Per CS24", total_battery_kwh_per_cs24, "Calculated"],
    ["Sellable kWh Per Cycle", sellable_kwh_per_cycle, "Calculated"],
    ["Daily kWh Per CS24", daily_kwh_per_cs24, "Calculated"],
    ["CS24 Hardware & Installation USD", cs24_hardware_installation_usd, "Editable"],
    ["Battery USD Per kWh", battery_usd_per_kwh, "Editable"],
    ["Down Payment %", down_payment_pct, "Editable"],
    ["Battery Life (Years)", battery_life_years, "Editable"],
    ["Tax Rate %", tax_rate_pct, "Editable"],
    ["Wh Per Km", wh_per_km, "Editable"],
    ["Active User % (Courier E-Scooter)", active_user_pct_courier_avg, "Override"],
    ["Active User % (Commuter E-Scooter)", active_user_pct_commuter_avg, "Override"],
    ["Heavy Duty User % (Courier Scooter)", heavy_duty_user_pct, "Override"],
    ["Standard Duty User % (Courier Scooter)", standard_duty_user_pct, "Override"],
    ["Light Duty User % (Courier Scooter)", light_duty_user_pct, "Override"],
    ["Heavy Duty Km Per Day", heavy_duty_km_per_day, "Override"],
    ["Standard Duty Km Per Day", standard_duty_km_per_day, "Override"],
    ["Light Duty Km Per Day", light_duty_km_per_day, "Override"],
    ["Weighted Avg Km Per Day (Courier E-Scooter)", weighted_avg_km_per_day_courier, "Calculated"],
    ["Km Per Day (System Avg)", km_per_day_system_avg, "Calculated"],
    ["kWh Per Motor Per Day", kwh_per_motor_per_day, "Calculated"],
]
assumptions_display_df = pd.DataFrame(assumption_rows, columns=["Parameter", "Value", "Type"])


# =========================================================
# DISPLAY FORMATTING
# =========================================================

assumptions_display_fmt = assumptions_display_df.copy()
assumptions_display_fmt["Value"] = assumptions_display_fmt["Value"].astype(object)
assumptions_display_fmt["Type"] = assumptions_display_fmt["Type"].astype(object)

for i, row in assumptions_display_fmt.iterrows():
    param = str(row["Parameter"])
    val = row["Value"]

    if "%" in param:
        assumptions_display_fmt.at[i, "Value"] = fmt_pct(val)
    elif "USD" in param:
        assumptions_display_fmt.at[i, "Value"] = fmt_num(val)
    elif "Price" in param or "Margin" in param:
        assumptions_display_fmt.at[i, "Value"] = fmt_try(val)
    else:
        assumptions_display_fmt.at[i, "Value"] = fmt_num(val)

dashboard_monthly_fmt = monthly_df.copy()
for c in dashboard_monthly_fmt.columns:
    dashboard_monthly_fmt[c] = dashboard_monthly_fmt[c].astype(object)

currency_cols = [
    "Monthly Revenue", "Monthly Energy Cost", "Monthly Energy Gross Profit (NewCo)",
    "Monthly Motor Gross Profit", "Monthly CS24 Other Gross Profit", "Monthly Total Gross Profit",
    "CS24 CapEx Outflow", "Raw Material Cost", "Total Personnel Cost", "Total OPEX",
    "Monthly EBITDA Proxy", "Monthly Interest Cost", "Monthly Tax", "Monthly Net Income",
    "Monthly Net Cash Flow", "Cumulative Cash Balance", "Raw Material Inventory",
    "Installed CS24 Asset Value", "Financing Liability", "Equity Proxy"
]
num_cols = [
    "Courier E-Scooter (Facelift) Units", "Courier E-Scooter (New) Units",
    "Commuter E-Scooter Units", "Monthly Production Units", "Units Sold", "Units Added To Stock",
    "Ending Stock Units", "Active E-Scooters", "USD/TRY (Month)", "Inflation Rate (Month)",
    "Energy Buy Price (Month)", "Energy Sell Price (Month)", "Monthly kWh Sold", "Required CS24",
    "New CS24", "Finished Goods Inventory"
]
plain_cols = ["Month", "Year"]

for c in plain_cols:
    if c in dashboard_monthly_fmt.columns:
        dashboard_monthly_fmt[c] = dashboard_monthly_fmt[c].map(fmt_plain_int)

for c in currency_cols:
    if c in dashboard_monthly_fmt.columns:
        dashboard_monthly_fmt[c] = dashboard_monthly_fmt[c].map(fmt_try)

for c in num_cols:
    if c in dashboard_monthly_fmt.columns:
        dashboard_monthly_fmt[c] = dashboard_monthly_fmt[c].map(fmt_num)

annual_fmt = annual_df.copy()
for c in annual_fmt.columns:
    annual_fmt[c] = annual_fmt[c].astype(object)

for c in annual_fmt.columns:
    if c == "Year":
        annual_fmt[c] = annual_fmt[c].map(fmt_plain_int)
    elif c in ["Monthly Production Units", "Units Sold", "Active E-Scooters", "Monthly kWh Sold", "Required CS24"]:
        annual_fmt[c] = annual_fmt[c].map(fmt_num)
    else:
        annual_fmt[c] = annual_fmt[c].map(fmt_try)

personnel_fmt = personnel_detail_df.copy()
if not personnel_fmt.empty:
    for c in personnel_fmt.columns:
        personnel_fmt[c] = personnel_fmt[c].astype(object)

    if "Month" in personnel_fmt.columns:
        personnel_fmt["Month"] = personnel_fmt["Month"].map(fmt_plain_int)
    if "Headcount" in personnel_fmt.columns:
        personnel_fmt["Headcount"] = personnel_fmt["Headcount"].map(fmt_num)
    if "Monthly Salary (Actual)" in personnel_fmt.columns:
        personnel_fmt["Monthly Salary (Actual)"] = personnel_fmt["Monthly Salary (Actual)"].map(fmt_try)
    if "Monthly Personnel Cost" in personnel_fmt.columns:
        personnel_fmt["Monthly Personnel Cost"] = personnel_fmt["Monthly Personnel Cost"].map(fmt_try)

cash_flow_fmt = cash_flow_df.copy()
for c in cash_flow_fmt.columns:
    cash_flow_fmt[c] = cash_flow_fmt[c].astype(object)

for c in cash_flow_fmt.columns:
    if c == "Month":
        cash_flow_fmt[c] = cash_flow_fmt[c].map(fmt_plain_int)
    else:
        cash_flow_fmt[c] = cash_flow_fmt[c].map(fmt_try)

balance_fmt = balance_sheet_df.copy()
for c in balance_fmt.columns:
    balance_fmt[c] = balance_fmt[c].astype(object)

for c in balance_fmt.columns:
    if c == "Month":
        balance_fmt[c] = balance_fmt[c].map(fmt_plain_int)
    elif c == "Finished Goods Inventory":
        balance_fmt[c] = balance_fmt[c].map(fmt_num)
    else:
        balance_fmt[c] = balance_fmt[c].map(fmt_try)


# =========================================================
# TABS
# =========================================================

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "Investor Dashboard",
    "Assumptions",
    "Monthly Financials",
    "Cash Flow Statement",
    "Balance Sheet",
    "Personnel",
    "Excel Preview",
])

with tab1:
    st.subheader("Investor Dashboard")

    last_row = monthly_df.iloc[-1]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Final Active E-Scooters", fmt_num(last_row["Active E-Scooters"]))
    c2.metric("Final Monthly kWh Sold", fmt_num(last_row["Monthly kWh Sold"]))
    c3.metric("Final Monthly Revenue", fmt_try(last_row["Monthly Revenue"]))
    c4.metric("Final Monthly EBITDA Proxy", fmt_try(last_row["Monthly EBITDA Proxy"]))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Final Required CS24", fmt_num(last_row["Required CS24"]))
    c6.metric("Final Cumulative Cash", fmt_try(last_row["Cumulative Cash Balance"]))
    c7.metric("Raw Material Lag", fmt_num(raw_material_required_before_months))
    c8.metric("Stock-In Rate", fmt_pct(stock_in_rate_pct))

    st.markdown("### Annual Summary")
    st.dataframe(annual_fmt, use_container_width=True, hide_index=True)

    st.markdown("### Monthly Trend")
    st.line_chart(monthly_df.set_index("Month")[["Monthly Revenue", "Monthly EBITDA Proxy", "Monthly Net Cash Flow"]])

with tab2:
    st.subheader("Assumptions")
    st.dataframe(assumptions_display_fmt, use_container_width=True, hide_index=True)
    st.caption("Editable rows are changed with arrows. Override rows are changed via text inputs. Calculated rows are formula-driven and read-only.")

with tab3:
    st.subheader("Monthly Financials")
    st.dataframe(dashboard_monthly_fmt, use_container_width=True, hide_index=True)

    st.markdown("### Monthly Revenue / EBITDA / Net Cash Flow")
    st.line_chart(monthly_df.set_index("Month")[["Monthly Revenue", "Monthly EBITDA Proxy", "Monthly Net Cash Flow"]])

with tab4:
    st.subheader("Cash Flow Statement")
    st.dataframe(cash_flow_fmt, use_container_width=True, hide_index=True)

    st.markdown("### Cash Flow Trend")
    st.line_chart(cash_flow_df.set_index("Month")[["Monthly Net Cash Flow", "Cumulative Cash Balance"]])

with tab5:
    st.subheader("Balance Sheet")
    st.dataframe(balance_fmt, use_container_width=True, hide_index=True)

    st.markdown("### Balance Sheet Trend")
    st.line_chart(balance_sheet_df.set_index("Month")[[
        "Cumulative Cash Balance", "Raw Material Inventory",
        "Installed CS24 Asset Value", "Financing Liability", "Equity Proxy"
    ]])

with tab6:
    st.subheader("Personnel Monthly View")
    if personnel_fmt.empty:
        st.info("No personnel data available.")
    else:
        st.dataframe(personnel_fmt, use_container_width=True, hide_index=True)

with tab7:
    st.subheader("Excel Preview")

    if assumptions_df is not None:
        st.markdown("### Assumptions")
        st.dataframe(assumptions_df, use_container_width=True)

    if monthly_inputs_df is not None:
        st.markdown("### Monthly Inputs")
        st.dataframe(monthly_inputs_df, use_container_width=True)

    if macro_df is not None:
        st.markdown("### Macro Assumptions")
        st.dataframe(macro_df, use_container_width=True)

    if opex_df is not None:
        st.markdown("### OPEX")
        st.dataframe(opex_df, use_container_width=True)

    if personnel_base_df is not None:
        st.markdown("### Personnel Base")
        st.dataframe(personnel_base_df, use_container_width=True)

    if personnel_rules_df is not None:
        st.markdown("### Personnel Monthly Rules")
        st.dataframe(personnel_rules_df, use_container_width=True)

    if bom_df is not None:
        st.markdown("### BoM List")
        st.dataframe(bom_df, use_container_width=True)

    if product_capex_df is not None:
        st.markdown("### Product CapEx")
        st.dataframe(product_capex_df, use_container_width=True)
