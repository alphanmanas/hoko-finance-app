import streamlit as st
import pandas as pd

st.set_page_config(layout="wide")

st.title("HOKO Mobility Financial Copilot v2")

# ---------------------
# PARAMETERS
# ---------------------

daily_km = st.sidebar.number_input("Km/day", value=150)
wh_per_km = st.sidebar.number_input("Wh/km", value=40)
sell_price = st.sidebar.number_input("Sell price (TL/kWh)", value=15.0)
buy_price = st.sidebar.number_input("Buy price (TL/kWh)", value=3.5)
company_share = st.sidebar.slider("Company share", 0.0, 1.0, 0.8)

cycle = st.sidebar.number_input("Cycle", value=2.0)

battery_kwh = 3.24
slots = 24
soc = 0.85

# ---------------------
# SALES SCHEDULE
# ---------------------

sales = {}

for m in range(1, 6):
    sales[m] = 0

sales[6] = 2000
sales[7] = 2500
sales[8] = 3000
sales[9] = 3000
sales[10] = 3250
sales[11] = 3250
sales[12] = 3250

for m in range(13, 19):
    sales[m] = 3500

for m in range(19, 49):
    sales[m] = 4000

# ---------------------
# CALCULATION
# ---------------------

data = []

active = 0

daily_kwh_per_motor = daily_km * wh_per_km / 1000
cs_capacity = slots * battery_kwh
usable = cs_capacity * soc
daily_cs_kwh = usable * cycle

margin = (sell_price - buy_price) * company_share

for m in range(1, 49):

    sold = sales[m]
    active += sold

    daily_kwh = active * daily_kwh_per_motor
    monthly_kwh = daily_kwh * 30

    revenue = monthly_kwh * sell_price
    cost = monthly_kwh * buy_price
    profit = monthly_kwh * margin

    cs_needed = int(daily_kwh / daily_cs_kwh) + 1

    data.append([
        m, sold, active,
        daily_kwh, monthly_kwh,
        revenue, cost, profit,
        cs_needed
    ])

df = pd.DataFrame(data, columns=[
    "Month", "Sold", "Active",
    "Daily kWh", "Monthly kWh",
    "Revenue", "Cost", "Profit",
    "CS24 Needed"
])

# ---------------------
# TABS
# ---------------------

tab1, tab2 = st.tabs(["Dashboard", "Cash Flow"])

# ---------------------
# DASHBOARD
# ---------------------

with tab1:

    st.subheader("Key Metrics")

    col1, col2, col3 = st.columns(3)

    col1.metric("Active Motors (Final)", df["Active"].iloc[-1])
    col2.metric("Monthly kWh (Final)", f"{df['Monthly kWh'].iloc[-1]:,.0f}")
    col3.metric("Monthly Profit (Final)", f"₺{df['Profit'].iloc[-1]:,.0f}")

# ---------------------
# CASH FLOW
# ---------------------

with tab2:

    st.subheader("Monthly Cash Flow")

    st.dataframe(df)

    st.line_chart(df.set_index("Month")[["Profit"]])
