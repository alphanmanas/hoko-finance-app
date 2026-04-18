import streamlit as st
import math

st.set_page_config(page_title="HOKO Mobility Financial Copilot", layout="wide")

st.title("HOKO Mobility Financial Copilot")

# -------------------------
# SIDEBAR INPUTS
# -------------------------

st.sidebar.header("Parametreler")

motors = st.sidebar.number_input("Motor Sayısı", value=6000)
daily_km = st.sidebar.number_input("Km / Gün", value=150)
wh_per_km = st.sidebar.number_input("Wh / km", value=40)

sell_price = st.sidebar.number_input("Satış Fiyatı (TL/kWh)", value=15.0)
buy_price = st.sidebar.number_input("Elektrik Alış (TL/kWh)", value=3.5)

company_share = st.sidebar.slider("Bizim Pay (%)", 0, 100, 80) / 100

cycle = st.sidebar.number_input("Cycle (CS24)", value=2.0)

battery_kwh = 3.24
slots = 24
soc = 0.85

usd_try = st.sidebar.number_input("USD/TRY", value=45.0)
cs_cost_usd = st.sidebar.number_input("CS24 Donanım ($)", value=5000.0)
battery_cost = st.sidebar.number_input("Battery $/kWh", value=200.0)

opex = st.sidebar.number_input("Yıllık OPEX (TL)", value=20000)

motor_profit = st.sidebar.number_input("Motor Brüt Kâr (TL)", value=80000)

# -------------------------
# HESAPLAR
# -------------------------

# Motor enerji
daily_kwh_per_motor = daily_km * wh_per_km / 1000
total_daily_kwh = motors * daily_kwh_per_motor
annual_kwh = total_daily_kwh * 365

# CS24 kapasite
cs_capacity = slots * battery_kwh
usable_per_cycle = cs_capacity * soc
daily_cs_kwh = usable_per_cycle * cycle

cs_needed = math.ceil(total_daily_kwh / daily_cs_kwh)

# Marj
margin = sell_price - buy_price
company_margin = margin * company_share

# Gelir
annual_energy_profit = annual_kwh * company_margin

# CapEx
battery_total_usd = cs_capacity * battery_cost
cs_total_try = (cs_cost_usd + battery_total_usd) * usd_try

down_payment = cs_total_try * 0.25
our_investment = cs_total_try * 0.75

# CS gelir
annual_cs_kwh = daily_cs_kwh * 365
cs_profit = annual_cs_kwh * company_margin

cs_net = cs_profit - opex

roi = cs_net / our_investment
payback = our_investment / cs_net

# IRR approx
irr = roi * 0.9  # basit approx

# Motor gelir
motor_total = motors * motor_profit

# TOPLAM
total_profit = motor_total + annual_energy_profit

# -------------------------
# OUTPUT
# -------------------------

col1, col2, col3, col4 = st.columns(4)

col1.metric("CS24 Sayısı", cs_needed)
col2.metric("Yıllık kWh", f"{annual_kwh:,.0f}")
col3.metric("Enerji Kârı", f"₺{annual_energy_profit:,.0f}")
col4.metric("Toplam Kâr", f"₺{total_profit:,.0f}")

st.divider()

st.subheader("CS24 Ekonomi")

col1, col2, col3 = st.columns(3)

col1.metric("CS24 ROI", f"{roi*100:.1f}%")
col2.metric("Payback (yıl)", f"{payback:.2f}")
col3.metric("IRR (approx)", f"{irr*100:.1f}%")

st.divider()

st.subheader("Detaylar")

st.write(f"Günlük enerji: {total_daily_kwh:,.0f} kWh")
st.write(f"CS24 başı günlük enerji: {daily_cs_kwh:,.0f} kWh")
st.write(f"Toplam CS24: {cs_needed}")

st.write(f"CS24 yatırım: ₺{cs_total_try:,.0f}")
st.write(f"Bizim yatırım: ₺{our_investment:,.0f}")

st.write(f"CS24 yıllık net: ₺{cs_net:,.0f}")