"""
Startup Runway Calculator — Free Web Tool by ClearMetric
https://clearmetric.gumroad.com

Helps founders calculate how long their money lasts and plan fundraising timing.
"""

import streamlit as st
import plotly.graph_objects as go
import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Startup Runway Calculator — ClearMetric",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Custom CSS — Navy theme
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .main .block-container { padding-top: 2rem; max-width: 1200px; }
    .stMetric { background: #f0f4f8; border-radius: 8px; padding: 12px; border-left: 4px solid #1e3a5f; }
    h1 { color: #1e3a5f; }
    h2, h3 { color: #2c3e50; }
    .runway-number { font-size: 2.5rem; font-weight: bold; color: #1e3a5f; text-align: center; }
    .alive { color: #27ae60; font-weight: bold; }
    .dead { color: #e74c3c; font-weight: bold; }
    .cta-box {
        background: linear-gradient(135deg, #1e3a5f 0%, #2c5282 100%);
        color: white; padding: 24px; border-radius: 12px; text-align: center;
        margin: 20px 0;
    }
    .cta-box a { color: #90cdf4; text-decoration: none; font-weight: bold; font-size: 1.1rem; }
    div[data-testid="stSidebar"] { background: #f0f4f8; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown("# 🚀 Startup Runway Calculator")
st.markdown("**How long does your money last?** — Plan fundraising, model hires, and know when you'll break even.")
st.markdown("---")

# ---------------------------------------------------------------------------
# Sidebar — User inputs
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## Your Startup Finances")

    st.markdown("### Cash & Revenue")
    current_cash = st.number_input(
        "Current cash in bank ($)",
        value=500_000,
        min_value=0,
        step=10_000,
        format="%d",
    )
    monthly_revenue = st.number_input(
        "Monthly revenue ($)",
        value=10_000,
        min_value=0,
        step=1_000,
        format="%d",
    )
    revenue_growth_pct = st.number_input(
        "Monthly revenue growth rate (%)",
        value=10.0,
        min_value=0.0,
        max_value=100.0,
        step=1.0,
        format="%.1f",
    )

    st.markdown("### Monthly Burn (by category)")
    salaries = st.number_input("Salaries & payroll ($)", value=30_000, min_value=0, step=1_000, format="%d")
    office = st.number_input("Office/workspace ($)", value=3_000, min_value=0, step=500, format="%d")
    software = st.number_input("Software & tools ($)", value=2_000, min_value=0, step=500, format="%d")
    marketing = st.number_input("Marketing ($)", value=5_000, min_value=0, step=500, format="%d")
    legal = st.number_input("Legal & accounting ($)", value=1_500, min_value=0, step=500, format="%d")
    other = st.number_input("Other expenses ($)", value=2_000, min_value=0, step=500, format="%d")

    st.markdown("### Planned Hires")
    num_hires = st.number_input("Number of new hires", value=2, min_value=0, step=1)
    avg_salary = st.number_input("Avg salary each ($/mo)", value=6_000, min_value=0, step=500, format="%d")
    hire_start_month = st.number_input("Start month", value=4, min_value=0, step=1)

    st.markdown("### One-time & Fundraising")
    one_time = st.number_input("One-time expenses ($)", value=0, min_value=0, step=5_000, format="%d")
    target_raise = st.number_input("Target raise amount ($)", value=0, min_value=0, step=50_000, format="%d")
    close_month = st.number_input("Expected close month", value=0, min_value=0, step=1)

# ---------------------------------------------------------------------------
# Core calculations
# ---------------------------------------------------------------------------
total_burn = salaries + office + software + marketing + legal + other
net_burn = total_burn - monthly_revenue
revenue_growth = revenue_growth_pct / 100

# Gross margin (simplified: revenue - COGS; for SaaS often high; we use revenue vs expenses)
gross_margin = (monthly_revenue - 0) / monthly_revenue * 100 if monthly_revenue > 0 else 0

# Month-by-month projection (36 months)
def project_months(
    cash_start,
    rev,
    rev_growth,
    burn_base,
    num_hires_val,
    salary_per_hire,
    hire_month,
    one_time_val,
    raise_amt,
    close_m,
    include_hires=True,
    include_fundraising=True,
):
    months = 36
    cash = cash_start
    rows = []
    burn = burn_base
    for m in range(1, months + 1):
        rev_m = rev * (1 + rev_growth) ** (m - 1)
        if include_hires and m >= hire_month and num_hires_val > 0:
            burn = burn_base + num_hires_val * salary_per_hire
        if include_fundraising and m == close_m and raise_amt > 0:
            cash += raise_amt
        if m == 1 and one_time_val > 0:
            cash -= one_time_val
        net = rev_m - burn
        cash += net
        rows.append({
            "Month": m,
            "Cash": max(0, cash),
            "Revenue": rev_m,
            "Expenses": burn,
            "Net": net,
        })
    return pd.DataFrame(rows)

# Current runway (no hires, no fundraising)
df_base = project_months(
    current_cash, monthly_revenue, revenue_growth, total_burn,
    num_hires, avg_salary, hire_start_month, one_time, 0, 0,
    include_hires=False, include_fundraising=False,
)
runway_months_base = None
for i, row in df_base.iterrows():
    if row["Cash"] <= 0:
        runway_months_base = int(row["Month"])
        break
if runway_months_base is None:
    runway_months_base = 36

# Runway with hires
df_hires = project_months(
    current_cash, monthly_revenue, revenue_growth, total_burn,
    num_hires, avg_salary, hire_start_month, one_time, 0, 0,
    include_hires=True, include_fundraising=False,
)
runway_months_hires = None
for i, row in df_hires.iterrows():
    if row["Cash"] <= 0:
        runway_months_hires = int(row["Month"])
        break
if runway_months_hires is None:
    runway_months_hires = 36

# Runway with fundraising
df_fundraise = project_months(
    current_cash, monthly_revenue, revenue_growth, total_burn,
    num_hires, avg_salary, hire_start_month, one_time, target_raise, close_month,
    include_hires=True, include_fundraising=True,
)
runway_months_fundraise = None
for i, row in df_fundraise.iterrows():
    if row["Cash"] <= 0:
        runway_months_fundraise = int(row["Month"])
        break
if runway_months_fundraise is None:
    runway_months_fundraise = 36

# Break-even month
break_even_month = None
for m in range(1, 37):
    rev_m = monthly_revenue * (1 + revenue_growth) ** (m - 1)
    burn_m = total_burn
    if m >= hire_start_month and num_hires > 0:
        burn_m += num_hires * avg_salary
    if rev_m >= burn_m:
        break_even_month = m
        break

# Default Alive analysis (Paul Graham)
# "Default alive" = if current revenue growth and expense trends continue, will you reach profitability?
# Simplified: compare revenue trajectory to expense trajectory
rev_at_12 = monthly_revenue * (1 + revenue_growth) ** 11
burn_at_12 = total_burn + (num_hires * avg_salary if hire_start_month <= 12 else 0)
default_alive = rev_at_12 >= burn_at_12

# When to start fundraising (6 months before cash runs out)
runway_used = runway_months_base
fundraise_start_month = max(1, runway_used - 6) if runway_used < 36 else 0

# Critical cash threshold (e.g., 3 months of burn)
critical_cash = total_burn * 3

# Scenario toggle
scenario = st.radio(
    "Scenario",
    ["Current (no hires, no fundraising)", "With planned hires", "With fundraising"],
    horizontal=True,
)

if scenario == "Current (no hires, no fundraising)":
    df_display = df_base
    runway_display = runway_months_base
elif scenario == "With planned hires":
    df_display = df_hires
    runway_display = runway_months_hires
else:
    df_display = df_fundraise
    runway_display = runway_months_fundraise

# ---------------------------------------------------------------------------
# Display — Key Metrics
# ---------------------------------------------------------------------------
st.markdown("## Key Metrics")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Runway (months)", f"{runway_display}", help="Months until cash runs out")
c2.metric("Net burn ($/mo)", f"${net_burn:,.0f}", help="Monthly expenses minus revenue")
c3.metric("Break-even month", f"{break_even_month}" if break_even_month else "N/A", help="When revenue >= expenses")
c4.metric("Monthly revenue growth", f"{revenue_growth_pct:.1f}%", help="Assumed monthly growth rate")

st.markdown("---")

# ---------------------------------------------------------------------------
# Area chart — Cash balance over time (with vs without fundraising)
# ---------------------------------------------------------------------------
st.markdown("## Cash Balance Over Time")

fig_cash = go.Figure()
fig_cash.add_trace(go.Scatter(
    x=df_base["Month"],
    y=df_base["Cash"],
    fill="tozeroy",
    name="Without fundraising",
    line=dict(color="#1e3a5f", width=2),
    fillcolor="rgba(30, 58, 95, 0.3)",
))
fig_cash.add_trace(go.Scatter(
    x=df_fundraise["Month"],
    y=df_fundraise["Cash"],
    fill="tozeroy",
    name="With fundraising",
    line=dict(color="#2c5282", width=2),
    fillcolor="rgba(44, 82, 130, 0.3)",
))
fig_cash.add_hline(y=0, line_dash="dash", line_color="#e74c3c", line_width=1)
fig_cash.add_hline(y=critical_cash, line_dash="dot", line_color="#f39c12", line_width=1,
                   annotation_text=f"Critical: ${critical_cash:,.0f}")
fig_cash.update_layout(
    title="Cash Balance — With vs Without Fundraising",
    xaxis_title="Month",
    yaxis_title="Cash ($)",
    yaxis_tickformat="$,.0f",
    height=400,
    template="plotly_white",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    margin=dict(t=60, b=40),
)
st.plotly_chart(fig_cash, use_container_width=True)

# ---------------------------------------------------------------------------
# Stacked bar — Monthly expense breakdown
# ---------------------------------------------------------------------------
st.markdown("## Monthly Expense Breakdown")

expense_cats = {
    "Salaries & payroll": salaries,
    "Office/workspace": office,
    "Software & tools": software,
    "Marketing": marketing,
    "Legal & accounting": legal,
    "Other": other,
}
colors_exp = ["#1e3a5f", "#2c5282", "#3182ce", "#6baed6", "#9ecae1", "#c6dbef"]

fig_exp = go.Figure()
fig_exp.add_trace(go.Bar(
    name="Expenses",
    x=list(expense_cats.keys()),
    y=list(expense_cats.values()),
    marker_color=colors_exp,
    text=[f"${v:,.0f}" for v in expense_cats.values()],
    textposition="outside",
))
fig_exp.update_layout(
    title="Monthly Burn by Category",
    xaxis_title="Category",
    yaxis_title="Amount ($)",
    yaxis_tickformat="$,.0f",
    height=350,
    template="plotly_white",
    showlegend=False,
    margin=dict(t=40, b=80),
)
st.plotly_chart(fig_exp, use_container_width=True)

# ---------------------------------------------------------------------------
# Line chart — Revenue vs Expenses trajectory (with break-even marker)
# ---------------------------------------------------------------------------
st.markdown("## Revenue vs Expenses Trajectory")

months_arr = np.arange(1, 37)
rev_traj = monthly_revenue * (1 + revenue_growth) ** (months_arr - 1)
exp_traj = np.full(36, total_burn)
for m in range(36):
    if m + 1 >= hire_start_month and num_hires > 0:
        exp_traj[m] = total_burn + num_hires * avg_salary

fig_rev = go.Figure()
fig_rev.add_trace(go.Scatter(x=months_arr, y=rev_traj, mode="lines", name="Revenue", line=dict(color="#27ae60", width=3)))
fig_rev.add_trace(go.Scatter(x=months_arr, y=exp_traj, mode="lines", name="Expenses", line=dict(color="#e74c3c", width=3)))
if break_even_month:
    fig_rev.add_vline(x=break_even_month, line_dash="dash", line_color="#f39c12", line_width=2,
                      annotation_text=f"Break-even: Month {break_even_month}")
fig_rev.update_layout(
    title="Revenue vs Expenses Over Time",
    xaxis_title="Month",
    yaxis_title="Amount ($)",
    yaxis_tickformat="$,.0f",
    height=400,
    template="plotly_white",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    margin=dict(t=60, b=40),
)
st.plotly_chart(fig_rev, use_container_width=True)

# ---------------------------------------------------------------------------
# Default Alive verdict
# ---------------------------------------------------------------------------
st.markdown("## Default Alive?")
st.caption("Paul Graham's framework: if revenue growth and expense trends continue, will you reach profitability?")

if default_alive:
    st.success(f"""
    **✅ Default Alive** — At current trends, your revenue will exceed expenses by month 12.
    You're on a path to profitability without needing to change course.
    """)
else:
    st.error(f"""
    **❌ Default Dead** — At current trends, revenue won't cover expenses by month 12.
    You'll need to either: grow revenue faster, cut burn, or raise capital before runway runs out.
    """)

# Key dates
st.markdown("### Key Dates")
st.markdown(f"- **Start fundraising by:** Month {fundraise_start_month} (6 months before cash runs out)")
st.markdown(f"- **Critical cash threshold:** ${critical_cash:,.0f} (3 months of burn)")
if break_even_month:
    st.markdown(f"- **Break-even:** Month {break_even_month}")

st.markdown("---")

# ---------------------------------------------------------------------------
# Year-by-year table (expandable)
# ---------------------------------------------------------------------------
with st.expander("📊 Month-by-Month Projection", expanded=False):
    df_show = df_display.copy()
    df_show["Cumulative Net"] = df_show["Net"].cumsum()
    st.dataframe(
        df_show.style.format({
            "Cash": "${:,.0f}",
            "Revenue": "${:,.0f}",
            "Expenses": "${:,.0f}",
            "Net": "${:,.0f}",
            "Cumulative Net": "${:,.0f}",
        }),
        use_container_width=True,
        height=400,
    )

# ---------------------------------------------------------------------------
# CTA — Excel version
# ---------------------------------------------------------------------------
st.markdown("---")
st.markdown("""
<div class="cta-box">
    <h3 style="color: white; margin: 0 0 8px 0;">Want the Full Runway Spreadsheet?</h3>
    <p style="margin: 0 0 16px 0;">
        Get the <strong>ClearMetric Startup Runway Calculator</strong> — a downloadable Excel template with:<br>
        ✓ 36-month projection table (editable)<br>
        ✓ 3-scenario comparison (current, lean, aggressive growth)<br>
        ✓ Default alive verdict & break-even analysis<br>
        ✓ Print-ready for investor updates<br>
    </p>
    <a href="https://clearmetric.gumroad.com" target="_blank">
        Get It on Gumroad — $14.99 →
    </a>
</div>
""", unsafe_allow_html=True)

# Cross-sell
st.markdown("### More from ClearMetric")
cx1, cx2, cx3 = st.columns(3)
with cx1:
    st.markdown("""
    **📊 Budget Planner** — $13.99
    Track income, expenses, savings with the 50/30/20 framework.
    [Get it →](https://clearmetric.gumroad.com)
    """)
with cx2:
    st.markdown("""
    **🔥 FIRE Calculator** — $14.99
    Find your FIRE number, compare scenarios, plan early retirement.
    [Get it →](https://clearmetric.gumroad.com)
    """)
with cx3:
    st.markdown("""
    **📈 Stock Portfolio Tracker** — $17.99
    Track stocks, dividends, sector allocation, performance.
    [Get it →](https://clearmetric.gumroad.com)
    """)

# Footer
st.markdown("---")
st.caption("© 2026 ClearMetric | [clearmetric.gumroad.com](https://clearmetric.gumroad.com) | "
           "This tool is for educational purposes only. Not financial advice.")
