import streamlit as st
import pandas as pd
from docx import Document
from datetime import datetime
import os

st.title("📊 Weekly Sales Report Generator")

# Section 1: User Inputs
st.header("🔹 1. MTD & Weekly Performance Data")

budget = st.number_input("Monthly Budget (KSH)", value=113998325)
mtd_revenue = st.number_input("MTD Revenue (KSH)", value=93415418)
weekly_budget = st.number_input("Weekly Budget (KSH)", value=26479125)
current_week_revenue = st.number_input("Current Week Revenue (KSH)", value=20943811)
previous_week_revenue = st.number_input("Previous Week Revenue (KSH)", value=20353938)
short_supplies = st.number_input("Short Supplies (KSH)", value=1266460)
returns = st.number_input("Returns (KSH)", value=193615)

st.header("🔹 2. Revenue Projection Inputs")
historical_trend = st.number_input("Historical Trend (KSH)", value=89936015)
linear_extrap = st.number_input("Linear Extrapolation (KSH)", value=99857861)
blended_estimate = st.number_input("Blended Conservative Estimate (KSH)", value=93904753)

highlight_may_25 = st.checkbox("✅ May 25 sales exceeded trend", value=True)
parmesan_price_increase = st.checkbox("✅ Due to Parmesan price increase", value=True)

st.header("🔹 3. Upload Supplementary Data")
short_supply_file = st.file_uploader("Upload Top 10 Short Supplied Items (Excel)", type=["xlsx"])
market_return_file = st.file_uploader("Upload Top 10 Market Returns (Excel)", type=["xlsx"])

# Calculations
revenue_gap = budget - mtd_revenue
achievement_pct = (mtd_revenue / budget) * 100
weekly_variance = current_week_revenue - weekly_budget
growth_rate = ((current_week_revenue - previous_week_revenue) / previous_week_revenue) * 100
closing_pct = (blended_estimate / budget) * 100

# Generate Report
if st.button("📝 Generate Report"):
    doc = Document()
    doc.add_heading("Week 22 Sales Report", 0)
    doc.add_paragraph("Generated on: " + datetime.now().strftime("%Y-%m-%d %H:%M"))

    doc.add_heading("MTD Sales Revenue Update", level=1)
    doc.add_paragraph(f"Budget: KSH {budget:,.0f}")
    doc.add_paragraph(f"MTD Revenue: KSH {mtd_revenue:,.0f}")
    doc.add_paragraph(f"Achievement vs Budget: {achievement_pct:.0f}%")
    doc.add_paragraph(f"Revenue Gap to Budget: KSH {revenue_gap:,.0f} ({100 - achievement_pct:.0f}%)")
    doc.add_paragraph(f"Closing Revenue Estimate: KSH {blended_estimate:,.0f} ({closing_pct:.0f}% of Budget)")

    doc.add_heading("Current Week Performance", level=1)
    doc.add_paragraph(f"Weekly Budget: KSH {weekly_budget:,.0f}")
    doc.add_paragraph(f"Current Week Revenue: KSH {current_week_revenue:,.0f}")
    doc.add_paragraph(f"Variance to Budget: KSH {weekly_variance:,.0f} ({(weekly_variance / weekly_budget * 100):+.0f}%)")
    doc.add_paragraph(f"Previous Week Revenue: KSH {previous_week_revenue:,.0f}")
    doc.add_paragraph(f"Growth Rate: {growth_rate:.2f}%")

    doc.add_heading("Operational Insights", level=1)
    doc.add_paragraph(f"Short Supplies: KSH {short_supplies:,.0f} (~{(short_supplies / current_week_revenue * 100):.1f}% impact)")
    doc.add_paragraph(f"Returns: KSH {returns:,.0f} (~{(returns / current_week_revenue * 100):.1f}% impact)")

    doc.add_heading("Key Highlights", level=1)
    doc.add_paragraph(f"Week 22 showed a {growth_rate:+.2f}% growth over Week 21, though still under budget.")
    doc.add_paragraph(f"MTD Revenue is now at {achievement_pct:.0f}% of budget with KSH {revenue_gap:,.0f} to go.")

    if highlight_may_25:
        doc.add_paragraph("• May 25 sales exceeded the historical trend.", style="List Bullet")
        if parmesan_price_increase:
            doc.add_paragraph("• This was likely due to an increase in Parmesan price.", style="List Bullet")

    doc.add_heading("Closing Estimates Summary", level=1)
    table = doc.add_table(rows=1, cols=2)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Scenario'
    hdr_cells[1].text = 'Estimate (KSH)'
    data_rows = [
        ('Historical Trend', historical_trend),
        ('Linear Extrapolation', linear_extrap),
        ('Blended Estimate', blended_estimate)
    ]
    for label, value in data_rows:
        row_cells = table.add_row().cells
        row_cells[0].text = label
        row_cells[1].text = f"{value:,.0f}"

    if short_supply_file:
        doc.add_heading("Top 10 Short Supplied Items", level=1)
        df_short = pd.read_excel(short_supply_file)
        doc.add_paragraph(df_short.to_string(index=False))

    if market_return_file:
        doc.add_heading("Top 10 Market Returns", level=1)
        df_returns = pd.read_excel(market_return_file)
        doc.add_paragraph(df_returns.to_string(index=False))

    # Save and Download
    filename = f"Week22_Sales_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    filepath = os.path.join(".", filename)
    doc.save(filepath)

    with open(filepath, "rb") as file:
        st.download_button("📥 Download Word Report", file, file_name=filename)

    st.success("Report generated and ready to download!")

