import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl

# -------------------------------
# Replace with your GitHub raw CSV URL
url = "https://raw.githubusercontent.com/scswork/multiplier_app/refs/heads/main/multiplier_2021_data.csv"
# -------------------------------

# Load dataset
df = pd.read_csv(url)

st.title("Economic Impact Report")

# Sidebar layout
st.sidebar.header("Filters")

# GEO filter (multiple select)
geo_filter = st.sidebar.multiselect("Select Provinces (GEO)", options=sorted(df['GEO'].unique()), default=[])

# Geographical coverage filter (single select)
coverage_filter = st.sidebar.selectbox("Select Geographical Coverage", options=["All"] + sorted(df['Geographical coverage'].unique()))

# CAPEX and OPEX dropdowns
capex_industry = st.sidebar.selectbox("Select CAPEX Industry", options=df['Industry'].unique())
opex_industry = st.sidebar.selectbox("Select OPEX Industry", options=df['Industry'].unique())

# Type and Variable filters
type_filter = st.sidebar.multiselect("Filter by Type", options=df['Multiplier type'].unique(), default=list(df['Multiplier type'].unique()))
variable_filter = st.sidebar.multiselect("Filter by Variable", options=df['Variable'].unique(), default=list(df['Variable'].unique()))

st.sidebar.markdown("---")
st.sidebar.subheader("Enter 15-Year Investment Data")

# Editable table for CAPEX/OPEX values
years = list(range(1, 16))
initial_data = pd.DataFrame({
    "Year": years,
    "Value": [0]*15,
    "Type": ["CAPEX"]*15
})

edited_data = st.sidebar.data_editor(initial_data, num_rows="dynamic", use_container_width=True)

# Main panel
st.subheader("Impact Report")

if st.button("Generate Report"):
    # Convert Value to numeric
    edited_data['Value'] = pd.to_numeric(edited_data['Value'], errors='coerce').fillna(0)

    # Calculate totals
    capex_total = edited_data.loc[edited_data['Type'] == "CAPEX", 'Value'].sum()
    opex_total = edited_data.loc[edited_data['Type'] == "OPEX", 'Value'].sum()

    # Apply GEO and coverage filters
    filtered_df = df.copy()
    if geo_filter:
        filtered_df = filtered_df[filtered_df['GEO'].isin(geo_filter)]
    if coverage_filter != "All":
        filtered_df = filtered_df[filtered_df['Geographical coverage'] == coverage_filter]

    # Filter database for selected industries
    capex_db = filtered_df[filtered_df['Industry'] == capex_industry]
    opex_db = filtered_df[filtered_df['Industry'] == opex_industry]

    # Merge CAPEX and OPEX data
    report = pd.merge(capex_db, opex_db, on=['Multiplier type', 'Variable'], suffixes=('_CAPEX', '_OPEX'))
    report = report[(report['Multiplier type'].isin(type_filter)) & (report['Variable'].isin(variable_filter))]

    # Impact calculations
    def calculate_impact(row, total, value_col):
        if "Jobs" in row['Variable']:
            return (row[value_col] * (total / 1e6))
        else:
            return row[value_col] * total

    report['CAPEX_Impact'] = report.apply(lambda r: calculate_impact(r, capex_total, 'VALUE_CAPEX'), axis=1)
    report['OPEX_Impact'] = report.apply(lambda r: calculate_impact(r, opex_total, 'VALUE_OPEX'), axis=1)

    # Format values
    report['VALUE_CAPEX'] = report['VALUE_CAPEX'].round(4)
    report['VALUE_OPEX'] = report['VALUE_OPEX'].round(4)
    report['CAPEX_Impact'] = report.apply(lambda r: f"{int(r['CAPEX_Impact']):,}" if "Jobs" in r['Variable'] else f"${int(r['CAPEX_Impact']):,}", axis=1)
    report['OPEX_Impact'] = report.apply(lambda r: f"{int(r['OPEX_Impact']):,}" if "Jobs" in r['Variable'] else f"${int(r['OPEX_Impact']):,}", axis=1)

    st.dataframe(report)

    # Download Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        report.to_excel(writer, index=False, sheet_name='Impact Report')
    st.download_button(
        label="Download Excel Report",
        data=output.getvalue(),
        file_name="impact_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
