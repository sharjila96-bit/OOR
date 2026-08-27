import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="OOR Auto Report", layout="wide")
st.title("📦 OOR Pickup - Auto Report Generator")

file = st.file_uploader("Yaha .xls file upload karo", type=['xls','xlsx'])

if file:
    xls = pd.ExcelFile(file)
    main_sheet = max(xls.sheet_names, key=lambda s: len(xls.parse(s)))
    df = pd.read_excel(file, sheet_name=main_sheet)
    st.success(f"Loaded: {len(df)} shipments")
    station_wise = df.groupby('station_code').agg(Total=('tracking_id','count'), RVP=('is_rvp_enabled', lambda x: (x=='RVP').sum()), NON_RVP=('is_rvp_enabled', lambda x: (x=='NON RVP').sum())).reset_index().sort_values('Total', ascending=False)
    ctl_wise = df.groupby(['CTL','station_code']).size().reset_index(name='Total').sort_values('Total', ascending=False)
    st.dataframe(station_wise)
    st.dataframe(ctl_wise)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        station_wise.to_excel(writer, sheet_name='Station Wise', index=False)
        ctl_wise.to_excel(writer, sheet_name='CTL Wise', index=False)
        df.to_excel(writer, sheet_name='Raw', index=False)
    st.download_button("📥 Final Report Download", output.getvalue(), file_name="Final_OOR_Report.xlsx")
