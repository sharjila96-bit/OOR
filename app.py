import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="OOR Report", layout="wide")
st.title("📦 OOR Pickup - Auto Report Generator")
st.write("Yaha .xls file upload karo")

file = st.file_uploader("File chuno")

if file:
    try:
        # Try all engines
        try:
            df = pd.read_excel(file, engine='xlrd')
        except:
            file.seek(0)
            df = pd.read_excel(file, engine='openpyxl')
        
        st.success(f"Loaded: {len(df)} rows, {len(df.columns)} cols")
        st.write("Columns:", list(df.columns)[:20])
        st.dataframe(df.head())

        # Try to make report if columns exist
        if 'station_code' in df.columns:
            station_wise = df.groupby('station_code').size().reset_index(name='Total').sort_values('Total', ascending=False)
            st.subheader("Station Wise")
            st.dataframe(station_wise)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                station_wise.to_excel(writer, sheet_name='Station Wise', index=False)
                df.to_excel(writer, sheet_name='Raw', index=False)
            st.download_button("📥 Final Report Download", output.getvalue(), file_name="Final_OOR_Report.xlsx")
        else:
            st.warning("station_code column nahi mila. Upar columns dekho, naam alag ho sakta hai.")
            
    except Exception as e:
        st.error(f"Error: {e}")
