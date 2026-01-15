import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Data Mapper", layout="wide")

st.markdown("### 📊 Excel Data Mapping Tool")

# Sidebar එකේ file upload කරන්න ඉඩ ලබා දීම
st.sidebar.header("Upload Files")
main_file = st.sidebar.file_uploader("Upload Copy File (Main.xlsx)", type=['xlsx'])
trans_file = st.sidebar.file_uploader("Upload Paste File (Transport.xlsx)", type=['xlsx'])

if main_file and trans_file:
    df_main = pd.read_excel(main_file)
    df_trans = pd.read_excel(trans_file)
    
    main_cols = df_main.columns.tolist()
    trans_cols = df_trans.columns.tolist()

    st.write("---")
    st.markdown("#### Select & Map Columns")
    st.caption("Select the column to be copied on the left and select which column on the right it should be pasted into.")

    # Table එකක් වගේ පෙනෙන්නට Header එකක් හදමු
    h1, h2, h3 = st.columns([0.5, 2, 2])
    h1.write("**Select**")
    h2.write("**Copy From (Main.xlsx)**")
    h3.write("**Paste To (Transport.xlsx)**")

    mapping_results = {}

    # Main file එකේ තීරු ලැයිස්තුව පෙන්වීම
    for i, m_col in enumerate(main_cols):
        c1, c2, c3 = st.columns([0.5, 2, 2])
        
        with c1:
            is_selected = st.checkbox("", key=f"check_{i}")
        
        with c2:
            st.info(f"{m_col}") # Copy කරන column එකේ නම
            
        with c3:
            if is_selected:
                # දකුණු පැත්තේ ඇති Heading එක තෝරාගැනීම
                target_col = st.selectbox(
                    f"Map {m_col} to:", 
                    options=trans_cols, 
                    key=f"target_{i}",
                    label_visibility="collapsed"
                )
                mapping_results[m_col] = target_col
            else:
                st.text_input("Not Selected", disabled=True, key=f"dis_{i}", label_visibility="collapsed")

    st.write("---")
    
    # Paste Button
    if st.button("🚀 Execute Paste Operation", use_container_width=True):
        if mapping_results:
            try:
                # දත්ත මාරු කිරීම (Heading වෙනස් නොවේ)
                for m_col, t_col in mapping_results.items():
                    df_trans[t_col] = df_main[m_col]
                
                # ගොනුව සකස් කිරීම
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_trans.to_excel(writer, index=False)
                
                st.success("Data transfer successful! Download the file using the button below.")
                st.download_button("📥 Download Updated Transport File", output.getvalue(), "Updated_Transport.xlsx")
            except Exception as e:
                st.error(f"Error: {e}")
        else:
            st.warning("Please select at least one column.")

else:
    st.info("To get started, upload the two Excel files from the sidebar on the left.")