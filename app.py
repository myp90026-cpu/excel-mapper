import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Smart Header Mapper", layout="wide")
st.title("Excel Smart Header Detection Tool")

# Sidebar - File Upload
st.sidebar.header("📂 Upload Files")
main_file = st.sidebar.file_uploader("Upload Main.xlsx (Copy)", type=['xlsx'])
trans_file = st.sidebar.file_uploader("Upload Transport.xlsx (Paste)", type=['xlsx'])

def get_df_with_correct_header(file, label):
    """මුල් පේළි 5 පරීක්ෂා කර Heading එක සොයා දෙන Function එක"""
    xl = pd.ExcelFile(file)
    sheet = st.selectbox(f"Select Sheet for {label}", xl.sheet_names, key=f"sheet_{label}")
    
    # මුල් පේළි 10 කියවා බලන්න (Header එක කොහේ තිබුණත් සොයා ගැනීමට)
    preview_df = pd.read_excel(file, sheet_name=sheet, header=None, nrows=10)
    
    # පේළි අංකය තෝරා ගැනීමට (Default එක විදියට 0 දෙනවා, නමුත් ඔබට වෙනස් කළ හැක)
    header_row = st.number_input(f"{label} හි Heading එක ඇති පේළිය (Row Index 0 සිට)", 
                                 min_value=0, max_value=10, value=0, key=f"row_{label}")
    
    # තෝරාගත් පේළිය අනුව Dataframe එක නැවත කියවීම
    df = pd.read_excel(file, sheet_name=sheet, header=header_row)
    return df

col_a, col_b = st.columns(2)

with col_a:
    if main_file:
        st.subheader("1. Source Settings")
        df_main = get_df_with_correct_header(main_file, "Main File")
        st.write("Headers found:", list(df_main.columns))

with col_b:
    if trans_file:
        st.subheader("2. Destination Settings")
        df_trans = get_df_with_correct_header(trans_file, "Transport File")
        st.write("Headers found:", list(df_trans.columns))

st.divider()

# දෙකම තෝරාගෙන තිබේ නම් පමණක් ඉදිරියට
if main_file and trans_file:
    # --- Filter Section ---
    st.subheader("🎯 3. Row Filter")
    enable_filter = st.checkbox("දත්ත Filter කිරීමට මෙය Tick කරන්න")
    
    active_df = df_main.copy()
    if enable_filter:
        f_col1, f_col2 = st.columns(2)
        with f_col1:
            filter_col = st.selectbox("Filter කළ යුතු Column එක", df_main.columns)
        with f_col2:
            unique_vals = df_main[filter_col].dropna().unique().tolist()
            selected_vals = st.multiselect(f"'{filter_col}' හි අගයන්", unique_vals)
        
        if selected_vals:
            active_df = df_main[df_main[filter_col].isin(selected_vals)]
            st.info(f"Filter කළ පේළි ගණන: {len(active_df)}")

    st.divider()

    # --- Mapping Section ---
    st.subheader("🔗 4. Map Columns")
    mapping_dict = {}
    
    # Table Header
    h1, h2, h3 = st.columns([0.5, 2, 2])
    h1.write("**Select**")
    h2.write("**Source Column**")
    h3.write("**Destination Column**")

    for i, m_col in enumerate(active_df.columns):
        c1, c2, c3 = st.columns([0.5, 2, 2])
        with c1:
            checked = st.checkbox("", key=f"chk_{i}")
        with c2:
            st.info(m_col)
        with c3:
            if checked:
                target = st.selectbox(f"Map {i}", df_trans.columns, key=f"tar_{i}", label_visibility="collapsed")
                mapping_dict[m_col] = target

    # --- Process Button ---
    if st.button("🚀 Process & Download", use_container_width=True):
        if mapping_dict:
            try:
                # දත්ත ගලපා අලුත් DataFrame එකක් සෑදීම
                # Transport එකේ Heading එක වෙනස් නොවී දත්ත පමණක් update වේ
                final_df = pd.DataFrame(columns=df_trans.columns)
                for src, dest in mapping_dict.items():
                    final_df[dest] = active_df[src].values
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False)
                
                st.success("සාර්ථකව දත්ත සකස් කරන ලදී!")
                st.download_button("📥 Download Excel", output.getvalue(), "Updated_Data.xlsx")
            except Exception as e:
                st.error(f"Error: {e}")
