import streamlit as st
import tempfile
import os
from python_script import (
    compare_requested_series_from_github,  # ✅ added
    update_master_series,
    delete_from_master_series,
    update_series_rules,
    delete_from_series_rules
)

from helper import (
    load_file,
    similarity_ratio,
    normalize_series

)


# ===== Streamlit Config =====
st.set_page_config(page_title="Series Comparison Tool", layout="wide")
st.title("📊 Enhanced Series Comparison Tool")

# ===== Sidebar Menu =====
menu = st.sidebar.radio("Select Operation", [
    "Compare Series",
    "Update Master Series History",
    "Delete from Master Series History",
    "Update Series Rules",
    "Delete from Series Rules"
])

def save_uploaded_file(uploaded_file):
    """Save uploaded file to a temporary location and return the path."""
    if uploaded_file is None:
        return None
    temp_dir = tempfile.mkdtemp()
    path = os.path.join(temp_dir, uploaded_file.name)
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path

# ===== Compare Series =====
if menu == "Compare Series":
    st.subheader("Compare Requested Series (GitHub Master & Rules)")
    comparison_file = st.file_uploader("Upload Comparison file", type=["xlsx", "xls", "csv"])
    top_n = st.number_input("Top N series to show", min_value=1, max_value=10, value=2)

    if st.button("Run Comparison"):
        if comparison_file:
            comparison_path = save_uploaded_file(comparison_file)
            df_result = compare_requested_series_from_github(comparison_path, top_n=top_n)

            st.success("✅ Comparison completed!")
            st.dataframe(df_result)

            output_path = os.path.join(tempfile.mkdtemp(), "comparison_result.xlsx")
            df_result.to_excel(output_path, index=False)
            with open(output_path, "rb") as f:
                st.download_button("Download Result Excel", f, file_name="comparison_result.xlsx")
        else:
            st.error("Please upload a Comparison file.")
# ===== Update Master =====
elif menu == "Update Master Series History":
    st.subheader("Update Master Series History")
    update_file = st.file_uploader("Upload Update file", type=["xlsx", "xls", "csv"])
    master_file = st.file_uploader("Upload Master file", type=["xlsx", "xls", "csv"])

    if st.button("Run Update"):
        if update_file and master_file:
            update_path = save_uploaded_file(update_file)
            master_path = save_uploaded_file(master_file)
            update_master_series(update_path, master_path)
            st.success("✅ Update completed!")
        else:
            st.error("Please upload both Update and Master files.")

# ===== Delete Master =====
elif menu == "Delete from Master Series History":
    st.subheader("Delete from Master Series History")
    delete_file = st.file_uploader("Upload Delete file", type=["xlsx", "xls", "csv"])
    master_file = st.file_uploader("Upload Master file", type=["xlsx", "xls", "csv"])

    if st.button("Run Deletion"):
        if delete_file and master_file:
            delete_path = save_uploaded_file(delete_file)
            master_path = save_uploaded_file(master_file)
            delete_from_master_series(delete_path, master_path)
            st.success("✅ Deletion completed!")
        else:
            st.error("Please upload both Delete and Master files.")

# ===== Update Rules =====
elif menu == "Update Series Rules":
    st.subheader("Update Series Rules")
    update_file = st.file_uploader("Upload Update file", type=["xlsx", "xls", "csv"])
    rules_file = st.file_uploader("Upload Rules file", type=["xlsx", "xls", "csv"])

    if st.button("Run Update"):
        if update_file and rules_file:
            update_path = save_uploaded_file(update_file)
            rules_path = save_uploaded_file(rules_file)
            update_series_rules(update_path, rules_path)
            st.success("✅ Rules update completed!")
        else:
            st.error("Please upload both Update and Rules files.")

# ===== Delete Rules =====
elif menu == "Delete from Series Rules":
    st.subheader("Delete from Series Rules")
    delete_file = st.file_uploader("Upload Delete file", type=["xlsx", "xls", "csv"])
    rules_file = st.file_uploader("Upload Rules file", type=["xlsx", "xls", "csv"])

    if st.button("Run Deletion"):
        if delete_file and rules_file:
            delete_path = save_uploaded_file(delete_file)
            rules_path = save_uploaded_file(rules_file)
            delete_from_series_rules(delete_path, rules_path)
            st.success("✅ Rules deletion completed!")
        else:
            st.error("Please upload both Delete and Rules files.")
