import pandas as pd
import re
import streamlit as st
import time
from concurrent.futures import ThreadPoolExecutor

def extract_sales_data(file_content):
    lines = file_content.split("\n")
    shop_id, shop_name = None, None
    
    for line in lines:
        match = re.search(r"(\d{5,})[_\s-]+(.+)", line)  # Support underscores and dashes
        if match:
            shop_id, shop_name = match.groups()
            shop_name = shop_name.strip()
            break
    
    sales_data = {
        "CARD": 0, "CASH": 0, "COD": 0, "CREDIT": 0, "MOBI KWIK": 0,"PAYBYLINK": 0,"GIFT VOUCHER":0,
        "PAYTM CARD": 0, "PAYTM DQRC": 0, "QR CODE": 0, "RELIGARE": 0, "UPI": 0, "POS SALES": 0,
        "TENDER WISE SALES SUMMARY": 0, "CORPORATE CODE WISE SALES SUMMARY": 0
    }
    
    def safe_extract(value):
        value = value.strip().replace(',', '')
        try:
            return float(value) if '.' in value else int(value)
        except ValueError:
            return 0
    
    for line in lines:
        # Handle explicit TOTALAMOUNT (no space) -> TENDER WISE SALES SUMMARY
        if re.search(r"TOTALAMOUNT", line, re.IGNORECASE):
            values = re.findall(r"[-+]?[0-9,]*\.?[0-9]+", line)
            if values:
                sales_data["TENDER WISE SALES SUMMARY"] = safe_extract(values[-1])
                continue

        # Handle explicit TOTAL AMOUNT (with space) -> CORPORATE CODE WISE SALES SUMMARY
        if re.search(r"TOTAL\s*AMOUNT", line, re.IGNORECASE):
            values = re.findall(r"[-+]?[0-9,]*\.?[0-9]+", line)
            if values:
                sales_data["CORPORATE CODE WISE SALES SUMMARY"] = safe_extract(values[-1])
                continue

        for key in list(sales_data.keys()):
            # Skip TENDER WISE SALES SUMMARY as handled above
            if key == "TENDER WISE SALES SUMMARY":
                continue
            pattern = rf"^{re.escape(key)}\b"
            if re.search(pattern, line, re.IGNORECASE):
                values = re.findall(r"[-+]?[0-9,]*\.?[0-9]+", line)
                if values:
                    sales_data[key] = safe_extract(values[-1])
    
    # Fallback: search reversed lines for TOTALAMOUNT or TOTAL AMOUNT
    for line in reversed(lines):
        if re.search(r"TOTALAMOUNT", line, re.IGNORECASE):
            values = re.findall(r"[-+]?[0-9,]*\.?[0-9]+", line)
            if values:
                sales_data["TENDER WISE SALES SUMMARY"] = safe_extract(values[-1])
                break
        if re.search(r"TOTAL\s*AMOUNT", line, re.IGNORECASE):
            values = re.findall(r"[-+]?[0-9,]*\.?[0-9]+", line)
            if values:
                sales_data["CORPORATE CODE WISE SALES SUMMARY"] = safe_extract(values[-1])
                break
    
    return shop_id, shop_name, sales_data

def process_files(uploaded_files):
    data_list = []
    total_files = len(uploaded_files)
    
    with ThreadPoolExecutor() as executor:
        results = list(executor.map(lambda f: extract_sales_data(f.getvalue().decode("utf-8", errors="ignore")), uploaded_files))
    
    for idx, (shop_id, shop_name, sales_data) in enumerate(results):
        if shop_id and shop_name:
            row_data = {"Shop id": shop_id, "Shop Name": shop_name, **sales_data}
            data_list.append(row_data)
        st.progress((idx + 1) / total_files)
    
    return data_list

def convert_to_excel(data_list):
    df = pd.DataFrame(data_list)
    # Drop POS SALES if present
    if "POS SALES" in df.columns:
        df = df.drop(columns=["POS SALES"])

    # Ensure uppercase TENDER WISE SALES SUMMARY and CORPORATE CODE WISE SALES SUMMARY exist
    if "TENDER WISE SALES SUMMARY" not in df.columns:
        df["TENDER WISE SALES SUMMARY"] = 0
    else:
        df["TENDER WISE SALES SUMMARY"] = df["TENDER WISE SALES SUMMARY"].fillna(0)

    if "CORPORATE CODE WISE SALES SUMMARY" not in df.columns:
        df["CORPORATE CODE WISE SALES SUMMARY"] = 0
    else:
        df["CORPORATE CODE WISE SALES SUMMARY"] = df["CORPORATE CODE WISE SALES SUMMARY"].fillna(0)

    # Reorder columns so TENDER WISE SALES SUMMARY and CORPORATE CODE WISE SALES SUMMARY are the last two columns
    cols = list(df.columns)
    # Keep Shop id and Shop Name at the front if present
    front = [c for c in ["Shop id", "Shop Name"] if c in cols]
    # Build working list excluding front columns
    working = [c for c in cols if c not in front]
    # Remove TENDER WISE SALES SUMMARY and CORPORATE CODE WISE SALES SUMMARY from working so they can be appended at the end
    for end_col in ["TENDER WISE SALES SUMMARY", "CORPORATE CODE WISE SALES SUMMARY"]:
        if end_col in working:
            working.remove(end_col)
    end_cols = [c for c in ["TENDER WISE SALES SUMMARY", "CORPORATE CODE WISE SALES SUMMARY"] if c in df.columns]
    df = df[front + working + end_cols]

    file_path = "Sales_Report.xlsx"
    df.to_excel(file_path, index=False)
    return file_path, df

st.title("Text File to Excel Converter")
uploaded_files = st.file_uploader("Upload your text files", type=["txt"], accept_multiple_files=True)

if uploaded_files:
    st.info("Processing files... Please wait.")
    start_time = time.time()
    
    data_list = process_files(uploaded_files)
    
    if data_list:
        output_file_path, df = convert_to_excel(data_list)
        st.write("### Extracted Data Preview")
        st.dataframe(df)
        
        st.success("Processing Completed! Excel file generated successfully.")
        st.write(f"Total Processing Time: {time.time() - start_time:.2f} seconds")
        
        with open(output_file_path, "rb") as f:
            st.download_button("Download Excel File", f, file_name="Sales_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("Failed to process the files.")


