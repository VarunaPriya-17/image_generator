import streamlit as st
import pandas as pd
import requests
import re
import zipfile
from io import BytesIO

st.set_page_config(page_title="Product Image Downloader", layout="wide")
st.title("📦 Product Image Downloader from Excel (Preserve Order)")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Read Excel without changing row order
    df = pd.read_excel(uploaded_file, dtype=str)  # Keep as string to avoid issues

    # Validate required columns
    required_cols = {"Product Name", "Image URL"}
    if not required_cols.issubset(df.columns):
        st.error(f"❌ Excel must contain columns: {required_cols}")
    else:
        st.success("✅ File uploaded successfully. Downloading images in Excel order...")

        # Memory ZIP file
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for idx, row in df.iterrows():  # Keeps Excel order
                product_name = str(row["Product Name"]).strip()
                image_url = str(row["Image URL"]).strip()

                # Clean filename
                safe_name = re.sub(r'[\\/*?:"<>|]', "", product_name)

                try:
                    # Download image
                    response = requests.get(image_url, timeout=10)
                    response.raise_for_status()

                    # Write into ZIP (with index to preserve order in filename)
                    zipf.writestr(f"{idx+1:03d}_{safe_name}.jpg", response.content)

                    st.write(f"✅ {idx+1}. Downloaded: {safe_name}.jpg")

                except Exception as e:
                    st.write(f"❌ Failed for {product_name}: {e}")

        # Save ZIP in memory
        zip_buffer.seek(0)

        # Download button
        st.download_button(
            label="⬇ Download All Images (ZIP)",
            data=zip_buffer,
            file_name="product_images.zip",
            mime="application/zip"
        )
