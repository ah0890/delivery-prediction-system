import streamlit as st
import pandas as pd
import os
import re
import time
import zipfile
import io
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
import logging

# Add after imports
import os

# Check if running in Docker/Render
IN_DOCKER = os.path.exists('/.dockerenv')

# Update Chrome options in generate_tracking_pdf()
options = webdriver.ChromeOptions()
if IN_DOCKER:
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.binary_location = "/usr/bin/google-chrome"
else:
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    
# ================= CONFIGURATION =================
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tracking Data")
os.makedirs(OUTPUT_DIR, exist_ok=True)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
DATE_PATTERN = re.compile(r'^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}')

# ================= PAGE SETUP =================
st.set_page_config(
    page_title="BuffaloEx Tracker",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ================= HELPER FUNCTIONS (Reused from your script) =================

def parse_tracking_format(tracking_num):
    match = re.match(r"^([A-Za-z]+)(\d+)([A-Za-z]+)$", tracking_num)
    if not match:
        raise ValueError(f"Invalid format: '{tracking_num}'")
    return match.group(1), int(match.group(2)), match.group(3), len(match.group(2))

def generate_tracking_range(start_num, end_num):
    prefix_s, num_s, suffix_s, len_s = parse_tracking_format(start_num)
    prefix_e, num_e, suffix_e, len_e = parse_tracking_format(end_num)
    if prefix_s != prefix_e or suffix_s != suffix_e:
        raise ValueError("Start and end must have SAME prefix and suffix")
    if num_s > num_e:
        raise ValueError("Start number cannot be greater than end number")
    return [f"{prefix_s}{str(i).zfill(len_s)}{suffix_s}" for i in range(num_s, num_e + 1)]

def generate_tracking_pdf(tracking_num, output_dir):
    driver = None
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-extensions")
        
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 30)
        
        driver.get("https://buffaloex.co.za/track.html")
        dropdown = wait.until(EC.presence_of_element_located((By.TAG_NAME, "select")))
        Select(dropdown).select_by_visible_text("Tracking No.")
        
        tracking_input = wait.until(EC.presence_of_element_located((By.ID, "order-no")))
        tracking_input.clear()
        tracking_input.send_keys(tracking_num)
        
        search_btn = wait.until(EC.element_to_be_clickable((By.ID, "submit-btn3")))
        driver.execute_script("arguments[0].click();", search_btn)
        wait.until(EC.visibility_of_element_located((By.ID, "search-result")))
        
        result_content = driver.find_element(By.ID, "result-content")
        all_text = result_content.text
        lines = [line.strip() for line in all_text.split('\n') if line.strip()]
        
        receive_city, sign_status, logistics_pairs, all_descriptions = "", "Not signed", [], []
        i = 0
        while i < len(lines):
            line = lines[i]
            if 'Receivecity' in line or 'Receive city' in line:
                receive_city = line.split(':')[-1].strip() if ':' in line else line
                i += 1; continue
            if 'Sign in picture' in line:
                sign_status = lines[i+1].strip() if i+1 < len(lines) else "Not signed"
                i += 2; continue
            if 'Logistics records' in line.lower():
                i += 1; continue
            if DATE_PATTERN.match(line):
                timestamp, description = line, ""
                if i + 1 < len(lines):
                    next_line = lines[i+1]
                    if not DATE_PATTERN.match(next_line) and not any(h in next_line.lower() for h in ['tracking', 'receivecity', 'sign', 'logistics']):
                        description = next_line
                        i += 2
                    else:
                        i += 1
                else:
                    i += 1
                logistics_pairs.append((timestamp, description))
                all_descriptions.append(description.lower())
            else:
                i += 1
        
        is_delivered = any('client received' in d or 'delivered' in d or 'received' in d for d in all_descriptions)
        final_status = "Delivered ✓" if is_delivered else "In Transit "
        status_color = colors.green if is_delivered else colors.orange
        
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        pdf_name = f"Tracking_{tracking_num}_{timestamp_str}.pdf"
        pdf_path = os.path.join(output_dir, pdf_name)
        
        doc = SimpleDocTemplate(pdf_path, rightMargin=0.75*inch, leftMargin=0.75*inch, topMargin=0.75*inch, bottomMargin=0.75*inch)
        styles = getSampleStyleSheet()
        header_style = ParagraphStyle('CustomHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, spaceAfter=4)
        normal_style = ParagraphStyle('CustomNormal', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=12, spaceAfter=0)
        status_style = ParagraphStyle('FinalStatus', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, textColor=status_color, spaceAfter=12, spaceBefore=6)
        
        elements = []
        elements.append(Paragraph(f"<b>Tracking Number:</b> {tracking_num}", header_style))
        elements.append(Spacer(1, 4))
        if receive_city:
            elements.append(Paragraph(f"<b>Receivecity:</b> {receive_city}", header_style))
            elements.append(Spacer(1, 4))
        elements.append(Paragraph(f"<b>Sign in picture:</b>", header_style))
        elements.append(Paragraph(f"  {sign_status}", normal_style))
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f"<b>Logistics records:</b>", header_style))
        elements.append(Spacer(1, 6))
        for ts, desc in logistics_pairs:
            if ts.strip():
                elements.append(Paragraph(f"  {ts.strip()}", normal_style))
                elements.append(Spacer(1, 4))
            if desc.strip():
                elements.append(Paragraph(f"  {desc.strip()}", normal_style))
            elements.append(Spacer(1, 8))
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f"<b>Final Status:</b>", header_style))
        elements.append(Paragraph(f"  {final_status}", status_style))
        doc.build(elements)
        return pdf_path
    except TimeoutException:
        logging.warning(f"⏱️ Timeout for {tracking_num}")
        return None
    except Exception as e:
        logging.error(f"❌ Error processing {tracking_num}: {str(e)}")
        return None
    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass

# ================= SIDEBAR =================
with st.sidebar:
    st.image("https://buffaloex.co.za/images/logo.png", width=200)
    st.markdown("### 📦 BuffaloEx Tracker")
    st.info("Enter a range of tracking numbers to fetch data and generate PDF reports.")
    st.markdown("---")
    st.markdown("**Tips:**")
    st.markdown("- Tracking numbers must have same prefix/suffix")
    st.markdown("- Example: `BUFZA5120042001YQ` to `BUFZA5120042010YQ`")
    st.markdown("- Processing ~2 seconds per number")
    st.markdown("---")
    st.caption("Built with ❤️ using Streamlit + Selenium")

# ================= MAIN APP =================
st.title("📦 BuffaloEx Batch Tracker")
st.markdown("Fetch tracking data, generate PDF reports, and analyze delivery performance — all from your browser.")

# Input Form
col1, col2 = st.columns(2)
with col1:
    start_num = st.text_input("🔹 Start Tracking Number", placeholder="BUFZA5120042001YQ")
with col2:
    end_num = st.text_input("🔹 End Tracking Number", placeholder="BUFZA5120042010YQ")

# Progress & Results
progress_bar = st.progress(0)
status_text = st.empty()
results_container = st.container()

# Generate Button
if st.button("🚀 Generate Reports", type="primary", disabled=not (start_num and end_num)):
    try:
        # Validate & generate range
        tracking_list = generate_tracking_range(start_num.strip(), end_num.strip())
        st.success(f"✅ Found {len(tracking_list)} tracking numbers to process")
        
        results = []
        for i, num in enumerate(tracking_list, 1):
            status_text.text(f"🔄 Processing {i}/{len(tracking_list)}: {num}")
            pdf_path = generate_tracking_pdf(num, OUTPUT_DIR)
            
            if pdf_path:
                results.append({
                    "Tracking Number": num,
                    "Status": "✅ Success",
                    "PDF": os.path.basename(pdf_path),
                    "Path": pdf_path
                })
            else:
                results.append({
                    "Tracking Number": num,
                    "Status": "❌ Failed",
                    "PDF": None,
                    "Path": None
                })
            
            progress_bar.progress(i / len(tracking_list))
            time.sleep(0.1)  # Smooth UI updates
        
        status_text.text("✨ All done!")
        progress_bar.empty()
        
        # Display Results Table
        if results:
            df = pd.DataFrame(results)
            st.subheader("📋 Results Summary")
            st.dataframe(
                df[["Tracking Number", "Status", "PDF"]],
                use_container_width=True,
                hide_index=True
            )
            
            # Download Buttons
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                # Download single PDF (first successful)
                successful = [r for r in results if r["PDF"]]
                if successful:
                    with open(successful[0]["Path"], "rb") as f:
                        st.download_button(
                            label="📄 Download Sample PDF",
                            data=f.read(),
                            file_name=successful[0]["PDF"],
                            mime="application/pdf"
                        )
            
            with col_dl2:
                # Download ZIP of all PDFs
                if any(r["PDF"] for r in results):
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                        for r in results:
                            if r["Path"] and os.path.exists(r["Path"]):
                                zip_file.write(r["Path"], os.path.basename(r["Path"]))
                    
                    st.download_button(
                        label="📦 Download All PDFs (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"BuffaloEx_Tracking_{datetime.now().strftime('%Y%m%d')}.zip",
                        mime="application/zip"
                    )
        
        # Summary Stats
        success_count = sum(1 for r in results if r["Status"] == "✅ Success")
        st.markdown(f"""
        <div style='background-color: #f0f2f6; padding: 1rem; border-radius: 0.5rem; margin-top: 1rem;'>
            <b>📊 Session Summary:</b><br>
            ✅ Successful: {success_count}<br>
            ❌ Failed: {len(results) - success_count}<br>
            📁 Saved to: <code>{OUTPUT_DIR}</code>
        </div>
        """, unsafe_allow_html=True)
        
    except ValueError as ve:
        st.error(f"❌ Format Error: {ve}")
        st.info("💡 Make sure start/end numbers have the same prefix and suffix (e.g., BUFZA...YQ)")
    except Exception as e:
        st.error(f"❌ Unexpected Error: {e}")
        st.exception(e)  # Show full traceback in dev mode

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666; font-size: 0.9rem;'>"
    "Built with Streamlit • Selenium • ReportLab • "
    f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    "</div>",
    unsafe_allow_html=True
)