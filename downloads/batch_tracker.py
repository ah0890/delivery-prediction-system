from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from datetime import datetime
import os
import logging
import re
import time

# ================= CONFIGURATION =================
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tracking Data")
os.makedirs(OUTPUT_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Pattern to detect logistics date lines: YYYY-MM-DD HH:MM
DATE_PATTERN = re.compile(r'^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}')


def parse_tracking_format(tracking_num):
    """Extracts prefix, numeric part, suffix, and numeric length from tracking number."""
    match = re.match(r"^([A-Za-z]+)(\d+)([A-Za-z]+)$", tracking_num)
    if not match:
        raise ValueError(f"Invalid format: '{tracking_num}'. Expected format like 'BUFZA5120042008YQ'")
    return match.group(1), int(match.group(2)), match.group(3), len(match.group(2))


def generate_tracking_range(start_num, end_num):
    """Generates a list of tracking numbers between start and end (inclusive)."""
    prefix_s, num_s, suffix_s, len_s = parse_tracking_format(start_num)
    prefix_e, num_e, suffix_e, len_e = parse_tracking_format(end_num)

    if prefix_s != prefix_e or suffix_s != suffix_e:
        raise ValueError("Start and end tracking numbers must have the SAME prefix and suffix.")
    if num_s > num_e:
        raise ValueError("Start number cannot be greater than end number.")

    tracking_list = []
    for i in range(num_s, num_e + 1):
        padded_num = str(i).zfill(len_s)
        tracking_list.append(f"{prefix_s}{padded_num}{suffix_s}")
    return tracking_list


def generate_tracking_pdf(tracking_num, output_dir):
    driver = None
    try:
        # ================= DRIVER SETUP =================
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

        # ================= OPEN & INTERACT =================
        driver.get("https://buffaloex.co.za/track.html")

        dropdown = wait.until(EC.presence_of_element_located((By.TAG_NAME, "select")))
        Select(dropdown).select_by_visible_text("Tracking No.")

        tracking_input = wait.until(EC.presence_of_element_located((By.ID, "order-no")))
        tracking_input.clear()
        tracking_input.send_keys(tracking_num)

        search_btn = wait.until(EC.element_to_be_clickable((By.ID, "submit-btn3")))
        driver.execute_script("arguments[0].click();", search_btn)

        wait.until(EC.visibility_of_element_located((By.ID, "search-result")))

        # ================= SCRAPE & PAIR DATA =================
        result_content = driver.find_element(By.ID, "result-content")
        all_text = result_content.text
        lines = [line.strip() for line in all_text.split('\n') if line.strip()]

        tracking_no = tracking_num
        receive_city = ""
        sign_status = "Not signed"
        logistics_pairs = []
        all_descriptions = []

        i = 0
        while i < len(lines):
            line = lines[i]

            if 'Tracking No' in line:
                tracking_no = line.split(':')[-1].strip()
                i += 1; continue
            if 'Receivecity' in line or 'Receive city' in line:
                receive_city = line.split(':')[-1].strip() if ':' in line else line
                i += 1; continue
            if 'Sign in picture' in line:
                sign_status = lines[i+1].strip() if i+1 < len(lines) else "Not signed"
                i += 2; continue
            if 'Logistics records' in line.lower():
                i += 1; continue

            if DATE_PATTERN.match(line):
                timestamp = line
                description = ""
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

        # ================= DETERMINE FINAL STATUS =================
        is_delivered = any('client received' in d or 'delivered' in d or 'received' in d for d in all_descriptions)
        final_status = "Delivered ✓" if is_delivered else "In Transit "
        status_color = colors.green if is_delivered else colors.orange

        # ================= CREATE PDF =================
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        pdf_name = f"Tracking_{tracking_num}_{timestamp_str}.pdf"
        pdf_path = os.path.join(output_dir, pdf_name)

        doc = SimpleDocTemplate(
            pdf_path,
            rightMargin=0.75*inch, leftMargin=0.75*inch,
            topMargin=0.75*inch, bottomMargin=0.75*inch
        )
        
        styles = getSampleStyleSheet()
        header_style = ParagraphStyle('CustomHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, spaceAfter=4)
        normal_style = ParagraphStyle('CustomNormal', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=12, spaceAfter=0)
        status_style = ParagraphStyle('FinalStatus', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, textColor=status_color, spaceAfter=12, spaceBefore=6)
        
        elements = []
        elements.append(Paragraph(f"<b>Tracking Number:</b> {tracking_no}", header_style))
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
            except Exception:
                pass


# ================= MAIN EXECUTION =================
if __name__ == "__main__":
    print("📦 BuffaloEx Batch Tracker")
    print("="*40)
    try:
        start_input = input("🔹 Enter START tracking number: ").strip()
        end_input = input("🔹 Enter END tracking number:   ").strip()

        if not start_input or not end_input:
            print("❌ Both start and end numbers are required.")
            exit(1)

        print("\n🔄 Generating tracking range...")
        tracking_list = generate_tracking_range(start_input, end_input)
        print(f"✅ Found {len(tracking_list)} tracking numbers to process.")
        print(f"📂 Output folder: {OUTPUT_DIR}\n")

        success_count = 0
        fail_count = 0

        for i, num in enumerate(tracking_list, 1):
            print(f"[{i}/{len(tracking_list)}] Processing: {num}")
            pdf_path = generate_tracking_pdf(num, OUTPUT_DIR)
            
            if pdf_path:
                success_count += 1
                print(f"✅ Saved: {os.path.basename(pdf_path)}")
            else:
                fail_count += 1
                print(f"❌ Failed to generate PDF for {num}")
            
            # Polite delay to avoid server rate limits
            if i < len(tracking_list):
                time.sleep(2)

        print("\n" + "="*40)
        print(f"🎉 BATCH COMPLETE!")
        print(f"✅ Successful: {success_count}")
        print(f"❌ Failed: {fail_count}")
        print(f"📂 All files saved in: {OUTPUT_DIR}")
        print("="*40)

    except KeyboardInterrupt:
        print("\n⚠️ Process interrupted by user.")
    except ValueError as ve:
        print(f"\n❌ Format Error: {ve}")
    except Exception as e:
        print(f"\n❌ Unexpected Error: {e}")