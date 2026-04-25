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

# ================= CONFIGURATION =================
TRACKING_NUMBER = "BUFZA5120042008YQ"
DOWNLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Pattern to detect logistics date lines: YYYY-MM-DD HH:MM
DATE_PATTERN = re.compile(r'^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}')


def generate_tracking_pdf(tracking_num):
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
        logistics_pairs = []  # Store as (timestamp, description)
        all_descriptions = []  # Store all descriptions to check for delivery

        i = 0
        while i < len(lines):
            line = lines[i]

            # Extract headers
            if 'Tracking No' in line:
                tracking_no = line.split(':')[-1].strip()
                i += 1
                continue
            if 'Receivecity' in line or 'Receive city' in line:
                receive_city = line.split(':')[-1].strip() if ':' in line else line
                i += 1
                continue
            if 'Sign in picture' in line:
                sign_status = lines[i+1].strip() if i+1 < len(lines) else "Not signed"
                i += 2
                continue
            if 'Logistics records' in line.lower():
                i += 1
                continue

            # If line starts with a date, pair it with the next line (if it's a description)
            if DATE_PATTERN.match(line):
                timestamp = line
                description = ""
                if i + 1 < len(lines):
                    next_line = lines[i+1]
                    # Skip if next line is another date or a header
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
        is_delivered = False
        for desc in all_descriptions:
            if 'client received' in desc or 'delivered' in desc or 'received' in desc:
                is_delivered = True
                break

        # UPDATED: Use a standard text checkmark to avoid "square block" issues in PDF fonts
        final_status = "Delivered ✓" if is_delivered else "In Transit "
        status_color = colors.green if is_delivered else colors.orange

        logging.info(f"✅ Extracted {len(logistics_pairs)} logistics records")
        logging.info(f"📦 Final Status: {final_status}")

        # ================= CREATE PDF =================
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pdf_name = f"Tracking_{tracking_num}_{timestamp}.pdf"
        pdf_path = os.path.join(DOWNLOAD_DIR, pdf_name)

        doc = SimpleDocTemplate(
            pdf_path,
            rightMargin=0.75*inch, leftMargin=0.75*inch,
            topMargin=0.75*inch, bottomMargin=0.75*inch
        )
        
        styles = getSampleStyleSheet()
        
        # Custom styles for clean formatting
        header_style = ParagraphStyle(
            'CustomHeader',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=10,
            spaceAfter=4
        )
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=10,
            leading=12,
            spaceAfter=0
        )
        
        # Style for final status
        status_style = ParagraphStyle(
            'FinalStatus',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=12,
            textColor=status_color,
            spaceAfter=12,
            spaceBefore=6
        )
        
        elements = []

        # Tracking Info
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

        # Logistics: Timestamp on one line, description on next, spacing after each step
        for ts, desc in logistics_pairs:
            if ts.strip():
                elements.append(Paragraph(f"  {ts.strip()}", normal_style))
                elements.append(Spacer(1, 4))  # Space between timestamp and description
            if desc.strip():
                elements.append(Paragraph(f"  {desc.strip()}", normal_style))
            elements.append(Spacer(1, 8))  # Line space after every step

        # Final Status Section - AT THE END
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f"<b>Final Status:</b>", header_style))
        elements.append(Paragraph(f"  {final_status}", status_style))

        # Build PDF
        doc.build(elements)
        logging.info(f"✅ PDF saved: {pdf_path}")
        return pdf_path

    except TimeoutException:
        logging.error("⏱️ Timeout waiting for page elements.")
    except Exception as e:
        logging.error(f" Unexpected error: {str(e)}", exc_info=True)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
    return None


# ================= RUN =================
if __name__ == "__main__":
    print(f"🔄 Fetching tracking data for: {TRACKING_NUMBER}")
    pdf_path = generate_tracking_pdf(TRACKING_NUMBER)
    
    if pdf_path:
        print(f"📄 Success! PDF saved to: {pdf_path}")
    else:
        print("❌ Failed to generate PDF.")