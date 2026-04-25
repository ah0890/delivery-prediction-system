from flask import Flask, request, send_file, jsonify
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime
import os
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

app = Flask(__name__)

# Absolute path for downloads folder (works reliably in VS Code/terminal)
DOWNLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)


@app.route('/')
def home():
    return jsonify({
        "message": "BuffaloEx Tracking API",
        "usage": "GET /track?number=YOUR_TRACKING_NUMBER"
    }), 200


@app.route('/track', methods=['GET'])
def track():
    tracking_number = request.args.get("number", "").strip()
    if not tracking_number:
        return jsonify({"error": "Tracking number is required"}), 400

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

        # ================= OPEN WEBSITE =================
        driver.get("https://buffaloex.co.za/track.html")

        # ================= SELECT DROPDOWN =================
        dropdown = wait.until(EC.presence_of_element_located((By.TAG_NAME, "select")))
        Select(dropdown).select_by_visible_text("Tracking No.")

        # ================= ENTER TRACKING =================
        tracking_input = wait.until(EC.presence_of_element_located((By.ID, "order-no")))
        tracking_input.clear()
        tracking_input.send_keys(tracking_number)

        # ================= CLICK SEARCH =================
        search_btn = wait.until(EC.element_to_be_clickable((By.ID, "submit-btn3")))
        driver.execute_script("arguments[0].click();", search_btn)

        # Wait for result container
        wait.until(EC.visibility_of_element_located((By.ID, "search-result")))

        # ================= SCRAPE BASIC INFO =================
        basic_info = []
        info_blocks = driver.find_elements(By.CSS_SELECTOR, "#result-content p.record-ul")
        for block in info_blocks:
            text = block.text.strip()
            if text:
                basic_info.append(text)

        if not basic_info:
            logging.warning(f"No tracking records found for {tracking_number}")
            return jsonify({"error": "No tracking data found. Verify tracking number or website structure."}), 404

        # ================= SIGN STATUS =================
        sign_status = driver.find_element(
            By.CSS_SELECTOR, "#result-content div.record-data"
        ).text.strip()

        # ================= LOGISTICS =================
        logistics = []
        records = driver.find_elements(
            By.XPATH, "//div[@id='result-content']//div[contains(@class,'record-data')]"
        )

        for rec in records[1:]:  # Skip first if it's the status block
            text = rec.text.strip().replace("\n", " ")
            if text:
                logistics.append(text)

        # ================= CREATE PDF =================
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pdf_name = f"Tracking_{tracking_number}_{timestamp}.pdf"
        pdf_path = os.path.join(DOWNLOAD_DIR, pdf_name)

        doc = SimpleDocTemplate(
            pdf_path,
            rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=18
        )
        styles = getSampleStyleSheet()
        elements = []

        elements.append(Paragraph("<b>BuffaloEx Tracking Report</b>", styles['Title']))
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f"<b>Tracking Number:</b> {tracking_number}", styles['Normal']))
        elements.append(Spacer(1, 10))

        for line in basic_info:
            elements.append(Paragraph(line, styles['Normal']))
            elements.append(Spacer(1, 6))

        elements.append(Spacer(1, 10))
        elements.append(Paragraph(f"<b>Sign Status:</b> {sign_status}", styles['Normal']))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph("<b>Logistics Records:</b>", styles['Heading2']))
        elements.append(Spacer(1, 10))

        for record in logistics:
            clean = record.replace("[", " (").replace("]", ")")
            elements.append(Paragraph(f"• {clean}", styles['Normal']))
            elements.append(Spacer(1, 6))

        doc.build(elements)
        logging.info(f"✅ PDF generated: {pdf_name}")

        # ================= RETURN FILE =================
        return send_file(
            pdf_path,
            as_attachment=True,
            download_name=pdf_name,
            mimetype='application/pdf'
        )

    except TimeoutException:
        logging.error("️ Timeout waiting for page elements. Website may be slow or structure changed.")
        return jsonify({"error": "Request timed out. The tracking page may be slow or structure changed."}), 504
    except Exception as e:
        logging.error(f"❌ Unexpected error: {str(e)}", exc_info=True)
        return jsonify({"error": "Internal server error. Check server logs for details."}), 500
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass


if __name__ == "__main__":
    # debug=False recommended for headless/terminal use
    app.run(host="0.0.0.0", port=5000, debug=False)