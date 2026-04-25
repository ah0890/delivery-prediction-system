# 📦 BuffaloEx Delivery Prediction System

An end-to-end automation pipeline that scrapes, analyzes, and predicts delivery performance for BuffaloEx shipments. Combines web scraping, data visualization, and machine learning to generate professional PDF reports and delay predictions.

---

## ✨ Features

- 🕷️ **Automated Web Scraping** – Fetches real-time tracking data from BuffaloEx
- 📄 **PDF Report Generation** – Clean, formatted tracking reports for single or batch numbers
- 📊 **Excel/CSV Export** – Structured data with delivery metrics, timestamps, and delay flags
- 📈 **Data Visualization** – Pie charts, bar graphs, histograms, and city-wise performance
- 🤖 **ML Delay Prediction** – Random Forest model classifies shipments as `On-Time` or `Delayed`
- 📑 **Prediction PDF Reports** – Auto-generated professional reports with confidence scores
- ⚙️ **Batch Processing** – Handles ranges of tracking numbers with progress tracking & error resilience

---

## 🛠️ Tech Stack

| Category        | Tools/Libraries                          |
|-----------------|------------------------------------------|
| **Language**    | Python 3.9+                              |
| **Scraping**    | Selenium, Chrome WebDriver               |
| **Data/ML**     | Pandas, Scikit-learn, Joblib, NumPy      |
| **Visualization**| Matplotlib, Seaborn                      |
| **PDF/Excel**   | ReportLab, Openpyxl, XlsxWriter          |
| **Environment** | Virtualenv (`.venv`)                     |

---

## 📦 Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/ah0890/delivery-prediction-system.git
   cd delivery-prediction-system