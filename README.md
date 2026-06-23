# FYW Pending Order Report

A Streamlit web app that generates the daily FYW (For You Wear) Pending Order Report — formatted Excel workbook with pivot table, brand summaries, and per-channel order details.

---

## 📁 Files

| File | Description |
|------|-------------|
| `app.py` | Streamlit web application |
| `report_engine.py` | Core data processing & Excel generation logic |
| `requirements.txt` | Python dependencies |

---

## 🚀 Running Locally

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/fyw-report-app.git
cd fyw-report-app

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the app
streamlit run app.py
```

Then open http://localhost:8501 in your browser.

---

## ☁️ Deploy on Streamlit Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** → select your repo → set `app.py` as the main file
4. Click **Deploy** — done!

---

## 📋 How to Use

### Daily Workflow
1. Download the FYW dashboard export (`ALL-DD-Mon-YYYY.csv`)
2. Download Shopee seller centre files for Melissa, Ipanema, CSpace (if available)
3. Upload all files in the sidebar
4. Review the pivot table and order summary on screen
5. Click **Download Excel Report** to get the formatted `.xlsx`

### Input Files

| File | Required | Notes |
|------|----------|-------|
| `ALL-[date].csv` | ✅ Yes | FYW dashboard export |
| Melissa Shopee `.xlsx` | ⚠️ Optional | For MP SLA date lookup |
| Ipanema Shopee `.xlsx` | ⚠️ Optional | For MP SLA date lookup |
| CSpace Shopee `.xlsx` | ⚠️ Optional | For MP SLA date lookup |

> Shopee files are optional but recommended — without them, Shopee MP SLA dates will be blank.

---

## 📊 Output Excel Sheets

| Sheet | Contents |
|-------|----------|
| **PIVOT** | Order count by Nickname × FYW SLA date (colour-coded: 🔴 past / 🟠 yesterday / 🟢 today) |
| **SUMMARY** | Orders, items, MYR value per brand |
| **ALL PENDING ORDERS** | Full detail list |
| **MELISSA** | Melissa-only orders |
| **IPANEMA** | Ipanema-only orders |
| **CSPACE** | CSpace-only orders |

---

## ⚙️ MP SLA Logic

| Channel | MP SLA Calculation |
|---------|-------------------|
| Shopee | Estimated Ship Out Date from Seller Centre file |
| TikTok | Order Date + 1 day |
| Lazada | Order Date + 1 day |
| Zalora | Order Date + 3 days |

---

## 🏷️ Pending Order Statuses

Orders with these FYW statuses are included:
- `ACCEPTED/PICKED`
- `READY TO SHIP`
- `NEW`
- `CANCEL REQUESTED`
