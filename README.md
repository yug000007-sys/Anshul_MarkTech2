# 📊 Commission Intelligence Agent — MAR1

An AI-powered Streamlit app that parses, reconciles, and reports on MARCTECH2 / American Bright Optoelectronics commission files.

---

## ✨ Features

| Feature | Description |
|---|---|
| **Parse** | Reads both PDF (`Comm_from_Payment`) and Excel (`Comm_Report`) files automatically |
| **Reconcile** | Cross-checks PDF Part I vs Excel Part I for every month and flags discrepancies |
| **Report** | Consolidated view of all months — regular sales + distributor sales + grand totals |
| **Ask AI** | Chat with Claude about your data — ask about trends, discrepancies, totals |

---

## 🧠 How the files work

```
PDF  (Comm_from_Payment_MAR1_XXX.pdf)
  └── Part I: Regular Sales — payments received in that month

Excel (Comm_Report_MAR1_XXX.xlsx)
  ├── Sheet "Comm Report"
  │     ├── Part I: Regular Sales for CURRENT month  ← should match the PDF
  │     └── Part II: Distributor Sales for PREVIOUS month  ← 1-month lag is intentional
  └── Sheet "Disty Sales" — line-item detail for distributors
```

---

## 🚀 Quick Start (Local)

### 1. Clone the repo
```bash
git clone https://github.com/YOUR_USERNAME/commission-agent.git
cd commission-agent
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run
```bash
streamlit run app.py
```

### 4. Use
1. Enter your **Anthropic API key** in the sidebar (`sk-ant-…`)  
   → Get one at https://console.anthropic.com
2. Upload your PDF and Excel commission files
3. Click **⚡ Analyze & Reconcile**
4. Explore the Summary, Reconciliation, Full Report, and Ask AI tabs

---

## ☁️ Deploy to Streamlit Cloud (Free)

1. Push this repo to GitHub (see below)
2. Go to **https://streamlit.io/cloud** → Sign in with GitHub
3. Click **New app** → select your repo → set `app.py` as the main file
4. Under **Advanced settings → Secrets**, add:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-your-key-here"
   ```
5. Click **Deploy** — your app will be live in ~2 minutes

> **Tip:** If you add the secret, the app will auto-load the key. Otherwise users enter it in the sidebar.

---

## 📁 Project Structure

```
commission-agent/
├── app.py              ← Main Streamlit application
├── requirements.txt    ← Python dependencies
├── .gitignore          ← Keeps secrets and caches out of git
└── README.md           ← This file
```

---

## 🔐 API Key Security

- **Never** commit your API key to git
- Use Streamlit Cloud secrets (see above) for deployment
- The `.gitignore` already excludes `.env` files

---

## 📦 Push to GitHub

```bash
# First time
git init
git add .
git commit -m "Initial commit: Commission Intelligence Agent"
git remote add origin https://github.com/YOUR_USERNAME/commission-agent.git
git branch -M main
git push -u origin main

# Subsequent updates
git add .
git commit -m "Your update message"
git push
```

---

## 🛠 Tech Stack

- **Streamlit** — UI framework
- **Anthropic Claude** — AI analysis & chat (`claude-sonnet-4-6`)
- **pdfplumber** — PDF text extraction
- **openpyxl** — Excel file reading
- **pandas** — Data tables

---

## 📝 License

MIT — free to use and modify.
