# 📊 Commission Intelligence Agent — MAR1

AI-powered Streamlit app that parses, reconciles, and reports on MARCTECH2 / American Bright Optoelectronics commission files — powered by **Groq (Free)**.

---

## ✨ Features

| Feature | Description |
|---|---|
| **Parse** | Reads both PDF (`Comm_from_Payment`) and Excel (`Comm_Report`) files |
| **Reconcile** | Cross-checks PDF Part I vs Excel Part I — flags all discrepancies |
| **Report** | Consolidated view: regular sales + distributor sales + grand totals |
| **Ask AI** | Chat with LLaMA 3.3 70B about your data — trends, totals, issues |

---

## 🆓 Free API Key (Groq)

1. Go to **[console.groq.com](https://console.groq.com)**
2. Sign up (free, no credit card)
3. Click **API Keys → Create API Key**
4. Copy your key — starts with `gsk_…`

---

## 🚀 Quick Start (Local)

```bash
git clone https://github.com/YOUR_USERNAME/commission-agent.git
cd commission-agent
pip install -r requirements.txt
streamlit run app.py
```

Enter your Groq API key in the sidebar and upload your files.

---

## ☁️ Deploy to Streamlit Cloud (Free)

1. Push this repo to GitHub
2. Go to **[streamlit.io/cloud](https://streamlit.io/cloud)** → sign in with GitHub
3. Click **New app** → select repo → main file: `app.py`
4. Under **Advanced settings → Secrets**, add:
   ```toml
   GROQ_API_KEY = "gsk_your_key_here"
   ```
5. Click **Deploy** — live in ~2 minutes ✅

---

## 📁 Project Structure

```
commission-agent/
├── app.py                          ← Main Streamlit app
├── requirements.txt                ← Dependencies
├── .gitignore                      ← Keeps secrets out of git
├── .streamlit/
│   └── secrets.toml.example        ← API key template
└── README.md
```

---

## 🧠 How the files work

```
PDF  (Comm_from_Payment_MAR1_XXX.pdf)
  └── Part I: Regular Sales — payments received that month

Excel (Comm_Report_MAR1_XXX.xlsx)
  ├── Part I:  Regular Sales CURRENT month  ← must match PDF
  └── Part II: Distributor Sales PRIOR month ← 1-month lag (normal)
```

---

## 🛠 Tech Stack

- **Streamlit** — UI
- **Groq + LLaMA 3.3 70B** — AI analysis & chat (free)
- **pdfplumber** — PDF extraction
- **openpyxl** — Excel reading
- **pandas** — Data tables
