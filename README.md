# SME Loan Right-First-Time Pre-Screen

This project implements a rule-based SME loan pre-screen system using:

- Google Sheets (dataset + rule engine + dashboard)
- Google Apps Script (sidebar chatbot / input form)
- Simple document flags (KYC, Income, Business Proof)

## How it works

1. Loan officer opens the Google Sheet.
2. Uses menu **SME Portal → Open Chatbot** to launch the sidebar.
3. Enters applicant details (ID, industry, amount, SME category) and selects
   whether KYC, Income, and Business Proof documents are submitted (Yes/No).
4. Apps Script:
   - Writes a new application row to the **Loan Application DataSet** sheet.
   - Copies formulas in columns K–O to evaluate document completeness.
   - Computes a final status: **READY FOR APPRAISAL / HOLD / REJECT / Business Proof Needed**.
5. The **Dashboard** sheet shows aggregate stats and a pie chart for all applications.

## Files in this repo

- `Code.gs` – Apps Script backend (menu, chatbot handler, rule-engine link)
- `ChatbotUpload.html` – Sidebar UI shown in Google Sheets
- `sample-dataset.xlsx` – Example loan dataset (same structure as the live sheet)
- `screenshots/dashboard.png` – Result dashboard
- `screenshots/chatbot.png` – Chatbot sidebar
