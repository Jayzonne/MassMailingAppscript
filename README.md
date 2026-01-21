# 👋 Hi, I'm building structured tools on top of Google Workspace

I design **robust, scalable internal tools** using **Google Sheets, Apps Script (V8), and Google Docs** — with a strong focus on **safety, maintainability, and UX for non-technical users**.

---

## 🚀 Current Focus

### 📧 Google Sheets Mass Mailing Engine
A fully structured **mass mailing system** built on Google Sheets + Apps Script.

**Why?**  
Because most mass-mailing tools are either:
- too simple and unsafe, or
- too complex for business users.

This project sits in between:  
👉 *Spreadsheet-driven, but engineered like real software.*

---

## ✨ What this project does

- Send **personalized emails** from Google Sheets
- Use **Google Docs templates** with variables like `$Topic1$`
- Strong separation between:
  - 🟦 email configuration (To / CC / BCC / Reply-To / No Reply)
  - 🟩 template variables
- Built-in **test row** for safe validation
- **Immediate feedback** (rows marked `Sent` as emails go out)
- Strong **anti-spam throttling** (10–15s per email)
- One-click **template reconstruction** for clean onboarding

All of this is handled via a **custom menu inside Google Sheets**.

---

## 🧠 Engineering Principles

- **Object-Oriented Apps Script (V8)**
- Clear separation of responsibilities:
  - Table parsing
  - Email composition
  - Template rendering
  - Throttling & sending
- Defensive programming:
  - header validation
  - error isolation
  - no silent failures
- No spoofing, no hacks — only **Workspace-compliant behavior**

---

## 🛠 Tech Stack

- Google Apps Script (V8)
- Google Sheets
- Google Docs
- Gmail / Google Workspace
- ES6 classes & modular architecture

---

## 🎯 Use cases

- Internal communications
- Event invitations
- Controlled B2B outreach
- Customer notifications
- Any workflow where **data lives in Sheets but quality matters**

---

## 📌 Philosophy

> *Treat spreadsheets like a UI, not like code.*

Most spreadsheet automations fail because they grow without structure.  
This project proves you can build **real systems** on top of Google Workspace — if you treat them with the same discipline as backend code.

---

## 📬 Get in touch

If you're interested in:
- advanced Apps Script patterns
- safe mass-mailing architectures
- or building serious tools on top of Google Workspace

feel free to connect.
