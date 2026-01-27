# 📧 Google Sheets Mass Mailing Engine

A robust, maintainable mass-mailing system built on **Google Sheets + Apps Script**, designed for teams who need to send **personalized emails at scale** while keeping full control, traceability, and safety.

This project is **not** a simple mail merge: it is a structured engine with validation, throttling, per-row configuration, and a reproducible sheet template.

---

## ✨ Key Features

- ✅ Send personalized emails from Google Sheets  
- ✅ One email = one row (full control)  
- ✅ Per-row subject override + mandatory global subject  
- ✅ Google Docs template with `$Variable$` placeholders  
- ✅ CC / BCC / Reply-To / No-Reply support  
- ✅ Throttling (anti-spam & quota friendly)  
- ✅ Immediate status updates (`Sent` + `SentAt`)  
- ✅ Test email mode (safe, isolated row)  
- ✅ One-click template reconstruction  
- ✅ Clean, maintainable, object-oriented Apps Script architecture  

---

## 🧠 How It Works (High Level)

Google Sheet
↓
SheetTable (parses headers + rows)
↓
MailOrchestrator (validation + flow)
↓
EmailComposer (email options + vars)
↓
TemplateRenderer (Google Docs merge)
↓
MailSender (send + mark + throttle)


Everything is driven by **column headers**, not column positions.

---

## 📁 Project Structure

.
├── Main.gs
├── Config.gs
├── AppContext.gs
├── Utils.gs
├── Orchestrator.gs
├── Services_EmailComposer.gs
├── Services_TemplateRenderer.gs
├── Services_MailSender.gs
├── Services_SheetTable.gs
├── ReconstructTemplate.gs


---

## 🚀 Getting Started — Step-by-Step Tutorial

### 1️⃣ Create the Apps Script project

1. Open a Google Sheet  
2. Go to **Extensions → Apps Script**  
3. Paste all `.gs` files from this repository  
4. Save the project  

---

### 2️⃣ Reconstruct the mailing template

1. Reload the Google Sheet  
2. In the menu, click **Send email → Reconstruct mass mailing template**  
3. Confirm  

➡️ Your sheet is now rebuilt with:
- Correct columns
- Checkboxes
- Formatting
- Test row
- Configuration area

⚠️ This action **clears the sheet entirely** by design.

---

### 3️⃣ Prepare your Google Docs template

1. Create a Google Docs file  
2. Use placeholders that match **column headers**, for example:

```text
Hello $Name$,

Please visit $Topic1$.

Best regards.
Copy the document ID from the URL

Paste it into cell B6 (Template ID)

4️⃣ Configure the global subject
In cell B7, set a global subject

This value is mandatory — sending is blocked if it is empty

Example:

Pour me soutenir => une étoile sur mon GitHub
Each row can override this via the Subject column.

5️⃣ Fill your data rows
Starting from row 12:

Required

To send → checked

Email

Subject (or leave empty to use the global subject)

Optional

cc, bcc

replyTo

noReply

Template variables (Name, Topic1, etc.)

6️⃣ Send a test email (recommended)
Use row 10 (“Test Email Data”)

Menu → Send email → Test email

This sends only one email, without touching campaign rows.

7️⃣ Send the campaign
Check To send on desired rows

Menu → Send email → Send selected emails

Confirm

During sending:

Rows are updated immediately

Sent is checked

SentAt is filled (date + time)

Throttling is applied between emails

🧩 Column Semantics
Control / Email Columns (blue)
To send — user intent

Sent — system status

SentAt — system timestamp (read-only)

Subject — per-row override

Email, cc, bcc, replyTo, noReply

Template Variables (green)
Every non-reserved header becomes available in the template as:

$HeaderName$
Example:

Column Topic1 → $Topic1$

🔒 No-Reply Configuration
If noReply is checked:

Email is sent from APP_CONFIG.noReplyFromEmail

⚠️ This address must be configured as an alias in Gmail:

Gmail → Settings → Accounts → Send mail as

⏱ Throttling & Safety
Configured in Config.gs:

throttling: {
  secondsMin: 10,
  secondsMax: 15,
}
Why:

Avoid Gmail rate limits

Reduce spam-like behavior

Improve reliability on large batches

🛠 Maintenance & Customization
Change layout → ReconstructTemplate.gs

Change headers → Config.gs

Add new template fields → just add columns

Change throttling → config only

Add protections (optional) → Google Sheets protections

The system is designed so most changes do not require touching orchestration logic.

⚠️ Important Notes
❌ Email spoofing is not supported (by design)

✅ Only verified Gmail aliases can be used

❌ This tool does not bypass Gmail limits

✅ It works with Gmail rules, not against them

📜 License
MIT — use freely, modify responsibly.

⭐ Support
If this project helped you, consider starring the repo ❤️
👉 https://github.com/jayzonne
