# 📊 Fetch Report Add-on for Google Sheets

This Google Sheets Add-on allows users to fetch data from Marketcheck APIs and populate their sheets with structured CSV data. It supports advanced features like authentication, dynamic parameters, and execution history tracking.

---

## 🚀 Features

- ✅ Marketcheck API Key support
- 📑 Target multiple sheets at once
- 📝 Execution history log
- 📌 Cell A1 notes with summary info
- ✨ Append or overwrite mode

---

## 🛠 How to Use

1. **Install the Add-on** from the Google Workspace Marketplace *(link to be provided after approval)*.
2. Open any Google Sheet.
3. Click on the **Fetch Report → Run** menu.
4. Fill out the “Developer Settings” sheet with:
   - `Params[data_server_url]`
   - `Params[target_tabs]`
   - `Params[data_load_mode]` (`append` or `overwrite`)
   - Any other required `Params[...]`

5. Click **Run** to fetch and load your report.
6. Developer settings could be little bit tricky. If you have received a template from Marcketcheck, you can use it. In that case, you need to provide your inputs in tab named "User Settings".

---

## 📷 Screenshots

*(You can include screenshots here showing how the menu looks, the Developer Settings tab, and a populated sheet.)*

---

## 🔒 Privacy Policy

[View Privacy Policy](https://yourname.github.io/fetch-report-addon/privacy.html)

---

## 📃 Terms of Service

[View Terms](https://yourname.github.io/fetch-report-addon/terms.html)

---

## 🧠 Developer Info

This add-on is written in **Google Apps Script** and published via the **Google Workspace Marketplace**.


