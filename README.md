# 🤖 RPA Excel-to-Web Automation Bot (Python & Selenium)

Automated Robotic Process Automation (RPA) workflow built with **Python**, **Selenium WebDriver**, and **OpenPyXL**. This tool reads raw operational records from an Excel spreadsheet (`Reporte_Diario_Gallinas.xlsx`), sanitizes corrupted data types, automates web authentication and form submission into a target web portal (`portal_granja.html`), and writes back the generated confirmation IDs directly into the original Excel file.

---

## 🚀 Key Technical Features

* **Automated Excel Data Sanitization:** Detects and fixes common Excel auto-formatting errors (such as floating-point weights like `2.1` converted into `datetime` objects) before injecting data into the browser.
* **Fault-Tolerant Exception Handling:** Identifies incomplete rows (missing primary IDs), logs `"ERROR_ID_VACIO"` directly into the output column in Excel, and continues execution without crashing.
* **Dynamic Synchronization (`WebDriverWait`):** Uses explicit waits and custom `lambda` conditions to detect real-time DOM updates and extract unique confirmation tokens (`COMP-GS-XXXX`) after each form submission.
* **Business Rule Conditional Logic:** Evaluates health/status strings dynamically to toggle UI checkboxes (`alerta_sanitaria`) and dropdown selectors (`Select`) automatically.
* **Bidirectional Excel Logging:** Updates and saves the `.xlsx` workbook with the web-generated tracking folio for every processed row.

---

## 🛠️ Tech Stack

* **Language:** Python 3
* **Web Automation:** Selenium WebDriver (`webdriver-manager`)
* **Spreadsheet Manipulation:** `openpyxl`
* **Target Interface:** HTML5 / CSS3 / JS Local Portal (`portal_granja.html`)

---

## 📂 Project Structure

* `gallians.py` — Main RPA automation script with data cleaning, Selenium control, and Excel write-back logic.
* `Reporte_Diario_Gallinas.xlsx` — Input/Output dataset containing records, edge cases (null IDs, date-formatted floats), and recorded confirmation folios.
* `portal_granja.html` — Target web application featuring session login, form validation, and dynamic table generation.

---

## ⚙️ How to Run

1. Install dependencies:
   ```bash
   pip install selenium openpyxl webdriver-manager
