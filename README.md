# WiFi Login Fail Report Generator

## Overview for IBUS Tech
The **WiFi Login Fail Report Generator** is a specialized, browser-based utility designed to help the IBUS Tech team streamline the maintenance and analysis of WiFi authentication logs. It completely automates the tedious process of cross-referencing separate "Failure" and "Success" network logs, merging them into a single, comprehensive, and professionally formatted Excel document.

## Core Problem Solved
When maintaining WiFi networks, technicians often export raw login data from network controllers resulting in massive, separate CSV or Excel files for both successful and failed attempts. Manually finding which users failed to connect, why they failed, and whether they eventually succeeded is a time-consuming spreadsheet task. This tool automates the intersection of this data to instantly highlight problematic user journeys.

## Key Capabilities & Workflows

*   **"Fail-First" MAC Address Matching:** The tool cross-references reports using device MAC addresses. It filters the Success report to *only* include devices that initially appeared in the Failure report. This isolates users who struggled to connect, removing the noise of standard, seamless logins.
*   **Intelligent Error Standardization:** It automatically cleans and translates verbose system logs into concise, easily readable reasons. For example, it intelligently compares Failure and Success data to deduce "Wrong Room Number" or "Wrong Last Name", and maps timeout errors to "Connection Timeout".
*   **Seamless Template Injection:** The tool leverages ExcelJS to inject the matched data directly into a pre-approved blank template. It perfectly preserves the template's cell formatting, custom borders, and styling, eliminating the need for manual post-export formatting.
*   **Smart Auto-Mapping:** The tool auto-detects headers in the raw network exports and uses synonym-matching (e.g., linking "User" to "Username" or "Time" to "Date") to automatically map data to the correct columns in the template.
*   **Secure & Local Processing:** The application runs entirely client-side in the browser. No sensitive user data, MAC addresses, or network logs are ever uploaded to a remote server, ensuring full compliance with privacy and security standards.

## Value to the IBUS Tech Team
By standardizing error reasons and automating data intersection, this utility transforms raw, messy data into actionable insights in seconds. It eliminates manual spreadsheet comparison, reduces human error, and allows the tech team to spend less time formatting reports and more time resolving actual network and authentication issues.

https://huntersouvik.github.io/Login-fail-report-generator/

## 🚀 Features

- **Smart Data Merging:** Combines raw failure and success logs into a single workflow.
- **MAC Address Matching:** Implements a "Fail-First" logic where success records are only included if the device (MAC) also appears in the failure report. This helps identify users who struggled before successfully connecting.
- **Template-Based Export:** Uses **ExcelJS** to inject processed data into a pre-styled `.xlsx` template, preserving existing cell formatting, borders, and headers.
- **Intelligent Column Mapping:**
  - Auto-detects headers in uploaded files.
  - Suggests column mappings based on synonyms (e.g., matching "User" to "Username" or "Time" to "Date").
- **Data Cleaning:** Automatically standardizes common error messages (e.g., converts verbose "Room not match" errors to "Wrong Room Number").
- **Modern UI:**
  - Dark/Light mode support (persisted via LocalStorage).
  - Drag-and-drop file uploads.
  - Live clock and interactive stepper interface.
  - Keyboard shortcuts (Alt+T to toggle theme).

## 🛠️ Tech Stack

- **Frontend:** HTML5, CSS3, Vanilla JavaScript.
- **Libraries:**
  - SheetJS (xlsx) - For robust reading of input Excel/CSV files.
  - ExcelJS - For writing data to the output template while maintaining styles.
- **Fonts:** Inter (via Google Fonts).

## 📖 How to Use

1. **Launch the Tool:**
   Open `index.html` in any modern web browser. No server installation is required.

2. **Step 1: Upload Fail Report**
   Drag and drop your raw failure log (Excel or CSV). The tool will scan for headers.

3. **Step 2: Upload Success Report**
   Drag and drop your raw success log.

4. **Step 3: Upload Template**
   Upload the `.xlsx` template file.
   > **Note:** The template must contain two distinct table header rows (one for Failures, one for Successes) for the tool to detect where to insert data.

5. **Step 4: Map & Generate**
   - Select the **MAC Address** column for both reports to enable cross-referencing.
   - Map source columns to the template columns (the tool will try to auto-select the best matches).
   - Click **Generate**.

6. **Step 5: Results**
   View statistics on how many records were processed and matched, and download the generated file automatically.

## ⚙️ Logic Details

### Header Detection
The tool uses a heuristic approach to find header rows in raw data files, looking for rows with unique text values while skipping common report titles (like "LOGIN-Failure").

### Data Intersect
1. The tool builds a Set of all MAC addresses found in the **Fail Report**.
2. It iterates through the **Success Report**.
3. A success record is kept **only if** its MAC address exists in the Fail Report set.
4. This ensures the final report focuses on users who experienced issues.

## 📦 Installation

Since this is a client-side application, you can simply clone the repository and open the file directly:

```bash
git clone https://github.com/yourusername/wifi-fail-report-generator.git
cd wifi-fail-report-generator
# Open index.html in your browser
```

## 👤 Credits

**Designed & Developed by:** Souvik Sarkar

## 📄 License

This project is available for use and modification.
