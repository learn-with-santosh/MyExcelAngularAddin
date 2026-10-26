# Microsoft Excel Modern Analytics Add-in

Built with **Angular 19**, **Office.js**, **Kendo UI for Angular**, **Custom Functions**, **Task Pane**, and **Ribbon Commands**.

---

## 🌟 Key Features

1. **Angular 19 Standalone Architecture**:
   - Modern Angular 19 standalone component pattern with async animation providers.
   - Initialized via `Office.onReady()` with fallback for standalone browser preview.

2. **Kendo UI for Angular Components**:
   - **Kendo Grid**: Interactive data grid displaying sales records with selection sync to the active Excel sheet.
   - **Kendo Chart**: Visual column chart displaying product category revenue distribution.
   - **Kendo Buttons & Theming**: Styled with `@progress/kendo-theme-default` and Microsoft Excel brand styling.

3. **Office.js Integration (`src/app/services/office.service.ts`)**:
   - `readSelectedRange()`: Reads active user selection directly from Excel sheet.
   - `writeToRange()`: Writes and formats tabular data directly into the active sheet (with automatic header styling and autofit).
   - `highlightRow()`: Highlights and selects corresponding row in Excel when clicked in Kendo Grid.
   - `populateSampleSalesData()`: Injects demo sales dataset into Excel with a single click.

4. **Excel Custom Functions (`public/functions.json` & `public/functions.js`)**:
   - `=CONTOSO.ADD(first, second)`: Adds two numbers directly inside Excel formulas.
   - `=CONTOSO.MULTIPLY(first, second)`: Computes product of numbers.
   - `=CONTOSO.STREAMTIMER(interval)`: Streaming asynchronous function that updates clock time in the cell in real-time.

5. **Excel Ribbon Commands (`public/commands.html` & `public/commands.js`)**:
   - Adds custom **Excel Analytics** ribbon tab in Excel.
   - **Show Taskpane**: Opens the Kendo UI Angular taskpane.
   - **Timestamp**: Inserts current date & time into active cell without opening taskpane.
   - **Highlight**: Accents the selected range with yellow fill.

---

## 🚀 Running the Project

### 1. Start the Dev Server
The development server runs over HTTPS on port 3000 (required for Excel Add-ins):
```bash
npm start
```
Open [https://localhost:3000](https://localhost:3000) in your browser. (Accept the self-signed certificate if prompted).

### 2. Sideloading into Microsoft Excel

#### Option A: Excel on the Web (Fastest)
1. Open [Excel for the Web (Office.com)](https://office.live.com/start/Excel.aspx) in your browser.
2. Open a blank workbook.
3. Go to the **Insert** tab > **Add-ins** (or **Get Add-ins**).
4. Click **Manage My Add-ins** > **Upload My Add-in**.
5. Browse and select [`manifest.xml`](file:///d:/All%20In%20Projects/My%20Projects/MyExcelAngularAddin/manifest.xml) from this project folder.
6. The **Excel Analytics** tab will appear on your Excel ribbon!

#### Option B: Excel on Windows Desktop
1. Share the project folder (or a dedicated folder containing `manifest.xml`) as a Windows network share:
   - Right-click folder > **Properties** > **Sharing** > **Share...**
2. In Excel Desktop:
   - Go to **File** > **Options** > **Trust Center** > **Trust Center Settings...** > **Trusted Add-in Catalogs**.
   - In **Catalog URL**, enter the UNC path (e.g. `\\localhost\MyExcelAddin`).
   - Check **Show in Menu**, then click **Add Catalog** and **OK**.
3. Restart Excel > **Insert** tab > **My Add-ins** > **Shared Folder** tab > Select **Excel Analytics Add-in**.

---

## 📁 Key File Structure

```text
├── manifest.xml                     <- Classic XML manifest for Office & Excel
├── public/
│   ├── commands.html / commands.js  <- Office ribbon command action handlers
│   ├── functions.html / functions.js<- Custom functions runtime script
│   ├── functions.json               <- Excel custom function metadata
│   └── assets/                      <- Add-in icons (16px, 32px, 80px)
├── src/
│   ├── main.ts                      <- Office.onReady bootstrapping
│   ├── styles.css                   <- Kendo UI theme & Excel styling
│   └── app/
│       ├── app.component.ts         <- Standalone Angular component with Kendo imports
│       ├── app.component.html       <- Taskpane UI with Kendo Grid & Charts
│       ├── app.component.css        <- Taskpane styling
│       └── services/
│           └── office.service.ts    <- Excel.run context wrapper
└── package.json                     <- Angular 19 & Kendo UI v20 dependencies
```
