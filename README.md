# 📊 Inventory & Sales Dashboard

A full-featured **Shiny dashboard** for managing and visualizing inventory, stock flow, and sales performance.  
It supports **Excel uploads**, **interactive visualizations**, **real-time KPI tracking**, and **data exports** (CSV, Excel, PDF).  

---

## 🚀 Features

### 🔹 Data Upload & Management
- Upload Excel files containing sheets:  
  - `sales_data`: Records of items sold  
  - `storage_data`: Inbound stock (purchases or restocks)  
  - `goods_data` : Static goods master list  
- Add new storage (inbound) or sale (outbound) records manually  
- Automatically cleans and normalizes inconsistent column names  

### 🔹 KPI Dashboard (Stock Snapshot)
- Displays top-selling, fastest-selling, and overstocked items  
- Calculates total profit, profit margin, and total stock value  
- Identifies unsold and low-stock items  

### 🔹 Inventory Overview
- View and filter current stock per item  
- Automatically flags items with stock below threshold (default: < 5 units)  

### 🔹 Inbound & Outbound Logs
- Interactive tables (using **DT**) to review:
  - Sales transactions  
  - Storage (restock) activities  

### 🔹 COGS & Profit Analysis
- FIFO-based **Cost of Goods Sold (COGS)** calculation  
- Generates profit per sale, total revenue, and tax-adjusted net profit  
- Top and worst-selling product summaries  

### 🔹 Charts & Visual Insights
- **Daily sales**, **profit margin trends**, and **cumulative profit** plots  
- Interactive **date range sliders** for time-based filtering  
- All charts rendered using **ggplot2** with a dark modern theme  

### 🔹 Exports & Reports
- Export data as:
  - CSV (updated logs)
  - Excel summary (`.xlsx`)
  - PDF report (sales overview)

---

## 🧩 Tech Stack

| Component | Description |
|------------|-------------|
| **R Shiny** | Core web framework |
| **shinydashboard** | Dashboard layout and theme |
| **DT** | Interactive tables |
| **ggplot2** | Data visualization |
| **readxl / writexl** | Excel I/O |
| **dplyr / tidyr** | Data cleaning and wrangling |
| **shinyWidgets** | Enhanced UI controls (sliders, etc.) |
| **lubridate** | Date/time formatting |
| **rmarkdown** | PDF report generation |
| **shinyBS** | Additional UI feedback elements (e.g., tooltips, modals) |

---

## 📂 Project Structure

📦 Inventory-Sales-Dashboard
│
├── app.R # Main Shiny application (UI + Server integrated)
├── README.md # Documentation (this file)
├── example_data.xlsx # Optional demo data (sales_data, storage_data, goods_data)
└── /www # (Optional) Folder for custom CSS, logos, etc.

## ⚙️ Setup & Installation

### 1️⃣ Install Required Packages

```r
install.packages(c(
  "shiny", "shinydashboard", "readxl", "dplyr", "tidyr", 
  "ggplot2", "DT", "scales", "shinyWidgets", "writexl", 
  "rmarkdown", "lubridate", "shinyBS"
))
```

2️⃣ Run the App
```r
library(shiny)
runApp("path/to/your/app.R")
```

or simply open app.R in RStudio and click Run App.


📊 Excel File Format
You can upload an Excel workbook with these sheets:
Sheet	Required Columns	Description
sales_data	Date, Item, Quantity, Price, Total	Records each sale
storage_data	Date, Item, Stock_Added, Price	Tracks inbound stock
goods_data (optional)	Item, Price, Stock, Date_Added	Static item details

⚠️ Ensure sheet names match exactly as above (case-insensitive).

🎨 Theme & UI
Fully dark-mode interface with blue accent highlights.
Styled tables, dropdowns, pagination, and hover effects.
Responsive layout powered by shinydashboard and custom CSS.

💾 Data Flow Overview
Excel Upload → Reactive DataFrames (sales_log, storage_log, goods)
         ↓
      Data Cleaning & Normalization
         ↓
  KPIs / COGS / Charts / Tables
         ↓
   Export: CSV / Excel / PDF Reports

🧠 Key Reactives & Components
Reactive	Description
- sales_log()	Outbound sales transactions
- storage_log()	Inbound stock updates
- goods()	Static product list
- current_stock()	Real-time stock balance
- cogs_data()	Full COGS + profit computation
- sales_profit_data()	Profit per sale item

🧾 Example Use Case
- Upload example_data.xlsx with sample data.
- Add a few manual entries (new sale, restock).
- Explore KPIs and visual charts.
- Export an updated Excel or PDF report.

🛠️ Future Improvements
- Multi-user login & authentication
- Barcode/QR integration for item scanning
- Automated email reports
- Integration with SQL databases (PostgreSQL/MySQL)

👤 Author
Developed by: Triston Aloyssius Marta
📧 tristonmarta@yahoo.com.sg
💼 Data Science | Analytics
