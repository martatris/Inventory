# Full dashboard-style Shiny app (integrated, ready-to-run)
library(shiny)
library(shinydashboard)
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(DT)
library(scales)
library(shinyWidgets)
library(writexl)
library(rmarkdown)
library(lubridate)
library(shinyBS)

# ---- UI ----
ui <- dashboardPage(
  skin = "blue",
  dashboardHeader(title = "Inventory & Sales Dashboard", titleWidth = 300),
  dashboardSidebar(
    width = 300,
    sidebarMenu(
      id = "sidebar",
      menuItem("Upload & Previews", tabName = "upload_tab", icon = icon("upload")),
      menuItem("Stock Snapshot", tabName = "snapshot_tab", icon = icon("boxes")),
      menuItem("Inventory Overview", tabName = "inventory_tab", icon = icon("warehouse")),
      menuItem("Inbound & Outbound Logs", tabName = "logs_tab", icon = icon("exchange-alt")),
      menuItem("Sales & COGS Summary", tabName = "cogs_tab", icon = icon("file-invoice-dollar")),
      menuItem("Sales & Profit Charts", tabName = "charts_tab", icon = icon("chart-line")),
      menuItem("Exports", tabName = "export_tab", icon = icon("download"))
    )
  ),
  dashboardBody(
    tags$head(tags$style(HTML("
    /* --- GLOBAL DARK THEME --- */
    body, .content-wrapper { background-color: #0d1117; color: #d8e2ef; }
    .box { background-color: #1b2230 !important; border: 1px solid #2f3a4a; }
    .box-header { background: linear-gradient(90deg, #005f9e, #007bff); color: #fff !important; }
    .main-header .logo { background-color: #002d5f !important; color: #fff !important; font-weight: bold; }
    .sidebar-menu > li > a { color: #aab8c2 !important; }
    h1,h2,h3,h4,h5 { color: #ffffff !important; }

    /* --- TABLES --- */
    table.dataTable, table {
      color: #f8f9fa !important;
      background-color: #1b2230 !important;
    }
    table.dataTable th, table th {
      background-color: #005f9e !important;
      color: #ffffff !important;
      text-align: center;
    }
    table.dataTable tbody tr:hover, table tr:hover {
      background-color: #2a3b54 !important;
    }
    table.dataTable td, table td {
      border-color: #2f3a4a !important;
    }

    /* --- DATATABLE CONTROLS (Search + Entries) --- */
    div.dataTables_length label, div.dataTables_filter label {
      color: #ffffff !important;
      font-weight: 500;
    }
    div.dataTables_length select {
      background-color: #2a3b54 !important;
      color: #f8f9fa !important;
      border: 1px solid #007bff !important;
      border-radius: 6px;
      padding: 3px 6px;
    }
    div.dataTables_filter input {
      background-color: #2a3b54 !important;
      color: #f8f9fa !important;
      border: 1px solid #007bff !important;
      border-radius: 6px;
      padding: 3px 6px;
    }
    div.dataTables_filter input:focus,
    div.dataTables_length select:focus {
      outline: none !important;
      border-color: #00bfff !important;
      box-shadow: 0 0 5px #00bfff;
    }

    /* --- PAGINATION BUTTONS --- */
    .dataTables_wrapper .dataTables_paginate .paginate_button {
      background-color: #1b2230 !important;
      color: #f8f9fa !important;
      border: 1px solid #2f3a4a !important;
    }
    .dataTables_wrapper .dataTables_paginate .paginate_button.current {
      background-color: #007bff !important;
      color: #fff !important;
    }
    .dataTables_wrapper .dataTables_paginate .paginate_button:hover {
      background-color: #005f9e !important;
      color: #fff !important;
    }

    /* --- DATATABLE FOOTER INFO (Showing entries text) --- */
    .dataTables_wrapper .dataTables_info {
      color: #80cfff !important;
      font-weight: 500;
      padding-top: 10px;
    }
  "))),
    tabItems(
      # Upload
      tabItem(tabName = "upload_tab",
              fluidRow(
                box(title = "Upload Excel", status = "primary", solidHeader = TRUE, width = 6,
                    fileInput("excel_upload", "Upload Excel File", accept = c(".xlsx", ".xls")),
                    helpText("Expected sheets: sales_data, storage_data, goods_data, ..."),
                    br(),
                    actionButton("reload_demo", "Load Demo (if available)", icon = icon("redo"))
                ),
                box(title = "Recent Previews", status = "info", solidHeader = TRUE, width = 6,
                    h5("Recent Sales Log (Preview)"),
                    div(style = "overflow-x: auto; max-height: 200px;", tableOutput("sales_preview")),
                    hr(),
                    h5("Recent Storage Log (Preview)"),
                    div(style = "overflow-x: auto; max-height: 200px;", tableOutput("storage_preview"))
                )
              ),fluidRow(
                box(
                  title = "Add Storage Record (Inbound)",
                  status = "primary", solidHeader = TRUE, width = 6,
                  textInput("storage_item", "Item Name"),
                  numericInput("storage_qty", "Stock Added (Quantity)", value = 0, min = 0, step = 1),
                  numericInput("storage_price", "Price per Unit", value = 0, min = 0, step = 1),
                  actionButton("add_storage_btn", "Add Storage Record", icon = icon("plus"))
                ),
                box(
                  title = "Record Outbound Sale",
                  status = "warning", solidHeader = TRUE, width = 6,
                  textInput("sale_item", "Item Name"),
                  numericInput("sale_qty", "Quantity Sold", value = 0, min = 0, step = 1),
                  numericInput("sale_price", "Selling Price per Unit", value = 0, min = 0, step = 1),
                  actionButton("add_sale_btn", "Add Sale Record", icon = icon("shopping-cart"))
                )
              )
      ),
      # Stock Snapshot
      tabItem(tabName = "snapshot_tab",
              fluidRow(
                box(title = "KPIs", status = "primary", solidHeader = TRUE, width = 12,
                    fluidRow(
                      column(4, verbatimTextOutput("top_selling_item")),
                      column(4, verbatimTextOutput("low_stock_count")),
                      column(4, verbatimTextOutput("total_stock_value"))
                    ),
                    fluidRow(
                      column(4, verbatimTextOutput("total_profit_text")),
                      column(4, verbatimTextOutput("profit_margin_text")),
                      column(4, verbatimTextOutput("top_profit_item"))
                    ),
                    fluidRow(
                      column(4, verbatimTextOutput("fastest_selling_item")),
                      column(4, verbatimTextOutput("overstocked_item")),
                      column(4, verbatimTextOutput("unsold_items"))
                    )
                )
              ),
              fluidRow(
                box(title = "Current Stock Bar Plot", status = "info", solidHeader = TRUE, width = 12,
                    plotOutput("stock_bar_plot", height = "300px"))
              ),
              fluidRow(
                box(width = 12, title = "Goods Table", DT::dataTableOutput("goods_table"))
              )
      ),
      # Inventory Overview
      tabItem(tabName = "inventory_tab",
              fluidRow(
                box(title = "Filter by Item", status = "info", solidHeader = TRUE, width = 4,
                    uiOutput("item_selector")),
                box(title = "Current Stock Summary in Warehouse", status = "primary", solidHeader = TRUE, width = 8,
                    tableOutput("stock_summary"))
              ),
              fluidRow(
                box(title = "Low Stock Alert (Stock < 5)", status = "warning", solidHeader = TRUE, width = 12,
                    tableOutput("low_stock_table"))
              )
      ),
      # Logs
      tabItem(tabName = "logs_tab",
              fluidRow(
                box(title = "Filter Logs by Item", status = "info", solidHeader = TRUE, width = 4,
                    uiOutput("log_item_selector")),
                box(title = "Outbound Sales Records", status = "primary", solidHeader = TRUE, width = 8,
                    DTOutput("sales_log_table"))
              ),
              fluidRow(
                box(title = "Inbound Stock Records", status = "success", solidHeader = TRUE, width = 12,
                    DTOutput("storage_log_table"))
              )
      ),
      # COGS & Sales summary
      tabItem(tabName = "cogs_tab",
              fluidRow(
                box(width = 6, title = "Sales Summary", DT::dataTableOutput("sales_summary_table")),
                box(width = 6, title = "COGS & Inventory Summary", DT::dataTableOutput("cogs_summary"))
              ),
              fluidRow(
                box(title = "Stock Added (Purchase-based)", status = "success", solidHeader = TRUE, width = 12,
                    tableOutput("stock_added_table"))
              ),
              fluidRow(
                box(title = "Top / Worst Selling", status = "warning", solidHeader = TRUE, width = 12,
                    column(6, tableOutput("top_selling_table")), column(6, tableOutput("worst_selling_table")))
              )
      ),
      # Charts
      tabItem(tabName = "charts_tab",
              fluidRow(
                box(title = "Daily Sales", status = "primary", solidHeader = TRUE, width = 12,
                    uiOutput("slider_sales_ui"), plotOutput("sales_plot", height = "300px"))
              ),
              fluidRow(
                box(title = "Profit Margin per Day", status = "info", solidHeader = TRUE, width = 12,
                    uiOutput("slider_profit_margin_ui"), plotOutput("profit_margin_plot", height = "300px"))
              ),
              fluidRow(
                box(title = "Cumulative Sales & Profit Over Time", status = "success", solidHeader = TRUE, width = 12,
                    uiOutput("slider_cumulative_ui"), plotOutput("cumulative_plot", height = "350px"))
              )
      ),
      # Exports
      tabItem(tabName = "export_tab",
              fluidRow(
                box(title = "Downloads & Summary", status = "primary", solidHeader = TRUE, width = 6,
                    downloadButton("download_data", "Download Updated CSV"),
                    br(), br(),
                    downloadButton("download_excel", "Download Excel Summary (.xlsx)"),
                    br(), br(),
                    downloadButton("download_pdf", "Download Sales Report (PDF)")
                ),
                box(title = "Total Sales (quick)", status = "info", solidHeader = TRUE, width = 6,
                    verbatimTextOutput("total_sales_text"))
              )
      )
    )
  )
)

# ---- Server ----
server <- function(input, output, session) {
  # reactive holders
  goods <- reactiveVal(NULL)
  sales_log <- reactiveVal(data.frame(Date = character(), Item = character(), Quantity = numeric(), Price = numeric(), Total = numeric(), stringsAsFactors = FALSE))
  storage_log <- reactiveVal(data.frame(Date = character(), Item = character(), Stock_Added = numeric(), Price = numeric(), Location = character(), stringsAsFactors = FALSE))
  
  # Helper: safe read excel sheets into a named list (returns empty list if no file)
  read_uploaded_excel <- function(path) {
    sheets <- tryCatch(readxl::excel_sheets(path), error = function(e) character(0))
    if (length(sheets) == 0) return(list())
    out <- setNames(lapply(sheets, function(sh) {
      tryCatch(readxl::read_excel(path, sheet = sh), error = function(e) NULL)
    }), sheets)
    # remove NULL sheets
    out[!sapply(out, is.null)]
  }
  
  # Observe excel_upload
  observeEvent(input$excel_upload, {
    req(input$excel_upload)
    path <- input$excel_upload$datapath
    data_list <- read_uploaded_excel(path)
    
    # sales_data and storage_data are required for core functionality
    if (!("sales_data" %in% names(data_list)) || !("storage_data" %in% names(data_list))) {
      showNotification("Please ensure uploaded Excel includes sheets: sales_data and storage_data", type = "error")
      return()
    }
    
    # normalize and store sales & storage
    sd <- as.data.frame(data_list[["sales_data"]], stringsAsFactors = FALSE)
    st <- as.data.frame(data_list[["storage_data"]], stringsAsFactors = FALSE)
    
    # Ensure expected columns exist and fix naming differences if any
    # sales_data: Item, Quantity, Price, Total, Date
    if (!"Date" %in% names(sd) && "date" %in% names(sd)) sd$Date <- sd$date
    if (!"Total" %in% names(sd) && all(c("Quantity", "Price") %in% names(sd))) sd$Total <- sd$Quantity * sd$Price
    if (!"Quantity" %in% names(sd) && "Qty" %in% names(sd)) sd$Quantity <- sd$Qty
    # convert Date to character for consistent storage
    if ("Date" %in% names(sd)) sd$Date <- as.character(sd$Date)
    
    # storage_data: Item, Price, Stock_Added, Date
    if (!"Stock_Added" %in% names(st) && "Quantity" %in% names(st)) st <- st %>% rename(Stock_Added = Quantity)
    if (!"Date" %in% names(st) && "date" %in% names(st)) st$Date <- st$date
    if ("Date" %in% names(st)) st$Date <- as.character(st$Date)
    
    # goods_data optional
    if ("goods_data" %in% names(data_list)) {
      gd <- as.data.frame(data_list[["goods_data"]], stringsAsFactors = FALSE)
      if (!"Date_Added" %in% names(gd) && "date_added" %in% names(gd)) gd$Date_Added <- gd$date_added
      if ("Date_Added" %in% names(gd)) gd$Date_Added <- as.character(gd$Date_Added)
      goods(gd)
      # update buy_item choices if UI had them (keeps backward compatibility)
      if ("Item" %in% names(gd)) updateSelectInput(session, "buy_item", choices = unique(gd$Item))
    }
    
    # set logs
    sales_log(sd)
    storage_log(st)
    
    showNotification("Excel uploaded and parsed.", type = "message")
  })
  
  # Previews
  output$sales_preview <- renderTable({
    df <- sales_log()
    if (is.null(df) || nrow(df) == 0) return(data.frame())
    
    # Ensure all columns exist and consistent
    if (!"Type" %in% names(df)) df$Type <- "Sale"
    
    df <- df %>%
      mutate(
        Type = ifelse(is.na(Type), "Sale", Type),
        Date = format(as.POSIXct(Date), "%Y-%m-%d %H:%M:%S")
      )
    
    head(df, 8)
  })
  
  output$storage_preview <- renderTable({
    df <- storage_log()
    if (is.null(df) || nrow(df) == 0) return(data.frame())
    
    if (!"Type" %in% names(df)) df$Type <- "Stock Added"
    
    df <- df %>%
      mutate(
        Type = ifelse(is.na(Type), "Stock Added", Type),
        Date = format(as.POSIXct(Date), "%Y-%m-%d %H:%M:%S")
      )
    
    head(df, 8)
  })
  
  # --- ADD STORAGE RECORD (INBOUND) ---
  observeEvent(input$add_storage_btn, {
    req(input$storage_item, input$storage_qty > 0)
    
    new_row <- data.frame(
      Date = as.character(Sys.time()),  # convert to character for consistency
      Item = input$storage_item,
      Type = "Stock Added",
      Stock_Added = as.integer(input$storage_qty),
      Price = as.integer(input$storage_price),
      stringsAsFactors = FALSE
    )
    
    old_df <- tryCatch(storage_log(), error = function(e) data.frame())
    # ensure same column types before binding
    if ("Date" %in% names(old_df)) old_df$Date <- as.character(old_df$Date)
    updated_df <- dplyr::bind_rows(old_df, new_row)
    storage_log(updated_df)
    
    showNotification("✅ New storage record added.", type = "message")
  })
  
  
  # --- ADD OUTBOUND SALE ---
  observeEvent(input$add_sale_btn, {
    req(input$sale_item, input$sale_qty > 0)
    
    new_row <- data.frame(
      Date = as.character(Sys.time()),  # convert to character for consistency
      Item = input$sale_item,
      Type = "Sale",
      Quantity = as.integer(input$sale_qty),
      Price = as.integer(input$sale_price),
      Total = as.integer(input$sale_qty * input$sale_price),
      stringsAsFactors = FALSE
    )
    
    old_df <- tryCatch(sales_log(), error = function(e) data.frame())
    if ("Date" %in% names(old_df)) old_df$Date <- as.character(old_df$Date)
    updated_df <- dplyr::bind_rows(old_df, new_row)
    sales_log(updated_df)
    
    showNotification("✅ New outbound sale recorded.", type = "message")
  })
  
  # Helper: cost per unit via average from storage_log
  cost_per_unit_table <- reactive({
    df <- storage_log()
    if (is.null(df) || nrow(df) == 0) return(data.frame())
    df %>% group_by(Item) %>% summarise(cost_per_unit = mean(Price, na.rm = TRUE), .groups = "drop")
  })
  
  # FIFO cost function used in some places (robust to column names)
  fifo_cost_per_unit <- function(item, quantity_needed, log_df = storage_log()) {
    log <- log_df
    if (is.null(log) || nrow(log) == 0) return(NA_real_)
    if (!"Stock_Added" %in% names(log) && "Quantity" %in% names(log)) log$Stock_Added <- log$Quantity
    # Use Price or Unit_Cost if present
    price_col <- if ("Price" %in% names(log)) "Price" else if ("Unit_Cost" %in% names(log)) "Unit_Cost" else NA
    if (is.na(price_col)) return(NA_real_)
    relevant <- log %>% filter(Item == item) %>% arrange(as.POSIXct(Date))
    if (nrow(relevant) == 0) return(NA_real_)
    remaining <- relevant$Stock_Added
    total_cost <- 0
    total_qty <- 0
    for (i in seq_len(nrow(relevant))) {
      if (quantity_needed <= 0) break
      use_qty <- min(quantity_needed, remaining[i])
      unit_cost <- as.numeric(relevant[[price_col]][i])
      total_cost <- total_cost + use_qty * unit_cost
      total_qty <- total_qty + use_qty
      quantity_needed <- quantity_needed - use_qty
    }
    if (total_qty == 0) return(NA_real_)
    total_cost / total_qty
  }
  
  # SALES PROFIT DATA
  sales_profit_data <- reactive({
    s <- sales_log()
    cp <- cost_per_unit_table()
    if (is.null(s) || nrow(s) == 0 || is.null(cp) || nrow(cp) == 0) return(NULL)
    merged <- left_join(s, cp, by = "Item")
    merged <- merged %>% mutate(total_profit = (Price - cost_per_unit) * Quantity)
    merged
  })
  
  output$total_profit_text <- renderText({
    df <- sales_profit_data()
    if (is.null(df) || nrow(df) == 0) return("Total Profit: Rp 0")
    total_profit <- sum(df$total_profit, na.rm = TRUE)
    paste0("Total Profit: Rp ", format(round(total_profit, 0), big.mark = ","))
  })
  
  output$profit_margin_text <- renderText({
    df <- sales_profit_data()
    if (is.null(df) || nrow(df) == 0) return("Profit Margin: N/A")
    total_revenue <- sum(df$Price * df$Quantity, na.rm = TRUE)
    total_profit <- sum(df$total_profit, na.rm = TRUE)
    margin <- ifelse(total_revenue == 0, 0, total_profit / total_revenue)
    paste0("Profit Margin: ", round(margin * 100, 2), " %")
  })
  
  output$top_profit_item <- renderText({
    df <- sales_profit_data()
    if (is.null(df) || nrow(df) == 0) return("Top Profit Item: N/A")
    top <- df %>% group_by(Item) %>% summarise(total_profit = sum(total_profit, na.rm = TRUE), .groups = "drop") %>% arrange(desc(total_profit)) %>% slice(1)
    paste0("Top Profit Item: ", top$Item)
  })
  
  # CURRENT STOCK (from storage and sales)
  current_stock <- reactive({
    st <- storage_log()
    s <- sales_log()
    if (is.null(st) || nrow(st) == 0) {
      return(data.frame(Item = character(), Quantity = numeric()))
    }
    if (!"Stock_Added" %in% names(st) && "Quantity" %in% names(st)) st$Stock_Added <- st$Quantity
    st_sum <- st %>% group_by(Item) %>% summarise(Stocked = sum(Stock_Added, na.rm = TRUE), .groups = "drop")
    if (is.null(s) || nrow(s) == 0) {
      return(st_sum %>% transmute(Item = Item, Quantity = Stocked))
    }
    s_sum <- s %>% group_by(Item) %>% summarise(Sold = sum(Quantity, na.rm = TRUE), .groups = "drop")
    cur <- full_join(st_sum, s_sum, by = "Item") %>%
      mutate(Stocked = replace_na(Stocked, 0), Sold = replace_na(Sold, 0), Quantity = Stocked - Sold) %>%
      filter(Quantity > 0) %>% select(Item, Quantity)
    cur
  })
  
  # TOP / LOW / MISC TEXTS
  output$fastest_selling_item <- renderText({
    df <- sales_log()
    if (is.null(df) || nrow(df) == 0) return("Fastest Selling Item: N/A")
    top <- df %>% group_by(Item) %>% summarise(total_qty = sum(Quantity, na.rm = TRUE), .groups = "drop") %>% arrange(desc(total_qty)) %>% slice(1)
    paste0("Fastest Selling Item: ", top$Item, " (", top$total_qty, ")")
  })
  
  output$overstocked_item <- renderText({
    cur <- current_stock()
    if (is.null(cur) || nrow(cur) == 0) return("Most Overstocked Item: N/A")
    thr <- median(cur$Quantity, na.rm = TRUE) * 2
    over <- cur %>% filter(Quantity > thr)
    if (nrow(over) == 0) return("No significantly overstocked items")
    top <- over %>% arrange(desc(Quantity)) %>% slice(1)
    paste0("Most Overstocked Item: ", top$Item, " (", top$Quantity, " — over ", round(thr), ")")
  })
  
  output$unsold_items <- renderPrint({
    st <- storage_log()
    s <- sales_log()
    if (is.null(st) || nrow(st) == 0) return("Unsold Items: N/A")
    stocked_items <- unique(st$Item)
    sold_items <- if (is.null(s) || nrow(s) == 0) character(0) else unique(s$Item)
    unsold <- setdiff(stocked_items, sold_items)
    if (length(unsold) == 0) cat("All items have been sold at least once.") else {
      cat(paste0("Unsold Items (", length(unsold), "):\n"))
      for (it in unsold) {
        qty <- sum(st$Stock_Added[st$Item == it], na.rm = TRUE)
        cat(paste0("- ", it, " (", qty, " in stock)\n"))
      }
    }
  })
  
  # stock summary for goods() if available
  stock_summary_table <- reactive({
    df <- goods()
    if (is.null(df) || nrow(df) == 0) return(data.frame())
    if (!"Stock" %in% names(df) && "Quantity" %in% names(df)) df <- df %>% rename(Stock = Quantity)
    if (!"Price" %in% names(df) && "UnitPrice" %in% names(df)) df$Price <- df$UnitPrice
    df %>% group_by(Item, Price) %>% summarise(Stock = sum(Stock, na.rm = TRUE), .groups = "drop")
  })
  
  # restock suggestions (from goods_data)
  restock_suggestions <- reactive({
    gd <- goods()
    if (is.null(gd) || nrow(gd) == 0) return(data.frame())
    if (!"Stock" %in% names(gd) && "Quantity" %in% names(gd)) gd <- gd %>% rename(Stock = Quantity)
    gd %>% filter(Stock < 5) %>% transmute(Goods = ifelse("Goods" %in% names(.), Goods, Item), Stock)
  })
  
  # unit_sold_table uses cogs_data (computed below)
  unit_sold_table <- reactive({
    gd <- goods()
    cogsd <- cogs_data()
    if (is.null(gd) || nrow(gd) == 0 || is.null(cogsd) || nrow(cogsd) == 0) return(data.frame())
    cogs_details <- cogsd[cogsd$Type == "COGS Detail", ]
    # normalise names for matching
    if (!"Date_Added" %in% names(gd) && "Date" %in% names(gd)) gd$Date_Added <- gd$Date
    gd$Unit_Sold <- 0
    for (i in seq_len(nrow(gd))) {
      row <- gd[i, ]
      matches <- cogs_details %>% filter(Item == row$Item & Price == row$Price & as.character(Date) == as.character(row$Date_Added))
      gd$Unit_Sold[i] <- sum(-1 * matches$Quantity, na.rm = TRUE)
    }
    gd$Date_Added <- format(as.POSIXct(gd$Date_Added), "%Y-%m-%d %H:%M:%S")
    gd %>% transmute(Date_Added, Item, Price, Stock, Unit_Sold)
  })
  
  item_stock_summary_table <- reactive({
    gd <- goods()
    if (is.null(gd) || nrow(gd) == 0) return(NULL)
    if (!"Stock" %in% names(gd) && "Quantity" %in% names(gd)) gd <- gd %>% rename(Stock = Quantity)
    gd %>% group_by(Item) %>% summarise(Stock = sum(Stock, na.rm = TRUE), .groups = "drop")
  })
  
  # COGS builder (FIFO approach) — returns combined frame
  cogs_data <- reactive({
    sales <- sales_log()
    storage <- storage_log()
    if (is.null(sales) || is.null(storage) || nrow(sales) == 0 || nrow(storage) == 0) return(NULL)
    
    # ensure storage has Stock_Added
    if (!"Stock_Added" %in% names(storage) && "Quantity" %in% names(storage)) storage$Stock_Added <- storage$Quantity
    
    tax_rate <- 0.1
    combined <- list()
    
    # stock added rows
    for (i in seq_len(nrow(storage))) {
      row <- storage[i, ]
      combined[[length(combined) + 1]] <- data.frame(
        Date = row$Date,
        Item = row$Item,
        Type = "Stock Added",
        Quantity = row$Stock_Added,
        Price = row$Price,
        Revenue = NA_real_,
        `Cost per Unit` = round(row$Price, 2),
        COGS = NA_real_,
        Tax = NA_real_,
        `Net Profit` = NA_real_,
        stringsAsFactors = FALSE,
        check.names = FALSE
      )
    }
    
    # for each item, process sales in chronological order and allocate storage FIFO
    for (it in unique(c(sales$Item, storage$Item))) {
      item_sales <- sales %>% filter(Item == it) %>% arrange(as.POSIXct(Date))
      item_storage <- storage %>% filter(Item == it) %>% arrange(as.POSIXct(Date))
      
      if (nrow(item_sales) == 0) next
      
      for (si in seq_len(nrow(item_sales))) {
        sale_row <- item_sales[si, ]
        sale_date <- as.POSIXct(sale_row$Date)
        usable_storage <- item_storage %>% filter(as.POSIXct(Date) <= sale_date)
        if (nrow(usable_storage) == 0) {
          # no prior stock -> mark sale but can't compute COGS
          combined[[length(combined) + 1]] <- data.frame(
            Date = sale_row$Date,
            Item = sale_row$Item,
            Type = "Sale",
            Quantity = sale_row$Quantity,
            Price = sale_row$Price,
            Revenue = round(sale_row$Total, 2),
            `Cost per Unit` = NA_real_,
            COGS = NA_real_,
            Tax = round(sale_row$Total * tax_rate, 2),
            `Net Profit` = NA_real_,
            stringsAsFactors = FALSE,
            check.names = FALSE
          )
          next
        }
        qty_needed <- sale_row$Quantity
        storage_copy <- usable_storage$Stock_Added
        ptr <- 1
        cogs_total <- 0
        while (qty_needed > 0 && ptr <= length(storage_copy)) {
          avail <- storage_copy[ptr]
          if (avail <= 0) {
            ptr <- ptr + 1
            next
          }
          take_qty <- min(qty_needed, avail)
          unit_cost <- usable_storage$Price[ptr]
          cogs_part <- take_qty * unit_cost
          combined[[length(combined) + 1]] <- data.frame(
            Date = sale_row$Date,
            Item = sale_row$Item,
            Type = "COGS Detail",
            Quantity = -take_qty,
            Price = unit_cost,
            Revenue = NA_real_,
            `Cost per Unit` = round(unit_cost, 2),
            COGS = round(cogs_part, 2),
            Tax = NA_real_,
            `Net Profit` = NA_real_,
            stringsAsFactors = FALSE,
            check.names = FALSE
          )
          cogs_total <- cogs_total + cogs_part
          storage_copy[ptr] <- storage_copy[ptr] - take_qty
          qty_needed <- qty_needed - take_qty
        }
        total_cost <- sum(usable_storage$Stock_Added * usable_storage$Price, na.rm = TRUE)
        total_qty <- sum(usable_storage$Stock_Added, na.rm = TRUE)
        avg_cost <- ifelse(total_qty > 0, total_cost / total_qty, NA_real_)
        tax <- sale_row$Total * tax_rate
        net_profit <- sale_row$Total - cogs_total - tax
        combined[[length(combined) + 1]] <- data.frame(
          Date = sale_row$Date,
          Item = sale_row$Item,
          Type = "Sale",
          Quantity = sale_row$Quantity,
          Price = sale_row$Price,
          Revenue = round(sale_row$Total, 2),
          `Cost per Unit` = round(avg_cost, 2),
          COGS = round(cogs_total, 2),
          Tax = round(tax, 2),
          `Net Profit` = round(net_profit, 2),
          stringsAsFactors = FALSE,
          check.names = FALSE
        )
      }
    }
    
    final <- bind_rows(combined)
    final$Date <- as.POSIXct(final$Date)
    final <- final %>% arrange(Date)
    final
  })
  
  # sidebar totals quick
  output$total_sales_text <- renderPrint({
    df <- sales_log()
    if (is.null(df) || nrow(df) == 0) return("Total Sales Revenue: Rp 0")
    total <- sum(df$Total, na.rm = TRUE)
    paste0("Total Sales Revenue: Rp ", format(round(total, 0), big.mark = ","))
  })
  
  # BUY / ADD actions (keeping backward compatibility with earlier UI bits if present)
  observeEvent(input$buy_button, {
    req(input$buy_item, input$buy_quantity_input)
    gd <- goods()
    if (is.null(gd) || nrow(gd) == 0) {
      showNotification("No goods available for sale.", type = "error"); return()
    }
    idx <- which(gd$Item == input$buy_item)
    if (length(idx) == 0) { showNotification("Item not found", type = "error"); return() }
    total_stock <- sum(gd$Stock[idx], na.rm = TRUE)
    if (input$buy_quantity_input > total_stock) { showNotification("Not enough stock", type = "error"); return() }
    price <- ifelse(is.na(input$custom_buy_price) || is.null(input$custom_buy_price), gd$Price[idx[1]], input$custom_buy_price)
    total_price <- input$buy_quantity_input * price
    # append sale
    new_sale <- data.frame(Date = as.character(format(Sys.time(), "%Y-%m-%d %H:%M:%S")), Item = input$buy_item, Quantity = input$buy_quantity_input, Price = price, Total = total_price, stringsAsFactors = FALSE)
    sales_log(rbind(sales_log(), new_sale))
    # deduct from goods (first rows)
    for (i in idx) {
      if (input$buy_quantity_input <= 0) break
      used <- min(gd$Stock[i], input$buy_quantity_input)
      gd$Stock[i] <- gd$Stock[i] - used
      input$buy_quantity_input <- input$buy_quantity_input - used
    }
    goods(gd)
    showNotification("Sale recorded", type = "message")
  })
  
  observeEvent(input$add_btn, {
    if (is.null(input$new_item) || input$new_item == "" || input$new_stock <= 0) {
      showNotification("Please provide item, price and positive stock", type = "error"); return()
    }
    new_entry <- data.frame(Item = input$new_item, Price = input$new_price, Stock = input$new_stock, Date_Added = as.character(format(Sys.time(), "%Y-%m-%d %H:%M:%S")), stringsAsFactors = FALSE)
    gd <- goods()
    if (is.null(gd) || nrow(gd) == 0) goods(new_entry) else goods(bind_rows(gd, new_entry))
    # append storage log
    new_storage <- data.frame(Date = as.character(format(Sys.time(), "%Y-%m-%d %H:%M:%S")), Item = input$new_item, Stock_Added = input$new_stock, Price = input$new_price, Location = NA_character_, stringsAsFactors = FALSE)
    storage_log(bind_rows(storage_log(), new_storage))
    showNotification("New storage added", type = "message")
  })
  
  # PREVIEW / TABLE outputs
  output$goods_table <- DT::renderDataTable({
    df <- goods()
    if (is.null(df) || nrow(df) == 0) return(data.frame())
    
    DT::datatable(
      df,
      options = list(pageLength = 10, scrollX = TRUE),
      rownames = FALSE,
      class = "display nowrap cell-border stripe hover"
    )
  })
  
  # Stock bar plot
  output$stock_bar_plot <- renderPlot({
    cur <- current_stock()
    if (is.null(cur) || nrow(cur) == 0) return(NULL)
    ggplot(cur %>% arrange(desc(Quantity)) %>% slice_head(n = 30), aes(x = reorder(Item, Quantity), y = Quantity)) +
      geom_col(fill = "#0072B2") +
      coord_flip() +
      labs(title = "Current Stock (top 30)", x = "", y = "Quantity") +
      theme_minimal()
  })
  
  # sales_summary_table and cogs_summary table outputs
  output$sales_summary_table <- DT::renderDataTable({
    df <- cogs_data()
    if (is.null(df) || nrow(df) == 0) return(data.frame())
    
    df_sales <- df %>% filter(Type == "Sale")
    if (nrow(df_sales) == 0) return(data.frame())
    
    df_sales %>% mutate(Date = format(Date, "%Y-%m-%d %H:%M:%S"))
  }, options = list(scrollX = TRUE, pageLength = 5), rownames = FALSE)
  
  output$cogs_summary <- DT::renderDataTable({
    df <- cogs_data()
    if (is.null(df) || nrow(df) == 0) return(data.frame())
    
    df %>% mutate(Date = format(Date, "%Y-%m-%d %H:%M:%S"))
  }, options = list(scrollX = TRUE, pageLength = 5), rownames = FALSE)
  
  output$stock_added_table <- renderTable({
    df <- cogs_data()
    if (is.null(df) || nrow(df) == 0) return(data.frame())
    
    df_added <- df %>% filter(Type == "Stock Added")
    if (nrow(df_added) == 0) return(data.frame())
    
    df_added %>%
      mutate(
        Date = format(Date, "%Y-%m-%d %H:%M:%S"),
        Quantity = as.integer(Quantity)   # ✅ no decimals
      ) %>%
      transmute(
        Date,
        Item,
        Type,
        Quantity,
        Price,
        `Cost per Unit`,
        Inventory.Value = as.integer(Quantity * Price)
      )
  })
  
  # Top & worst selling
  output$top_selling_table <- renderTable({ top_selling_items() }, sanitize.text.function = function(x) x)
  output$worst_selling_table <- renderTable({ worst_selling_items() }, sanitize.text.function = function(x) x)
  
  top_selling_items <- reactive({
    df <- sales_log()
    if (is.null(df) || nrow(df) == 0) return(data.frame())
    df %>%
      group_by(Item) %>%
      summarise(Total_Quantity_Sold = as.integer(sum(Quantity, na.rm = TRUE)), Total_Revenue = sum(Total, na.rm = TRUE), .groups = "drop") %>%
      arrange(desc(Total_Quantity_Sold)) %>%
      slice_head(n = 5)
  })
  worst_selling_items <- reactive({
    df <- sales_log()
    if (is.null(df) || nrow(df) == 0) return(data.frame())
    df %>%
      group_by(Item) %>%
      summarise(Total_Quantity_Sold = as.integer(sum(Quantity, na.rm = TRUE)), Total_Revenue = sum(Total, na.rm = TRUE), .groups = "drop") %>%
      arrange(Total_Quantity_Sold) %>%
      slice_head(n = 5)
  })
  
  # Sales & profit charts (sliders)
  output$slider_sales_ui <- renderUI({
    df <- sales_log()
    if (is.null(df) || nrow(df) == 0) return(NULL)
    df$Date <- as.Date(as.POSIXct(df$Date))
    uniq <- sort(unique(df$Date))
    shinyWidgets::sliderTextInput("sales_date_slider", "Select Date Range (Sales)", choices = uniq, selected = c(min(uniq), max(uniq)), grid = TRUE, force_edges = TRUE)
  })
  
  output$sales_plot <- renderPlot({
    df <- sales_log()
    if (is.null(df) || nrow(df) == 0) return(NULL)
    df$Date <- as.Date(as.POSIXct(df$Date))
    if (!is.null(input$sales_date_slider)) {
      df <- df %>% filter(Date >= as.Date(input$sales_date_slider[1]) & Date <= as.Date(input$sales_date_slider[2]))
    }
    daily <- df %>% group_by(Date) %>% summarise(Total_Sales = sum(Total, na.rm = TRUE), .groups = "drop")
    ggplot(daily, aes(x = Date, y = Total_Sales)) + geom_line(color = "#0072B2") + geom_point() + labs(x = "Date", y = "Total Sales (Rp)") + theme_minimal() + scale_y_continuous(labels = scales::comma)
  })
  
  # profit margin slider + plot
  output$slider_profit_margin_ui <- renderUI({
    df <- cogs_data()
    if (is.null(df) || nrow(df) == 0) return(NULL)
    s <- df %>% filter(Type == "Sale")
    if (nrow(s) == 0) return(NULL)
    s$Date <- as.Date(s$Date)
    uniq <- sort(unique(s$Date))
    shinyWidgets::sliderTextInput("margin_date_slider", "Select Date Range (Profit Margin)", choices = uniq, selected = c(min(uniq), max(uniq)), grid = TRUE, force_edges = TRUE)
  })
  
  output$profit_margin_plot <- renderPlot({
    df <- cogs_data()
    if (is.null(df) || nrow(df) == 0) return(NULL)
    s <- df %>% filter(Type == "Sale")
    if (nrow(s) == 0) return(NULL)
    s$Date <- as.Date(s$Date)
    if (!is.null(input$margin_date_slider)) s <- s %>% filter(Date >= as.Date(input$margin_date_slider[1]) & Date <= as.Date(input$margin_date_slider[2]))
    s <- s %>% mutate(Profit_Margin = ifelse(Revenue > 0, `Net Profit` / Revenue, NA))
    daily <- s %>% group_by(Date) %>% summarise(Avg_Margin = mean(Profit_Margin, na.rm = TRUE), .groups = "drop")
    ggplot(daily, aes(x = Date, y = Avg_Margin)) + geom_line(color = "forestgreen") + geom_point(color = "forestgreen") + scale_y_continuous(labels = scales::percent) + labs(x = "Date", y = "Profit Margin") + theme_minimal()
  })
  
  # cumulative slider + plot
  output$slider_cumulative_ui <- renderUI({
    df <- cogs_data()
    if (is.null(df) || nrow(df) == 0) return(NULL)
    s <- df %>% filter(Type == "Sale")
    s$Date <- as.Date(s$Date)
    uniq <- sort(unique(s$Date))
    shinyWidgets::sliderTextInput("cumulative_date_slider", "Select Date Range (Cumulative)", choices = uniq, selected = c(min(uniq), max(uniq)), grid = TRUE, force_edges = TRUE)
  })
  
  output$cumulative_plot <- renderPlot({
    df <- cogs_data()
    if (is.null(df) || nrow(df) == 0) return(NULL)
    s <- df %>% filter(Type == "Sale")
    s$Date <- as.Date(s$Date)
    if (!is.null(input$cumulative_date_slider)) s <- s %>% filter(Date >= as.Date(input$cumulative_date_slider[1]) & Date <= as.Date(input$cumulative_date_slider[2]))
    daily <- s %>% group_by(Date) %>% summarise(Revenue = sum(Revenue, na.rm = TRUE), Profit = sum(`Net Profit`, na.rm = TRUE), .groups = "drop") %>% arrange(Date) %>% mutate(Cumulative_Revenue = cumsum(Revenue), Cumulative_Profit = cumsum(Profit))
    ggplot(daily, aes(x = Date)) + geom_line(aes(y = Cumulative_Revenue, color = "Revenue")) + geom_line(aes(y = Cumulative_Profit, color = "Profit")) + labs(y = "Rp") + theme_minimal() + scale_y_continuous(labels = scales::comma)
  })
  
  # Item selector, log selectors
  output$item_selector <- renderUI({
    gd <- goods()
    if (is.null(gd) || nrow(gd) == 0) return(NULL)
    items <- unique(gd$Item)
    selectInput("category_filter", "Item:", choices = c("All", sort(items)), selected = "All")
  })
  
  output$stock_summary <- renderTable({
    storage_df <- storage_log()
    sales_df <- sales_log()
    goods_df <- goods()
    
    # Safety check
    if (is.null(storage_df) && is.null(sales_df) && is.null(goods_df)) {
      return(data.frame(Item = character(), Stock = numeric()))
    }
    
    # --- Prepare storage data (inbound) ---
    if (!is.null(storage_df) && nrow(storage_df) > 0) {
      storage_df <- storage_df %>%
        mutate(Item = as.character(Item),
               Stock_Added = as.numeric(Stock_Added)) %>%
        group_by(Item) %>%
        summarise(Added = sum(Stock_Added, na.rm = TRUE), .groups = "drop")
    } else {
      storage_df <- data.frame(Item = character(), Added = numeric())
    }
    
    # --- Prepare sales data (outbound) ---
    if (!is.null(sales_df) && nrow(sales_df) > 0) {
      sales_df <- sales_df %>%
        mutate(Item = as.character(Item),
               Quantity = as.numeric(Quantity)) %>%
        group_by(Item) %>%
        summarise(Sold = sum(Quantity, na.rm = TRUE), .groups = "drop")
    } else {
      sales_df <- data.frame(Item = character(), Sold = numeric())
    }
    
    # --- Combine both to get live stock ---
    stock_summary <- full_join(storage_df, sales_df, by = "Item") %>%
      mutate(
        Added = replace_na(Added, 0),
        Sold = replace_na(Sold, 0),
        Stock = Added - Sold
      ) %>%
      select(Item, Stock) %>%
      arrange(Item)
    
    # --- Merge with goods info to include category (if available) ---
    if (!is.null(goods_df) && "Category" %in% names(goods_df)) {
      stock_summary <- left_join(stock_summary, goods_df %>% select(Item, Category), by = "Item") %>%
        distinct(Item, .keep_all = TRUE)
    }
    
    # --- Apply category filter (if input exists) ---
    if (!is.null(input$category_filter) && input$category_filter != "All" &&
        "Category" %in% names(stock_summary)) {
      stock_summary <- stock_summary %>% filter(Category == input$category_filter)
    }
    
    stock_summary
  })
  
  output$low_stock_table <- renderTable({
    gd <- goods()
    if (is.null(gd) || nrow(gd) == 0) return(data.frame())
    if (!"Stock" %in% names(gd) && "Quantity" %in% names(gd)) gd <- gd %>% rename(Stock = Quantity)
    gd %>% group_by(Item) %>% summarise(Stock = as.integer(sum(Stock, na.rm = TRUE)), .groups = "drop") %>% filter(Stock < 5)
  })
  
  output$log_item_selector <- renderUI({
    gd <- goods(); sl <- sales_log()
    items <- unique(c(if (!is.null(gd)) gd$Item else character(0), if (!is.null(sl)) sl$Item else character(0)))
    if (length(items) == 0) return(NULL)
    selectInput("log_item_filter", "Item:", choices = c("All", sort(items)), selected = "All")
  })
  
  output$sales_log_table <- renderDT({
    sl <- sales_log()
    if (is.null(sl) || nrow(sl) == 0) return(datatable(data.frame()))
    if (!is.null(input$log_item_filter) && input$log_item_filter != "All") sl <- sl %>% filter(Item == input$log_item_filter)
    datatable(sl, options = list(pageLength = 10), rownames = FALSE)
  })
  
  output$storage_log_table <- renderDT({
    st <- storage_log()
    if (is.null(st) || nrow(st) == 0) return(datatable(data.frame()))
    if (!is.null(input$log_item_filter) && input$log_item_filter != "All") st <- st %>% filter(Item == input$log_item_filter)
    datatable(st, options = list(pageLength = 10), rownames = FALSE)
  })
  
  # restock_table DT
  output$restock_table <- renderDT({
    rs <- restock_suggestions()
    if (is.null(rs) || nrow(rs) == 0) return(datatable(data.frame()))
    datatable(rs, options = list(pageLength = 10), rownames = FALSE)
  })
  
  # unit_sold_table render
  output$unit_sold_table <- renderTable({
    unit_sold_table()
  })
  
  # cogs_summary (human-friendly) already above as renderTable
  
  # Total stock value, low stock count, top selling item
  output$top_selling_item <- renderText({
    sl <- sales_log()
    if (is.null(sl) || nrow(sl) == 0) return("Top Selling Item: N/A")
    top <- sl %>% count(Item, sort = TRUE) %>% slice(1) %>% pull(Item)
    paste0("Top Selling Item: ", top)
  })
  output$low_stock_count <- renderText({
    gd <- goods()
    if (is.null(gd) || nrow(gd) == 0) return("Missing Stock or Item column")
    if (!"Stock" %in% names(gd) && "Quantity" %in% names(gd)) gd <- gd %>% rename(Stock = Quantity)
    df <- gd %>% group_by(Item) %>% summarise(total_stock = sum(Stock, na.rm = TRUE), .groups = "drop") %>% filter(total_stock < 5)
    if (nrow(df) == 0) "No items are low in stock." else paste0("Low in Stock (", nrow(df), " item(s))")
  })
  output$total_stock_value <- renderText({
    gd <- goods()
    if (is.null(gd) || nrow(gd) == 0) return("Missing Stock or Price column")
    if (!all(c("Stock", "Price") %in% names(gd))) {
      if ("Quantity" %in% names(gd)) gd <- gd %>% rename(Stock = Quantity)
      if ("UnitPrice" %in% names(gd)) gd$Price <- gd$UnitPrice
    }
    total_value <- sum((gd$Stock * gd$Price), na.rm = TRUE)
    paste0("Total Stock Value: Rp ", format(round(total_value, 0), big.mark = ","))
  })
  
  # low stock items reactive (used elsewhere)
  low_stock_items <- reactive({
    gd <- goods()
    if (is.null(gd) || nrow(gd) == 0) return(data.frame())
    if (!"Stock" %in% names(gd) && "Quantity" %in% names(gd)) gd <- gd %>% rename(Stock = Quantity)
    gd %>% group_by(Item) %>% summarise(Total_Stock = as.integer(sum(Stock, na.rm = TRUE)), .groups = "drop") %>% filter(Total_Stock < 5)
  })
  
  output$low_stock_table <- renderTable({ low_stock_items() })
  
  # Download CSV updated goods
  output$download_data <- downloadHandler(
    filename = function() paste0("updated_goods_", Sys.Date(), ".csv"),
    content = function(file) {
      gd <- tryCatch(goods(), error = function(e) data.frame())
      if (is.null(gd) || nrow(gd) == 0) {
        write.csv(data.frame(Message = "No data available"), file, row.names = FALSE)
      } else {
        write.csv(gd, file, row.names = FALSE)
      }
    }
  )
  
  # Download multi-sheet Excel
  output$download_excel <- downloadHandler(
    filename = function() paste0("sales_storage_summary_", Sys.Date(), ".xlsx"),
    content = function(file) {
      s <- tryCatch(sales_log(), error = function(e) data.frame())
      st <- tryCatch(storage_log(), error = function(e) data.frame())
      gd <- tryCatch(goods(), error = function(e) data.frame())
      # ensure column names fit expected output
      if (!"Stock_Added" %in% names(st) && "Quantity" %in% names(st)) st <- st %>% rename(Stock_Added = Quantity)
      if ("Date" %in% names(s)) s$Date <- as.character(s$Date)
      if ("Date" %in% names(st)) st$Date <- as.character(st$Date)
      if ("Date_Added" %in% names(gd)) gd$Date_Added <- as.character(gd$Date_Added)
      # other tables
      stock_summary <- tryCatch(stock_summary_table(), error = function(e) data.frame())
      unit_sold <- tryCatch(unit_sold_table(), error = function(e) data.frame())
      item_stock <- tryCatch(item_stock_summary_table(), error = function(e) data.frame())
      cogs_sum <- tryCatch(cogs_data(), error = function(e) data.frame())
      sheets <- list(
        sales_data = s,
        storage_data = st,
        goods_data = gd,
        current_stock_summary = stock_summary,
        units_sold_by_batch = unit_sold,
        current_stock_per_item = item_stock,
        cogs_and_inventory_summary = cogs_sum
      )
      writexl::write_xlsx(sheets, path = file)
    }
  )
  
  # ====================================
  # Download PDF (uses sales_report.Rmd if available)
  # ====================================
  output$download_pdf <- downloadHandler(
    filename = function() paste0("sales_report_", Sys.Date(), ".pdf"),
    content = function(file) {
      tempReport <- tempfile(fileext = ".Rmd")
      
      if (file.exists("sales_report.Rmd")) {
        # Use existing custom Rmd
        file.copy("sales_report.Rmd", tempReport, overwrite = TRUE)
      } else {
        # Create a minimal Rmd dynamically if missing
        cat(
          "---
title: \"Sales Report\"
output: pdf_document
params:
sales_data: NULL
storage_data: NULL
---

```{r setup, include=FALSE}
knitr::opts_chunk$set(echo = FALSE)
library(knitr)
Sales Summary
sd <- params$sales_data
st <- params$storage_data

if (is.null(sd) || nrow(sd) == 0) {
cat('No sales data available.')
} else {
kable(head(sd, 20))
}
",
      file = tempReport
        )
      }
    # Render the PDF report
    rmarkdown::render(
      input = tempReport,
      output_file = basename(file),
      output_dir = dirname(file),
      envir = new.env(parent = globalenv()),
      params = list(
        sales_data = sales_log(),
        storage_data = storage_log()
      ),
      quiet = TRUE
    )
    }
  )
  #  ====================================
  #    Reload Demo Excel File (Optional)
  #  ====================================
  observeEvent(input$reload_demo, {
    demo_file <- "sales_storage_summary_2025-09-17.xlsx"
    if (file.exists(demo_file)) {
      # Read and load demo Excel file
      data_list <- read_uploaded_excel(demo_file)
      
      if ("sales_data" %in% names(data_list) && "storage_data" %in% names(data_list)) {
        sales_log(as.data.frame(data_list[["sales_data"]], stringsAsFactors = FALSE))
        storage_log(as.data.frame(data_list[["storage_data"]], stringsAsFactors = FALSE))
        
        if ("goods_data" %in% names(data_list)) {
          goods(as.data.frame(data_list[["goods_data"]], stringsAsFactors = FALSE))
        }
        
        showNotification("Demo Excel loaded successfully.", type = "message")
      } else {
        showNotification("Demo file missing required sheets.", type = "error")
      }
    } else {
      showNotification("No demo Excel file found in app directory.", type = "error")
    }
  })
}
#====================================
#  Run Shiny App
#====================================
shinyApp(ui, server)