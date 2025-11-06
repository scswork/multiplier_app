
library(shiny)
library(rhandsontable)
library(dplyr)
library(tidyr)
library(openxlsx)
library(ggplot2)
library(readxl)

# Load Excel file safely
if (file.exists("Multiplier2021DB.xlsx")) {
  db <- read_excel("Multiplier2021DB.xlsx")
} else {
  db <- data.frame(Industry = character(), Type = character(), Variable = character(), Value = numeric())
}

ui <- fluidPage(
  titlePanel("Economic Impact Report"),
  
  sidebarLayout(
    sidebarPanel(
      selectInput("capex_industry", "Select CAPEX Industry",
                  choices = unique(db$Industry), selected = NULL),
      selectInput("opex_industry", "Select OPEX Industry",
                  choices = unique(db$Industry), selected = NULL),
      checkboxGroupInput("type_filter", "Filter by Type",
                         choices = unique(db$Type), selected = unique(db$Type)),
      checkboxGroupInput("variable_filter", "Filter by Variable",
                         choices = unique(db$Variable), selected = unique(db$Variable)),
      br(),
      h4("Enter 15-Year Investment Data"),
      rHandsontableOutput("input_table"),
      br(),
      actionButton("generate", "Generate Report"),
      br(),
      downloadButton("download_excel", "Download Excel Report")
    ),
    
    mainPanel(
      h4("Impact Report"),
      tableOutput("impact_table")
    )
  )
)

server <- function(input, output, session) {
  initial_data <- data.frame(Year = 1:15, Value = rep(0, 15), Type = rep("CAPEX", 15), stringsAsFactors = FALSE)
  
  output$input_table <- renderRHandsontable({
    rhandsontable(initial_data) %>%
      hot_col("Type", type = "dropdown", source = c("CAPEX", "OPEX")) %>%
      hot_col("Value", format = "0")
  })
  
  report_data <- reactiveVal(NULL)
  
  observeEvent(input$generate, {
    data <- hot_to_r(input$input_table)
    data$Value <- as.numeric(data$Value)
    
    capex_total <- sum(data$Value[data$Type == "CAPEX"], na.rm = TRUE)
    opex_total <- sum(data$Value[data$Type == "OPEX"], na.rm = TRUE)
    
    capex_db <- db %>% filter(Industry == input$capex_industry)
    opex_db <- db %>% filter(Industry == input$opex_industry)
    
    report <- full_join(capex_db, opex_db, by = c("Type", "Variable"), suffix = c("_CAPEX", "_OPEX")) %>%
      filter(Type %in% input$type_filter, Variable %in% input$variable_filter) %>%
      mutate(
        CAPEX_Impact = ifelse(grepl("Jobs", Variable),
                              (Value_CAPEX * (capex_total / 1e6)) / sum(data$Type == "CAPEX" & data$Value > 0),
                              Value_CAPEX * capex_total),
        OPEX_Impact = ifelse(grepl("Jobs", Variable),
                             (Value_OPEX * (opex_total / 1e6)) / sum(data$Type == "OPEX" & data$Value > 0),
                             Value_OPEX * opex_total)
      ) %>%
      mutate(
        Value_CAPEX = round(Value_CAPEX, 4),
        Value_OPEX = round(Value_OPEX, 4),
        CAPEX_Impact = ifelse(grepl("Jobs", Variable),
                              format(round(CAPEX_Impact), big.mark = ","),
                              paste0("$", format(round(CAPEX_Impact, 0), big.mark = ","))),
        OPEX_Impact = ifelse(grepl("Jobs", Variable),
                             format(round(OPEX_Impact), big.mark = ","),
                             paste0("$", format(round(OPEX_Impact, 0), big.mark = ",")))
      )
    
    report_data(report)
    output$impact_table <- renderTable(report)
  })
  
  output$download_excel <- downloadHandler(
    filename = function() { "impact_report.xlsx" },
    content = function(file) { write.xlsx(report_data(), file) }
  )
}

shinyApp(ui, server)
