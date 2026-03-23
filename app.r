source("renv/activate.R")

###############################################
# app.R — Förenklad Excel-struktur

rm(list = ls())
graphics.off()
cat("\014")


library(shiny)
library(readxl)
library(writexl)
library(janitor)
library(dplyr)
library(lubridate)
library(stringr)
library(rhandsontable)

# ===================== KONFIG =====================

source("R/config.R",               encoding = "UTF-8")
source("R/helpers_core.R",         local = environment(), encoding = "UTF-8")
source("R/helpers_ids_labels.R",   local = environment(), encoding = "UTF-8")
source("R/helpers_hot.R",          local = environment(), encoding = "UTF-8")
source("R/helpers_history.R",      local = environment(), encoding = "UTF-8")

# ===================== ENSURE SHEETS =====================

source("R/helpers_ensure.R",       local = environment(), encoding = "UTF-8")
source("R/helpers_report.R",       local = environment(), encoding = "UTF-8")
source("R/interval_report.R",      local = environment(), encoding = "UTF-8")

# ===================== DISPLAY HELPERS =====================

source("R/helpers_display.R",      local = environment(), encoding = "UTF-8")

# ===================== UI =====================

source("R/ui.R",                   encoding = "UTF-8")

# ===================== SERVER =====================

source("R/server_refresh.R",       local = environment(), encoding = "UTF-8")
source("R/server_add_handlers.R",  local = environment(), encoding = "UTF-8")
source("R/server.R",               encoding = "UTF-8")

shinyApp(ui, server)
