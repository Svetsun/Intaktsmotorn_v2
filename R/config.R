# ===================== KONFIG =====================
# Central configuration for the Roban Shiny app.
# All environment-specific constants live here.
# Source this file at the top of app.R.

# Path to the Excel workbook used as the data store.
# Relative to the app working directory (project root when launched via shiny::runApp).
TARGET_XLSX <- file.path("data", "Base_data.xlsx")

# Revenue threshold above which the bonus percentage applies (SEK / month).
BONUS_THRESHOLD <- 100000

# Date columns per sheet — used by normalize_wb_dates() to coerce date fields
# after reading the workbook from disk.
DATE_COLS_MASTER <- list(
  "Konsulter"              = c("startdatum", "slutdatum"),
  "Uppdrag"                = c("startdatum", "slutdatum"),
  "Uppgift"                = c("startdatum", "slutdatum", "created_at"),
  "Tidrapportering"        = c("startdatum", "slutdatum", "created_at"),
  "TimprisHistory"         = c("created_at"),
  "GrundlonHistory"        = c("created_at"),
  "BonusHistory"           = c("created_at"),
  "GroupBonusHistory"      = c("created_at"),
  "SalesBonusHistory"      = c("created_at"),
  "Maklare"                = c("created_at"),
  "Kunder"                 = c("created_at"),
  "ArbetstimmarGrund"      = c("date"),
  "FaktureringInformation" = c()
)
