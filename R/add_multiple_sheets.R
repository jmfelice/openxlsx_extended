
add_multiple_sheets <- function(.wb, .sheet_names, ...) {
  a <- list("wb" = .wb, ...)
  lapply(.sheet_names, function(x) {
    do.call(openxlsx::addWorksheet, c(a, "sheetName" = x))
  })
}