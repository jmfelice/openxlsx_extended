#' Verify Worksheet Name
#'
#' @param .wb workbook
#' @param .name character
#'
verify_sheet_name <- function(wb, name) {
  if (!name %in% wb$sheet_names) { stop("sheet is not a worksheet name!") }
}

#' Find Worksheet Name
#'
#' @param .wb workbook
#' @param .index numeric
#'
find_worksheet_name <- function(wb, index) {
  if (!is.numeric(index)) { stop("sheet is not numeric!") }
  wb$sheet_names[index]
}

#' Find Worksheet Index
#'
#' @param .wb workbook
#' @param .name character
#'
find_worksheet_index <- function(.b, name) {
  if (!is.character(name))  { stop("sheet is not character!") }
  verify_sheet_name(wb, name)
  which(wb$sheet_names %in% name)
}