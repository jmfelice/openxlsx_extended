#' @title Add Data to Worksheet: Vertical
#'
#' @param .wb workbook
#' @param .df data.frame
#' @param .sheet character or numeric (default)
#' @param .col_names character
#' @param .title character
#' @param .header character
#' @param .indent_rows numeric
#' @param .indent_cols numeric
#' @param .bold_rows numeric
#' @param .bold_cols numeric
#' @param .cols_to_adjust_w numeric
#' @param .col_widths numeric
#' @param .h_border_rows numeric
#' @param .with_filter logical
#' @param .auto_width logical
#'
#' @return NULL
#' @export
#'
#'
add_data_to_worksheet_vertical <- function(wb,
                                           df,
                                           sheet            = 1,
                                           starting_col     = 1,
                                           indent_cols      = 1,
                                           indent_rows      = NULL,
                                           bold_cols        = 1,
                                           bold_rows        = NULL,
                                           cols_to_adjust_w = 1,
                                           col_widths       = 25,
                                           with_filter      = FALSE,
                                           auto_width       = TRUE,
                                           col_names        = NULL,
                                           title            = NULL,
                                           header           = NULL) {

  # identify sheet
  if (is.character(sheet)) {sheet <- find_worksheet_index(wb, sheet) }
  wb_sheet <- wb$worksheets[[sheet]]

  # set row count
  # identify if any other tables exist in the spreadsheet
  # if so, the starting row is equal to the last row used, plus 2
  # so we can maintain a space between tables.
  # Otherwise the first row used should be row 1
  data_count <- wb_sheet$sheet_data$data_count
  r <- ifelse(data_count > 0, max(wb_sheet$sheet_data$rows) + 2, 1)

  # add title
  if (!is.null(title)) {
    ## add title if requesed
    openxlsx::writeData(wb, sheet, title, startCol = 1, startRow = r)
    openxlsx::mergeCells(wb, sheet, cols = 1:ncol(.df), rows = r)
    openxlsx::addStyle(wb, sheet, title_sty, cols = 1:ncol(df), rows = r)
    r <- r + 1
  }

  # add headers
  if (!is.null(header)) {
    hdr <- t(as.data.frame(header))
    openxlsx::writeData(
      wb,
      sheet,
      hdr,
      colNames = FALSE,
      startCol = 1,
      startRow = r
    )
    vctrs <- rle_vectors(c(hdr))
    sapply(vctrs, function(x) {
      openxlsx::mergeCells(wb, sheet, cols = x, rows = r)
    })
    openxlsx::addStyle(
      wb,
      sheet,
      header_sty,
      cols = 1:ncol(df),
      rows = r
    )
    r <- r + 1
  }

  # table names
  if (!is.null(col_names)) {
    names(df) <- col_names
  }

  ## write data
  openxlsx::writeData(
    wb          = wb,
    sheet       = sheet,
    x           = df,
    headerStyle = col_names_sty,
    startCol    = starting_col,
    startRow    = r,
    borders     = "all",
    withFilter  = with_filter
  )

  ## indent rows
  if (!is.null(indent_rows)) {
    openxlsx::addStyle(
      wb    = wb,
      sheet = sheet,
      rows  = indent_rows + r,
      cols  = indent_cols,
      style = indent_sty,
      stack = TRUE
    )
  }

  ## bold rows
  if (!is.null(bold_rows)) {
    openxlsx::addStyle(
      wb    = wb,
      sheet = sheet,
      rows  = bold_rows + r,
      cols  = bold_cols,
      style = bold_sty,
      stack = TRUE
    )
  }

  ## column widths
  width      <- ifelse(auto_with, "auto", col_widths)
  width_cols <- ifelse(auto_width, 1:ncol(df), cols_to_adjust_w)
  openxlsx::setColWidths(
    wb                = wb,
    sheet             = sheet,
    cols              = width_cols,
    widths            = width,
    ignoreMergedCells = TRUE
  )
}

