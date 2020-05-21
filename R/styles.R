#' title_sty
#'
title_sty <- function() {
  openxlsx::createStyle(
    fontSize = 20,
    halign = "center",
    valign = "center",
    textDecoration = "bold",
    border = "TopBottomLeftRight"
  )
}

#' col_names_sty
#'
col_names_sty <- function() {
  openxlsx::createStyle(
    fontSize = 12,
    fontColour = "#ffffff",
    fgFill = "#ED7D31",
    halign = "center",
    valign = "center",
    textDecoration = "bold",
    border = "TopBottomLeftRight",
    wrapText = TRUE
  )
}

#' header_sty
#'
header_sty <- function() {
  openxlsx::createStyle(
    fontSize = 15,
    fontColour = "#ffffff",
    fgFill = "#b67e58",
    halign = "center",
    valign = "center",
    border = "TopBottomLeftRight",
    textDecoration = "bold",
    wrapText = TRUE
  )
}

#' indent_sty
#'
indent_sty <- function() { openxlsx::createStyle(indent = 1) }

#' bold_sty
#'
bold_sty <- function() { openxlsx::createStyle(textDecoration = "bold") }
