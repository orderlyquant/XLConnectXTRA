#' Add attributes required for "reporting" workflow
#'
#' Returns a "reportable" object. A "reportable" object has attributes that
#' allow it to be properly reported in Excel via `report_reportables()`.
#'
#' @param obj Vector of length 1, data.frame or ggplot object
#' @param name String used to create range names
#' @param incl_header Include header in Excel? If `TRUE` column names
#'   will be written to first row of the named range
#' @param dims Optional numeric vector of length 2 containing the dimensions
#'   (height, width) of the ggplot2 in inches.
#'
#' @returns
#' An object of the same type as `obj` containing all required attributes.
#'
#' @examples
#' # Subset mtcars and make reportable to a named range: "mtcars_data",
#' # including header
#' mtc <- head(mtcars)
#' mtc <- make_reportable(mtc, "mtcars_data", TRUE)
#'
#' @export
make_reportable <- function(
    obj, name, incl_header = TRUE, dims = NULL,
    os = NULL, scale = NULL
) {
  attr(obj, "name") <- name
  attr(obj, "type") <- get_base_class(obj)
  attr(obj, "incl_header") <- incl_header
  attr(obj, "dims") <- dims
  attr(obj, "os") <- os
  attr(obj, "scale") <- scale

  return(obj)
}
