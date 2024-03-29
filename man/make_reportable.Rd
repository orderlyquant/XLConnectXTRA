% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/report.R
\name{make_reportable}
\alias{make_reportable}
\title{Add attributes required for "reporting" workflow}
\usage{
make_reportable(
  obj,
  name,
  incl_header = TRUE,
  dims = NULL,
  os = FALSE,
  scale = NULL
)
}
\arguments{
\item{obj}{Vector of length 1, data.frame or ggplot object}

\item{name}{String used to create range names}

\item{incl_header}{Include header in Excel? If \code{TRUE} column names
will be written to first row of the named range}

\item{dims}{Optional numeric vector of length 2 containing the dimensions
(height, width) of the ggplot2 in inches.}
}
\value{
An object of the same type as \code{obj} containing all required attributes.
}
\description{
Returns a "reportable" object. A "reportable" object has attributes that
allow it to be properly reported in Excel via \code{report_reportables()}.
}
\examples{
# Subset mtcars and make reportable to a named range: "mtcars_data",
# including header
mtc <- head(mtcars)
mtc <- make_reportable(mtc, "mtcars_data", TRUE)

}
