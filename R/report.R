make_reportable <- function(obj, name, incl_header = NULL, dims = NULL) {
  attr(obj, "name") <- name
  attr(obj, "type") <- get_base_class(obj)
  attr(obj, "incl_header") <- incl_header
  attr(obj, "dims") <- dims

  return(obj)
}
