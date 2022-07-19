get_base_class <- function(x) {
  cls <- class(x)

  return(
    cls[length(cls)]
  )
}
