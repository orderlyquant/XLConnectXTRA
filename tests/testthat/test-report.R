test_that("make_reportable(obj) adds all required attributes to obj", {
  expect_equal(
    all(
      c("name", "type", "incl_header", "dims") %in%
        names(attributes(make_reportable(iris, "iris", TRUE, c(1, 1))))
      ),
    TRUE
  )
})
