test_that("get_base_class() returns a length 1 character vector", {
  expect_equal(class(get_base_class(iris)), "character")
  expect_length(get_base_class(iris), 1)
})
