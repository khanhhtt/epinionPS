#' Get path to epinionRWeighting example
#'
#' epinionRWeighting comes bundled with a number of sample files in its `inst/extdata`
#' directory. This function make them easy to access
#'
#' @param file Name of file. If `NULL`, the example files will be listed.
#' @export
#' @examples
#' example()
#' example("sample_test.xlsx")

example <- function(file = NULL) {
  if (is.null(file)) {
    dir(system.file("extdata", package = "epinionPS"))
  } else {
    system.file("extdata", file, package = "epinionPS", mustWork = TRUE)
  }
}
