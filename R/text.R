

#' Replace matching text by text
#'
#' Looks through all shapes and finds a shape with matching text pattern.
#' The text in the shape is replaced.
#'
#' @param ppt   The ppt object as used in \pkg{R2PPT}.
#' @param what  Text pattern to match against.  
#' @param replace  Text to replace pattern with.
#' @param ... Not evaluated.
#' @export
#' @example inst/examples/PPT.ReplaceTextByText.R
#'
PPT.ReplaceTextByText <- function(ppt, what, replace, ...)
{
  slides <- ppt$pres[["Slides"]]
  ss <- slides_retrieve_shapes(slides, what)   # get all shape objects that match text pattern 
  if (length(ss) == 0)
    warning("No shape with matching text pattern was not found.", call. = FALSE)
  if (length(ss) > 1)
    warning("More than one shape with matching text pattern found and replaced.", call. = FALSE)
  
  for (s in ss) {               # delete from last to first
    s[["Textframe"]][["TextRange"]]$Replace(FindWhat=what, ReplaceWhat = replace)
  }  
  ppt
}