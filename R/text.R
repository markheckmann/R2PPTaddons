

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


# . ----
autosize_text_range <- function(shp) 
{
  has_text <- shp[["TextFrame2"]][["HasText"]] == -1
  if (has_text) {
    text_frame <- shp[["TextFrame2"]]
    autosize <- text_frame[["AutoSize"]] == 1   # msoAutoSizeTextToFitShape = 1
    if (autosize) {
      text_frame[["AutoSize"]] <- 0   #turn off an on to trigger autoresizing
      text_frame[["AutoSize"]] <- 1   # msoAutoSizeTextToFitShape = 1
      # Alterative:
      # tr <- shp[["TextFrame2"]][["TextRange"]]
      # tr[["Text"]] <- tr[["Text"]]    # trick to trigger autoresizing
    }
  }
  shp
}

update_all_autosize_text_ranges_current_slide <- function(ppt) 
{
  shps <- PPT.ShapesOnCurrentSlide(ppt)
  l <- lapply(shps, autosize_text_range)
  invisible(NULL)
}


#' Autosize textboxes
#' 
#' Sometimes text boxes which are set to autosize are not properly resized
#' when the file is opened. The autosizing is triggered by the function.
#' 
#' @param ppt   The ppt object as used in \pkg{R2PPT}.
#' @param slide_index Slides indexes for which to update textboxes (default `NULL` = all slides).
#' @export
#' 
PPT.UpdateAutosizedTextboxes <- function(ppt, slide_index = NULL) 
{
  n_slides <- ppt$pres[["Slides"]][["Count"]]
  if (is.null(slide_index)) {
    slide_index <- seq_len(n_slides)
  }
  for (i in slide_index) {
    ppt <- PPT.UpdateCurrentSlide(ppt, i = i)
    update_all_autosize_text_ranges_current_slide(ppt)
  }
  ppt
}

