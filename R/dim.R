#' Get width and height of active slide
#'
#' @param ppt  An \pkg{R2PPT} presentation object.
#' @export
#' @rdname slide_dim
#' @examples \dontrun{
#'  p <- PPT.Init(visible=TRUE)
#'  p <- PPT.AddBlankSlide(p)
#'  slide_width(p)
#'  slide_height(p)
#' }
#' 
slide_width <- function(ppt){
  ppt$pres[["PageSetup"]][["SlideWidth"]] 
}


#' @inheritParams slide_width
#' @export
#' @rdname slide_dim
#'
slide_height <- function(ppt){
  ppt$pres[["PageSetup"]][["SlideHeight"]]
}