#' Update the current slide stored in R2PPT object
#'
#' R2PPT uses an object to store the current slide amongst other things. 
#' Unfortunately the current slide is only set when a new slide is inserted. It
#' is NULL when a file is loaded. This will cause errors sometimes, hence we may
#' need to set it manually.
#' 
#' @param ppt   The ppt object as used in \pkg{R2PPT}.
#' @param i     Slide index.
#' @param slide A slide object as alternative to setting the index.
#' @author Mark Heckmann
#' @export
#'
PPT.UpdateCurrentSlide <- function(ppt, i=NULL, slide=NULL)
{
  if (!is.null(i))
    slide <- ppt$pres[["Slides"]]$Item(i)
  ppt$Current.Slide <- slide
  ppt
}


#' Get pointers to all shapes on current slide
#' 
#' @param ppt The ppt object as used in \pkg{R2PPT}.
#' @export
#' @return List of pointers to shapes
#' 
PPT.ShapesOnCurrentSlide <- function(ppt)
{
  shapes <- ppt$Current.Slide[["Shapes"]]  
  l <- list()
  for ( i in 1:shapes[["Count"]] ) {
    l[[i]] <- shapes$Item(i)
  }
  l
}



# Get current slide
PPT.NumberOfCurrentSlide <- function(ppt)
{
  ppt$Current.Slide[["SlideNumber"]]
}



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


#' @export
#' @rdname slide_dim
#'
slide_height <- function(ppt){
  ppt$pres[["PageSetup"]][["SlideHeight"]]
}



# get position of shape
# (left, top, width, height)
#
get_shape_position <- function(shape)
{
  list(top = shape[["Top"]],
       left = shape[["Left"]],
       width = shape[["Width"]],
       height = shape[["Height"]]
  )
}


# get selected shape properties
#
get_shape_properties <- function(shape)
{
  list(ShapeName = shape[["Name"]],  # "ShapeType Number"
       ShapeId = shape[["Id"]],
       Type = shape[["Type"]],
       AutoShapeType = shape[["AutoShapeType"]]
       # HasTextFrame = shape[["HasTextFrame"]]
  )
}


# z -order -----------------------------------------------

# Shape.ZOrder method: 
#
# Moves the specified shape in front of or behind other shapes in the 
# collection (that is, changes the shape's position in the z-order).
#
# MsoZOrderCmd Enum:

# msoBringToFront	0 Bring shape to the front.
# msoSendToBack	1	Send shape to the back.
# msoBringForward	2	Bring shape forward.
# msoSendBackward	3 Send shape backward.
# msoBringInFrontOfText	4	Bring shape in front of text.
# msoSendBehindText 5 Send shape behind text.


# x: A pointer to a shape object
z_order <- function(x) {
  x[["ZOrderPosition"]]
}


# x: A pointer to a shape object.
# value: An integer corresponding to the MsoZOrderCmd Enum.
`z_order<-` <- function(x, value) {
  if (!value %in% 0:5) {
    warning("Value must correspond to MsoZOrderCmd Enum number.", call. = FALSE)
    return(NULL)
  }
  tryCatch({
    x$ZOrder(value)  
  }, error = function(e) {
    warning("Setting the z-order failed")
  })
}



# MISC ------------------------------------------------------


# Works but currently not needed.
# get pointers to all shapes on slide as a list
# slide:  pointer to slide
#
# get_slide_shape_pointers <- function(slide)
# {
#   shapes <- slide[["Shapes"]]               # get all shapes on slide
#   nshapes <- shapes[["Count"]]              # number of shapes
#   l <- list()
#   if (nshapes == 0)
#     return(l)
#   for (i in 1L:nshapes) {
#     l[[i]] <- shapes$Item(i)
#   }
#   l
# }




#' Generate some png graphics
#'
#' @param n      Number of graphics to be generated.
#' @param dev    dev of output device (\code{"png", "bmp", "jpeg"}).
#' @param ...    Arguments passed on to tze specific graphic 
#'               device function (\code{png, bmp, jpeg})