
# Attribut: Shape.ZOrderPosition (read only)
# Method: Shape.ZOrder-Methode
#
# MsoZOrderCmd enumertaion:
#   msoBringForward	       2	 Bring shape forward
#   msoBringInFrontOfText	 4	 Bring shape in front of text
#   msoBringToFront	       0	 Bring shape to the front
#   msoSendBackward	       3	 Send shape backward
#   msoSendBehindText	     5	 Send shape behind text
#   msoSendToBack	         1	 Send shape to the back


#' Set z-order of shape
#' 
#' @param z Shape pointer.
#' @param zcmd MsoZOrderCmd enumeration determines change in z-order. 
#' Either a number or a text abbreviation (e.g. \code{"front"}).
#' @param zposition Set z-position directly. This is useful to restore 
#' a shapes z-position if a shape is replaced. 
#' 
#' \tabular{ll}{
#' 0, "front" \tab Bring shape to the front \cr
#' 1, "back" \tab Send shape to the back \cr
#' 2, "forward" \tab Bring shape forward \cr
#' 3, "backward" \tab	Send shape backward \cr
#' 4 \tab	Bring shape in front of text \cr
#' 5 \tab	Send shape behind text
#' }
#' @return \code{NULL}
#' @keywords internal
#' @export
#' 
set_shape_zorder <- function(s, zcmd = 0, zposition = NULL)
{
  
  n_shps<- s[["Parent"]][["Shapes"]][["Count"]]
  zpos <- s[["ZOrderPosition"]] 
  
  n <- shps_on_slide[["Count"]]
  
  # convert to numeric MsoZOrderCmd
  if (is.character(zcmd)) {
    zcmd <- tolower(zcmd)
    zchar <- c("front", "back", "forward", "backward")
    zcmd <- match.arg(zcmd, zchar)
    zcmd <- match(zcmd, zchar) - 1
  }
  
  # invole change of order    
  s$ZOrder(ZOrderCmd = zcmd)
  invisible(NULL)
}




library(R2PPT)
devtools::load_all(".")


#### textframe ####

# Presentation
p <- PPT.Init(visible=T, method = "RDCOMClient")

p <- PPT.AddBlankSlide(p)
# add a rectangle
p <- PPT.AddShape(p, shape.type= 1, height=.4, width =.4, left=.0, top=.0, fill.color = "red")
# add a rounded rectangle
p <- PPT.AddShape(p, shape.type= 5, height=.4, width =.4, left=.15, top=.15, fill.color = "green")
# add a triangle 
p <- PPT.AddShape(p, shape.type= 7, height=.4, width =.4, left=.3, top=.3)

# move
ss <- PPT.ShapesOnCurrentSlide(p)
get_shape_properties(ss[[1]])$zorder
set_shape_zorder(ss[[1]], 0)









