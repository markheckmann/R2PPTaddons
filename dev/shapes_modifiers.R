
## Hide / unhide all shapes with mathcing pattern


# add graphic to slides with matching and remoce text
library(R2PPTaddons)

# Search shape and replace with graphic
devtools::load_all(".")



#' Get pointers to all shapes in presentation
#' 
#' @param ppt   The ppt object as used in \pkg{R2PPT}.
#' 
#' @return List of pointers to shape object
#' 
# shapes <- function()
#   
# 
# PPT.GetShapes <- function(ppt) {
#   
# }



get_slides <- function(ppt) 
{
  ppt$pres[["Slides"]]  
}


get_shapes_on_slide <- function(slide) 
{
  l <- list()
  shapes <- slide[["Shapes"]]  
  for (i in 1:shapes[["Count"]] ) {
    l[[i]] <- shapes$Item(i)
  }
  l
}


get_shapes <- function(ppt)
{
  slides <- get_slides(ppt)
  nslides <- slides[["Count"]]
  
  l <- list()
  for (i in 1L:nslides) {
    sld <- slides$Item(i)
    l <- append(l, get_shapes_on_slide(sld))
  }  
  class(l) <- "ppt_shape_list"
  l
}


print.ppt_shape_list <- function(x, ...) 
{
  cat("List of pointers to Shape objects: n =", length(x))
}


shape_info <- function(l)
{
  # get shapes properties for data frame
  slide_num <- get_slide_number(l)
  shp_type <- sapply(l, `[[`, "Type")  # Type property
  shp_name <- sapply(l, `[[`, "Name")  # Name property
  
  mso <- enumeration$MsoShapeType
  d <- tibble(slide = slide_num, 
              name = shp_name, 
              type = shp_type)
  d %<>% left_join(mso, by = c("type" = "Value"))
  d$pointer <- l
  d
}


visible <- function(x) {
  sapply(x, `[[`, "visible")
}

`visible<-` <- function(x, value) 
{
  for (shp in x) {
    shp[["visible"]] <- value
  }
  x
}


# l : shp_list
get_slide_number <- function(l)
{
  sapply(l, function(shp) {
    shp[["Parent"]][["SlideNumber"]]   
  })
}


ppt <- PPT.Open("inst/template.pptx", method = "RDCOMClient")
l <- get_shapes(ppt)
shape_info(l)

visible(l)
visible(l[2:3]) <- 0





