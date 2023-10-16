
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


is_com_object <- function(x) {
  isS4(x) | inherits(x, "COMIDispatch")  
}


stop_if_not_is_com_object <- function(x) {
  if (!is_com_object(x)) {
    stop("You must supply a pointer to a COM object.")
  }
}


get_attribute <- function(x, att, default = NA) 
{
  stop_if_not_is_com_object(x)
  
  # opts <- options(show.error.messages = FALSE)
  # on.exit(options(opts))
  tryCatch({
    x[[att]]
  }, error = function(e) {
    warning(e)
    default
  })
}


ppt <- PPT.Open("inst/template.pptx", method = "RDCOMClient")
l <- get_shapes(ppt)
shape_info(l)

visible(l)
visible(l[2:3]) <- 0


get_attribute(s, "sx")

l <- get_shapes(ppt)
s <- l[[1]]
# shape attributes
atts <- c("id", "Name", "ZOrderPosition", "Width", "Height", 
          "Left", "Top", "Title", "Type", "Visible", 
          "HasTextFrame", "TextFrame", "TextFrame2", "Fill",
          "Parent", "Table")
att_list <- list()
for (att in atts) {
  obj <- get_attribute(s, att)
  att_list[att] <- list(obj)
} 


# draw shapes on side ------------------------------------



slides <- get_slides(ppt)
sld <- slides[[4]]
shapes <- get_shapes_on_slide(sld)
get_shape_position(shapes[[1]])

# draw slide and all shapes on it
# get slide

ppt$ppt[["Width"]]
ppt$ppt[["Height"]]
ppt$ppt[["ActivePresentation"]][["PageSetup"]][["SlideWidth"]]
ppt$ppt[["ActivePresentation"]][["PageSetup"]][["SlideHeight"]]

lapply(shapes, z_order)
s <- shapes[[3]]
z_order(s) <- 0
