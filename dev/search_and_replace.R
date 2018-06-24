
# add graphic to slides with matching and remoce text
library(R2PPTaddons)

ppt <- PPT.Open("inst/template.pptx", method="RDCOMClient")

what <- "[[tag 1]]"
PPT.ReplaceTextByGraphic(ppt, what, file = "inst/image_1.png")

# Note that the text appears twice and the graphic is inserted twice with a warning.
what <- "[[tag 2]]"
PPT.ReplaceTextByGraphic(ppt, what, file = "inst/image_2.png")


# Currently, a figure is added on any slide



# TODO
# Search shape and replace with graphic
devtools::load_all(".")


# shape object type property
#
# MsoShapeType:
# msoAutoShape	1	AutoForm
# msoPicture	13	Grafik
# msoTextBox	17	Textfeld
# msoTable	19	Tabelle

# Autoshape type
# # MsoAutoShapeType:
# msoShapeRectangle	1	Rechteck.
# msoShapeParallelogram	2	Parallelogramm.
# msoShapeRoundedRectangle	5	Abgerundetes Rechteck.


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


# Add an image and fit it inside a given rectangle shape
# x either one of "left", "center", "right" or a number between 0 (for top) and 1 (for bottom)
# y either one of "top", "center", "bottom" or a number between 0 (for top) and 1 (for bottom)


PPT.FitGraphicInShape2 <- function(ppt, 
                                   file, 
                                   shp,       # shape to place inside
                                   width=1,   # usually not necessary to change 
                                   height=1,
                                   x="center", # position of image inside shape
                                   y="center", 
                                   x.offset=0, # offset
                                   y.offset=0, 
                                   proportional=TRUE, 
                                   maxscale=1)
{
  pos <- get_shape_position(shp)
  sld <- shp[["Parent"]]        # get shape's slide
  ppt <- PPT.UpdateCurrentSlide(ppt, slide=sld)   # to insert graphic on correct slide
  
  
  # make calculations to allow exact positioning of graphic inside shape
  
  
  # fit graphic inside shape using standard graphic function
  PPT.AddGraphicstoSlide2(ppt, 
                          file, 
                          newslide=FALSE,
                          width = pos$width, 
                          height = pos$height,
                          x = x, 
                          y = y, 
                          x.offset = pos$left, 
                          y.offset = pos$top)
}




#### textframe ####

# Presentation
p <- PPT.Init(visible=T, method = "RDCOMClient")
p <- PPT.AddBlankSlide(p)

slides <- p$pres[["Slides"]]
ss <- slides_retrieve_shapes(slides, "GRAFIK")   # get all shape objects that match text pattern 
#str(lapply(ss, get_type_properties))

# get left, top, width, height
shp <- ss[[1]]
get_shape_position(shp)
#get_type_properties(shp)

ss <- slides_retrieve_shapes(slides, "GRAFIK")   # get all shape objects that match text pattern 
shp <- ss[[1]]
shp$Select()
file <- "inst/image_1.png"
PPT.FitGraphicInShape2(p, file, shp)
shp$Delete()





