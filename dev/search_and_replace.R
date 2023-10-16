
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


slides <- p$pres[["Slides"]]
ss <- slides_retrieve_shapes(slides, "GRAFIK")   # get all shape objects that match text pattern 
#str(lapply(ss, get_type_properties))

file <- "inst/image_1.png"
PPT.AddGraphicstoSlide2(p, file, display.frame = T)
PPT.AddGraphicstoSlide2(p, file, left = .5, width = .3, 
                         display.frame = T, vjust="top", hjust="right", newslide = T)
PPT.AddGraphicstoSlide2(p, file, left = .05, width = .45, 
                         top=.1, height=.4, display.frame = T, newslide = F, vjust=0)


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





