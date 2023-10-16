\dontrun{

# open PPT template
file <- system.file("template.pptx", package = "R2PPTaddons")
ppt <- PPT.Open(file, method="RDCOMClient")


#### CASE 1: ADD GRAPHIC AT TEXTBOX ####

# add graphic to slides with matching text and remove text. Note that by default
# only text in text boxes is replaced. Text inside rectangles, for example, is
# not touched unless explicitly prompted. Note also that additional arguments
# are passed on to PPT.AddGraphicstoSlide2 to specify the position and size of
# the shape.
what <- "[[tag 1]]"
PPT.ReplaceTextByGraphic(ppt, what, file = "inst/image_1.png", width = .6)

# Note that the text appears twice and the graphic is inserted twice with a warning.
what <- "[[tag 2]]"
PPT.ReplaceTextByGraphic(ppt, what, file = "inst/image_2.png")



#### CASE 2: REPLACE SHAPE BY GRAPHIC ####

# replace rectangle with matching text by graphic and move to back
what <- "[[tag 1]]"
PPT.ReplaceShapeByGraphic(ppt, what, file = "inst/image_1.png", z.order = 1)

# Note  that additional arguments are passed on to PPT.FitGraphicIntoShape and 
# subsequently to PPT.AddGraphicstoSlide2, e.g. to specify the appearance of the image.
what <- "[[tag 2]]"
PPT.ReplaceShapeByGraphic(ppt, what, file = "inst/image_2.png", line.size=1)

}




