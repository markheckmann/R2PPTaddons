\dontrun{

# create new PPT object and add one slide
p <- PPT.Init(visible=T, method = "RDCOMClient")
p <- PPT.AddBlankSlide(p)

# add two rectangle shapes to slide
p <- PPT.AddRectangle(p, width = .4, left=.1)                      # add a shape to slide
p <- PPT.AddRectangle(p, width = .4, left=.5, height=.45, top=.5)  # add a shape to slide

# get all shapes on current slide and replace by image
# shapes are not destroyed to see the image placement
s <- PPT.ShapesOnCurrentSlide(p)   
file <- system.file("image_1.png", package = "R2PPTaddons")  # get image
p <- PPT.FitGraphicIntoShape(p, file, shp=s[[1]], vjust="top", delete.shape =FALSE)
p <- PPT.FitGraphicIntoShape(p, file, shp=s[[2]], vjust="bottom", delete.shape =FALSE)

}




