



library(R2PPT)
devtools::load_all(".")


 p <- PPT.Init(visible=T, method = "RDCOMClient")
 p <- PPT.AddBlankSlide(p)
 p <- PPT.AddShape(p, shape.type= 1, height=.4, width =.4, left=.0, top=.0, fill.color = "red")
 p <- PPT.AddShape(p, shape.type= 1, height=.4, width =.4, left=.1, top=.1, fill.color = "green", zcmd="back")
 p <- PPT.AddShape(p, shape.type= 1, height=.4, width =.4, left=.2, top=.2, zpos=-1)

 # change z-order of shape
 ss <- PPT.ShapesOnCurrentSlide(p)
 s <- ss[[1]]  # get first shape
 get_shape_properties(s)$zorder  # current z-position
 set_shape_zorder(s, "up")
 set_shape_zorder(s, zpos = 3)

 
 PPT.ReplaceShapeByGraphic()
 
 # create new PPT object and add one slide
 p <- PPT.Init(visible=T, method = "RDCOMClient")
 p <- PPT.AddBlankSlide(p)
 
 # add two rectangle shapes to slide
 p <- PPT.AddShape(p, width = .4, left=.1)                      # add a shape to slide
 p <- PPT.AddShape(p, width = .4, left=.5, height=.45, top=.5)  # add a shape to slide
 
 # get all shapes on current slide and replace by image
 # shapes are not destroyed to see the image placement
 s <- PPT.ShapesOnCurrentSlide(p)   
 file <- system.file("image_1.png", package = "R2PPTaddons")  # get image
 p <- PPT.FitGraphicIntoShape(p, file, shp=s[[1]], vjust="top", delete.shape =T)
 p <- PPT.FitGraphicIntoShape(p, file, shp=s[[2]], vjust="bottom", delete.shape =FALSE)

 PPT.ReplaceShapeByGraphic(p, "placeholder", file, zposition = 4)



