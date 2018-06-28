\dontrun{
  
  #### EXAMPLE ####
  
  # add PNG that comes with the packages
  file <- system.file("image_1.png", package = "R2PPTaddons")
  p <- PPT.Init(visible=TRUE, method="RDCOMClient")
  p <- PPT.AddGraphicstoSlide2(p, file)

  # the argument file is vectorized, so it takes multiple images
  folder <- system.file(package = "R2PPTaddons")
  files <- list.files(folder, pattern = ".png", full.names = T)
  p <- PPT.AddGraphicstoSlide2(p, files)    
  
  
  #### MORE EXAMPLES ####
  
  p <- PPT.Init(visible=TRUE, method="RDCOMClient")
  file <- system.file("image_1.png", package = "R2PPTaddons")
  
  # the image is placed within a frame. To get a better understanding
  # what happens first only the frame is displayed. Afterwards the image is added.
  p <- PPT.AddGraphicstoSlide2(p, file, display.frame = TRUE, display.image = FALSE) 
  p <- PPT.AddGraphicstoSlide2(p, file, newslide=F)    
  
  # width and height values smaller than maxscale (default is 1) 
  # are interpreted as a proportions of the slide width/height. Values 
  # greater than 1 are taken as absolute pixel width/height.
  p <- PPT.AddGraphicstoSlide2(p, file, width=.5)
  p <- PPT.AddGraphicstoSlide2(p, file, height=.5)

  # using pixel width/height instead of proportions of 
  # available slide width/height
  p <- PPT.AddGraphicstoSlide2(p, file, width=400)
  p <- PPT.AddGraphicstoSlide2(p, file, width=400, left=0)
  
  # one can also fit the image completely to the frame potentially 
  # destroying the image's aspect ration, i.e. distorting it
  p <- PPT.AddGraphicstoSlide2(p, file, width=400, height=100, proportional=FALSE)
  
  # position the image on slide
  p <- PPT.AddGraphicstoSlide2(p, file, left=0, width =.5)
  p <- PPT.AddGraphicstoSlide2(p, file, top=0, height =.5)
  
  # aligment of image inside the frame
  p <- PPT.AddGraphicstoSlide2(p, file, width=400, display.frame = T)
  p <- PPT.AddGraphicstoSlide2(p, file, width=400, vjust="top",display.frame = T)
  p <- PPT.AddGraphicstoSlide2(p, file, width=400, vjust=0,display.frame = T)
  p <- PPT.AddGraphicstoSlide2(p, file, width=400, vjust="bottom",display.frame = T)
  p <- PPT.AddGraphicstoSlide2(p, file, width=400, vjust=1, display.frame = T)
  p <- PPT.AddGraphicstoSlide2(p, file, width=400, vjust=400, display.frame = T)
}
