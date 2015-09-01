\dontrun{
  
  #### EXAMPLE ####
  
  #add PNG that comes with the packages
  file <- "inst/image_1.png"      # comes with package
  p <- PPT.Init(visible=TRUE, method="RDCOMClient")
  p <- PPT.AddGraphicstoSlide2(p, file)

  # the 'file' argument is vectorized, so it can take multiple images
  files <- c("inst/image_2.png", "inst/image_3.png")
  p <- PPT.AddGraphicstoSlide2(p, files)    
  
  
  #### MORE EXAMPLES ####
  
  # width and height values smaller than maxscale (default is 1) 
  # are interpreted as a proportions of the slide width/height. Values 
  # greater than 1 are taken as absolute pixel width/height.
  
  file <- "inst/image_1.png"      # comes with package
  p <- PPT.Init(visible=TRUE, method="RDCOMClient")
  p <- PPT.AddGraphicstoSlide2(p, file, width=.5)
  p <- PPT.AddGraphicstoSlide2(p, file, height=.9)
  p <- PPT.AddGraphicstoSlide2(p, file, width=.5, height=.9)
  p <- PPT.AddGraphicstoSlide2(p, file, width=.5, height=.9, proportional=FALSE)

  # using pixel width/height instead of proportions of 
  # available slide width/height
  p <- PPT.AddGraphicstoSlide2(p, file, width=400)
  p <- PPT.AddGraphicstoSlide2(p, file, height=100)
  p <- PPT.AddGraphicstoSlide2(p, file, width=400, height=100)
  p <- PPT.AddGraphicstoSlide2(p, file, width=400, height=100, proportional=FALSE)
  
  # changing x and y placement
  p <- PPT.AddGraphicstoSlide2(p, file, x="left")
  p <- PPT.AddGraphicstoSlide2(p, file, x="right", y="bottom")
  
  # adding offset for finer placement control  
  p <- PPT.AddGraphicstoSlide2(p, file, width=400, x="left")
  p <- PPT.AddGraphicstoSlide2(p, file, width=400, x="left", x.offset=10)
}
