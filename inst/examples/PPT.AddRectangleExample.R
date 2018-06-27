\dontrun{
  
# add rectangle shapes to slide
p <- PPT.Init(visible=T, method = "RDCOMClient")
p <- PPT.AddBlankSlide(p)
p <- PPT.AddRectangle(p, height=.4)
p <- PPT.AddRectangle(p, top = .5, width = .4, height=.45, 
                      fill.color="blue", fill.transparency = .8) 
p <- PPT.AddRectangle(p, top = .5, width = .4, height=.45, left=.55, 
                      line.type=4, line.color="red", line.size = 3) 

}

