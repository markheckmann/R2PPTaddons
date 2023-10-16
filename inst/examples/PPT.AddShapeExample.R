\dontrun{

# open new PPT presentation
p <- PPT.Init(visible=T, method = "RDCOMClient")  


## EXAMPLE 1 ##

# Look up the shape type number in the "MsoAutoShapeType enumeration". 
# It is one integer, e.g. rectangle = 1, just google it.

p <- PPT.AddBlankSlide(p)
# add a rectangle
p <- PPT.AddShape(p, shape.type= 1, height=.4, width =.4, left=.05, top=.05)
# add a rounded rectangle
p <- PPT.AddShape(p, shape.type= 5, height=.4, width =.4, left=.05, top=.55)
# add a triangle 
p <- PPT.AddShape(p, shape.type= 7, height=.4, width =.4, left=.55, top=.05)
# add a smiley
p <- PPT.AddShape(p, shape.type= 17, height=.4, width =.4, left=.55, top=.55)



## EXAMPLE 2 ##

# add many type of shapes on one slide in two loops to get an 
# overview what shapes exist. The fill and line parameters are 
# randomly modified to get a variety of shapes.
p <- PPT.AddBlankSlide(p)
cols <- colors()   # all named R colors
i <- 0             # counter
set.seed(0)        # make sampling redproducible

# loop over left and top to alter position of shape
for ( left in seq(.05, .85, by=.1) ) {
  for ( top in seq(.05, .85, by=.1) ) {
    i <- i + 1
    p <- PPT.AddShape(p, 
                      shape.type= i,                               # go though shapes one by one 1,2,....
                      height=.08, width =.08,                      # height and width of shape
                      left=left, top=top,                          # shape position is modified in each loop 
                      fill.transparency = sample(c(0,.3, .7), 1),  # use different transparencies
                      fill.color = cols[i],                        # go through all colors for filling
                      line.color = sample(cols, 1),                # random line color
                      line.size = sample(1:30/10, 1),              # random line size
                      line.type = sample(1:8, 1)                   # random line type
                      )
  }
}


}

