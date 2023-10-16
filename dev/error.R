
library(gap)
library(R2PPTaddons)
library(stringr)


visible <- T
file.template <- "dev/template.pptx"
file.out <- str_replace(file.template, ".pptx$", "_filled.pptx")
#scale.width <- 90
e <- 16                                             # m_einrichtung_id
bb <- c(10, 180)
b <- 180
# e = 16
# b = 180

###################  LOAD  ###################  

p <- PPT.Open(file.template, method="RDCOMClient")

###################  BARCHARTS  ###################  

cat("\n\nBARCHARTS\n\n")
# placeholder <- "[[type=barchart e=16 b=180]]"
# file <- "output/a_barcharts/png/barchart__einrichtung_16__graphic_180.png"
type= "barchart"
i <- 0
  i <- i + 1
  cat("\rGraphic", i, "in", length(bb))
  placeholder <- paste0("[[", "type=", type, " ", "b=", b, "]]")
  b.dir <- "dev"
  b.name <- paste0("barchart__einrichtung_", e, "__graphic_", f(b), ".png") 
  file <- file.path(b.dir, b.name)
  file.exists(file)
  p <- PPT.ReplaceTextByGraphic(p, placeholder, file)
p$pres



#### fit image in box ####

# Find box shape with text pattern
what <- "[[type=barchart b=180]]"
slides <- p$pres[["Slides"]]
ss <- slides_retrieve_shapes(slides, what)
s <- ss[[1]]
s$Select()

# insert graphic
file <- "dev/barchart__einrichtung_16__graphic_180.png"
sld <- s[["Parent"]]        # get shape's slide
p <- PPT.UpdateCurrentSlide(p, slide=sld)   # to insert graphic on correct slide


# get shape size and position properties
h <- s[["Height"]]
w <- s[["Width"]]
left <- s[["Left"]]
top <- s[["Top"]]

# check if type is the correct one (actually not needed) 
s[["Type"]] == 1    # MsoShapeType Enumeration:  msoAutoShape = 1 (e.g. rectangle), msoTextBox = 17

# calculate the new position within the box (box)

# options:
# position: 0-1 (0 top, )
sld[["Shapes"]]
PPT.AddGraphicstoSlide2_(p, file, newslide=FALSE,
                         width=w, height=h, 
                         x = "left", y = "top",
                         x.offset=left,
                         y.offset=top)



#### sxs ####

# function to get width and height of scaled graphic

# get shape size and position properties
h <- s[["Height"]]
w <- s[["Width"]]
left <- s[["Left"]]
top <- s[["Top"]]

# function to position image
width = .9 
height = .9
x = "center"
y = "center"
x.offset = 0
z.offset = 0

shp.width
shp.height
shp.left
shp.top

box.width
box.height
box.left
box.right









