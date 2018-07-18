\dontrun{

# open new PPT presentation
p <- PPT.Init(visible=T, method = "RDCOMClient")  

#### EXAMPLE 1 ####

p <- PPT.AddBlankSlide(p)
m <- mtcars[1:2, ]

# no column names
p <- PPT.AddTable(p, m, colnames = F)
# change position and font size
p <- PPT.AddTable(p, m, top=.25, font.size=12)
p <- PPT.AddTable(p, m, colnames = F, rownames = T, top=.45, font.color="red")
p <- PPT.AddTable(p, m, rownames = T, top=.7)

#### EXAMPLE: FONTS ####

p <- PPT.AddBlankSlide(p)

# font and text arguments are vectorized and will be recycled 
# to fit the number of columns
p <- PPT.AddTable(p, m, 
                  text.align = 1:3, 
                  font.color=c("cyan3", "#4400ff"),
                  font.bold = c(F,F,T), 
                  font.italic = c(T,F),
                  font.size = c(10,14,18))

p <- PPT.AddTable(p, m, cell.height = .01)

#### MISC EXAMPLES  ####

#~ 3 x speed up when setting show.actions to FALSE

m <- mtcars[1:10, ]
p <- PPT.AddBlankSlide(p)
system.time(
  p <- PPT.AddTable(p, m)
)

p <- PPT.AddBlankSlide(p)
system.time(
  p <- PPT.AddTable(p, m, show.actions=F)
)

}

