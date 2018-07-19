\dontrun{

# open new PPT presentation
p <- PPT.Init(visible=T, method = "RDCOMClient")  

#### EXAMPLE 1 ####

p <- PPT.AddBlankSlide(p)
m2 <- mtcars[1:2, ]
m4 <- mtcars[1:4, ]
m10 <- mtcars[1:10, ]

# no column names
p <- PPT.AddTable(p, m2, colnames = F)
# change position and font size
p <- PPT.AddTable(p, m2, top=.25, font.size=12)
p <- PPT.AddTable(p, m2, colnames = F, rownames = T, top=.45, font.color="red")
p <- PPT.AddTable(p, m2, rownames = T, top=.7)

#### FONTS ####

# font and text arguments are vectorized and will be recycled 
# to fit the number of columns
p <- PPT.AddBlankSlide(p)
p <- PPT.AddTable(p, m4, 
                  text.align = 1:3, 
                  font.color=c("cyan3", "#4400ff"),
                  font.bold = c(F,F,T), 
                  font.italic = c(T,F),
                  font.size = c(10,14,18))

#### ROW HEIGHT AND COLUMN WIDTH ####

# big header, smaller rows
p <- PPT.AddTable(p, m4, header.height = 80, row.height = 20)
# alternating small and big row heights
p <- PPT.AddTable(p, m4, header.height = 80, row.height = c(20, 60))
# big and small columns
p <- PPT.AddTable(p, m4, column.width = c(40, 80))


#### TEXT ALIGNMENT  ####

# different vertical alignments for header and rows
p <- PPT.AddTable(p, m4,
                  header.v.align = "top", 
                  text.v.align = "bottom",
                  row.height = 60)
# alternate aligment of rows
p <- PPT.AddTable(p, m4,
                  header.v.align = "top", 
                  text.v.align = c("top", "bottom"),   # alternate rows
                  row.height = 60)
# different horizontal alignments for columns
p <- PPT.AddTable(p, m4,
                  header.h.align = c("right","left"), 
                  text.h.align = c("left", "right"),   # aternate columns
                  row.height = 60)


#### MISC EXAMPLES  ####

#~ 3 x speed up when setting show.actions to FALSE
p <- PPT.AddBlankSlide(p)
system.time( p <- PPT.AddTable(p, m10) )
p <- PPT.AddBlankSlide(p)
system.time( p <- PPT.AddTable(p, m10, show.actions=F) )

}

