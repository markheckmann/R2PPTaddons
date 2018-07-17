\dontrun{

# open new PPT presentation
p <- PPT.Init(visible=T, method = "RDCOMClient")  

## EXAMPLE 1 ##

p <- PPT.AddBlankSlide(p)
m <- mtcars[1:2, ]
p <- PPT.AddTable(p, m, colnames = F, rownames = F, top=.1)
p <- PPT.AddTable(p, m, colnames = T, rownames = F, top=.25)
p <- PPT.AddTable(p, m, colnames = F, rownames = T, top=.45)
p <- PPT.AddTable(p, m, colnames = T, rownames = T, top=.7)

}

