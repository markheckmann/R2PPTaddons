library(R2PPT)
devtools::load_all(".")


#### textframe ####

# Presentation
p <- PPT.Init(visible=T, method = "RDCOMClient")
p <- PPT.AddTitleSlide(p, title="Test", subtitle=NULL)
p <- PPT.AddBlankSlide(p)

# align textbox
txt = c("Line 1", "Line 2", "Line 3")
p <- PPT.AddTextBox(p, txt, 
                       x = .5, 
                       y = .5, 
                       xy.format = "perc",
                        font.size=30, 
                        font.color = "blue",
                        font.bold = T,
                        font.italic = T,
                        font.name = "Courier New",
                       width = .2,
                       x.text.align = "left", 
                       bullet.points = "unnumbered",
                       bullet.points.color = "darkgreen")

p <- PPT.AddBlankSlide(p)
shps <- p$Current.Slide[["Shapes"]]
s <- shps$Item(1)
s$Select()
tmp <- s[["TextFrame"]][["TextRange"]][["Font"]]


for ( x.offset in seq(0, .9, .1) ) {
  p <- PPT.AddTextBox(p, txt, 
                         x = "left", 
                         y = "top", 
                         width = .1,
                         x.offset = x.offset,
                         text.color = "darkgreen",
                         border.color = "darkgreen",
                         fill.color = "white",
                         offset.format = "percent",
                         x.text.align = "left", 
                         bullet.points = "unnum",
                         bullet.points.color = "darkgreen")
}


p <- PPT.AddBlankSlide(p)
for ( x.offset in seq(0, .9, .15) ) {
  p <- PPT.AddTextBox(p, txt, 
                         x = "left", 
                         y = "top", 
                         width = .15,
                         x.offset = x.offset,
                         offset.format = "percent",
                         x.text.align = "left", 
                         bullet.points = "none")
}


