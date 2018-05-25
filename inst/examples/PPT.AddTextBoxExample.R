\dontrun{
  
# add textbox to middle of slide
p <- PPT.Init(visible=T, method = "RDCOMClient")
p <- PPT.AddBlankSlide(p)
txt = c("Line 1", "Line 2", "Line 3")
p <- PPT.AddTextBox(p, txt) 

}

