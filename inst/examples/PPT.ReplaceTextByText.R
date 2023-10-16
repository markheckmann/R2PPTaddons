\dontrun{
  
  # replace text on slide
  ppt <- PPT.Open("inst/template.pptx", method="RDCOMClient")

  what <- "[[tag 1]]"
  replace <- "This has been replaced"
  PPT.ReplaceTextByText(ppt, what, replace)
}
