
# add graphic to slides with matching and remoce text
library(R2PPTaddons)

ppt <- PPT.Open("inst/template.pptx", method="RDCOMClient")

what <- "[[tag 1]]"
PPT.ReplaceTextByGraphic(ppt, what, file = "inst/image_1.png")

# Note that the text appears twice and the graphic is inserted twice with a warning.
what <- "[[tag 2]]"
PPT.ReplaceTextByGraphic(ppt, what, file = "inst/image_2.png")



