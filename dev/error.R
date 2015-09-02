
library(gap)
library(R2PPTaddons)
#library(R.utils)
library(stringr)


visible <- T
file.template <- "dev/template.pptx"
file.out <- str_replace(file.template, ".pptx$", "_filled.pptx")
#scale.width <- 90
e <- 16                                             # m_einrichtung_id
bb <- c(10, 180)
b <- 10
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




