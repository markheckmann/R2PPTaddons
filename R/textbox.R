#### Insert graphic ####


# tests

# # correct
# is_percent("1%")
# is_percent(".52%")
# is_percent("10.1%")
# 
# # incorrect
# is_percent("%")
# is_percent("a%")
#
# is_percent("a%")
# check if string has the format "FLOAT%"
is_percent_format <- function(x)
{
  # regex to check if string has the format "FLOAT%"
  stringr::str_detect(x, "^([0-9]*[.])?[0-9]+%$")  
}

# convert percent string into numeric
percent_format_to_numeric <- function(x)
{
  # regex to get numeric part of percent string
  num <- stringr::str_extract(x, "^([0-9]*[.])?[0-9]+")  
  as.numeric(num)
}




# Add and position a textframe on the slide
# TODO:
# - x/y offset as percentage of width / height
# - autoresize
# PpAutoSize Enumeration
# ppAutoSizeMixed	-2	Mixed size.
# ppAutoSizeNone	0	Does not change size.
# ppAutoSizeShapeToFitText	1	Auto sizes the shape to fit the text.
#
# add bullet list? (0 = No, 1 = Normal, 2 = Numbered)
PPT.AddTextFrame_ <- function(ppt, 
                              txt, 
                              width=.9, 
                              height=.9,
                              x="center", 
                              y="center", 
                              x.offset=0, 
                              y.offset=0, 
                              offset.format = "points",  # either points or percent
                              x.text.align = "center",
                              bullet.points = "none", 
                              bullet.points.color = 0,
                              text.color = NA,
                              fill.color = NA,   # fill color
                              border.color = NA,  # border color
                              #proportional=TRUE, 
                              newslide=FALSE, 
                              maxscale=1,
                              autosize = TRUE)
{    
  # collapse text if vector longer 1. Textrange "Text" property only allows
  # a single string.
  if (length(txt) > 1) {
    txt <- paste0(txt, collapse="\n")
  }
    
  # checking arguments
  x.sel <- c("center", "left", "right")
  y.sel <- c("center", "top", "bottom")
  offset.sel <- c("points", "percent")  # offset value in points or percent
  
  if (is.character(x))
    x <- x.sel[pmatch(tolower(x), x.sel, duplicates.ok=FALSE)]  
  if (is.character(y))
    y <- y.sel[pmatch(tolower(y), y.sel, duplicates.ok=FALSE)]  
  if (is.character(y))
    offset.format <- offset.sel[pmatch(tolower(offset.format), offset.sel, duplicates.ok=FALSE)]  
  
  if (is.na(x))
    stop("x must be numeric or 'center', 'left' or 'right'", call. = FALSE)
  if (is.na(y))
    stop("x must be numeric or 'center', 'top' or 'bottom'", call. = FALSE)

  
  
  # Adding a new slide before adding textbox if promted
  if (newslide)
    ppt <- PPT.AddBlankSlide(ppt)  
  # if the current slide object is not set, an error will occur
  if (!newslide & is.null(ppt$Current.Slide)) {  
    warning("No current slide defined. Slide 1 ist selected.", call. = FALSE)
    ppt <- PPT.UpdateCurrentSlide(ppt, i=1)
  }
  
  # TODO: problem here because current slide is not updated when changing focus interactively
  #browser()
  #ppt$pres Application.ActiveWindow.View.Slide
  shapes <- ppt$Current.Slide[["Shapes"]]
  slide.width <- ppt$pres[["PageSetup"]][["SlideWidth"]] 
  slide.height <- ppt$pres[["PageSetup"]][["SlideHeight"]]
  
  # include shape with a pixel size not too small. I do not know why, but
  # size 1,1 would not work and will produce blurry images.
  # For an unknown reason the size has to be reasonably big, here 90 percent
  # of the slide dimensions are used.
  shp <- shapes$AddTextBox(Orientation = 1,    # MsoTextOrientation enumeration
                           Left = 1, 
                           Top = 1, 
                           Width = slide.width,   # initially full slide width
                           Height = slide.height)
  # Add textframe with text
  txt_frame <- shp[["TextFrame"]]
  txt_range <- txt_frame[["TextRange"]]
  txt_range[["Text"]] <- txt

  
  ## TEXT ALIGNMENT ##
  
  # Align text
  # PpParagraphAlignment Enumeration:
  #  2 =	Center align
  #  5	= Distribute
  #  4	= Justify
  #  7	= Low justify
  #  1	= Left aligned
  # -2 = Mixed alignment
  #  3	= Right-aligned
  #  6	= Thai distributed
  #
  x.text.align.sel <- c("center", "left", "right")  # allowed values
  x.text.align <- x.text.align.sel[pmatch(tolower(x.text.align), 
                                          x.text.align.sel, duplicates.ok=FALSE)] 
  # map alignment string to numeric enumeration
  x.text.align.num <- switch(x.text.align,       
                             "center" = 2,
                             "left" = 1,
                             "right" = 3 )
  p <- txt_range[["ParagraphFormat"]]
  p[["Alignment"]] <- x.text.align.num  
  
  
  ## BULLET POINTS ##
  
  # bullet points if requested:
  # PpBulletType enumeration:
  # -2	= Mixed bullets
  #  0	= No bullets
  #  1	= Unnumbered bullets
  #  2	= Numbered bullets
  #  3	= Bullets with an image
  #
  bullet.points.sel <- c("none", "unnumbered", "numbered")  # allowed values
  bullet.points <- bullet.points.sel[pmatch(tolower(bullet.points), 
                                          bullet.points.sel, duplicates.ok=FALSE)] 
  bullet.points.num <- switch(bullet.points,       
                             "none" = 0,
                             "unnumbered" = 1,
                             "numbered" = 2 )
  bullet <- p[["Bullet"]]
  bullet[["Type"]] <- bullet.points.num 
  
  

  ## COLORS  ##
  
  # bullet point color
  if (bullet.points != "none" & !is.na(bullet.points.color) ) {
    f <- bullet[["Font"]][["Color"]]
    f[["RGB"]] <- color_to_integer(bullet.points.color)  
  }
  
  
  # text color
  if ( !is.na(text.color) ) {
    #f <- shape[["TextFrame"]][["TextRange"]][["Font"]][["Color"]]
    f <- txt_range[["Font"]][["Color"]]
    f[["RGB"]] <- color_to_integer(text.color)
  }
  
  
  # fill color
  if ( !is.na(fill.color) ) {
    f <- shp[["Fill"]]
    f[["Visible"]] <- 1  # make filling visible
    f[["ForeColor"]][["RGB"]] <- color_to_integer(fill.color)    
  }

  # border color
  if ( !is.na(border.color) ) {
    l <- shp[["Line"]]
    l[["Visible"]] <- 1     # make border line visible
    l[["ForeColor"]][["RGB"]] <- color_to_integer(border.color)
  }
  

  
  # calculate optimal scaling for picture to fit slide
  # if width and height are supplied, the graphic is rescaled so that the
  # condition (img.width <=width & img.height <= height) is satisfied
  shp.width <- shp[["width"]]
  shp.height <- shp[["height"]]

  # convert abspolute into proprtional width and height
  if (!is.na(width) & width > maxscale)
    width <- width / slide.width
  if (!is.na(height) & height > maxscale)
    height <- height / slide.height

  # factor to rescale shape by
  rescale.width.by <- width * slide.width / shp.width
  rescale.height.by <- height * slide.height / shp.height
  width.na <- is.na(width)
  height.na <- is.na(height)

  if (!width.na & height.na)
    rescale.height.by <- rescale.width.by
  if (width.na & !height.na)
    rescale.width.by <- rescale.height.by
  # if (!width.na & !height.na & proportional) {
  #   m <- min(rescale.height.by, rescale.width.by)
  #   rescale.width.by <- m
  #   rescale.height.by <- m
  # }

  # new width and height
  shp[["width"]] = shp.width * rescale.width.by
  shp[["height"]] = shp.height * rescale.height.by

  # resize box to textsize
  if (autosize) {
    txt_frame[["Autosize"]] <- 1
    tf <- shp[["TextFrame2"]]
    tf[["Autosize"]] <- 2
  }
  
  # locate shape horizontally
  if (x == "center") 
    x.left <- slide.width / 2 - shp[["Width"]] / 2
  if (x == "left") 
    x.left <- 0
  if (x == "right")
    x.left <- slide.width - shp[["Width"]]
  if (is.numeric(x))
    x.left <- x
  # TODO: Percent format verarbeiten
  
  # locate pic vertically
  if (y == "center") 
    y.top <- slide.height / 2 - shp[["Height"]] / 2
  if (y == "top") 
    y.top <- 0
  if (y == "bottom")
    y.top <- slide.height - shp[["Height"]]
  if (is.numeric(y))
    y.top <- y

  # convert width/height to points if offset is passed as 
  # percentage
  if (offset.format == "percent") {
    x.offset = x.offset * slide_width(ppt)  # convert to points
    y.offset = y.offset * slide_height(ppt) # convert to points
  }
  
  # position shape
  shp[["Left"]] <- x.left + x.offset
  shp[["Top"]] <- y.top + y.offset
  
  
  invisible(ppt)
}



