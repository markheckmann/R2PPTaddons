
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




# TODO:
# vjust, hjust offset as fraction of shape width 

#' Adding a textbox to a slide.
#'
#' Add a textbox to a slide. YOu can easily position it and modify a limited
#' number of aspects of its appearance (color, bullet points, etc.)
#' 
#' @param ppt The ppt object as used in \pkg{R2PPT}.
#' @param txt Text to put into the textbox. A vector with length greater 1
#'   is collapsed using linebreak characters.
#' @param width Width of graphic. For values smaller than \code{maxscale}
#'   (default is \code{1}) this refers to a proportion of the current slide's
#'   width. Values bigger than \code{maxscale} are interpreted as pixels.If
#'   \code{NA} only the height argument is used for sclaing.
#' @param height Height of graphic. For values smaller than \code{maxscale}
#'   (default is \code{1}) this refers to a proportion of the current slide's
#'   height. Values bigger than \code{maxscale} are interpreted as pixels. If
#'   \code{NA} only the width argument is used for scaling.
#' @param x Horizontal placement of the textbox. Either a string
#'   (\code{"center", "left", "right"}) or a numerical value. Numerical values
#'   are interpreted as absolute position in pixels counted from the left of the
#'   slide.
#' @param y Vertical placenment of the textbox. Either a string (\code{"center",
#'   "top", "bottom"}) or a numerical value. Numerical values are interpreted as
#'   absolute position in pixels counted from the top of the slide.
#' @param xy.format The numeric x and y input will be interpreted either as
#'   \code{"pixels"} (default) or \code{"percent"} of the slide's total
#'   width/height. Character input will not be affected.
#' @param x.offset  Additional horizontal offset in pixel. Used for finetuning
#'   position on slide.
#' @param y.offset  Additional horizontal offset in pixel or as percent (see .
#'   Used for finetuning position on slide.
#' @param offset.format The offset will be interpreted either as \code{"pixels"}
#'   (default) or \code{"percent"} of the slide's total width/height.
#' @param x.text.align Horizontal alignment of text (\code{"left", "center",
#'   "right"}).
#' @param bullet.points Whether to treat each new line and vector element as a
#'   bullet point (\code{"none", "unnumbered", "numbered"}).
#' @param bullet.points.color Color of bullet points either as hex value or
#'   color name.
#' @param font.color Color of text either as hex value or color name.
#' @param font.size Text size (default 16).
#' @param font.bold Bold text (default \code{FALSE}).
#' @param font.italic Italic text (default \code{FALSE}).
#' @param fill.color Background color either as hex value or color name.
#' @param border.color Border line color either as hex value or color name.
#' @param newslide  Logical (default is \code{TRUE}) Whether the graphic will be
#'   placed on a new slide.
#' @param maxscale  Threshold below which values are interpreted as proportional
#'   scaling factors for the \code{width} and \code{height} argument. Above the
#'   threshold values are interpreted as pixels.
#'                  
#' @author Mark Heckmann
#' @export
#' @example inst/examples/PPT.AddTextBoxExample.R
#'
PPT.AddTextBox <- function( ppt, 
                            txt, 
                            width=.9, 
                            height=.9,
                            x="center", 
                            y="center",
                            xy.format = "points",  # either points or percent
                            x.offset=0, 
                            y.offset=0, 
                            offset.format = "points",  # either points or percent
                            x.text.align = "center",
                            bullet.points = "none", 
                            bullet.points.color = 0,
                            font.size = 16,
                            font.color = "black",
                            font.bold = FALSE,
                            font.italic = FALSE,
                            fill.color = NA,   # fill color
                            border.color = NA,  # border color
                            newslide=FALSE, 
                            maxscale=1,
                            autosize = TRUE)
{    
  # collapse text if vector longer 1. Textrange "Text" property only allows
  # a single string.
  if (length(txt) > 1) {
    txt <- paste0(txt, collapse="\n")
  }
  
  # get width / heught of slides
  sld.width = slide_width(ppt)
  sld.height = slide_height(ppt)
  
  
  # checking arguments and implement partial matching of input
  x.sel <- c("center", "left", "right")
  y.sel <- c("center", "top", "bottom")
  xy.sel <- c("points", "percent")      # offset value in points or percent
  offset.sel <- c("points", "percent")  # offset value in points or percent
  
  if (is.character(x))
    x <- x.sel[pmatch(tolower(x), x.sel, duplicates.ok=FALSE)]  
  if (is.na(x))
    stop("x must be numeric or 'center', 'left' or 'right'", call. = FALSE)
  
  if (is.character(y))
    y <- y.sel[pmatch(tolower(y), y.sel, duplicates.ok=FALSE)]  
  if (is.na(y))
    stop("x must be numeric or 'center', 'top' or 'bottom'", call. = FALSE)
  
  xy.format <- xy.sel[pmatch(tolower(xy.format), xy.sel, duplicates.ok=FALSE)]  
  if (is.na(xy.format))
    stop("xy.format must be 'points' or 'percent'", call. = FALSE)
 
   offset.format <- offset.sel[pmatch(tolower(offset.format), offset.sel, duplicates.ok=FALSE)]  
  if (is.na(offset.format))
    stop("offset.format must be 'points' or 'percent'", call. = FALSE)  


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
  
  
  ## FONT PROPERTIES ##
  
  f <- txt_range[["Font"]]
  f[["Size"]] <- font.size
  f[["Bold"]] <- font.bold
  f[["Italic"]] <- font.italic
  # color
  fc <- f[["Color"]]
  fc[["RGB"]] <- color_to_integer(font.color)
  
  
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
  
  # bullet point color
  if (bullet.points != "none" & !is.na(bullet.points.color) ) {
    f <- bullet[["Font"]][["Color"]]
    f[["RGB"]] <- color_to_integer(bullet.points.color)  
  }
  

  ## Shape properties ##
  
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

  # resize textbox to textsize
  # NOTE: not sure if it's needed
  if (autosize) {
    txt_frame[["Autosize"]] <- 1
    tf <- shp[["TextFrame2"]]
    tf[["Autosize"]] <- 2
  }
  
  # if x,y is percentage it must not be character
  if (xy.format == "percent" & ( !is.numeric(x) | !is.numeric(y) ) ) {
    stop("If xy.format = 'percent' x and y must be numeric.", call. = FALSE)
  }
    
  # convert xy to points if passed as percentage
  if (xy.format == "percent") {
    x.offset = x * sld.width  # convert to points
    y.offset = y * sld.height # convert to points
  }
  
  # position textbox horizontally
  if (x == "center") 
    x.left <- slide.width / 2 - shp[["Width"]] / 2
  if (x == "left") 
    x.left <- 0
  if (x == "right")
    x.left <- slide.width - shp[["Width"]]
  if (is.numeric(x))
    x.left <- x

  # position textbox vertically
  if (y == "center") 
    y.top <- slide.height / 2 - shp[["Height"]] / 2
  if (y == "top") 
    y.top <- 0
  if (y == "bottom")
    y.top <- slide.height - shp[["Height"]]
  if (is.numeric(y))
    y.top <- y

  # convert offset to points if passed as percentage
  if (offset.format == "percent") {
    x.offset = x.offset * sld.width  # convert to points
    y.offset = y.offset * sld.height # convert to points
  }
  
  # position shape
  shp[["Left"]] <- x.left + x.offset
  shp[["Top"]] <- y.top + y.offset
  
  # return PPT object
  invisible(ppt)
}




#' Add Rectangle shape
#'
#' Add a rectangle to a slide. YOu can position it and modify a limited number
#' of aspects of its appearance (color etc.)
#'
#' @param ppt The ppt object as used in \pkg{R2PPT}.
#' @param width,height Dimensions of shape. For values smaller than
#'   \code{maxscale} (default is \code{1}) this refers to a proportion of the
#'   current slide's width or height. Values bigger than \code{maxscale} are
#'   interpreted as pixels.
#' @param top,left Vertical and horizontal placement of the shape. Either as
#'   fraction of slides dimensions or as pixel value. Values bigger than
#'   \code{maxscale} are interpreted as pixels.#'
#' @param line.color Color of text either as hex value or color name.
#' @param line.type \code{1} = solid (default), \code{2-8}= dots, dashes and
#'   mixtures. See MsoLineDashStyle Enumeration for details.
#' @param line.size Thickness of line (default\code{1}).
#' @param fill.color Background color either as hex value or R color name.
#' @param fill.transparency Transparency of filling (\code{[0,1]}, 
#' default is \code{0} = opaque.).
#' @param newslide  Logical (default is \code{TRUE}) Whether the graphic will be
#'   placed on a new slide.
#' @param maxscale  Threshold below which values are interpreted as proportional
#'   scaling factors for the \code{width} and \code{height} argument. Above the
#'   threshold values are interpreted as pixels.
#' @author Mark Heckmann
#' @export
#' @example inst/examples/PPT.AddRectangleExample.R
#'
PPT.AddRectangle <- function(ppt, 
                             top = .05,
                             left = .05,
                             width = .9,
                             height= .9,
                             fill.color="grey", 
                             fill.transparency = 0, 
                             line.color = "black",
                             line.type = 1,
                             line.size = 1,
                             maxscale = 1,
                             newslide = FALSE)
{
  # Adding a new slide before adding textbox if promted
  if (newslide)
    ppt <- PPT.AddBlankSlide(ppt)  
  # if the current slide object is not set, an error will occur
  if (!newslide & is.null(ppt$Current.Slide)) {  
    warning("No current slide defined. Slide 1 ist selected.", call. = FALSE)
    ppt <- PPT.UpdateCurrentSlide(ppt, i=1)
  }
  
  # prepare coordinates and get shape collection
  shapes <- ppt$Current.Slide[["Shapes"]]
  slide.width <- ppt$pres[["PageSetup"]][["SlideWidth"]] 
  slide.height <- ppt$pres[["PageSetup"]][["SlideHeight"]]
  
  # convert fractions to pixels
  if (!is.na(width) & width <= maxscale) {
    width <- width * slide.width      # frame width as fraction of slide width
  }
  if (!is.na(height) & height <= maxscale) {
    height <- height * slide.height   # frame height as fraction of slide height
  }
  if (!is.na(top) & top <= maxscale) {
    top <- top * slide.height         # top as fraction of slide height
  }
  if (!is.na(left) & left <= maxscale) {
    left <- left * slide.width        # left as fraction of slide width
  }
  

  # print ccordinates for debugging
  l <- list(slide.width = slide.width,
            slide.height = slide.height,
            top = top, 
            left = left, 
            width = width, 
            height = height)
  print(l)
  
  # add rectangle
  rect <- shapes$AddShape( Type = 1,  # msoShapeRectangle
                           Left = left, 
                           Top = top, 
                           Width =width, 
                           Height = height)
  
  # format rectangle
  obj <- rect[["Fill"]] 
  obj[["ForeColor"]][["RGB"]] = color_to_integer(fill.color)
  obj[["Transparency"]] = fill.transparency
  
  # format border line
  obj <- rect[["Line"]] 
  obj[["DashStyle"]] = line.type  # dashed, see: MsoLineDashStyle enumeration
  obj[["ForeColor"]][["RGB"]] = color_to_integer(line.color)
  obj[["Weight"]] = line.size
  
  invisible(p)
}
