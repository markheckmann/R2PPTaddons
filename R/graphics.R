

#' Update the current slide stored in R2PPT object
#'
#' R2PPT uses an object to store the current slide amongst other things. 
#' Unfortunately the current slide is only set when a new slide is inserted. It
#' is NULL when a file is loaded. This will cause errors sometimes, hence we may
#' need to set it manually.
#' 
#' @param ppt   The ppt object as used in \pkg{R2PPT}.
#' @param i     Slide index.
#' @param slide A slide object as alternative to setting the index.
#' @author Mark Heckmann
#' @export
#' @example inst/examples/PPT.ReplaceTextByGraphicExample.R
#'
PPT.UpdateCurrentSlide <- function(ppt, i=NULL, slide=NULL)
{
  if (!is.null(i))
    slide <- ppt$pres[["Slides"]]$Item(i)
  ppt$Current.Slide <- slide
  ppt
}


#### Insert graphic ####


# This it the workhorse, arguments are explained in function 
# PPT.AddGraphicstoSlide2 below.
#
PPT.AddGraphicstoSlide2_ <- function(ppt, file, width=.9, height=.9,
                                      x="center", y="center", 
                                      x.offset=0, y.offset=0, 
                                      proportional=TRUE, newslide=FALSE, 
                                      maxscale=1)
{    
  # checking arguments
  x.sel <- c("center", "left", "right")
  y.sel <- c("center", "top", "bottom")
  if (is.character(x))
    x <- x.sel[pmatch(tolower(x), x.sel, duplicates.ok=FALSE)]  
  if (is.character(y))
    y <- y.sel[pmatch(tolower(y), y.sel, duplicates.ok=FALSE)]  
  if (is.na(x))
    stop("x must be numeric or 'center', 'left' or 'right'", call. = FALSE)
  if (is.na(y))
    stop("x must be numeric or 'center', 'top' or 'bottom'", call. = FALSE)
 
  # Adding a new slide before adding graphic
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
  
  file <- PPT.getAbsolutePath(file)         # absolute paths must be supplied to COM object
  #file <- R.utils::getAbsolutePath(file)   # absolute paths must be supplied to COM object
  file <- gsub("/", "\\\\", file)
  
  img <- shapes$AddPicture(FileName = file, 
                           LinkToFile = 0, 
                           SaveWithDocument = -1, 
                           Left = 1, 
                           Top = 1, 
                           Width = slide.width, 
                           Height = slide.height)
  
  # rescale picture to full size initial size
  img$ScaleHeight(1, -1)
  img$ScaleWidth(1, -1)
  
  # calculate optimal scaling for picture to fit slide
  # if width and height are supplied, the graphic is rescaled so that the
  # condition (img.width <=width & img.height <= height) is satisfied 
  img.width <- img[["width"]]
  img.height <- img[["height"]]

  if (!is.na(width) & width > maxscale)
    width <- width/slide.width
  if (!is.na(height) & height > maxscale)
    height <- height/slide.height  
  
  rescale.width.by <- width * slide.width / img.width
  rescale.height.by <- height * slide.height / img.height
  width.na <- is.na(width)
  height.na <- is.na(height)
  
  if (!width.na & height.na) 
    rescale.height.by <- rescale.width.by
  if (width.na & !height.na)
    rescale.width.by <- rescale.height.by
  if (!width.na & !height.na & proportional) {
    m <- min(rescale.height.by, rescale.width.by)
    rescale.width.by <- m
    rescale.height.by <- m 
  }
 
  img$ScaleHeight(rescale.width.by, -1)
  img$ScaleWidth(rescale.height.by, -1)

  # locate pic horizontally
  if (x == "center") 
    x.left <- slide.width / 2 - img[["Width"]] / 2
  if (x == "left") 
    x.left <- 0
  if (x == "right")
    x.left <- slide.width - img[["Width"]]  
  if (is.numeric(x))
    x.left <- x
  
  # locate pic vertically
  if (y == "center") 
    y.top <- slide.height / 2 - img[["Height"]] / 2
  if (y == "top") 
    y.top <- 0
  if (y == "bottom")
    y.top <- slide.height - img[["Height"]]
  if (is.numeric(y))
    y.top <- y
    
  img[["Left"]] <- x.left + x.offset
  img[["Top"]] <- y.top + y.offset
  invisible(ppt)
}



#' Adding a graphic to a slide.
#'
#' \code{PPT.AddGraphicstoSlide2} is a modified version of 
#' \code{PPT.AddGraphicstoSlide} from \pkg{R2PPT}. 
#' The PPT.AddGraphicstoSlide function has the drawback that it requires
#' the size of the graphic to be supplied by the user. If the size is not
#' supplied correctly, the graphic is deformed.
#'
#' @param ppt       The ppt object as used in \pkg{R2PPT}.
#' @param file      Path to the graphic file.
#' @param width     Width of graphic. For values smaller than \code{maxscale}
#'                  (default is \code{1}) this refers to a proportion of the 
#'                  current slide's width. Values bigger than \code{maxscale} 
#'                  are interpreted as pixels.If \code{NA} only the height 
#'                  argument is used for sclaing.
#' @param height    Height of graphic. For values smaller than \code{maxscale}
#'                  (default is \code{1}) this refers to a proportion of the 
#'                  current slide's height. Values bigger than \code{maxscale} 
#'                  are interpreted as pixels. If \code{NA} only the width 
#'                  argument is used for scaling.
#' @param x         Horizontal placement of graphic. Either a string (\code{"center", 
#'                  "left", "right"}) or a numerical value. Numerical values are 
#'                  interpreted as absolute position in pixels counted
#'                  from the left of the slide.
#' @param y         Vertical placenment of graphic.Either a string (\code{"center", 
#'                  "top", "bottom"}) or a numerical value. Numerical values are 
#'                  interpreted as absolute position in pixels counted from the 
#'                  top of the slide.
#' @param x.offset  Additional horizontal offset in pixel. Used for finetuning
#'                  position on slide.
#' @param y.offset  Additional horizontal offset in pixel. Used for finetuning
#'                  position on slide.
#' @param proportional  Logical (default \code{TRUE}). Whether scaling preserves
#'                      the aspect ratio of the graphic. See details section for
#'                      additional information.
#' @param newslide  Logical (default is \code{TRUE}) Whether the graphic will
#'                  be placed on a new slide.
#' @param maxscale  Threshold below which values are interpreted as proportional 
#'                  scaling factors for the \code{width} and \code{height} argument.
#'                  Above the threshold values are interpreted as pixels.
#' @note            The common use case is to add graphics and scale them 
#'                  while preserving their aspect ratio. In the case this 
#'                  this is not wanted the argument 
#'                  \code{proportional} can be set to \code{FALSE}. 
#'                  When the aspect ratio is preserved and both arguments 
#'                  \code{width} and \code{height} are supplied, 
#'                  the graphic is scaled so that it will not
#'                  exceed the size given by one of these values. This is useful when adding a lot of 
#'                  graphics of different size. Only using \code{width} may result
#'                  in graphics exceeding the slide vertically and vice versa. using 
#'                  both parameters (the default) will cause the graphic to be scaled 
#'                  so it will fit the slide regardless of its dimensions.
#'                  
#' @author Mark Heckmann
#' @export
#' @example inst/examples/PPT.AddGraphicstoSlideExample.R
#'
PPT.AddGraphicstoSlide2 <- function(ppt, file, width=.9, height=.9,
                                     x="center", y="center", 
                                     x.offset=0, y.offset=0, 
                                     proportional=TRUE, newslide=TRUE, 
                                     maxscale=1)
{
  for (f in file) {
    ppt <- PPT.AddGraphicstoSlide2_(ppt, file, width, height, x, y, 
                                    x.offset, y.offset, 
                                    proportional, newslide, maxscale)
  }
  invisible(ppt)
}
# Quasi-vectorization as mapply and Vectorize cannot be applied as
# the ppt hanlde would need to be updated in between. At least using mapply 
# it throws an error so I chose this version which will suffice 
# for most of my use cases. 



#### Find text and replace by graphic ####


# search string on all slides and replace it by graphic


# Detect the presence or absence of text pattern in a shape object
#
# does the shape contain the text that is searched for
# shp: poiter to shape
# what: text that is searched for
#
shape_detect_text <- function(shape, what)
{
  has.text <- FALSE
  textframe <- shape[["HasTextFrame"]]                # does the shape contain text?
  if (textframe == -1) {                              # msoTriState Constant: msoFalse =0, msoTrue=-1
    textRange <- shape[["TextFrame"]][["TextRange"]]  # get textrange from textframe
    f.textRange <- textRange$Find(FindWhat = what)    # search in tectrange for text
    txt <- f.textRange[["Text"]]                      # retrieve matched text, NULL if no matches
    if (!is.null(txt)) { # NULL if text was not found
      has.text <- TRUE
    }                    
  }
  has.text
}


# Detect the presence or absence of text pattern in each shape of shapes collection
#
# sld: pointer to slide
# what: text that is searched for
# Returns indexes of shapes that contain the text pattern
#
shapes_detect_text <- function(shapes, what)
{
  #shapes <- slide[["Shapes"]]               # get all shapes on slide
  nshapes <- shapes[["Count"]]              # number of shapes
  if (nshapes == 0)
    return(integer(0))
  res <- rep(NA, nshapes)
  for (i in 1L:nshapes) {
    shp <- shapes$Item(i)
    res[i] <- shape_detect_text(shp, what)         
  }
  which(res)
}



# retrieve shape objects with matching text pattern
# slide: pointer to slide
# what: text that is searched for
#
slide_retrieve_shapes <- function(slide, what)
{
  shapes <- slide[["Shapes"]]
  ii <- shapes_detect_text(shapes, what)
  l <- list()
  for (i in ii){
    l[[i]] <- shapes$Item(i)
  }
  l
}


# retrieve shape objects with matching text pattern across all slides
# slide: pointer to slide
# what: text that is searched for
#
slides_retrieve_shapes <- function(slides, what)
{
  nslides <- slides[["Count"]]
  r <- list()
  for (i in 1L:nslides){
    sld <- slides$Item(i)
    l <- slide_retrieve_shapes(sld, what)
    r <- c(r, l)
  }
  r
}



#' Replace matching text by graphic
#'
#' Looks through all shapes and finds a shape with mathicng text pattern.
#' The shape is deleted and a graphic is inseretd on the shape's parent slide. 
#'
#' @param ppt   The ppt object as used in \pkg{R2PPT}.
#' @param what  Text pattern to match against.  
#' @param file  Path to the graphic file.
#' @param ... Arguments passed on to \code{\link{PPT.AddGraphicstoSlide2}}.
#' @author Mark Heckmann
#' @export
#' @example inst/examples/PPT.ReplaceTextByGraphicExample.R
#'
PPT.ReplaceTextByGraphic <- function(ppt, what, file, ...)
{
  slides <- ppt$pres[["Slides"]]
  ss <- slides_retrieve_shapes(slides, what)   # get all shape objects that match text pattern 
  if (length(ss) == 0)
    warning("No shape with matching text pattern was not found.", call. = FALSE)
  if (length(ss) > 1)
    warning("More than one shape with matching text pattern found and replaced.", call. = FALSE)
  
  for (s in ss) {               # delete from last to first
    sld <- s[["Parent"]]        # get shape's slide
    #sld$Select()                # shape select throws error if focus is not on shape's slide, so select parent first
    s$Delete()                  # delete shape
    ppt <- PPT.UpdateCurrentSlide(ppt, slide=sld)    # PPT.AddGraphicstoSlide2 needs ppt$CurrentSlide to be set
    PPT.AddGraphicstoSlide2(ppt, file, newslide=FALSE)
  }  
}


























  
  



