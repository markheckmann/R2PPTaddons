# This is the workhorse, arguments are explained in the function 
# PPT.AddGraphicstoSlide2 below.
#
PPT.AddGraphicstoSlide2_ <- function(ppt, 
                                     file, 
                                     width=.9, 
                                     height=.9,
                                     x="center", 
                                     y="center", 
                                     x.offset=0, 
                                     y.offset=0, 
                                     # the frame of reference inside which the positioning happens. 
                                     # default is the corners of the PPT slide 
                                     frame = list(top=NA,    
                                                  left=NA, 
                                                  width=NA,
                                                  height=NA),
                                     proportional=TRUE, 
                                     newslide=FALSE, 
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
  # it is set when a new slide is added but not when an existing file is opened
  if (is.null(ppt$Current.Slide)) {  
    warning("No current slide defined. Slide 1 is used.\n", 
            "Use 'PPT.UpdateCurrentSlide' to set a slide.", call. = FALSE)
    ppt <- PPT.UpdateCurrentSlide(ppt, i=1)   # default slide to use
  }
  
  shapes <- ppt$Current.Slide[["Shapes"]]
  slide.width <- ppt$pres[["PageSetup"]][["SlideWidth"]] 
  slide.height <- ppt$pres[["PageSetup"]][["SlideHeight"]]
  
  # initialize reference frame to full slide area if not set.
  # the frame describes the area in which the graphic is placed.
  #
  if (is.na(frame$width)) {
    frame$width <- slide.width
  }
  if (is.na(frame$height)) {
    frame$height <- slide.height
  }
  if (is.na(frame$top)) {
    frame$top <- 0
  }
  if (is.na(frame$left)) {
    frame$left <- 0
  }
  
  # change slide width / height in case a different frame of reference is selected
  # i.e. not the whole slide but a smaller region. This feature is important if we want
  # to only use a subset of the whole slide for adding grahics
  # TODO: implement frame of reference here.
  
  
  
  # include shape with a pixel size not too small. I do not know why, but size
  # 1,1 would not work and will produce blurry images. For an unknown reason the
  # size has to be reasonably big. Here the slide's dimensions are used.
  
  file <- normalizePath(file)               # absolute paths must be supplied to COM object
  file <- gsub("/", "\\\\", file)           # backslashes must be used
  
  img <- shapes$AddPicture(FileName = file, 
                           LinkToFile = 0, 
                           SaveWithDocument = -1,   # msoTriState Constant: msoFalse =0, msoTrue=-1
                           Left = 1, 
                           Top = 1, 
                           Width = slide.width, 
                           Height = slide.height)
  
  # rescale picture to full size initial size
  img$ScaleHeight(1, -1)
  img$ScaleWidth(1, -1)
  
  # calculate optimal scaling for picture to always fit on slide
  # If width and height are supplied, the graphic is rescaled so that the
  # condition (img.width <= width & img.height <= height) is satisfied 
  img.width <- img[["width"]]
  img.height <- img[["height"]]
  
  # width can be passed as fraction of frame width/height, 
  # i.e. by default slide width/height. If this is the case, image 
  # width/height is converted to pixels 
  if (!is.na(width) & width > maxscale)
    width <- width / slide.width
  if (!is.na(height) & height > maxscale)
    height <- height / slide.height  
  
  # calculate factor to rescale the image width / height 
  # In case both width / hight are passed, the smaller factor is used
  # so the image will fit onto the slide or reference frame
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
  # perform rescaling
  img$ScaleHeight(rescale.width.by, -1)
  img$ScaleWidth(rescale.height.by, -1)
  
  # locate img horizontally
  if (x == "center") 
    x.left <- slide.width / 2 - img[["Width"]] / 2
  if (x == "left") 
    x.left <- 0
  if (x == "right")
    x.left <- slide.width - img[["Width"]]  
  if (is.numeric(x))
    x.left <- x
  
  # locate img vertically
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
PPT.AddGraphicstoSlide2 <- function(ppt, 
                                    file, 
                                    width=.9, 
                                    height=.9,
                                    x="center", 
                                    y="center", 
                                    x.offset=0, 
                                    y.offset=0, 
                                    frame = list(top=0,    
                                                 left=0, 
                                                 width=1,
                                                 height=1),
                                    proportional=TRUE, 
                                    newslide=TRUE, 
                                    maxscale=1)
{
  # iterate over all files
  for (f in file) {
    ppt <- PPT.AddGraphicstoSlide2_(ppt = ppt, 
                                    file = file, 
                                    width = width, 
                                    height = height, 
                                    x = x, 
                                    y = y, 
                                    x.offset = x.offset, 
                                    y.offset = y.offset, 
                                    frame = frame,
                                    proportional = proportional, 
                                    newslide = newslide, 
                                    maxscale = maxscale)
  }
  invisible(ppt)
}
# Quasi-vectorization as mapply and Vectorize cannot be applied as
# the ppt hanlde would need to be updated in between. At least using mapply 
# it throws an error so I chose this version which will suffice 
# for most of my use cases. 