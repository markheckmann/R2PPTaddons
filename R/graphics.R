

#### ____________________________ ####
#### -------------------INSERT IMAGE -------------------####


# This is the workhorse, arguments are explained in the function 
# PPT.AddGraphicstoSlide2 below.
# 
PPT.AddGraphicstoSlide2_ <- function(ppt, 
                                     file, 
                                     width=.9, 
                                     height=.8,
                                     left = .05,
                                     top = .1,
                                     hjust = .5,
                                     vjust = .5,
                                     # frame args can be passed as list to make it easier to pass shapes
                                     frame = list(),
                                     proportional=TRUE, 
                                     newslide=FALSE, 
                                     maxscale=1,
                                     display.frame = FALSE,  # show rectangle where graphic is fitted into for dev purposes  
                                     display.image = TRUE,   # add the image? Can be supressed to only add the frame 
                                     line.color = "black",
                                     line.type = 1,
                                     line.size = NA,
                                     shadow.visible = FALSE,
                                     shadow.type = 21,
                                     shadow.color = "black",
                                     shadow.x = 2,
                                     shadow.y = 2,
                                     shadow.transparency = .6,
                                     ...)
{    
  # frame in which the graphic is fitted
  frm <- list(top=top,    
              left=left, 
              width=width,
              height=height)
  
  # overwrite values passed in frame list
  f <- modifyList(frm, frame)

  # checking arguments: vjust, hjust
  hjust.sel <- c("center", "left", "right")
  vjust.sel <- c("center", "top", "bottom")
  if (is.character(hjust))
    hjust <- hjust.sel[pmatch(tolower(hjust), hjust.sel, duplicates.ok=FALSE)]  
  if (is.character(vjust))
    vjust <- vjust.sel[pmatch(tolower(vjust), vjust.sel, duplicates.ok=FALSE)]  
  if (is.na(hjust))
    stop("'hjust' must be numeric or 'center', 'left' or 'right'", call. = FALSE)
  if (is.na(vjust))
    stop("'vjust' must be numeric or 'center', 'top' or 'bottom'", call. = FALSE)
 
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
  
  slide.width <- ppt$pres[["PageSetup"]][["SlideWidth"]] 
  slide.height <- ppt$pres[["PageSetup"]][["SlideHeight"]]
  
  
  #### __ Get and add image frame ####
  
  # the frame describes the area in which the graphic is placed.
  # Convert pixel values to fractions of slide dimensions.
  # We will only operate in fractions of the slide afterwards.
  #
  if (!is.na(f$width) & f$width > maxscale) {
    f$width <- f$width / slide.width      # frame width as fraction of slide width
  }
  if (!is.na(f$height) & f$height > maxscale) {
    f$height <- f$height / slide.height   # frame height as fraction of slide height
  }
  if (!is.na(f$top) & f$top > maxscale) {
    f$top <- f$top / slide.height         # top as fraction of slide height
  }
  if (!is.na(f$left) & f$left > maxscale) {
    f$left <- f$left / slide.width        # left as fraction of slide width
  }
  
  if (!is.na(f$left) & f$left > maxscale) {
    f$left <- f$left / slide.width        # left as fraction of slide width
  }
  
  # display frame in which graphic is placed
  # for debugging purposes (default FALSE)
  if (display.frame) {
    ppt <- PPT.AddRectangle(ppt, 
                            top = f$top, 
                            left = f$left,
                            width = f$width,
                            height = f$height, 
                            line.type = 4,
                            line.color = "grey",
                            fill.color = "white")    
  }

  #### __ Add image and rescale ####

  # include shape with a pixel size not too small. I do not know why, but size
  # 1,1 would not work and will produce blurry images. For an unknown reason the
  # size has to be reasonably big. Here the slide's dimensions are used.

  file <- PPT.getAbsolutePath(file)         # absolute paths must be supplied to COM object
  file <- gsub("/", "\\\\", file)           # backslashes must be used

  # insert image if not supressed (default)
  if ( display.image ) {
    
  # add image with width/height of slide which may distort the
  # original image proportions as image will be fitted to slide
  shapes <- ppt$Current.Slide[["Shapes"]]
  img <- shapes$AddPicture(FileName = file,
                           LinkToFile = 0,
                           SaveWithDocument = -1,   # msoTriState Constant: msoFalse =0, msoTrue=-1
                           Left = 1,
                           Top = 1,
                           Width = slide.width,
                           Height = slide.height)

  # rescale image to recreate correct image proportions
  # this may cause the image to be bigger than slide.
  # Shape.ScaleHeight method: rescale by a factor
  # -1 = msoTrue: rescale with regard to original image size
  img$ScaleHeight(1, -1)
  img$ScaleWidth(1, -1)

  # calculate optimal scaling for picture to always fit in reference frame
  # If width and height are supplied, the graphic is rescaled so that the
  # condition (img.width <= width & img.height <= height) is satisfied
  img.width <- img[["width"]]
  img.height <- img[["height"]]

  # calculate factor to rescale the image width / height
  # In case both width/height are passed, the smaller factor is used
  # so the image will fit onto the slide or reference frame
  rescale.width.by <- f$width * slide.width / img.width
  rescale.height.by <- f$height * slide.height / img.height
  width.na <- is.na(f$width)
  height.na <- is.na(f$height)

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

  img[["Left"]] <- f$left * slide.width
  img[["Top"]] <- f$top * slide.height
  

  
  #### __ Align horiz / vert ####
  
  # vertically / horizontally align image inside frame 
  # this has  an effect, if the frame's dimensions are different 
  # from the image dimensions, which is usually the case.
  # To calculate the alignment we need to
  # calculate the size final of the image inside the frame first. 
  
  # convert hjust / vjust to fractions if passed as characters
  if ( is.character(hjust) ) {
    if (hjust == "left")
      hjust <- 0
    if (hjust == "center")
      hjust <- .5 
    if (hjust == "right")
      hjust <- 1
  }
  
  if ( is.character(vjust) ) {
    if (vjust == "top")
      vjust <- 0
    if (vjust == "center")
      vjust <- .5 
    if (vjust == "bottom")
      vjust <- 1
  }
  
  # initialize offset within frame
  hjust_pxl <- 0
  vjust_pxl <- 0
  
  # are h/v passed as pixels, i.e. bigger maxscale?
  hjust.is.pxl <- hjust > maxscale | hjust < -maxscale
  vjust.is.pxl <- vjust > maxscale | vjust < -maxscale
  
  # keep h/v just as is if they come as pixels
  if (hjust.is.pxl) {
    hjust_pxl <- hjust
  }
  if (vjust.is.pxl) {
    vjust_pxl <- vjust 
  }

  # current image size for calculations
  img.width <- img[["width"]]
  img.height <- img[["height"]]
  
  # calculate h/v remaining space between image and frame
  delta_width <- f$width * slide.width - img.width  
  delta_height <- f$height * slide.height - img.height  
  
  # allocate proportion of gap to left/top
  if (delta_width > 0 & !hjust.is.pxl) {
    hjust_pxl <- delta_width * hjust      
  }
  if (delta_height > 0 & !vjust.is.pxl) {
    vjust_pxl <- delta_height * vjust    
  }
  
  # frame position + just (in pxl)
  img.left <- f$left * slide.width + hjust_pxl
  img.top <- f$top * slide.height + vjust_pxl
  
  # position image at top left of frame by default
  img[["Left"]] <- img.left
  img[["Top"]] <- img.top 

  # issue warning if the image exceeds borders of slide
  if (img.left < 0 | 
      img.left + img.width > slide.width |
      img.top < 0 | 
      img.top + img.height > slide.height) {
      warning("Image exceeds borders of slide.", call. = FALSE)
  }
  
  
  
  #### __ Line and form effect ####
  
  ### format border line
   
  obj <- img[["Line"]]  # get line format object
  
  if ( !is.na(line.size) & line.size != 0 ) {
    obj[["DashStyle"]] = line.type  # dashed, see: MsoLineDashStyle enumeration
    obj[["ForeColor"]][["RGB"]] = color_to_integer(line.color)
    obj[["Weight"]] = line.size
    obj[["Visible"]] = T
  }

  ### shadow
  
  s <- img[["Shadow"]]  # get shadow format object
  
  if (shadow.visible) {
    s[["Visible"]] = shadow.visible
    s[["Type"]] = shadow.type
    s[["ForeColor"]][["RGB"]] = color_to_integer(shadow.color)
    s[["OffsetX"]] = shadow.x
    s[["OffsetY"]] = shadow.y
    s[["Transparency"]] = shadow.transparency
  }
  
  }  # end if display.image == TRUE
  
  invisible(ppt)
}



#' Adding a graphic to a slide.
#'
#' \code{PPT.AddGraphicstoSlide2} is a modified version of
#' \code{PPT.AddGraphicstoSlide} from \pkg{R2PPT}. The PPT.AddGraphicstoSlide
#' function has the drawback that it requires the size of the graphic to be
#' supplied by the user. If the size is not supplied correctly, the graphic is
#' deformed. This function keeps the apsect ration intact and offers many
#' additional features for placing the image on the slide.
#'
#' @param ppt The ppt object as used in \pkg{R2PPT}.
#' @param file Path to the graphic file.
#' @param width,height Width / height of graphic. For values smaller than
#'   \code{maxscale} (default is \code{1}) refers to a proportion of the current
#'   slide's width. Values bigger than \code{maxscale} are interpreted as
#'   pixels.
#' @param top,left Horizontal und vertical alignment of graphic inside frame.
#'   Either a string (\code{"center", "left", "right"}) or (\code{"center",
#'   "top", "bottom"}) or a numerical value. Numerical values bigger than
#'   \code{maxscale} are interpreted as absolute pixels, smaller ones as
#'   proportions.
#' @param hjust,vjust Horizontal und vertical alignment of image inside frame.
#'   Either a string (\code{"center", "left", "right"}) or (\code{"center",
#'   "top", "bottom"}) or a numerical value between \code{[0,1]}. Values bigger
#'   than \code{maxscale} are used for absolute horizontal und vertical offset.
#' @param proportional  Logical (default \code{TRUE}). Whether scaling preserves
#'   the aspect ratio of the graphic. See details section for additional
#'   information.
#' @param newslide  Logical (default is \code{TRUE}) Whether the graphic will be
#'   placed on a new slide.
#' @param maxscale  Threshold below which values are interpreted as proportional
#'   scaling factors for the \code{width} and \code{height} argument. Above the
#'   threshold values are interpreted as pixels.
#' @param display.frame  Add a shape representing the frame into which the image
#'   is placed. This makes it easier when developing new slides (default
#'   \code{FALSE}).
#' @param display.image  Whether or not the image should be  added (default
#'   \code{TRUE}).
#' @param line.color Color of text either as hex value or color name.
#' @param line.type \code{1} = solid (default), \code{2-8}= dots, dashes and
#'   mixtures. See MsoLineDashStyle Enumeration for details.
#' @param line.size Thickness of line (default\code{1}).
#' @param shadow.visible Show shadow (default \code{FALSE}).
#' @param shadow.type 1-20, see MsoShadowType enumeration (default \code{21}).
#' @param shadow.color Color of shadow (default \code{"black"}).
#' @param shadow.x,shadow.y Size of shadow.
#' @param shadow.transparency Shadow strength. 
#' 
#' @note The common use case is to add graphics and scale them while preserving
#'   their aspect ratio. In the case this this is not wanted the argument
#'   \code{proportional} can be set to \code{FALSE}. When the aspect ratio is
#'   preserved and both arguments \code{width} and \code{height} are supplied,
#'   the graphic is scaled so that it will not exceed the size given by one of
#'   these values. This is useful when adding a lot of graphics of different
#'   size. Only using \code{width} may result in graphics exceeding the slide
#'   vertically and vice versa. using both parameters (the default) will cause
#'   the graphic to be scaled so it will fit the slide regardless of its
#'   dimensions.
#'
#' @author Mark Heckmann
#' @export
#' @example inst/examples/PPT.AddGraphicstoSlideExample.R
#'   
PPT.AddGraphicstoSlide2 <- function(ppt, 
                                    file, 
                                    width=.9, 
                                    height=.8,
                                    left = .05,
                                    top = .1,
                                    hjust = .5,
                                    vjust = .5,
                                    # frame args can be passed as list to make it easier to pass shapes
                                    frame = list(),
                                    proportional=TRUE, 
                                    newslide=TRUE, 
                                    maxscale=1,
                                    display.frame = FALSE,  # show rectangle where graphic is fitted into for dev purposes  
                                    display.image = TRUE,
                                    # border line properties
                                    line.color = "black",
                                    line.type = 1,
                                    line.size = 0,
                                    # shadow
                                    shadow.visible = FALSE,
                                    shadow.type = 21,
                                    shadow.color = "black",
                                    shadow.x = 3,
                                    shadow.y = 3,
                                    shadow.transparency = .6
                                    )
{
  # iterate over all files
  for (f in file) {
    ppt <- PPT.AddGraphicstoSlide2_(ppt = ppt, 
                                    file = f, 
                                    width = width, 
                                    height = height, 
                                    left = left,
                                    top = top,
                                    hjust = hjust, 
                                    vjust = vjust, 
                                    frame = frame,
                                    proportional = proportional, 
                                    newslide = newslide, 
                                    maxscale = maxscale,
                                    display.frame = display.frame,
                                    display.image = display.image,
                                    line.color = line.color,
                                    line.type = line.type,
                                    line.size = line.size,
                                    shadow.visible = shadow.visible,
                                    shadow.type = shadow.type,
                                    shadow.color = shadow.color,
                                    shadow.x = shadow.x,
                                    shadow.y = shadow.y,
                                    shadow.transparency = shadow.transparency)
  }
  invisible(ppt)
}
# Quasi-vectorization as mapply and Vectorize cannot be applied as
# the ppt handle would need to be updated in between. At least using mapply 
# it throws an error so I chose this version which will suffice 
# for most of my use cases. 






####.####
#### ____________________________ ####
#### ---------------FIT IMAGE INTO SHAPE ------------------####


#' Fit an image into an existing shape
#'
#' Sometimes shapes serve as placeholders for an image. The function takes a
#' shape, fits an image in its position and deletes the placeholder shape
#' afterwards.
#'
#' @param ppt The ppt object as used in \pkg{R2PPT}.
#' @param file Path to the image file.
#' @param shp Pointer to the shape which the image is fitted into.
#' @inheritParams PPT.AddGraphicstoSlide2
#' @param delete.shape Whether to destroy the placeholder shape afterwards
#'   (default \code{TRUE}).
#' @author Mark Heckmann
#' @export
#' @example inst/examples/PPT.FitGraphicIntoShapeExample.R
#'   
PPT.FitGraphicIntoShape <- function(ppt, 
                                    file, 
                                    shp,        # shape to place inside
                                    hjust = "center",
                                    vjust = "center",
                                    proportional=TRUE, 
                                    maxscale=1,
                                    delete.shape = TRUE)
{
  # position of shape and pointer to shape's slide
  frm <- get_shape_position(shp) 
  sld <- shp[["Parent"]]
  
  # update current slide to insert graphic on correct slide
  ppt <- PPT.UpdateCurrentSlide(ppt, slide=sld)
  
  # add graphic using shapes position as the frame 
  # to fit the image into
  p <- PPT.AddGraphicstoSlide2(p, 
                               file, 
                               frame=frm, 
                               hjust = hjust,
                               vjust = vjust,
                               newslide = F,
                               maxscale=maxscale)
  # destroy shape the image was fitted onto
  if (delete.shape)
    shp$Delete()
  
  # return ppt object
  invisible(p)
}



####.####
#### ____________________________ ####
#### -------- FIND TEXT AND REPLACE BY GRAPHIC --------####


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
  #shapes <- slide[["Shapes"]]        # get all shapes on slide
  nshapes <- shapes[["Count"]]        # number of shapes
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
  r[!sapply(r, is.null)]        # erase NULL elements
}



#' Replace matching text by graphic
#'
#' Looks through all shapes and finds a shape with matching text pattern. The
#' shape is deleted and a graphic is inserted on the shape's parent slide.
#'
#' @param ppt   The ppt object as used in \pkg{R2PPT}.
#' @param what  Text pattern to match against.
#' @param file  Path to the graphic file.
#' @param shape.type Shape types in which the text pattern is searched for. By
#'   default only plain text fields (\code{shape.type = 17}) are searched. Other
#'   shapes, e.g. rectangles with text, are ignored. To search all shapes use
#'   \code{shape.type = NA}. The types are documented in the
#'   \code{MsoAutoShapeType} enumeration in Microsoft's MSDN docu.
#' @param ... Arguments passed on to \code{\link{PPT.AddGraphicstoSlide2}}.
#' @author Mark Heckmann
#' @export
#' @example inst/examples/PPT.ReplaceTextByGraphicExample.R
#'
PPT.ReplaceTextByGraphic <- function(ppt, what, file, shape.type = 17, ...)
{
  slides <- ppt$pres[["Slides"]]
  ss <- slides_retrieve_shapes(slides, what)   # get all shape objects that match text pattern 
  
  # only keep specified shape types
  if (!is.na(shape.type)) {
    ss_types <- sapply(ss, function(s) s[["Type"]] )  # get shape type property 
    ii <- ss_types %in% shape.type                    # only keep shapes of specified type to replace
    ss <- ss[ii]    
  }

  if (length(ss) == 0)
    warning("No shape with matching text pattern was not found.", call. = FALSE)
  if (length(ss) > 1)
    warning("More than one shape with matching text pattern found and replaced.", call. = FALSE)
  
  # loop over shapes and replace with image
  for (s in ss) {               # delete from last to first
    sld <- s[["Parent"]]        # get shape's slide
    #sld$Select()                # shape select throws error if focus is not on shape's slide, so select parent first
    s$Delete()                  # delete shape
    ppt <- PPT.UpdateCurrentSlide(ppt, slide=sld)    # PPT.AddGraphicstoSlide2 needs ppt$CurrentSlide to be set
    PPT.AddGraphicstoSlide2(ppt, file, newslide=FALSE, ...)
  }  
  
  invisible(ppt)
}


# Works but currently not needed.
# get pointers to all shapes on slide as a list
# slide:  pointer to slide
#
# get_slide_shape_pointers <- function(slide)
# {
#   shapes <- slide[["Shapes"]]               # get all shapes on slide
#   nshapes <- shapes[["Count"]]              # number of shapes
#   l <- list()
#   if (nshapes == 0)
#     return(l)
#   for (i in 1L:nshapes) {
#     l[[i]] <- shapes$Item(i)
#   }
#   l
# }
  








