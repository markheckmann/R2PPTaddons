\name{PPT.AddGraphicstoSlide2}
\alias{PPT.AddGraphicstoSlide2}
\title{Adding a graphic to a slide.}
\usage{
  PPT.AddGraphicstoSlide2(ppt, file, width = 0.9,
    height = 0.9, x = "center", y = "center", x.offset = 0,
    y.offset = 0, proportional = TRUE, newslide = FALSE,
    maxscale = 1)
}
\arguments{
  \item{ppt}{The ppt object as used in \pkg{R2PPT}.}

  \item{file}{Path to the graphic file.}

  \item{width}{Width of graphic. For values smaller than
  \code{maxscale} (default is \code{1}) this refers to a
  proportion of the current slide's width. Values bigger
  than \code{maxscale} are interpreted as pixels.If
  \code{NA} only the height argument is used for sclaing.}

  \item{height}{Height of graphic. For values smaller than
  \code{maxscale} (default is \code{1}) this refers to a
  proportion of the current slide's height. Values bigger
  than \code{maxscale} are interpreted as pixels. If
  \code{NA} only the width argument is used for scaling.}

  \item{x}{Horizontal placement of graphic. Either a string
  (\code{"center", "left", "right"}) or a numerical value.
  Numerical values are interpreted as absolute position in
  pixels counted from the left of the slide.}

  \item{y}{Vertical placenment of graphic.Either a string
  (\code{"center", "top", "bottom"}) or a numerical value.
  Numerical values are interpreted as absolute position in
  pixels counted from the top of the slide.}

  \item{x.offset}{Additional horizontal offset in pixel.
  Used for finetuning position on slide.}

  \item{y.offset}{Additional horizontal offset in pixel.
  Used for finetuning position on slide.}

  \item{proportional}{Logical (default \code{TRUE}).
  Whether scaling preserves the aspect ratio of the
  graphic. See details section for additional information.}

  \item{newslide}{Logical (default is \code{FALSE}) Whether
  te graphica will be placed on a new slide.}

  \item{maxscale}{Threshold below which values are
  interpreted as proportional scaling factors for the
  \code{width} and \code{height} argument.  Above the
  threshold values are interpreted as pixels.}
}
\description{
  \code{PPT.AddGraphicstoSlide2} is a modified version of
  \code{PPT.AddGraphicstoSlide} from \pkg{R2PPT}. The
  PPT.AddGraphicstoSlide function has the drawback that it
  requires the size of the graphic to be supplied by the
  user. If the size is not supplied correctly, the graphic
  is deformed.
}
\note{
  The common use case is to add graphics and scale them
  while preserving its' aspect ratio. In the case this this
  is not wanted the argument \code{proportional} can be set
  to \code{FALSE}.  When the aspect ratio is preserved and
  both arguments \code{width} and \code{height} are
  supplied the graphic is scaled so that it will not exceed
  any of these values. This is useful when adding a lot of
  graphics of different size. Only using \code{width} may
  result in graphics exceeding the slide vertically and
  vice versa. using both parameters (the default) will
  cause the graphic to be scaled so it will fit the slide
  regardless of its dimensions.
}
\examples{
\dontrun{

 # simple example
 file <- tempfile()
 png(file)
 pairs(iris[1:4], col=iris$Species)
 dev.off()

 myPres <- PPT.Init(visible=TRUE)
 myPres <- PPT.AddBlankSlide(myPres)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file)

 # some images
 myPres <- PPT.Init(visible=TRUE)
 files <- arbitrary_graphics(3, "png")
 for (f in files) {
   myPres <- PPT.AddBlankSlide(myPres)
   myPres <- PPT.AddGraphicstoSlide2(myPres, f)
 }


 # modifying graphics size and placement
 myPres <- PPT.Init(visible=TRUE)
 myPres <- PPT.AddBlankSlide(myPres)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file)

 # A common use case is to add graphics to an empty slide.
 # Setting newslide=TRUE will automatically add an empty slide.

 myPres <- PPT.AddGraphicstoSlide2(myPres, file, newslide=TRUE)

 # width and height values smaller than maxscale (default is 1)
 # are interpreted as a proportions of the slide width/height. Values
 # greater than 1 are taken as absolute pixel width/height.

 myPres <- PPT.AddGraphicstoSlide2(myPres, file, width=.5, newslide=TRUE)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file, height=.9, newslide=TRUE)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file, width=.5, height=.9, newslide=TRUE)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file, width=.5, height=.9,
                                   proportional=FALSE, newslide=TRUE)

 # using pixel width/height instead of proportions of
 # available slide width/height
 myPres <- PPT.AddBlankSlide(myPres)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file, width=400)
 myPres <- PPT.AddBlankSlide(myPres)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file, height=200)
 myPres <- PPT.AddBlankSlide(myPres)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file, width=400, height=200)
 myPres <- PPT.AddBlankSlide(myPres)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file, width=400, height=200, proportional=FALSE)

 # changing x and y placement
 myPres <- PPT.AddBlankSlide(myPres)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file, x="left")
 myPres <- PPT.AddBlankSlide(myPres)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file, x="right", y="bottom")

 # adding offset for finer placement control
 myPres <- PPT.AddBlankSlide(myPres)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file, width=400)
 myPres <- PPT.AddBlankSlide(myPres)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file, width=400, x="left")
 myPres <- PPT.AddBlankSlide(myPres)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file, width=400, x="left", x.offset=10)
 myPres <- PPT.AddBlankSlide(myPres)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file, width=500, height=200, proportional=FALSE)


 # disproportional scaling
 myPres <- PPT.AddBlankSlide(myPres)
 myPres <- PPT.AddGraphicstoSlide2(myPres, file, width=.5, height=1, proportional=F)


 # myPres <- PPT.Init(visible=TRUE)
 # d <- "R2PPTaddons/pics/"
 # files <- paste0(getwd(), "/", d, dir(d))
 # for (file in files) {
 #   myPres <- PPT.AddBlankSlide(myPres)
 #   myPres <- PPT.AddGraphicstoSlide2(myPres, file, .9)
 # }
}
}
\author{
  Mark Heckmann
}

