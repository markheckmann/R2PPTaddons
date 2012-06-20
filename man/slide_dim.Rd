\name{slide_width}
\alias{slide_height}
\alias{slide_width}
\title{Get width and height of active slide}
\usage{
  slide_width(ppt)

  slide_height(ppt)
}
\arguments{
  \item{ppt}{An \pkg{R2PPT} presentation object.}
}
\description{
  Get width and height of active slide
}
\examples{
\dontrun{
 p <- PPT.Init(visible=TRUE)
 p <- PPT.AddBlankSlide(p)
 slide_width(p)
 slide_height(p)
}
}

