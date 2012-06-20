\name{arbitrary_graphics}
\alias{arbitrary_graphics}
\title{Generate some png graphics}
\usage{
  arbitrary_graphics(n, dev = "png", ...)
}
\arguments{
  \item{n}{Number of graphics to be generated.}

  \item{...}{Arguments passed on to tze specific graphic
  device function (\code{png, bmp, jpeg})}
}
\value{
  A vector of temporary files.
}
\description{
  Generate some png graphics
}
\examples{
# generate some graphics
 files <- arbitrary_graphics(4, dev="png")
 files
}
\author{
  Mark Heckmann
}
\keyword{internal}

