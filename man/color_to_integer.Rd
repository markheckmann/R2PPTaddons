% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/color.R
\name{color_to_integer}
\alias{color_to_integer}
\title{Convert color value to long integer used in COM model for colors}
\usage{
color_to_integer(x)
}
\arguments{
\item{x}{Hex value (e.g. "#00FF00") or R color name (e.g., "blue").}
}
\value{
Numeric vector.
}
\description{
Internally, a color is representaed as a single integer in the PPT COM model.
The function takes hex values or R color names and convert them into the
correponding integer.
}
\examples{
color_to_integer("green")
color_to_integer(c("green", "#00FF00"))
color_to_integer(colors())

}
\author{
Mark Heckmann
}
\keyword{internal}
