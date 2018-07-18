

# convert hex or R color value to integer
# used to presnt colors in COM model
color_to_integer_ <- function(x)
{
  # x must be a hex of color name string
  if (!is.character(x)) {
    stop("Please supply a hex or R color name string", .Call = FALSE)
  }
  
  # stop if hex format is incorrect
  is_hex <- stringr::str_detect(x, "^#")     # starts with #?
  if (is_hex & !is_valid_hex(x)) {
    stop("The hex value '", x, "' does not have required format format '#rrggbb' or '#rgb'", 
         "Alpha values must not be specified.", .Call = FALSE)
  }
  
  # convert to long integer
  rgb_val <- as.vector(col2rgb(x))   # convert to RGB
  rgb_to_long(rgb_val)               # RGB to long integer
}


# conversion formula RGB to LONG:
# B * 256*256 + G * 256 + R
# x: rgb vector
rgb_to_long <- function(x)
{
  f <- c(r=1, g=256, b=256*256)
  sum(x * f)
}


# check for valid 6 element hex value
# alpha values (positions 7-8) are not allowed
#
is_valid_hex <- function(x)
{
  # https://stackoverflow.com/questions/1636350/how-to-identify-a-given-string-is-hex-color-format
  stringr::str_detect(x, stringr::regex("^#(?:[0-9a-fA-F]{3}){1,2}$"))
}


#' Convert color value to long integer used in COM model for colors
#'
#' Internally, a color is representaed as a single integer in the PPT COM model.
#' The function takes hex values or R color names and convert them into the
#' correponding integer.
#' 
#' @param x      Hex value (e.g. "#00FF00") or R color name (e.g., "blue").
#' @author Mark Heckmann
#' @export
#' @keywords internal
#' @return  Numeric vector.
#' @examples 
#' color_to_integer("green")
#' color_to_integer(c("green", "#00FF00"))
#' color_to_integer(colors())
#' 
color_to_integer <- Vectorize(color_to_integer_, USE.NAMES = FALSE)


