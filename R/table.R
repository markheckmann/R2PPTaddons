


#### ____________________________ ####
#### -------------------INSERT TABLE -------------------####


#' Adding a table to a slide.
#' 
#' @export
#' @param text.h.align Horizontal alignment of text. \code{1 = "left"}, 
#' \code{2= "center"} (default), and \code{3 = "right"}. Google \code{PpParagraphAlignment} 
#' enumeration for more options. Eizher a string or a numeric values can be used.
#' @param text.v.align Vertical alignment of text. \code{1 = "top"}, 
#' \code{3= "center"} (dafault), and \code{4 = "bottom"}. Google \code{PpParagraphAlignment} 
#' enumeration for more options.
#' @param show.actions All actions are visisble by default. Invisible actions 
#' will often speed things up.
#' @example inst/examples/PPT.AddTableExample.R
#' @section TODO: cell height, cell.width, header.height  
#' 
PPT.AddTable <- function(ppt, 
                         df, 
                         width=.9, 
                         height=.9,
                         header.height = row.height[1],
                         row.height = 20,
                         column.width = NA,
                         left = .05,
                         top = .1,
                         font.size = 18,
                         font.bold = 0,
                         font.italic = 0,
                         font.color = "black",
                         text.h.align = 2,   # per column
                         text.v.align = 3,   # per row
                         header.h.align = text.h.align,    # header column
                         header.v.align = text.v.align[1], # header row
                         colnames = TRUE,    # add coloumn names to table
                         rownames = FALSE,   # add rownames to table
                         newslide=FALSE, 
                         show.actions=TRUE,   # to updates inivisily for speed up
                         maxscale=1)
{
  ## check and process parameters
  
  # is column width NA (default)?
  column_width_is_na <- all(is.na(column.width))
  
  # recode header.h.align if character
  if (is.character(header.h.align)) {
    header.h.align <- match.arg(tolower(header.h.align), c('left', 'center', 'right'), TRUE)
    header.h.align <- recode(header.h.align, left = 1, center = 2, right =3)
    if (any(is.na(header.h.align)))
      stop("'header.h.align' requires numeric values",
           " or one of 'left', 'center', 'right'", call. = FALSE)
  }
  # recode text.v.align if character
  if (is.character(text.h.align)) {
    text.h.align <- match.arg(tolower(text.h.align), c('left', 'center', 'right'), TRUE)
    text.h.align <- recode(text.h.align, left = 1, center = 2, right =3)
    if (any(is.na(text.h.align)))
      stop("'text.h.align' requires numeric values",
           " or one of 'left', 'center', 'right'", call. = FALSE)
  }
  # recode text.v.align if character
  if (is.character(text.v.align)) {
    text.v.align <- match.arg(tolower(text.v.align), c('top', 'center', 'bottom'), TRUE)
    text.v.align <- recode(text.v.align, top = 1, center = 3, bottom =4)
    if (any(is.na(text.v.align)))
      stop("'text.v.align' requires numeric values",
           " or one of 'top', 'center', 'bottom'", call. = FALSE)
  }
  # recode header.v.align if character
  if (is.character(header.v.align)) {
    header.v.align <- match.arg(tolower(header.v.align), c('top', 'center', 'bottom'))
    header.v.align <- recode(header.v.align, top = 1, center = 3, bottom =4)
    if (any(is.na(header.v.align)))
      stop("'header.v.align' requires a numeric value",
           " or one of 'top', 'center', 'bottom'", call. = FALSE)
  }

  # get width / height of slides
  sld.width = slide_width(ppt)
  sld.height = slide_height(ppt)

  shapes <- ppt$Current.Slide[["Shapes"]]
  
  # Adding a new slide before adding textbox if promted
  if (newslide)
    ppt <- PPT.AddBlankSlide(ppt)  
  # if the current slide object is not set, an error will occur
  if (!newslide & is.null(ppt$Current.Slide)) {  
    warning("No current slide defined. Slide 1 ist selected.", call. = FALSE)
    ppt <- PPT.UpdateCurrentSlide(ppt, i=1)
  }
  
  # add row / column names as seperate row / column
  if (colnames) {
    df <- rbind(colnames(df), df)
    rownames(df)[1] <- ""
  }
  if (rownames) {
    df <- cbind(rownames(df), df)
    colnames(df)[1] <- ""
  }

  # size dataframe
  nr <- nrow(df)
  nc <- ncol(df)
  
  # if no column width is suplied use width of table
  if (column_width_is_na) column.width <- width / nc
    
  # recycle vectors to match nrows / ncols
  # columns
  text.h.align <- rep_len(text.h.align, nc)
  header.h.align <- rep_len(header.h.align, nc)
  column.width <- rep_len(column.width, nc)
  font.size <- rep_len(font.size, nc)
  font.bold <- rep_len(font.bold, nc)
  font.italic <- rep_len(font.italic, nc)
  font.color <- rep_len(font.color, nc)
  
  # rows
  
  # initialize row heights as identical for all rows
  row.height <- rep_len(row.height, nr)
  text.v.align <- rep_len(text.v.align, nr)
  
  # replace height of first row if columns names are shown
  if (colnames) {
    row.height[1] <- header.height
    text.v.align[1] <- header.v.align 
  }
  
  # convert all row.heights as fraction of slide height
  row.height <- ifelse(row.height <= maxscale, 
                       row.height, 
                       row.height / sld.height)
  column.width <- ifelse(column.width <= maxscale, 
                         column.width, 
                         column.width / sld.width)
  
  # size and position
  row.height <- row.height * sld.height
  column.width <- column.width * sld.width
  width_px <- sum(column.width)
  height_px <- sum(row.height)
  left_px <- left * sld.width
  top_px <- top * sld.height
  
  # Add empty table
  shp <- shapes$AddTable(NumRows = nr, 
                         NumColumns = nc, 
                         Left = left_px, 
                         Top = top_px, 
                         Width = width_px, 
                         Height = height_px)
  
  # hide shape while it is update
  if (!show.actions)
    shp[["Visible"]] <- FALSE
  
  # fill table with values from dataframe 
  for (i in 1L:nr) {
    for (j in 1L:nc) {
      ## cell and column actions
      cat("\rrow:", i, "col:", j)
      cell <- shp[["Table"]]$Cell(i, j)
      txt_rng <- cell[["Shape"]][["TextFrame"]][["TextRange"]]
      txt_rng[["Text"]] <- as.character(df[i, j])    # factors will cause an error
      
      ## FONT PROPERTIES ##
      f <- txt_rng[["Font"]]
      f[["Size"]] <- font.size[j]
      f[["Bold"]] <- font.bold[j]
      f[["Italic"]] <- font.italic[j]
     
       # color
      fc <- f[["Color"]]
      fc[["RGB"]] <- color_to_integer(font.color[j])
      
      # alignment horiz 
      if (colnames & i == 1) 
        h.align <- header.h.align
      else 
        h.align <- text.h.align
      p <- txt_rng[["ParagraphFormat"]]
      p[["Alignment"]] <- h.align[j]
      # alignment vertical
      txt_frm <- cell[["Shape"]][["TextFrame"]]
      txt_frm[["VerticalAnchor"]] <- text.v.align[i]
      
    }
  }
  
  ## row actions
  for (i in 1L:nr) {
    row <- shp[["Table"]][["Rows"]]$Item(i)
    row[["Height"]] <- row.height[i]
  }
  
  ## column actions
  for (j in 1L:nc) {
    column <- shp[["Table"]][["Columns"]]$Item(j)
    column[["Width"]] <- column.width[j]
  }
  
  # TODO:  
  # # set wdiths
  # ActivePresentation.Slides(2).Shapes(5).Table.Columns(1).Width = 80
 

  # # There is no enumeration for styles
  # shp[["Table"]]$ApplyStyle("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
  # shp[["Table"]]$ApplyStyle("{00A15C55-8517-42AA-B614-E9B94910E393}")
  # shp[["Table"]]$ApplyStyle("{93296810-A885-4BE3-A3E7-6D5BEEA58F35}")
  # shp[["Table"]][["Style"]][["Id"]]
  
  # show shape after finishing upading
  if (!show.actions)
    shp[["Visible"]] <- TRUE

  invisible(ppt)
}







