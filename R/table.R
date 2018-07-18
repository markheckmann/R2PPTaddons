


#### ____________________________ ####
#### -------------------INSERT TABLE -------------------####


#' Adding a table to a slide.
#' 
#' @export
#' @param text.align Horizontal alignment of text. \code{1 = left}, 
#' \code{2= center}, and \code{3 = right}. Google \code{PpParagraphAlignment} 
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
                         head.height =.01,
                         cell.height = .01,
                         #cell.width = .02,
                         left = .05,
                         top = .1,
                         font.size = 18,
                         font.bold = 0,
                         font.italic = 0,
                         font.color = "black",
                         text.align = 1,
                         colnames = TRUE,   # add coloumn names to table
                         rownames = FALSE,  # add rownames to table
                         newslide=FALSE, 
                         show.actions=TRUE,   # to updates inivisily for speed up
                         maxscale=1)
{

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
  
  # recycle vectors to match nrows / ncols
  # columns
  text.align <- rep_len(text.align, nc)
  font.size <- rep_len(font.size, nc)
  font.bold <- rep_len(font.bold, nc)
  font.italic <- rep_len(font.italic, nc)
  font.color <- rep_len(font.color, nc)
  # rows
  
  # TODO: cell.height now yet vectorized or used
  # best approach: use cell height and overwrite with 
  # header.height if header is present
  cell.height <- rep_len(cell.height, nr)
  
  # size and position
  width_px <- width * sld.width
  height_px <- sum(cell.height * sld.height)
  left_px <- left * sld.width
  top_px <- top * sld.height
  
  
  # Add empty table
  shp <- shapes$AddTable(NumRows = nr, NumColumns = nc, 
                         Left = left_px, Top = top_px, 
                         Width = width_px, Height = height_px)
  
  # hide shape while it is update
  if (!show.actions)
    shp[["Visible"]] <- FALSE
  
  # fill table with values from dataframe 
  for (i in 1L:nr) {
    for (j in 1L:nc) {
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
      
      p <- txt_rng[["ParagraphFormat"]]
      p[["Alignment"]] <- text.align[j]
      
    }
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







