


#### ____________________________ ####
#### -------------------INSERT TABLE -------------------####


#' Adding a table to a slide.
#' 
#' @export
#' @example inst/examples/PPT.AddTableExample.R
#' 
PPT.AddTable <- function(ppt, 
                         df, 
                         width=.9, 
                         height=.9,
                         cell.height = .01,
                         #cell.width = .02,
                         col.align = "left",
                         left = .05,
                         top = .1,
                         colnames = TRUE,   # add coloumn names to table
                         rownames = FALSE,  # add rownames to table
                         newslide=FALSE, 
                         maxscale=1)
{

  # get width / heught of slides
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
  
  # size and position
  width_px <- width * sld.width
  height_px <- nr * cell.height * sld.height
  left_px <- left * sld.width
  top_px <- top * sld.height
  
  # Add empty table
  shp <- shapes$AddTable(NumRows = nr, NumColumns = nc, 
                         Left = left_px, Top = top_px, 
                         Width = width_px, Height = height_px)
  # shp[["Visible"]] <- FALSE
  
  # fill table with values from dataframe 
  for (i in 1L:nr) {
    for (j in 1L:nc) {
      cat("\rrow:", i, "col:", j)
      cell <- shp[["Table"]]$Cell(i, j)
      txt_rng <- cell[["Shape"]][["TextFrame"]][["TextRange"]]
      txt_rng[["Text"]] <- as.character(df[i, j])    # factors will cause an error
    }
  }

  # TODO:  
  # # set wdiths
  # ActivePresentation.Slides(2).Shapes(5).Table.Columns(1).Width = 80
  # 
  # # set style
  # With tbl.Cell(3, 3).Shape.TextFrame.TextRange
  # .Font.Bold = msoTrue
  # .Font.Size = 24
  # End With
  
  # # There is no enumeration for styles
  # shp[["Table"]]$ApplyStyle("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
  # shp[["Table"]]$ApplyStyle("{00A15C55-8517-42AA-B614-E9B94910E393}")
  # shp[["Table"]]$ApplyStyle("{93296810-A885-4BE3-A3E7-6D5BEEA58F35}")
  # shp[["Table"]][["Style"]][["Id"]]
  
  # shp[["Visible"]] <- TRUE
  
  invisible(ppt)
}







