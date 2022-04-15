# NEWS

## TODO

* hide / unhide shapes by pattern
* allow wildcards in text matching for replacing?
* add PPT.AddTable function

## 0.3 (under development)

* Updated docs and twekaed DESCRIPTION (UTF8, no staged installation)
* Add `z.order` argument to `PPT.AddGraphicstoSlide2`, tweaked examples and sample PPTX 
  to show feature
* Add `MsoZOrderCmd` enumeration
* PPT.UpdateAutosizedTextboxes: Update size of text boxes with autosize set to true
* PPT.AddTextBox: Font name argument added
* PPT.ReplaceShapeByGraphic: Find shape using text pattern and replace by image
* PPT.ReplaceGraphic: the shapes to operate on (e.g. only text fields) can now be specified.
* PPT.AddGraphicstoSlideExample rewritten, not compatible with older version any more
* PPT.FitGraphicInShape2 function to place graphics inside shape
* Add frame of reference to PPT.AddGraphicstoSlide2 to allow better positioning
* PPT.AddShape: add arbitarry shapes and modify, position, size, filling and lines

## 0.2

* PPT.UpdateCurrentSlide to set the current slide
* PPT.ReplaceTextByGraphic: find text, delete it and place graphic on slide
* reworked PPT.AddGraphicstoSlide2

## 0.1

* PPT.AddGraphicstoSlide2: features for adding graphics to slides
