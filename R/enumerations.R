# enumerations

enumeration <- list()  

file <- system.file("extdata", "MsoShapeType.csv", package = "R2PPTaddons")
enumeration$MsoShapeType <- read.csv(file)

file <- system.file("extdata", "MsoZOrderCmd.csv", package = "R2PPTaddons")
enumeration$MsoZOrderCmd <- read.csv(file)