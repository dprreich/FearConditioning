require(tidyverse)
require(writexl)
require(janitor)
require(zoo)

#import file
filePath <- file.choose()

data = readxl::read_excel(filePath)
data = as.data.frame(data)

###Cleaning data
#remove unnecessary rows (header stuff)
data = data[-c(1:15), ]

#create column names
header.true = function(df) {
  names(df) = as.character(unlist(df[1,]))
  df[-1,]
}
data = header.true(data)

#Remove blank rows
data = data %>% filter(row_number() %% 7 != 0) ## Delete every 7th row starting from 0 (these rows are blank)

#Fill Animal Column
data$Animal = na.locf(na.locf(data$Animal),fromLast=TRUE)

##Reshape data into wide format
data = reshape(data, idvar = "Animal", timevar = "Component Name", direction = "wide")

empty_columns <- colSums(is.na(data) | data == "") == nrow(data)
data = data[, !empty_columns]

data = data[,-c(12:14, 22:24, 32:34, 42:44, 52:54) ]

#output
outputPath <- gsub(".xlsx", "_output.xlsx", filePath)
writexl::write_xlsx(data, outputPath)

