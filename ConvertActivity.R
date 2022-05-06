require(dplyr)
require(writexl)

filePath <- file.choose()

data <- readxl::read_excel(filePath)
data <- as.data.frame(data)

data$HZ.sum.Hr1 <- rowSums(data[grep('HACTV', names(data))])
data$Dist.sum.Hr1 <- rowSums(data[grep('TOTDIST', names(data))])
data$Fine.sum.Hr1 <- rowSums(data[grep('STRCNT', names(data))])
data$CenD.sum.Hr1 <- rowSums(data[grep('CTRDIST', names(data))])
data$CenT.sum.Hr1 <- rowSums(data[grep('CTRTIME', names(data))])
data$Rear.sum.Hr1 <- rowSums(data[grep('RMOVNO', names(data))])

data <- data %>% relocate(HZ.sum.Hr1, Dist.sum.Hr1, Fine.sum.Hr1, CenD.sum.Hr1, CenT.sum.Hr1, Rear.sum.Hr1, .after = CAGE)

outputPath <- gsub(".xlsx", "_output.xlsx", filePath)
writexl::write_xlsx(data, outputPath)
