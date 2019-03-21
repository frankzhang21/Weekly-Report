library(readxl)
library(writexl)
library(data.table)

ld <- read_excel("H:/Report/Top20 Analysis/AU/Top20/With_Destinations_2017-2019/DB_final.xlsx")

setDT(ld)

ld_2017 <- ld[`Publication Year`=="2017",]

