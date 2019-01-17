library(readxl)
library(tidyverse)
library(stringdist)
library(writexl)
library(data.table)

holiday <- read_excel(file.choose())
a <- holiday %>% 
  mutate(Sound=phonetic(Headline))
setDT(a)


a[,Count:=.N,by=Sound][]
write_xlsx(a,"H:/a.xlsx")

setDT(holiday)
holiday[,c("Sum_clicks","Sum_second_clicks"):=NULL]
holiday[!is.na(Sound),c("Sum_clicks","Sum_second_clicks"):=lapply(.SD,sum),by=Sound,.SDcols=c("Clicks","Secondary Clicks")]
with_dup <- holiday[!is.na(Sound),.SD[1],by=Sound]
without_dup <- holiday[is.na(Sound),]

write_xlsx(with_dup,"H:/with.xlsx")
write_xlsx(without_dup,"H:/without.xlsx")
