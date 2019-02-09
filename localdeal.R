library(data.table)
library(readxl)
library(writexl)


local_deal <- read_excel(file.choose(),sheet = "Database")

split_list <- split(local_deal,local_deal$DealCategoryName)

Entertainment <- split_list[[1]]
Getaway <- split_list[[2]]
Other <- split_list[[3]]
Restaurant <- split_list[[4]]
Spa <- split_list[[5]]

setDT(Entertainment)
setDT(Getaway)
setDT(Other)
setDT(Restaurant)
setDT(Spa)

create_summary <- function(x) {
  return(x[,.(sum_ma_revenue=sum(`Net revenue in USD`)),
                by=c("DealCategoryName","MerchantAgreementNumber","MerchantName")][order(-sum_ma_revenue)])
  
  
  
}
Entertainment_summary <- create_summary(Entertainment)
Getaway_summary <- create_summary(Getaway)
Other_summary <- create_summary(Other)
Restaurant_summary <- create_summary(Restaurant)
Spa_summary <- create_summary(Spa)

create_top50 <- function(x,y) {
  
}
entertainment_top50 <- merge(Entertainment_summary[1:50,],Entertainment,by = c("DealCategoryName","MerchantAgreementNumber","MerchantName"),all.y  = TRUE)

entertainment_final <- entertainment_top50[!is.na(sum_ma_revenue),.(MerchantName,MerchantAgreementNumber,DealPrice,RevenueShare,DealTitle,voucher_sold=sum(Successful),voucher_revenue=sum(`Net revenue in USD`),PublishedDateLocal,sum_ma_revenue),
                    by=c("DealCategoryName","MerchantAgreementNumber","MerchantName","DealPrice")][,.SD[1],by=c("DealCategoryName","MerchantAgreementNumber","MerchantName","DealPrice")]


write_xlsx(entertainment_top50[!is.na(sum_ma_revenue),.(MerchantName,MerchantAgreementNumber,DealPrice,RevenueShare,DealTitle,voucher_sold=sum(Successful),voucher_revenue=sum(`Net revenue in USD`),PublishedDateLocal,sum_ma_revenue),
                    by=c("DealCategoryName","MerchantAgreementNumber","MerchantName","DealPrice")][,.SD[1],by=c("DealCategoryName","MerchantAgreementNumber","MerchantName","DealPrice")],"E:/a.xlsx")

