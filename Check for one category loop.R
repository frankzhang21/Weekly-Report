library(tidyverse)
library(rvest)
library(writexl)
library(RSelenium)
library(readxl)
library(data.table)
library(lubridate)

# Pre-Settings ------------------------------------------------------------






week_index <- read_excel("C:/Users/fzhang/OneDrive - Travelzoo/Project/Weekly-Report/Week_Settings_DB.xlsx")





replace_comma <- partial(str_remove_all, pattern = ",")
to_number <- compose(as.numeric, replace_comma)

week_index$Start_Date <- str_c(day(week_index$Start_Date), month(week_index$Start_Date), year(week_index$Start_Date), sep = "/")
week_index$End_Date <- str_c(day(week_index$End_Date), month(week_index$End_Date), year(week_index$End_Date), sep = "/")

Start_date_list <- week_index$Start_Date
End_date_list <- week_index$End_Date

total_end_date <- End_date_list[nrow(week_index)]
current_week <- nrow(week_index)

# Settings-Rselenium ---------------------------------------------------------------


eCaps <- list(
  chromeOptions = list(
    prefs = list(
      "profile.default_content_settings.popups" = 0L,
      "download.prompt_for_download" = FALSE,
      "download.default_directory" = "C:/Users/fzhang/Desktop/Delivery"
    )
  )
)

remDr <- remoteDriver(
  remoteServerAddr =
    "127.0.0.1", port =
    4444, browserName =
    "chrome", extraCapabilities = eCaps
) # <U+8FDE><U+63A5>Server

remDr$open()
remDr$setTimeout(type = "page load", milliseconds = 200000)

# Start -------------------------------------------------------------------


remDr$navigate("https://intranet.travelzoo.com/common/production/DeliveryPeriodReport.aspx")

exclude_delivery_checkbox <- remDr$findElement(using = "xpath", '//*[@id="chkExcludeNoDelivery"]')
# Uncomment if want to exclude results with no delivery
exclude_delivery_checkbox$clickElement()

Country <- "AU"
product_name <- "CPC"
one_table <- week_index
setDT(one_table)
one_table[,c("Country","Product"):=.(Country,product_name)]
sum_number <- 0
for (i in 1:current_week) {

  
  start_date <- remDr$findElement(using = "xpath", '//*[@id="txtStartDate"]')
  start_date$clearElement()
  start_date$sendKeysToElement(list(Start_date_list[i]))
  
  end_date <- remDr$findElement(using = "xpath", '//*[@id="txtEndDate"]')
  end_date$clearElement()
  end_date$sendKeysToElement(list(End_date_list[i]))

  
  sbmt <- remDr$findElement(using = "xpath", '//*[@id="btnSubmit"]')
  sbmt$clickElement()
  
  pivot_table <- remDr$findElement(using = "xpath", '//*[@id="grdPivot_MT"]')
  
  pivot_table_text <- pivot_table$getElementText()[[1]]
  
  a <- pivot_table_text
  if (Country == "HK") {
    b <- str_match_all(a, "[:lower:]{1} \\d+,*\\d* HK[¥$]{1} (.+\\.\\d{2}) \\d")
  } else {
    b <- str_match_all(a, "[:lower:]{1} \\d+,*\\d*,*\\d* [¥$]{1} (.+\\.\\d{2}) \\d")
  }
  
  c <- data.table(b[[1]])
  d <- c[.N, 2]
  dev_amount <- d$V2
  
  if (dev_amount == "0.00") {
    final_dev_amount <- dev_amount
  } else {
    final_dev_amount <- str_sub(dev_amount, 1, str_locate(dev_amount, " ")[1] - 1)
  }
  

  print(str_c(Country,product_name,"of Week",week_index$Week_Index[i],"is",final_dev_amount,sep=" "))
  sum_number <- sum_number+to_number(final_dev_amount)
  one_table[Week_Index==week_index$Week_Index[i],Amount:=to_number(final_dev_amount)]
  if (i==current_week) {
    print(str_c(Country,product_name,"from Week",week_index$Week_Index[1],"to Week",week_index$Week_Index[i],"is",formatC(sum_number,format = "f",big.mark = ",",digits = 2),sep=" "))
    
  }
}

write_xlsx(one_table,str_c("C:/Users/fzhang/OneDrive - Travelzoo/Report/One_Category/",Country,"-",product_name," from Week ",first(week_index$Week_Index)," to Week ",last(week_index$Week_Index),".xlsx"))

