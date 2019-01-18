library(tidyverse)
library(rvest)
library(writexl)
library(RSelenium)
library(readxl)
library(data.table)
library(lubridate)

# Pre-Settings ------------------------------------------------------------

Country_list <- c("JP", "AU", "SG", "CN", "HK")

week_index <- read_excel("H:/Project/Weekly-Report/Week_Settings_DB.xlsx")
Production_query_table <- read_excel("H:/Project/Weekly-Report/Week_Settings_DB.xlsx", sheet = 2)

mother_production <- Production_query_table$Production[1]
sub_production <- Production_query_table$Sub_production

country_table <- data.table(Country=rep(Country_list,nrow(week_index)), Destination = 0,Week=1:nrow(week_index))


replace_comma <- partial(str_remove_all, pattern = ",")
to_number <- compose(as.numeric, replace_comma)

week_index$Start_Date <- str_c(day(week_index$Start_Date), month(week_index$Start_Date), year(week_index$Start_Date), sep = "/")
week_index$End_Date <- str_c(day(week_index$End_Date), month(week_index$End_Date), year(week_index$End_Date), sep = "/")

Start_date_list <- week_index$Start_Date
End_date_list <- week_index$End_Date
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

# Get buttons-Rselenium ---------------------------------------------------


remDr$navigate("https://intranet.travelzoo.com/common/production/DeliveryPeriodReport.aspx")

## Click exclude results with no delivery mannually

for (i in Country_list) {
  locale <- remDr$findElement(using = "xpath", '//*[@id="ddlLocales"]')
  locale$sendKeysToElement(list(i))

  for (n in 2:nrow(week_index)) {
    start_date <- remDr$findElement(using = "xpath", '//*[@id="txtStartDate"]')
    start_date$clearElement()
    start_date$sendKeysToElement(list(week_index$Start_Date[n]))

    end_date <- remDr$findElement(using = "xpath", '//*[@id="txtEndDate"]')
    end_date$clearElement()
    end_date$sendKeysToElement(list(week_index$End_Date[n]))

    # exclude_delivery_checkbox <- remDr$findElement(using = "xpath", '//*[@id="chkExcludeNoDelivery"]')
    # # Uncomment if want to exclude results with no delivery
    # exclude_delivery_checkbox$clickElement()

    # include_CPC <- remDr$findElement(using = "xpath",'//*[@id="chkIncludeOnlyCPC"]')
    # include_CPC$clickElement()

    product_group <- remDr$findElement(using = "xpath", '//*[@id="ddlProductGroup"]')
    product_group$sendKeysToElement(list(mother_production))

    

    for (g in sub_production) {
      sub_group <- remDr$findElement(using = "xpath", '//*[@id="ddlProductSubGroup"]')
      
      sub_group$sendKeysToElement(list(g))
      
      sbmt <- remDr$findElement(using = "xpath", '//*[@id="btnSubmit"]')
      sbmt$clickElement()
      # Sys.sleep(2)

      pivot_table <- remDr$findElement(using = "xpath",'//*[@id="grdPivot_MT"]')

      pivot_table_text <- pivot_table$getElementText()[[1]]
      
      a <- pivot_table_text
      if(i=="HK"){b <- str_match_all(a,"[:lower:]{1} \\d+,*\\d* HK[¥$]{1} (.+\\.\\d{2}) \\d")
      }else{b <- str_match_all(a,"[:lower:]{1} \\d+,*\\d* [¥$]{1} (.+\\.\\d{2}) \\d")}
      
      c <- data.table(b[[1]])
      d <- c[.N,2]
      dev_amount <- d$V2

      if(dev_amount=="0.00"){
        final_dev_amount <- dev_amount
      } else {
        final_dev_amount <- str_sub(dev_amount,1,str_locate(dev_amount," ")[1]-1)
      }
      
      if (str_sub(g, 1,1) == "D") {
        country_table[Country == i & Week==n, c("Destination") := .(Destination + to_number(final_dev_amount))]
      } else {
        country_table[Country == i & Week==n, c(g) := .(to_number(final_dev_amount))]
      }
    }
  }
}

sum_nm <- partial(sum,na.rm=TRUE)
sum_table <- country_table[,lapply(.SD, sum_nm),by=Country,.SDcols=c("Destination","Newflash (Flat fee)","Top 20 (Flat fee)","Website Placements")]
setcolorder(sum_table,c("Country" ,"Top 20 (Flat fee)","Newflash (Flat fee)","Website Placements","Destination"))
sum_table
