library(RSelenium)
library(tidyverse)
library(data.table)
library(readxl)
library(writexl)
library(lubridate)
library(rvest)


# Functions dim -----------------------------------------------------------

replace_comma <- partial(str_remove_all, pattern = ",")
to_interger <- compose(as.integer, replace_comma)

change_date <- function(x){
  return(str_c(day(x), month(x), year(x), sep = "/"))
}
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


# Data for Net members ------------------------------------------------------

remDr$navigate("https://intranet.travelzoo.com/dashboard/subscriber/")


# Top20 CTR Report --------------------------------------------------------

Country_list <- c("AE", "AU", "CA", "CN", "DE", "ES", "FR", "HK", "JP", "TW", "UK", "US")

Top20_CTR <- data.table(Country = Country_list, CTR = "cha", Open_rate = 100,Week_Index=sort(rep(c(2:52),12)))
setorder(Top20_CTR,Country)

week_list <- remDr$findElement(using = "xpath",'//*[@id="Select1"]')

week_list$sendKeysToElement(list('32'))
for (k in 2:52) {
  week_list <- remDr$findElement(using = "xpath",'//*[@id="Select1"]')
  
  week_list$sendKeysToElement(list(as.character(k)))
  print(k)
for (i in Country_list) {
  locale_button <- remDr$findElement(using = "xpath", '//*[@id="intranetHeaderBreadcrumbs"]/form/select')
  Country_name <- i
  locale_button$sendKeysToElement(list(Country_name))
  
  Sys.sleep(1.5)
  
  Response <- read_html(remDr$getPageSource()[[1]])
  production_table <- html_table(Response, fill = TRUE)[[7]]
  names(production_table) <- c("names", "data")
  Top20_CTR[Country == Country_name & Week_Index==k, c("CTR", "Open_rate") := .(
    str_remove_all(production_table$data[3], "%"),
    to_interger(str_remove_all(production_table$data[5], "%"))
  )]
}
}
Top20_CTR[, CTR := as.numeric(CTR)]

write_xlsx(as.data.frame(Top20_CTR), "H:/Report/Weekly/Report parts/Top20_CTR.xlsx")
