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

A <- read_excel('C:/Users/fzhang/OneDrive - Travelzoo/Report/Top20 Analysis/CN/Book3.xlsx')

setDT(A)
links <- str_c("https://intranet.travelzoo.com/office/Showclassified.aspx?id=",A$AD)





# Loop --------------------------------------------------------------------

for (i in 1:3775) {
  remDr$navigate(links[i])
  print(i)
  tryCatch({webElem <- remDr$findElements("css", "iframe")
  remDr$switchToFrame(webElem[[1]])
  webel <- remDr$findElement(using = "xpath", '//*[@id="deal-where"]')
  loc <- webel$getElementText()[[1]]
  A[AD==A$AD[i],c("location"):=.(as.character(loc))]},
  error = function(e){next},
  finally = next)
}

write_xlsx(A,'C:/Users/fzhang/OneDrive - Travelzoo/Report/Top20 Analysis/CN/result.xlsx')
