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

A <- read_excel(file.choose())

setDT(A)
links <- str_c("https://intranet.travelzoo.com/office/Showclassified.aspx?id=",A$AD)





# Loop --------------------------------------------------------------------

for (i in 1:1085) {
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

write_xlsx(A,'C:/Users/fzhang/OneDrive - Travelzoo/Report/Top20 Analysis/AU/2017_des.xlsx')

# Second Loop --------------------------------------------------------------------
i <- 1
for (i in 1:562) {
  remDr$navigate(links[i])
  print(i)
  tryCatch({
    hyper <- remDr$findElement(using = "xpath",'/html/body/table[2]/tbody/tr/td/a')
    linkss <- hyper$getElementAttribute("href")[[1]]
    remDr$navigate(linkss)
  webel <- remDr$findElement(using = "xpath", '//*[@id="deal-where"]')
  loc <- webel$getElementText()[[1]]
  A[AD==A$AD[i],c("location"):=.(as.character(loc))]},
  error = function(e){next},
  finally = next)
}

write_xlsx(A,'C:/Users/fzhang/OneDrive - Travelzoo/Report/Top20 Analysis/AU/second_des.xlsx')
