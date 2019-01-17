library(RSelenium)
library(tidyverse)
library(data.table)
library(readxl)
library(writexl)
library(lubridate)
library(rvest)
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

remDr <- remoteDriver(remoteServerAddr = "127.0.0.1" 
                      , port = 4444
                      , browserName = "chrome"
                      ,extraCapabilities = eCaps)#连接Server

remDr$open() 
remDr$setTimeout(type = "page load", milliseconds = 200000)


# Data for Net members ------------------------------------------------------

remDr$navigate("https://intranet.travelzoo.com/dashboard/subscriber/")

#Ending date for net members
net_member_end_date <- remDr$findElement(using = "xpath",'//*[@id="txtSSDateEnding"]')#Assign to var
net_member_end_date$clearElement()#Clear date input form
report_week_monday <- Sys.Date()-3
text_format_mondy <- str_c(day(report_week_monday),month(report_week_monday),year(report_week_monday),sep="/")
net_member_end_date$sendKeysToElement(list(text_format_mondy))#enter right date

Go_button <- remDr$findElement(using = "xpath",'//*[@id="btnSSSubmit"]')
Go_button$clickElement()
# Extract net members for each country
Response <- read_html(remDr$getPageSource()[[1]])
net_member_table <- html_table(Response,fill = TRUE)[[2]]
names(net_member_table) <- c("Country",str_c(c("Top20","Newsflash","Local_Deails"),"Subscriptions",sep="_"),"Subscribers")
new_member_by_country <- tibble(Country=net_member_table$Country[2:7],Net_Member=net_member_table$Subscribers[2:7])



