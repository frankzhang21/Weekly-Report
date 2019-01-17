library(RSelenium)
library(tidyverse)
library(data.table)
library(readxl)
library(writexl)
library(lubridate)
library(rvest)


# Functions dim -----------------------------------------------------------

replace_comma <- partial(str_remove_all,pattern=",")
to_interger <- compose(as.integer,replace_comma)

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
final_table <- data.table(new_member_by_country)

# Data for New&Unsubs -----------------------------------------------------

#junmp to new&unsubs

new_unsub_start_date <- "1/1/2019"
new_unsub_button <- remDr$findElement(using = "xpath",'//*[@id="rptTabs_ctl02_lnkTab"]')
new_unsub_button$clickElement()

#Set dates first
new_unsub_start_date_form <- remDr$findElement(using = "xpath",'//*[@id="txtSUTFrom"]')
new_unsub_start_date_form$clearElement()
new_unsub_start_date_form$sendKeysToElement(list(new_unsub_start_date))

new_unsub_end_date_form <- remDr$findElement(using = "xpath",'//*[@id="txtSUTTo"]')
new_unsub_end_date_form$clearElement()
new_unsub_end_date_form$sendKeysToElement(list(text_format_mondy))

#Loop over different country

for (i in final_table$Country) {
  

locale_button <- remDr$findElement(using = "xpath",'//*[@id="drpSUTLocale"]')
Country_name <- i
locale_button$sendKeysToElement(list(Country_name))



scale <- remDr$findElement(using = "xpath",'//*[@id="drpSUTScale"]')
scale$sendKeysToElement(list("monthly"))

Go_button <- remDr$findElement(using = "xpath",'//*[@id="btnSUTGo"]')
Go_button$clickElement()

Response <- read_html(remDr$getPageSource()[[1]])
new_unsub_table <- html_table(Response,fill = TRUE)[[2]]

final_table[Country==Country_name,c("New_sub","Un_sub"):=.(to_interger(new_unsub_table$Subscribes[1]),
            to_interger(new_unsub_table$`Total Unsubscribes`[1]))]
}
