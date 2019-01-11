library(tidyverse)
library(rvest)
library(writexl)
library(RSelenium)


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

# Get buttons-Rselenium ---------------------------------------------------


remDr$navigate("https://intranet.travelzoo.com/common/production/DeliveryPeriodReport.aspx")

locale <- remDr$findElement(using = "xpath",'//*[@id="ddlLocales"]')
locale$sendKeysToElement(list("US"))

start_date <- remDr$findElement(using = "xpath",'//*[@id="txtStartDate"]')
start_date$sendKeysToElement(list("01/08/2018"))

end_date <- remDr$findElement(using = "xpath",'//*[@id="txtEndDate"]')
end_date$sendKeysToElement(list("01/01/2019"))

exclude_delivery_checkbox <- remDr$findElement(using = "xpath",'//*[@id="chkExcludeNoDelivery"]')
#Uncomment if want to exclude results with no delivery
exclude_delivery_checkbox$clickElement()

include_CPC <- remDr$findElement(using = "xpath",'//*[@id="chkIncludeOnlyCPC"]')
include_CPC$clickElement()

product_group <- remDr$findElement(using = "xpath",'//*[@id="ddlProductGroup"]')
product_group$sendKeysToElement(list("Local Deals"))
