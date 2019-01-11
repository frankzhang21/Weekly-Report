library(tidyverse)
library(rvest)
library(writexl)
library(RSelenium)


# Rselenium ---------------------------------------------------------------


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

remDr$navigate("https://intranet.travelzoo.com/common/production/DeliveryPeriodReport.aspx")

