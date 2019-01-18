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

# Ending date for net members
net_member_end_date <- remDr$findElement(using = "xpath", '//*[@id="txtSSDateEnding"]') # Assign to var
net_member_end_date$clearElement() # Clear date input form
report_week_monday <- Sys.Date() - 3
text_format_mondy <- str_c(day(report_week_monday), month(report_week_monday), year(report_week_monday), sep = "/")
net_member_end_date$sendKeysToElement(list(text_format_mondy)) # enter right date

Go_button <- remDr$findElement(using = "xpath", '//*[@id="btnSSSubmit"]')
Go_button$clickElement()
# Extract net members for each country
Response <- read_html(remDr$getPageSource()[[1]])
net_member_table <- html_table(Response, fill = TRUE)[[2]]
names(net_member_table) <- c("Country", str_c(c("Top20", "Newsflash", "Local_Deails"), "Subscriptions", sep = "_"), "Subscribers")
new_member_by_country <- tibble(Country = net_member_table$Country[2:7], Net_Member = to_interger(net_member_table$Subscribers[2:7]))
final_table <- data.table(new_member_by_country)

# Data for New&Unsubs -----------------------------------------------------

# junmp to new&unsubs

new_unsub_start_date <- "1/1/2019"
new_unsub_button <- remDr$findElement(using = "xpath", '//*[@id="rptTabs_ctl02_lnkTab"]')
new_unsub_button$clickElement()

# Set dates first
new_unsub_start_date_form <- remDr$findElement(using = "xpath", '//*[@id="txtSUTFrom"]')
new_unsub_start_date_form$clearElement()
new_unsub_start_date_form$sendKeysToElement(list(new_unsub_start_date))

new_unsub_end_date_form <- remDr$findElement(using = "xpath", '//*[@id="txtSUTTo"]')
new_unsub_end_date_form$clearElement()
new_unsub_end_date_form$sendKeysToElement(list(text_format_mondy))

# Loop over different country

for (i in final_table$Country) {
  locale_button <- remDr$findElement(using = "xpath", '//*[@id="drpSUTLocale"]')
  Country_name <- i
  locale_button$sendKeysToElement(list(Country_name))



  scale <- remDr$findElement(using = "xpath", '//*[@id="drpSUTScale"]')
  scale$sendKeysToElement(list("monthly"))

  Go_button <- remDr$findElement(using = "xpath", '//*[@id="btnSUTGo"]')
  Go_button$clickElement()

  Response <- read_html(remDr$getPageSource()[[1]])
  new_unsub_table <- html_table(Response, fill = TRUE)[[2]]

  final_table[Country == Country_name, c("New_sub", "Un_sub") := .(
    to_interger(new_unsub_table$Subscribes[1]),
    to_interger(new_unsub_table$`Total Unsubscribes`[1])
  )]
}


# Data for production -----------------------------------------------------

remDr$navigate("https://intranet.travelzoo.com/office/top20/eu/?tzlocale=10")

## Manually adjust week number and show stats


for (i in final_table$Country) {
  locale_button <- remDr$findElement(using = "xpath", '//*[@id="intranetHeaderBreadcrumbs"]/form/select')
  Country_name <- i
  locale_button$sendKeysToElement(list(Country_name))

  Sys.sleep(1.5)

  Response <- read_html(remDr$getPageSource()[[1]])
  production_table <- html_table(Response, fill = TRUE)[[7]]
  clicks_table <- html_table(Response, fill = TRUE)[[6]]
  clicks_string <- names(clicks_table)[5]
  names(production_table) <- c("names", "data")
  final_table[Country == Country_name, c("Delivery", "Clicks", "Open_Count", "Primary", "Secondary") := .(
    to_interger(production_table$data[1]), # Newsletters Sent
    to_interger(production_table$data[2]), # Newsletters Clicks
    to_interger(production_table$data[4]), # Open Count
    to_interger(str_match(clicks_string, "Primary (.+)\\n")[, 2]), # Primary Clicks
    to_interger(str_match(clicks_string, "Secondary (.+)")[, 2])
  )] # Secondary Clicks
}



reshape_version <- melt(final_table, id.vars = "Country") %>%
  spread(Country, value) %>%
  setcolorder(c("variable", "JP", "AU", "CN", "HK", "TW", "AE"))

write_xlsx(as.data.frame(reshape_version), "H:/Report/Weekly/Report parts/Market_Production.xlsx")


# Top20 CTR Report --------------------------------------------------------

Country_list <- c("AE", "AU", "CA", "CN", "DE", "ES", "FR", "HK", "JP", "TW", "UK", "US")

Top20_CTR <- data.table(Country = Country_list, CTR = "cha", Open_rate = 100)

for (i in Country_list) {
  locale_button <- remDr$findElement(using = "xpath", '//*[@id="intranetHeaderBreadcrumbs"]/form/select')
  Country_name <- i
  locale_button$sendKeysToElement(list(Country_name))

  Sys.sleep(1.5)

  Response <- read_html(remDr$getPageSource()[[1]])
  production_table <- html_table(Response, fill = TRUE)[[7]]
  names(production_table) <- c("names", "data")
  Top20_CTR[Country == Country_name, c("CTR", "Open_rate") := .(
    str_remove_all(production_table$data[3], "%"),
    to_interger(str_remove_all(production_table$data[5], "%"))
  )]
}

Top20_CTR[, CTR := as.numeric(CTR)]

write_xlsx(as.data.frame(Top20_CTR), "H:/Report/Weekly/Report parts/Top20_CTR.xlsx")
