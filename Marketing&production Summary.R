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

# Ending date for net members
net_member_end_date <- remDr$findElement(using = "xpath", '//*[@id="txtSSDateEnding"]') # Assign to var
net_member_end_date$clearElement() # Clear date input form
report_week_monday <- Sys.Date() - 3
text_format_mondy <- change_date(report_week_monday)
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





# Top20 CTR Report --------------------------------------------------------


# Data for production -----------------------------------------------------

remDr$navigate("https://intranet.travelzoo.com/office/top20/eu/?tzlocale=10")

## Manually adjust week number and show stats

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

write_xlsx(as.data.frame(Top20_CTR), "C:/Users/fzhang/OneDrive - Travelzoo/Report/Weekly/Report parts/Top20_CTR.xlsx")


# Score Card --------------------------------------------------------------

remDr$navigate("https://intranet.travelzoo.com/common/marketing/QualityScoreGrid.aspx")

this_thursday <- Sys.Date()


sub_from <- remDr$findElement(using = "xpath",'//*[@id="txtSubscribedFrom"]')
sub_from$clearElement()
sub_from_date <- this_thursday-22

sub_from_date <- change_date(sub_from_date)
sub_from$sendKeysToElement(list(sub_from_date))

sub_to <- remDr$findElement(using = "xpath",'//*[@id="txtSubscribedTo"]')
sub_to$clearElement()
sub_to$sendKeysToElement(list(change_date(this_thursday-16)))

click_from <- remDr$findElement(using = "xpath",'//*[@id="txtClickedFrom"]')
click_from$clearElement()
click_from$sendKeysToElement(list(change_date(this_thursday-15)))

click_to <- remDr$findElement(using = "xpath",'//*[@id="txtClickedTo"]')
click_to$clearElement()
click_to$sendKeysToElement(list(change_date(this_thursday-2)))

score_card_smt <- remDr$findElement(using = "xpath",'//*[@id="Submit"]')
score_card_smt$clickElement()

score_card_country <- c("JP","AU","CN","HK")

for (k in score_card_country) {
  country_filter <- remDr$findElement(using = "xpath",'//*[@id="uwcHeader_ddlTZLocale"]')
  country_filter$sendKeysToElement(list(k))
  score_card_smt <- remDr$findElement(using = "xpath",'//*[@id="Submit"]')
  score_card_smt$clickElement()
  Sys.sleep(1.5)
  
  Response <- read_html(remDr$getPageSource()[[1]])
  score_table <- html_table(Response, fill = TRUE)[[7]]
  setDT(score_table)
  index <- score_table[,X6]
  final_table[Country == k, c("Score_card"):=as.numeric(index[length(index)-1])]
  
}
final_table[,c("CTR","NEWSOPENRATE","PRIMARY","SECONDARY"):=.(Clicks/Delivery,Open_Count/Delivery,Primary/Delivery,Secondary/Primary)]
setcolorder(final_table,c("Country","Net_Member","New_sub","Un_sub","Score_card","Delivery","Clicks","CTR","Open_Count","NEWSOPENRATE",
                          "Primary","PRIMARY","Secondary","SECONDARY"))

reshape_version <- melt(final_table, id.vars = "Country") %>%
  spread(Country, value) %>%
  setcolorder(c("variable", "JP", "AU", "CN", "HK", "TW", "AE"))

var_name <- c("Net Members","New_sub","Un_sub","Score_card","Top 20 Delivery","Newsletter Clicks",     
              "CTR %","Newsletter Open Count","Newsletter Open Rate %","Pri Clicks","%","Sec Clicks","%")
reshape_version[,variable:=var_name]

write_xlsx(as.data.frame(reshape_version), "H:/Report/Weekly/Report parts/Market_Production.xlsx")



