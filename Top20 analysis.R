library(readxl)
library(writexl)
library(data.table)
library(tidyverse)

top20 <- read_excel("C:/Users/fzhang/OneDrive - Travelzoo/Report/Top20 Analysis/AU/Top20/With_Destinations_2017-2019/DB_final.xlsx", sheet = 1)

setDT(top20)


by_quarter <- top20[, .SD[order(Clicks, decreasing = TRUE)]
                    [1:10, .(Year_Quarter, Publication_Week, Category, Des, Parent_Des, Headline, Source, IO, Clicks, Secondary_Clicks)], 
                    by = Year_Quarter]
by_quarter_category <- top20[, .SD[order(Clicks, decreasing = TRUE)]
                    [1:10, .(Year_Quarter, Publication_Week, Category, Des, Parent_Des, Headline, Source, IO, Clicks, Secondary_Clicks)], 
                    by = c("Year_Quarter","Category")]
by_year <- top20[, .SD[order(Clicks, decreasing = TRUE)]
                             [1:10, .(Year_Quarter, Publication_Week, Category, Des, Parent_Des, Headline, Source, IO, Clicks, Secondary_Clicks)], 
                             by = c("Publication Year")]
by_year_category <- top20[, .SD[order(Clicks, decreasing = TRUE)]
                 [1:10, .(Year_Quarter, Publication_Week, Category, Des, Parent_Des, Headline, Source, IO, Clicks, Secondary_Clicks)], 
                 by = c("Publication Year","Category")]
top10 <- top20[, .SD[order(Clicks, decreasing = TRUE)]
                          [1:10, .(Year_Quarter, Publication_Week, Category, Des, Parent_Des, Headline, Source, IO, Clicks, Secondary_Clicks)]]
top10_category <- top20[, .SD[order(Clicks, decreasing = TRUE)]
               [1:10, .(Year_Quarter, Publication_Week, Category, Des, Parent_Des, Headline, Source, IO, Clicks, Secondary_Clicks)],by=Category]
write_xlsx(by_quarter, "C:/Users/fzhang/OneDrive - Travelzoo/Report/Top20 Analysis/AU/Top20/With_Destinations_2017-2019/by_quarter.xlsx")
write_xlsx(by_quarter_category, "C:/Users/fzhang/OneDrive - Travelzoo/Report/Top20 Analysis/AU/Top20/With_Destinations_2017-2019/by_quarter_category.xlsx")
write_xlsx(by_year, "C:/Users/fzhang/OneDrive - Travelzoo/Report/Top20 Analysis/AU/Top20/With_Destinations_2017-2019/by_year.xlsx")
write_xlsx(by_year_category, "C:/Users/fzhang/OneDrive - Travelzoo/Report/Top20 Analysis/AU/Top20/With_Destinations_2017-2019/by_year_category.xlsx")
write_xlsx(top10, "C:/Users/fzhang/OneDrive - Travelzoo/Report/Top20 Analysis/AU/Top20/With_Destinations_2017-2019/top10.xlsx")
write_xlsx(top10_category, "C:/Users/fzhang/OneDrive - Travelzoo/Report/Top20 Analysis/AU/Top20/With_Destinations_2017-2019/top10_category.xlsx")

