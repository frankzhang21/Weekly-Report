library(readxl)
library(tidyverse)
library(stringdist)
library(writexl)
library(data.table)
library(pinyin)
library(phonics)

# cn2017 <- read_excel(file.choose())
# mypy <- pydic(method = c("toneless"), multi = FALSE,
#               only_first_letter = FALSE, dic = c("pinyin"))
# a <- cn2017 %>% 
#   mutate(pinyinaa=py(Headline,sep = "",other_replace = NULL,dic = mypy)) %>% 
#   mutate(pure_pinyin = str_remove_all(pinyinaa,"[^a-zA-Z]")) %>% 
#   mutate(final_pinyin=str_remove_all(pure_pinyin,"CNY")) %>% 
#   mutate(aa_pinyin=str_remove(final_pinyin,"qi")) %>% 
#   mutate(Sound=metaphone(aa_pinyin))


write_xlsx(a,"C:/Users/fzhang/OneDrive - Travelzoo/Report/Top20 Analysis/CN/2017_result111.xlsx")

au2017 <- read_excel(file.choose())
a <- au2017 %>% 
  mutate(pure_headline = str_remove_all(Headline,"[^a-zA-Z]")) %>% 
  mutate(Sound=metaphone(pure_headline))
write_xlsx(a,"C:/Users/fzhang/OneDrive - Travelzoo/Report/Top20 Analysis/AU/2017_result.xlsx")
