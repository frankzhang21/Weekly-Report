ma <- read_excel(file.choose())

setDT(ma)

ma[,link := str_c("https://intranet.travelzoo.com/office/accounts/io/Index.aspx?iodraftid=",MA,"&mode=view")]

for (i in 1:77) {
  

remDr$navigate(ma$link[i])

address_a <- remDr$findElement(using = "xpath", '//*[@id="IOForm"]/table/tbody/tr[4]/td[2]')
ma[i,text_c := address_a$getElementText()[[1]]]
}

write_xlsx(ma,"C:/Users/fzhang/OneDrive - Travelzoo/b.xlsx")
