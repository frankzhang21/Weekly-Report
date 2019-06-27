library(ggplot2)
a <- read_excel(file.choose())
setDT(a)
a <- a[a$`Price Point`!="NA",]
setDT(a)
a[,c("Price Point"):=.(as.numeric(a$`Price Point`))]
ggplot(data = a)+
  geom_point(aes(x=a$`Price Point`,y=a$Clicks))+
  facet_wrap(vars(a$`Deal Category`),scales = "free")

b <- a[a$`Deal Category`%in% c("Other Activities"),]
j <- ggplot(data = b)+
  geom_point(aes(x=b$`Price Point`,y=b$Clicks))
ggplotly(j)
