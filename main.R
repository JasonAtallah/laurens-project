## R Final Project
## Lauren Fino ACCT 6374
## 11/29/2023


# clear environment and set working directory
rm(list=ls())

working_directory<-'/Users/lauren/final-project'
filename <- "bankofamericafinalproject.xlsx"
file_path <- paste(working_directory, filename, sep = "/")
setwd(working_directory)
print(working_directory)


# Install and Load Libraries
install.packages("stargazer")
install.packages("XML")
install.packages("plyr")
install.packages("quantmod")
install.packages("finreportr")
install.packages("dplyr")
install.packages("ggplot2")
install.packages("writexl")
install.packages("readxl")
install.packages("Hmisc")
install.packages("corrplot")
install.packages("forecast", dependencies = TRUE)

#Load libraries
library("finreportr")
library("stargazer")
library("XML")
library("plyr")
library("quantmod")
library("dplyr")
library("ggplot2")
library("writexl")
library("readxl")
library("Hmisc")
library("corrplot")
library("forecast")


####Download the XLSX files, clean them and consolidate, and then reupload.
###If you pull information externally to clean and consolidate, start from here.

# reading in the data
mydata <- read_excel(file_path)
summary(mydata)

#Correlations
cor(mydata, use="complete.obs", method="pearson") 
#Compute Matrix and Create Plot
M<-cor(mydata, use="complete.obs", method="pearson") 
head(round(M,2))
pdf("Correlation Plot.pdf")
corrplot(M, method="number")
dev.off()
stargazer(M, type = "html", title="Correlation", align=TRUE, out = "Correlation.htm")

#Let's play with ratios
####5 required Ratios - plot them with bivariate plots

##Calculate working capital and Add it to MyData as Variable
currentassets<-mydata$`Assets, Current`
currentliabilities<-mydata$`Liabilities, Current`
workingcapital<-currentassets - currentliabilities
mydata$workingcapital<-workingcapital
summary(mydata)

#Plot WC (1)
pdf("Plot WC.pdf")
plot(mydata$workingcapital ~ mydata$endDate,
     main="WC $ by Period",
     ylab="WC $",
     xlab="Period")
lines(mydata$endDate[order(mydata$endDate)], mydata$workingcapital[order(mydata$endDate)], xlim=range(mydata$endDate), ylim=range(mydata$workingcapital), pch=16)
dev.off()

##Calculate current ratio and Add it to MyData as Variable
currentratio<-currentassets / currentliabilities
mydata$currentratio<-currentratio
summary(mydata)

#Plot Current Ratio (2)
pdf("Plot CurRat.pdf")
plot(mydata$currentratio ~ mydata$endDate,
     main="Current Ratio $ by Period",
     ylab="Current Ratio",
     xlab="Period")
lines(mydata$endDate[order(mydata$endDate)], mydata$currentratio[order(mydata$endDate)], xlim=range(mydata$endDate), ylim=range(mydata$currentratio), pch=16)
dev.off()

##Calculate Asset Turnover and Add it to MyData as Variable
netrevenue <-mydata$`Revenue, Net`
avgtotassets<-mean(mydata$Assets)
assetTO<-netrevenue/avgtotassets
mydata$assetto<-assetTO
summary(mydata)

#Plot Asset Turnover (3)
pdf("Plot AssetTO.pdf")
plot(mydata$assetto ~ mydata$endDate,
     main="Asset Turnover $ by Period",
     ylab="Asset Turnover",
     xlab="Period")
lines(mydata$endDate[order(mydata$endDate)], mydata$assetto[order(mydata$endDate)], xlim=range(mydata$endDate), ylim=range(mydata$assetto), pch=16)
dev.off()

##Calculate Operating Income Margin and Add it to MyData as Variable
netrevenue <-mydata$`Revenue, Net`
oi<-mydata$`Operating Income (Loss)`
oimargin<-oi/netrevenue
mydata$oimargin<-oimargin
summary(mydata)

#Plot OI Margin (4)
pdf("Plot OImargin.pdf")
plot(mydata$oimargin ~ mydata$endDate,
     main="OperIncMargin $ by Period",
     ylab="OperIncMargin",
     xlab="Period")
lines(mydata$endDate[order(mydata$endDate)], mydata$oimargin[order(mydata$endDate)], xlim=range(mydata$endDate), ylim=range(mydata$oimargin), pch=16)
dev.off()

##Calculate DuPont Return on Operating Assets and Add it to MyData as Variable
dupont<-oimargin*assetTO
mydata$dupont<-dupont
summary(mydata)

#Plot DuPont (5)
pdf("Plot Dupont.pdf")
plot(mydata$dupont ~ mydata$endDate,
     main="Dupont by Period",
     ylab="Dupont",
     xlab="Period")
lines(mydata$endDate[order(mydata$endDate)], mydata$dupont[order(mydata$endDate)], xlim=range(mydata$endDate), ylim=range(mydata$dupont), pch=16)
dev.off()

##Calculate Cash Ratio 
cashcashequivalents <-mydata$`Cash, Cash Equivalents`
cashratio<-cashcashequivalents/currentliabilities
mydata$cashratio<-cashratio
summary(mydata)

#Plot Cash Ratio (6)
pdf("Plot CashRatio.pdf")
plot(mydata$cashratio ~ mydata$endDate,
     main="Cash Ratio by Period",
     ylab="Cash Ratio",
     xlab="Period")
lines(mydata$endDate[order(mydata$endDate)], mydata$cashratio[order(mydata$endDate)], xlim=range(mydata$endDate), ylim=range(mydata$cashratio), pch=16)
dev.off()

##Calculate Debt Ratio 
totalliabilities <-mydata$`Liabilities`
totalassets <-mydata$'Assets'
debtratio<-totalliabilities/totalassets
mydata$debtratio<-debtratio
summary(mydata)

#Plot Debt Ratio (7)
pdf("Plot DebtRatio.pdf")
plot(mydata$debtratio ~ mydata$endDate,
     main="Debt Ratio by Period",
     ylab="Debt Ratio",
     xlab="Period")
lines(mydata$endDate[order(mydata$endDate)], mydata$debtratio[order(mydata$endDate)], xlim=range(mydata$endDate), ylim=range(mydata$debtratio), pch=16)
dev.off()

##Calculate ROCE
netincome <-mydata$`Net Income`
preferreddiv <-mydata$'Preferred Dividends'
avgcommonequity <-mydata$'Average Common Equity'
roce<-(netincome-preferreddiv)/avgcommonequity
mydata$roce<-roce
summary(mydata)

#Plot ROCE (8)
pdf("Plot ROCE.pdf")
plot(mydata$roce ~ mydata$endDate,
     main="ROCE by Period",
     ylab="ROCE Ratio",
     xlab="Period")
lines(mydata$endDate[order(mydata$endDate)], mydata$roce[order(mydata$endDate)], xlim=range(mydata$endDate), ylim=range(mydata$roce), pch=16)
dev.off()

##Calculate %EarningsRetained
retainedearnings <-mydata$`Retained Earnings`
netincome <-mydata$`Net Income`
percentearningsretained<-retainedearnings/netincome
mydata$percentearningsretained<-percentearningsretained
summary(mydata)

#Plot %EarningsRetained (9)
pdf("Plot EarningsRetained.pdf")
plot(mydata$percentearningsretained ~ mydata$endDate,
     main="%EarningsRetained by Period",
     ylab="%EarningsRetained",
     xlab="Period")
lines(mydata$endDate[order(mydata$endDate)], mydata$percentearningsretained[order(mydata$endDate)], xlim=range(mydata$endDate), ylim=range(mydata$percentearningsretained), pch=16)
dev.off()

##Calculate Dividend Payout Ratio
totaldividends <-mydata$`Total Dividends`
netincome <-mydata$`Net Income`
dividendpayout<-totaldividends/netincome
mydata$dividendpayout<-dividendpayout
summary(mydata)

#Plot Dividend Payout Ratio (10)
pdf("Plot DividendPayoutRatio.pdf")
plot(mydata$dividendpayout ~ mydata$endDate,
     main="Dividend Payout Ratio by Period",
     ylab="Dividend Payout Ratio",
     xlab="Period")
lines(mydata$endDate[order(mydata$endDate)], mydata$dividendpayout[order(mydata$endDate)], xlim=range(mydata$endDate), ylim=range(mydata$dividendpayout), pch=16)
dev.off()


#Check Dupont with Plot of OI to ATO
pdf("Plot DupontCheck.pdf")
plot(mydata$oimargin ~ mydata$assetto,
     main="OI Margin by Asset TO",
     ylab="OI Margin",
     xlab="Asset TO")
lines(mydata$assetto[order(mydata$assetto)], mydata$oimargin[order(mydata$assetto)], xlim=range(mydata$assetto), ylim=range(mydata$oimargin), pch=16)
dev.off()

##Plot Forecasted Key Ratios with time series
#Create the time series
pdf("Plot timeseries.pdf")
fcstdata<-data.frame(mydata$workingcapital, mydata$currentratio, mydata$assetto, mydata$oimargin, mydata$dupont)
fcstdata.ts<-ts(fcstdata,frequency = 1, start=c(2013,1))
fcstdata.ts
plot(fcstdata.ts)
dev.off()

#Create the forecast (average)
pdf("Plot Forecast.pdf")
fcstdata.forecast<-forecast(fcstdata.ts)
fcstdata.forecast
fcstdata.forecast$method 
plot(fcstdata.forecast)
dev.off()

##Apply the ratio forecast results to your 2015 ratio balances to predict 2016 through 2018. Check these against actual results.

#Stock Results through 2016
pdf("plot stock price.pdf")
Sys.setenv(TZ = "UTC")
getSymbols(c("JPM"), src="yahoo", periodicity="weekly",
           from=as.Date("2006-01-01"), to=as.Date("2021-12-31"))
names(JPM)
nrow(JPM)
head(JPM,3)
tail(JPM,3)
plot(JPM[,6])
dev.off()

##Forecast stock price
#Create the time series
pdf("Plot JPMstock.pdf")
fstock.ts<-ts(JPM,frequency = 52, start=c(2006,1))
fstock.ts
plot(fstock.ts)
dev.off()

#Create the forecast stock plot
pdf("Plot Forecast Stock.pdf")
fstock.forecast<-forecast(fstock.ts)
fstock.forecast
fstock.forecast$method 
plot(fstock.forecast)
dev.off()
print(fstock.forecast)
fstockfcst<-as.data.frame(fstock.forecast)
write_xlsx(x = fstockfcst, path = "fstock.xlsx", col_names = TRUE)
stargazer(fstockfcst, type = "html", title="Forecast Stock Results", align=TRUE, out = "Forecast Stock.htm")

###Export MyData into Excel Workbook
write_xlsx(x = mydata, path = "mydata.xlsx", col_names = TRUE)


#### Test hypotheses with data using OLS regression

#Let's look for a relationship via linear regression
##Let's see whether asset turnover or OI margin is more important to revenue
fit <- lm(mydata$`Revenue, Net` ~ mydata$assetto + mydata$oimargin, data=mydata) # fit is the object that holds all our regression estimates
summary(fit)
stargazer(fit, type = "html", title="Regression Results", align=TRUE, out = "Regression Results.htm")

##Serial Correlation issue - this regression is useless

