
#Importing Libraries 
library(knitr)
library(dplyr) 
library(tidyverse) 
library(ggplot2)
library(openxlsx)
library(lme4)
library(janitor)
library(lattice)
library(readxl)
library(stargazer)
library(Mcomp)
library(smooth)
library(greybox)
library(modelr)


## 1. IMPORTING AND CLEANING ####

#Importing dataset
data <- read_xlsx('data.xlsx')

# Getting rid of NULL and negative (returns) values in Sell.Out.Value column
data_new <- data %>% filter(!is.null(Sell.Out.Value) & Sell.Out.Value != 'NULL' 
                            & Sell.Out.Value > 0)

# Selecting columns we need
data_new <- data_new %>% select(1,2,3,5,6)

# Changing Column Names 
names(data_new) <- c('Store', 'Date', 'Quantity', 'Category', 'Value')

# Formatting Date Column 
data_new$Date <- as.Date(data_new$Date , format = "%d.%m.%Y")

# Adding 'Year' Column 
data_new$Year <- format(data_new$Date, format = "%Y")

# Filtering out 2020 because unpredictable year 
data_new <- data_new %>% filter(Year %in% c(2018, 2019))

# Formatting Value column as numeric 
data_new$Value <- as.numeric(sub(",", ".", data_new$Value, fixed = TRUE))

# Formatting Quantity column as numeric 
data_new$Quantity <- as.numeric(data_new$Quantity)

# Filter data to only contain Espresso
espresso <- data_new %>% filter(Category == '0418 Espresso')

# Removing number code from category names 
espresso$Category <- str_sub(espresso$Category, 6)

# Getting unit price
espresso <- espresso %>% mutate(Unit_price = Value/Quantity)

# Filtering price range to represent an espresso machine model of the most 
# popular price range
espresso <- espresso %>% filter(between(Unit_price, 200, 400))

# We will aggregate data per week
espresso$Week <-as.numeric(strftime(espresso$Date, format = '%V'))
espresso$Month <-as.numeric(strftime(espresso$Date, format = '%m'))

# Adding this otherwise the end of year is labelled as the first week of year
espresso<- espresso %>% mutate(Year_Adj = ifelse(Week ==1 & Month == 12, 
                                                 as.numeric(Year) + 1, 
                                                 as.numeric(Year)))

espresso<- espresso %>% mutate(Week_Adj = ifelse(Year_Adj == 2019, Week + 52, 
                                                 ifelse(Year_Adj == 2020, 
                                                        Week + 104, Week)))


# Aggregating per week
espresso_agg <- espresso %>% 
  select(1, 3, 7, 11) %>% 
  group_by(Store, Week_Adj) %>% 
  summarise(Units_Sold = sum(Quantity), Avg_Price = mean(Unit_price))


# Importing Salesman Presence Data 
presence_2018  <- read_csv('2018.csv') %>% select(-1) 
presence_2019  <- read_csv('2019.csv') %>% select(-1)

# Changing layout of dataframe to have a date column
presence_2018 <- pivot_longer(presence_2018, cols = !Store_ID, names_to = 'Date')
presence_2019 <- pivot_longer(presence_2019, cols = !Store_ID, names_to = 'Date')

# Adding both in one data frame 
presence <- rbind(presence_2018, presence_2019)

names(presence)[3] <- 'Salesperson'
names(presence)[1] <- 'Store'

presence$Store <- as.character(presence$Store)

# Aggregating data on a weekly basis as before
presence$Week <-as.numeric(strftime( presence$Date, format = '%V'))
presence$Month <-as.numeric(strftime(presence$Date, format = '%m'))
presence$Year <-as.numeric(strftime(presence$Date, format = '%Y'))

# Adding this otherwise the end of year is labelled as the first week of year
presence <- presence %>% mutate(Year_Adj = ifelse(Week ==1 & Month == 12, 
                                                  Year + 1, Year))

presence <- presence %>% mutate(Week_Adj = ifelse(Year_Adj == 2019, Week + 52, 
                                                  ifelse(Year_Adj == 2020, 
                                                         Week + 104, Week)))

# Aggregating per week
presence_agg <- presence %>% 
  select(1, 3, 8) %>% 
  group_by(Store, Week_Adj) %>% 
  summarise(Days = sum(Salesperson))

# Merging the two dataframes 
all_data <- inner_join(espresso_agg, presence_agg, by = c('Store', 'Week_Adj'), 
                       copy = FALSE)
all_data <- as.data.frame(all_data)

# Replacing Store SAP numbers with Store1, Store2, ... 
stores <- levels(factor(all_data$Store))

new_names <- c()

for (i in 1:length(stores)){
  new_names[i] <- paste('Store', i)
}

all_data$Store <- factor(all_data$Store)
levels(all_data$Store) <- new_names

names(all_data)[2] <- 'Week'

#renaming all_data as df 
df <- all_data



## 2. EXPLORATION ####

#Aggregate weekly sales across stores 
df_agg <- df %>% 
  select(2,3) %>% 
  group_by(Week) %>% 
  summarize(Units = sum(Units_Sold))

df_agg %>% 
  ggplot(aes(x=Week, y =Units)) + geom_line(col='blue') + theme_bw() + 
  ggtitle('Total Weekly Sales')


#Assess need for multilevel model 

# Null multilevel model
null <- lmer(Units_Sold ~ (1 | Store), data = df) 
# Random Effects plot
res <- ranef(null, condVar = TRUE) # calculating residuals from random effects of model
store <- as.factor(rownames(res[[1]])) # extracting store names
table <- cbind(store, res[[1]]) # adding stores and residuals in one table
colnames(table) <- c("store","residuals")
table <- table[order(table$residuals), ] # order by residuals
table <- cbind(table, c(1:dim(table)[1]))
colnames(table)[3] <- "rank" # add rank
table <- table[order(table$store), ] # order 

#plotting
plot(table$rank, table$residuals, type = "n", xlab = "stores (ranked)", 
     ylab = "Residuals")

points(table$rank, table$residuals, col = "blue")
abline(h = 0, col = "red")
title('Store deviation from Units_Sold mean')


#Exponential smoothing 
exp_smoothing <- ses(df_agg$Units, alpha=0.1, initial="simple")

plot(df_agg$Units, type = 'l', ylab='Units Sold', xlab='Week')
lines(exp_smoothing$fitted, col='red')
title('Exponentially Smoothed Weekly Sales across all Stores')

# Adding smoothed category baseline volume to data frame
fval <- data.frame(Week = seq(1:105), Index = exp_smoothing$fitted)
df <- left_join(df, fval, by = 'Week')


## 3. FEATURE ENGINEERING ####

# Adding Dummy Variables 
df <- df %>% mutate(XMAS = ifelse(Week == 52 | Week == 104, 1, 0))
df <- df %>% mutate(BF = ifelse(Week == 47 | Week == 100, 1, 0))
df <- df %>% mutate(FW = ifelse(Week == 1  | Week == 105, 1, 0))

# Creating columns for the log values 
df$Units_Sold_log <- log(df$Units_Sold)
df$Avg_Price_log <- log(df$Avg_Price)


## 4. MODELS ####

# Identify stores with the highest number of observations 
table(df$Store)[table(df$Store) == max(table(df$Store))]

# Filtering dataset to only include data for store 48
store_48 <- df %>% filter(Store == 'Store 48')

# Creating Models 
model_48 <- lm(Units_Sold_log ~ Avg_Price_log + Days + Index + XMAS + BF + FW, 
               data = store_48)


# Adding normalized column 
df <- df %>% mutate(Index_norm = Index/max(Index))

# Models
model1 <- lmer(Units_Sold_log ~ Avg_Price_log + Days + Index_norm + XMAS + BF + 
                 FW + (1 | Store), data=df)
model2 <- lmer(Units_Sold_log ~ Avg_Price_log + Days + Index_norm + XMAS + BF + 
                 FW + (Days| Store), data=df)

stargazer(model1, model2, type = 'html', title = 'Regression Results', 
          keep.stat="n")

print(paste('Model 1 =', round(rsquare(model1, df),2)))
print(paste('Model 2 =', round(rsquare(model2, df),2)))

print(paste('Likelihood Ratio Test Value = ', round(-2*(logLik(model1)-
                                                          logLik(model2)),2)))

print(paste('5% level of Chi Squared Distribution on 8 Degrees of Freedom =', 
            round(qchisq(.95, df=8),2)))


#running same model without days predictor
model_base <- lmer(Units_Sold_log ~ Avg_Price_log + Index_norm + XMAS + BF + FW 
                   + (1 | Store), data=df)
baseline_sales = round(fitted(model_base),0)
baseline = sum(baseline_sales)
print(paste('Baseline Sales =', baseline))

print(paste('Incremental Sales = Actual Sales - Baseline Sales = ', 
            sum(df$Units_Sold), '-', round(baseline,2), '=', 
            round(sum(df$Units_Sold) - baseline,2)))


#Estimate profitability
prof <- data.frame(Incremental_Sales = df$Units_Sold - baseline_sales, 
                   Avg_Price = df$Avg_Price, Days = df$Days)

prof <- prof %>% mutate(Costs = Days*255, Incremental_Revenue = 
                          Incremental_Sales*Avg_Price)

print(paste('Profitability = ', round(sum(prof$Incremental_Revenue) - 
                                        sum(prof$Costs),2)))





