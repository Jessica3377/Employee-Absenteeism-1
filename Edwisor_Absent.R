#remove all the objects stored
rm(list=ls())

#set current working directory
setwd("F:/Absenteeism")

library(xlsx)      # Super simple excel reader
library(mice)       # missing values imputation
library(naniar)     # visualize missing values
library(dplyr)
library(corrplot)
library(ggplot2)
library(tidyverse)
library(randomForest)
library(caret)
library(data.table)
library(Boruta)
library(rpart)

## Read the data
df_absent <- read.xlsx2("Absenteeism_at_work_Project.xls", sheetIndex = 1, header = TRUE, colClasses = NA)

################################# exploratory data analysis####################################### 

# it is found that  Month.of.absence ,  there are 13 months present in data, hence to replace the false data by NA
df_absent = transform(df_absent, Month.of.absence = 
                                  ifelse(Month.of.absence == 0, NA, Month.of.absence ))

str(df_absent)

#changing the contious variables to categorical variables for the ease of performance
df_absent$Reason.for.absence = as.factor(df_absent$Reason.for.absence)
df_absent$Month.of.absence = as.factor(df_absent$Month.of.absence)
df_absent$Day.of.the.week = as.factor(df_absent$Day.of.the.week)
df_absent$Seasons = as.factor(df_absent$Seasons)
df_absent$Service.time = as.factor(df_absent$Service.time)
df_absent$Hit.target = as.factor(df_absent$Hit.target)
df_absent$Disciplinary.failure = as.factor(df_absent$Disciplinary.failure)
df_absent$Education = as.factor(df_absent$Education)
df_absent$Son = as.factor(df_absent$Son)
df_absent$Social.drinker = as.factor(df_absent$Social.drinker)
df_absent$Social.smoker = as.factor(df_absent$Social.smoker)
df_absent$Pet = as.factor(df_absent$Pet)
df_absent$Work.load.Average.day  = as.numeric(df_absent$Work.load.Average.day )

################################## Outlier analysis ###################################

outlierKD <- function(dt, var) {
  var_name <- eval(substitute(var),eval(dt))
  na1 <- sum(is.na(var_name))
  m1 <- mean(var_name, na.rm = T)
  par(mfrow=c(2, 2), oma=c(0,0,3,0))
  boxplot(var_name, main="With outliers")
  hist(var_name, main="With outliers", xlab=NA, ylab=NA)
  outlier <- boxplot.stats(var_name)$out
  mo <- mean(outlier)
  var_name <- ifelse(var_name %in% outlier, NA, var_name)
  boxplot(var_name, main="Without outliers")
  hist(var_name, main="Without outliers", xlab=NA, ylab=NA)
  title("Outlier Check", outer=TRUE)
  na2 <- sum(is.na(var_name))
  cat("Outliers identified:", na2 - na1, "n")
  cat("Propotion (%) of outliers:", round((na2 - na1) / sum(!is.na(var_name))*100, 1), "n")
  cat("Mean of the outliers:", round(mo, 2), "n")
  m2 <- mean(var_name, na.rm = T)
  cat("Mean without removing outliers:", round(m1, 2), "n")
  cat("Mean if we remove outliers:", round(m2, 2), "n")
  response <- readline(prompt="Do you want to remove outliers and to replace with NA? [yes/no]: ")
  if(response == "y" | response == "yes"){
    dt[as.character(substitute(var))] <- invisible(var_name)
    assign(as.character(as.list(match.call())$dt), dt, envir = .GlobalEnv)
    cat("Outliers successfully removed", "n")
    return(invisible(dt))
  } else{
    cat("Nothing changed", "n")
    return(invisible(var_name))
  }
}


outlierKD(df_absent,Absenteeism.time.in.hours)# outliers detected and replaced by NA
outlierKD(df_absent,Transportation.expense) #no outliers
outlierKD(df_absent,Distance.from.Residence.to.Work) #no outliers
outlierKD(df_absent,Service.time) #no outliers
outlierKD(df_absent,Age) #no outliers
outlierKD(df_absent,Work.load.Average.day.) # 1 found and replaced with NA
outlierKD(df_absent,Hit.target) # 1 found and replaced with NA
#outlierKD(df_absent,Son) # no outliers
#outlierKD(df_absent,Pet) # no outliers
outlierKD(df_absent,Weight) # no outliers
outlierKD(df_absent,Height) # no outliers
outlierKD(df_absent,Body.mass.index) #no outliers


################################## Missing value analysis ##################################

missing_val = data.frame(apply(df_absent,2,function(x){sum(is.na(x))}))
missing_val$Columns = row.names(missing_val)
names(missing_val)[1] =  "Missing_percentage"
missing_val$Missing_percentage = (missing_val$Missing_percentage/nrow(df_absent)) * 100
missing_val = missing_val[order(-missing_val$Missing_percentage),]
row.names(missing_val) = NULL
missing_val = missing_val[,c(2,1)]
write.csv(missing_val, "Miising_perc.csv", row.names = F)

#ggplot analysis
ggplot(data = missing_val[1:3,], aes(x=reorder(Columns, -Missing_percentage),y = Missing_percentage))+
  geom_bar(stat = "identity",fill = "grey")+xlab("Parameter")+
  ggtitle("Missing data percentage (Train)") + theme_bw()

library(ggplot2)
#actual value =30
#df_absent[1,20]
#df_absent[1,20]= NA
# kNN Imputation=29.84314
#after various calculations, it is found that knn imputation method suits the best for the data. hence here we are applying knn imputation

library(DMwR)
df_absent = knnImputation(df_absent, k = 3)
sum(is.na(df_absent))


################################### BoxPlots - Distribution and Outlier Check ##################################

numeric_index = sapply(df_absent,is.numeric) #selecting only numeric
numeric_data = df_absent[,numeric_index]

cnames = colnames(numeric_data)
library(ggplot2)
for (i in 1:length(cnames))
{
  assign(paste0("gn",i), ggplot(aes_string(y = (cnames[i]), x = "responded"), data = subset(df_absent))+ 
           stat_boxplot(geom = "errorbar", width = 0.5) +
           geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,
                        outlier.size=1, notch=FALSE) +
           theme(legend.position="bottom")+
           labs(y=cnames[i],x="responded")+
           ggtitle(paste("Box plot of responded for",cnames[i])))
}

################################## feature selection ##################################
library(corrgram)
## Correlation Plot - to check multicolinearity between continous variables
corrgram(df_absent[,numeric_index], order = F,
         upper.panel=panel.pie, text.panel=panel.txt, main = "Correlation Plot")
df_absent$Absenteeism.time.in.hours = as.factor(df_absent$Absenteeism.time.in.hours)

## Chi-squared Test of Independence-to check the multicolinearity between categorical variables
factor_index = sapply(df_absent,is.factor)
factor_data = df_absent[,factor_index]

for (i in 1:12)
{
  print(names(factor_data)[i])
  print(chisq.test(table(factor_data$Absenteeism.time.in.hours,factor_data[,i])))
}

df_absent$Absenteeism.time.in.hours = as.numeric(df_absent$Absenteeism.time.in.hours)

################################### Feature reduction ##################################
## Dimension Reduction
df_absent = subset(df_absent, 
                             select = -c(Weight,Hit.target,Education,Social.smoker,Pet))
#Feature Scaling
#Normality check
qqnorm(df_absent$Absenteeism.time.in.hours )
hist(df_absent$Absenteeism.time.in.hours )
str(df_absent)
#Normalisation
cnames = c("ID","Transportation.expense","Distance.from.Residence.to.Work","Height","Age","Work.load.Average.day","Body.mass.index",
           "Absenteeism.time.in.hours")

for(i in cnames){
  print(i)
  df_absent[,i] = (df_absent[,i] - min(df_absent[,i]))/
    (max(df_absent[,i] - min(df_absent[,i])))
}
########################## Univariate Distribution and Analysis ###############################

# function for univariate analysis for continous variables
#       function inpus:
#       1. dataset - input dataset
#       2. variable - variable for univariate analysis
#       3. variableName - variable title in string 
#       example.   univariate_analysis(df_absent,Absenteeism.time.in.hours,
#                                                   "Absenteeism.time.in.hours")
univariate_analysis <- function(dataset, variable,variableName){
  
  var_name = eval(substitute(variable), eval(dataset))
  
  if(is.numeric(var_name)){
    
    print(summary(var_name))
    ggplot(df_absent, aes(var_name)) + 
      geom_histogram(aes(y=..density..,binwidth=.5,colour="black", fill="white"))+ 
      geom_density(alpha=.2, fill="#FF6666")+
      labs(x = variableName, y = "count") + 
      ggtitle(paste("count of ",variableName)) +
      theme(legend.position = "null")
    
  }else{
    print("This is categorical variable.")
    
  }
  
}


# function for univariate analysis for categorical variables
#       function inpus:
#       1. dataset - input dataset
#       2. variable - variable for univariate analysis
#       3. variableName - variable title in string 
#       example.   univariate_analysis(df_absent,ID,
#                                                   "ID")
univariate_catogrical <- function(dataset,variable, variableName){
  variable <- enquo(variable)
  
  percentage <- dataset %>%
    select(!!variable) %>%
    group_by(!!variable) %>%
    summarise(n = n()) %>%
    mutate(percantage = (n / sum(n)) * 100)
  print(percentage)
  
  dataset %>%
    count(!!variable) %>%
    ggplot(mapping = aes_(x = rlang::quo_expr(variable), 
                          y = quote(n), fill = rlang::quo_expr(variable))) +
    geom_bar(stat = 'identity',
             colour = 'white') +
    labs(x = variableName, y = "count") + 
    ggtitle(paste("count of ",variableName)) +
    theme(legend.position = "bottom") -> p
  plot(p)
  
}
################################### Univariate analysis of continous variables ################################
univariate_analysis(df_absent,Absenteeism.time.in.hours,"Absenteeism.time.in.hours")

univariate_analysis(df_absent,Transportation.expense,"Transportation.expense")

univariate_analysis(df_absent,Distance.from.Residence.to.Work,
                    "Distance.from.Residence.to.Work")

univariate_analysis(df_absent,Service.time,"Service.time")

univariate_analysis(df_absent,Age,"Age")

univariate_analysis(df_absent,Work.load.Average.day ,"Work.load.Average.day ")

#univariate_analysis(df_absent,Hit.target ,"Hit.target")

univariate_analysis(df_absent,Son ,"Son")

#univariate_analysis(df_absent,Pet ,"Pet")

#univariate_analysis(df_absent,Weight ,"Weight")

univariate_analysis(df_absent,Height ,"Height")

univariate_analysis(df_absent,Body.mass.index ,"Body.mass.index")


################################### univariate analysis of categorical variables ##################################

univariate_catogrical(df_absent,ID,"Id")
univariate_catogrical(df_absent,Reason.for.absence,"Reason.for.absence")
univariate_catogrical(df_absent,Month.of.absence,"Month.of.absence")
univariate_catogrical(df_absent,Day.of.the.week,"Day.of.the.week")
univariate_catogrical(df_absent,Seasons,"Seasons")
univariate_catogrical(df_absent,Disciplinary.failure,"Disciplinary.failure")
univariate_catogrical(df_absent,Education,"Education")
univariate_catogrical(df_absent,Social.drinker,"Social.drinker")
univariate_catogrical(df_absent,Social.smoker,"Social.smoker")

################################### Sampling ##################################
##Systematic sampling
#Function to generate Kth index
sys.sample = function(N,n)
{
  k = ceiling(N/n)
  r = sample(1:k, 1)
  sys.samp = seq(r, r + k*(n-1), k)
}
lis = sys.sample(740, 300) #select the repective rows
# #Create index variable in the data
df_absent$index = 1:740
# #Extract subset from whole data
systematic_data = df_absent[which(df_absent$index %in% lis),]

################################### Model Development ##################################
#Clean the environment
library(DataCombine)

rmExcept("df_absent")
#Divide data into train and test using stratified sampling method
set.seed(1234)
df_absent$description = NULL
library(caret)
train.index = createDataPartition(df_absent$Absenteeism.time.in.hours, p = .80, list = FALSE)
train = df_absent[ train.index,]
test  = df_absent[-train.index,]

#load libraries
library(rpart)
#decision tree analysis
#rpart for regression
fit = rpart(Absenteeism.time.in.hours ~ ., data = train, method = "anova")
#Predict for new test cases
predictions_DT = predict(fit, test[,-16])
#MAPE
#calculate MAPE
MAPE = function(y, yhat){
  mean(abs((y - yhat)/y))*100
}

MAPE(test[,16], predictions_DT)

#Random Forest
library(randomForest)
RF_model = randomForest(Absenteeism.time.in.hours ~ ., train, importance = TRUE, ntree = 1000)
#Extract rules fromn random forest
#transform rf object to an inTrees' format
library(RRF)
library(inTrees)
treeList <- RF2List(RF_model)
#Extract rules
exec = extractRules(treeList, train[,-16])  # R-executable conditions
ruleExec <- extractRules(treeList,train[,-16],digits=4)

#Make rules more readable:
readableRules = presentRules(exec, colnames(train))
readableRules[1:2,]
#Get rule metrics
ruleMetric = getRuleMetric(exec, train[,-16], train$Absenteeism.time.in.hours)  # get rule metrics
#Predict test data using random forest model
RF_Predictions = predict(RF_model, test[,-16])
#rmse calculation
#install.packages("Metrics")
library(Metrics)
rmse(test$Absenteeism.time.in.hours, RF_Predictions)
#rmse value for random forest is 0.2065729
rmse(test$Absenteeism.time.in.hours, predictions_DT)
#rmse value for decision tree is 0.222542
write.csv(RF_Predictions, file = 'RF_Predictions.csv')
write.csv(predictions_DT, file = 'predictions_DT.csv')