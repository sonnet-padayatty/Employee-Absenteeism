rm(list=ls(all=T))
setwd("E:/Project 2")
options(warn = -1)
#Load Libraries
x = c("ggplot2", "corrgram", "DMwR", "caret", "randomForest", "unbalanced", "C50", "dummies", "e1071", "Information",
      "MASS", "rpart", "gbm", "ROSE", 'sampling', 'DataCombine', 'inTrees')

#install.packages(x)
lapply(x, require, character.only = TRUE)
rm(x)

## Read the data
library(xlsx)
absent = read.xlsx("Absenteeism_at_work_Project.xls",sheetIndex = 1, header = T)
View(absent)
sum(is.na(absent))
name=subset(absent,select=-c(ID,Day.of.the.week,Seasons,Body.mass.index))
a=colnames(name)
###################################Missing Values Analysis######################
for(i in a){
absent[,i][is.na(absent[,i])] = median(absent[,i], na.rm = T)
}
absent$Body.mass.index[is.na(absent$Body.mass.index)]=mean(absent$Body.mass.index,na.rm = T)
str(absent)

absent$Reason.for.absence=as.factor(absent$Reason.for.absence)
absent$Seasons=as.factor(absent$Seasons)
absent$Social.drinker=as.factor(absent$Social.drinker)
absent$Social.smoker=as.factor(absent$Social.smoker)
absent$Disciplinary.failure=as.factor(absent$Disciplinary.failure)
absent$Day.of.the.week=as.factor(absent$Day.of.the.week)
absent$Education=as.factor(absent$Education)
absent$ID=as.factor(absent$ID)
absent$Month.of.absence=as.factor(absent$Month.of.absence)

###############Outlier Analysis###########################

numeric_index = sapply(absent,is.numeric) #selecting only numeric

numeric_data = absent[,numeric_index]

cnames = colnames(numeric_data)

for (i in 1:length(cnames))
{
  assign(paste0("gn",i), ggplot(aes_string(y = (cnames[i]), x = "Absenteeism.time.in.hours"), data = subset(absent))+ 
           stat_boxplot(geom = "errorbar", width = 0.5) +
           geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,
                        outlier.size=1, notch=FALSE) +
           theme(legend.position="bottom")+
           labs(y=cnames[i],x="Absenteeism.time.in.hours")+
           ggtitle(paste("Box plot of Absenteeism.time.in.hours for",cnames[i])))
}

## Plotting plots together
gridExtra::grid.arrange(gn1,gn2,gn3,gn4,ncol=4)
gridExtra::grid.arrange(gn5,gn6,gn7,gn8,ncol=4)
gridExtra::grid.arrange(gn9,gn10,gn11,gn12,ncol=4)

#### Removing the outliers and imputing the missing values####

numeric=cnames
numeric=numeric[-11:-12]
for(i in numeric){
  val = absent[,i][absent[,i] %in% boxplot.stats(absent[,i])$out]
  print(length(val))
  absent[,i][absent[,i] %in% val] = NA
  absent[,i][is.na(absent[,i])] = median(absent[,i], na.rm = T)
}
sum(is.na(absent))
val = absent$Body.mass.index[absent$Body.mass.index %in% boxplot.stats(absent$Body.mass.index)$out]
absent$Body.mass.index[absent$Body.mass.index %in% val] = NA
## Correlation Plot 
corrgram(absent[,numeric_index], order = F,
         upper.panel=panel.pie, text.panel=panel.txt, main = "Correlation Plot")

## Anova test
factor_index = sapply(absent,is.factor)
factor_data = absent[,factor_index]
View(factor_data)
colnames(factor_data)
res.aov = aov(formula=Absenteeism.time.in.hours ~ ID, data = absent)
summary(res.aov)
res.aov = aov(formula=Absenteeism.time.in.hours ~ Reason.for.absence, data = absent)
summary(res.aov)
res.aov = aov(formula=Absenteeism.time.in.hours ~ Month.of.absence, data = absent)
summary(res.aov)
res.aov = aov(formula=Absenteeism.time.in.hours ~ Day.of.the.week, data = absent)
summary(res.aov)
res.aov = aov(formula=Absenteeism.time.in.hours ~ Seasons, data = absent)
summary(res.aov)
res.aov = aov(formula=Absenteeism.time.in.hours ~ Disciplinary.failure, data = absent)
summary(res.aov)
res.aov = aov(formula=Absenteeism.time.in.hours ~ Education, data = absent)
summary(res.aov)
res.aov = aov(formula=Absenteeism.time.in.hours ~ Social.drinker, data = absent)
summary(res.aov)
res.aov = aov(formula=Absenteeism.time.in.hours ~ Social.smoker, data = absent)
summary(res.aov)

### Dimensionality reduction ###

abs_final=subset(absent,select=-c(ID,Pet,Age,Disciplinary.failure,Month.of.absence,Seasons,Education,Social.drinker,Social.smoker,Body.mass.index))

#Normalisation
cn=c("Transportation.expense" , "Distance.from.Residence.to.Work",
         "Service.time",  "Hit.target"  ,                   
         "Son" ,"Height" , "Weight","Absenteeism.time.in.hours")

#Normality check
for(i in cn){
  hist(abs_final[,i],main = i ,col = "blue",border = "red",xlab="Range")
}

for(i in cn){
  print(i)
  abs_final[,i] = (abs_final[,i] - min(abs_final[,i]))/
    (max(abs_final[,i] - min(abs_final[,i])))
}

### Splitting the data set into train and test 
set.seed(1234)
train_ind = sample(1:nrow(abs_final),0.8*nrow(abs_final))                                                                                      
train = abs_final[train_ind,]
test = abs_final[-train_ind,]

# Decision trees

fit = rpart(Absenteeism.time.in.hours~.,data = train, method = 'anova')
summary(fit)
prediction_dt = predict(fit,test[,-11])
regr.eval(test[,11],prediction_dt,stats = "rmse")
# error rate = 11% #accuracy 89%

# Random forest

rf_mod = randomForest(Absenteeism.time.in.hours~.,train,importance = TRUE,ntree = 500)
rf_pred = predict(rf_mod,test[,-11])
regr.eval(test[,11],rf_pred,stats = "rmse")
#error rate = 10 #accuracy 90

library(rpart)
library(MASS)
#Linear Regression
#check multicollearity
library(usdm)
#droplevels(abs_final$Reason.for.absence)

#run regression model
lm_model = lm(Absenteeism.time.in.hours ~., data = train)

#Summary of the model
summary(lm_model)

#Predict
predictions_LR = predict(lm_model, test[,-11])
regr.eval(test[,11],predictions_LR,stats = "rmse")
#error=10%, accuracy=90%
############Work Loss###############

dim(absent)
loss_cal=subset(absent,select=c("Month.of.absence","Absenteeism.time.in.hours",
                                "Work.load.Average.day.","Service.time"))
dim(loss_cal)
table(absent$Month.of.absence)
loss_cal$loss=with(loss_cal,(loss_cal$Work.load.Average.day.*loss_cal$Absenteeism.time.in.hours)/loss_cal$Service.time)
View(loss_cal)
month=aggregate(loss_cal$loss,by=list(loss_cal$Month.of.absence),sum)[2:13,]
colnames(month)[1]="month"
colnames(month)[2]="loss"
row.names(month)=NULL
View(month)
write.csv(month,"LOSSpermonth.csv",row.names = F)
