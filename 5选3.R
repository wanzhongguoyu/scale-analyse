# 全模型 5选3_中介调节_process
rm(list=ls())
# install.packages('readxl')

library("readxl")
library(gtools)
library(car)
library(tcltk)  
library('writexl')
library(do)
source('process.R')   #调用process 脚本


# 读取excel
all_chose <- permutations(5, 3, v=1:5)  # 枚举所有可能性
all_data <- read_excel("整理数据517.xlsx", sheet=1)

ismed=list()  # 中介
ismod=list()  # 调节

pb <- tkProgressBar(title="进度",label="已完成 %", min=0, max=100, initial = 0, width = 300)  # 进度条

for (i in 1:nrow(all_chose)) {
  use_data <- all_data[, all_chose[i, ] ] # 提取
  colnames(use_data) <- c("x", "m", 'y') # 改名
  
  
  # 中介
  result_med <-
    process(
      data = use_data,
      y = "y",x = "x",m = "m",
      normal = 1,model = 4,
      effsize = 1,total = 1,stand = 1,
      boot = 5000,modelbt = 1,
      save = 2,plot = 3,progress = 0)
  
  
  if (result_med[3, 4]<0.05 # X到m 显著
      && result_med[8, 4]<0.05 # m到y 显著
      && result_med[15, 4] <0.05 ) # xm到y 显著
  {ismed[i] <-1}  else  # 1 有中介  
  {ismed[i] <-0}   # 2 无中介  
  
  
  
  # 调节
  result_mod <- process(
    data = use_data,
    y = "y",x = "x", w = "m",
    normal = 1,model = 1,
    effsize = 1,total = 1,
    boot = 5000,modelbt = 1,
    save = 2,plot = 0,progress = 1)
  
  if (result_mod[3, 4]<0.05 # X到y 显著
      && result_mod[4, 4]<0.05 # w到y 显著
      &&result_mod[5, 4]<0.05  )  # x*w 显著
  {ismod[i] <-1}  else  # 1 有调节 
  {ismod[i] <-0}   # 2 无调节 
  
  
  
  info<- sprintf("已完成 %d%%", round(i*100/nrow(all_chose)))  
  setTkProgressBar(pb, value = i*100/nrow(all_chose), title = sprintf("进度 (%s)",info),label = info)  
  
}

close(pb)    #关闭进度条

all_chose<-data.frame(all_chose)
Replace(data =all_chose,from =1,to = 'EI')
Replace(data =all_chose,from =2,to = 'ASLEC')
Replace(data =all_chose,from =3,to = 'PSQI')
Replace(data =all_chose,from =4,to = 'ERQ')
Replace(data =all_chose,from =5,to = 'DEQ')
colnames(all_chose) <- c("x", "m", 'y') 



ismed<-data.frame(unlist(ismed))
ismod<-data.frame(unlist(ismod))

ismed<-ismed$unlist.ismed.
ismod<-ismod$unlist.ismod.

all_chose$ismed<-ismed
all_chose$ismod<-ismod

write_xlsx(x = all_chose,file = "result_mod_med.xlsx")
