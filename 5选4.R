# 全模型 5选4_链式 多重_process
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
all_chose <- permutations(5, 4, v=1:5)
all_data <- read_excel("整理数据517.xlsx", sheet=4)

isser_med=list()  # 链式
ispar_med=list() # 平行

pb <- tkProgressBar(title="进度",label="已完成 %", min=0, max=100, initial = 0, width = 300)  # 进度条

for (i in 1:nrow(all_chose)) {
  use_data <- all_data[, all_chose[i, ] ] # 提取
  colnames(use_data) <- c("x", "m1",'m2', 'y') # 改名
  
  # 链式
  result_ser_med<-process(data=use_data,y="y",x="x",m=c("m1",'m2'),normal=1,model=6,
                          effsize = 1,total=1,stand=1,
                          boot=5000,modelbt = 1,save=2,
                          plot=0,progress=0)
  
  if (  result_ser_med[26,3]*result_ser_med[26,4]>0 # x-m1-y
        &&result_ser_med[27,3]*result_ser_med[27,4]>0 # x-m2-y 显著
        &&result_ser_med[28,3]*result_ser_med[28,4]>0 # x-m1-m2-y 显著
 )
  {isser_med[i] <-1}  else  
  {isser_med[i] <-0}
  
 
  # 平行
  result_par_med <-
    process(
      data = use_data,
      y = "y",x = "x",m =c('m1','m2'),
      normal = 1,model = 4,
      effsize = 1,total = 1,stand = 1,
      boot = 5000,modelbt = 1,
      save = 2,plot = 3,progress = 0)
  
  if (  result_par_med[24,3]*result_par_med[24,4]>0 # x-m1-y
        &&result_par_med[25,3]*result_par_med[25,4]>0 # x-m2-y 显著
  )
  {ispar_med[i] <-1}  else  
  {ispar_med[i] <-0}
  
  
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
colnames(all_chose) <- c("x", "m1",'m2', 'y') 


isser_med<-data.frame(unlist(isser_med))
ispar_med<-data.frame(unlist(ispar_med))

isser_med<-isser_med$unlist.isser_med.
ispar_med<-ispar_med$unlist.ispar_med.

all_chose$isser_med<-isser_med
all_chose$ispar_med<-ispar_med

write_xlsx(x = all_chose,file = "result_ser_par_med.xlsx")

