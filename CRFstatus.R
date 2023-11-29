####################################################
#Step 0: First time to install the program packages
####################################################
install.packages("readxl")
install.packages("tidyverse")
install.packages("janitor")
install.packages("lubridate")
install.packages("formattable")
###################################################

library(readxl)
library(tidyverse)
library(janitor)
library(lubridate)
library(formattable) 

######################
#Step 1: Loading Data
######################

### the folder path: please note the "\" should be changed to "/"
setwd("C:/Users/kate0/Downloads/20230112")

### load CRF status
crf<-read_excel("EV71_Freezing_eCRF_List.xlsx")

### load subject visit status
sv<-read_excel("EV71_Summary_Subject_Visit_List.xls",col_types = "text")

### load subject list
sl<-read_excel("EV71_Subject_List.xls")


#####################################
#Step 2: Process the subject progress
#####################################

### retract the subject status
status<-sl %>% 
  select(Subject, Status)
sv<-merge(sv,status)


### process the real subject status
sv$status2<-sv$Status
sv$status2[is.na(sv$`Registration No.`) & sv$Status == "Ongoing"]<-"Ongoing-f"
sv$status2[is.na(sv$`Registration No.`) & sv$Status =="Withdrawal"]<-"Withdrawal-f"
sv$status2[sv$Status=="Screening Failure" & !is.na(sv$`Registration No.`)] <-"Screening Failure-w"
tabyl(sv,status2)


#################################
#Step 2-1: process visit schedule
#################################
sv$Visit1<-ymd(sv$Visit1)
sv$Visit2<-ymd(sv$Visit2)
sv$Visit3<-ymd(sv$Visit3)
sv$Visit4<-ymd(sv$Visit4)
sv$Visit5<-ymd(sv$Visit5)
sv$Visit6<-ymd(sv$Visit6)

### expected2 as the expected visit date for visit 2
sv$expected2<-vector(length = nrow(sv))
sv$expected2<-ymd(sv$expected2)
sv$expected2[sv$status2=="Ongoing"]<-sv$Visit1[sv$status2=="Ongoing"]+28
### expected3 as the expected visit date for visit 3
sv$expected3<-vector(length = nrow(sv))
sv$expected3<-ymd(sv$expected3)
sv$expected3[sv$status2=="Ongoing"]<-sv$Visit1[sv$status2=="Ongoing"]+56
### expected4 as the expected visit date for visit 4
sv$expected4<-vector(length = nrow(sv))
sv$expected4<-ymd(sv$expected4)
sv$expected4[sv$status2=="Ongoing"]<-sv$Visit1[sv$status2=="Ongoing"]+196

### expected5 as the expected visit date for visit 5
sv$expected5<-vector(length = nrow(sv))
sv$expected5<-ymd(sv$expected5)
sv$expected5[sv$status2=="Ongoing"]<-sv$Visit1[sv$status2=="Ongoing"]+396
### expected6 as the expected visit date for visit 6
sv$expected6<-vector(length = nrow(sv))
sv$expected6<-ymd(sv$expected6)
sv$expected6[sv$status2=="Ongoing"]<-sv$Visit1[sv$status2=="Ongoing"]+728

#########################################
#Step 2-2: Defining Missing/Future Visit 
#########################################

today<-today()
### using last friday as the cut-off for missing or future visit!
last_fri<-today-wday(today + 1)

sv$Visit2c<-as.character(sv$Visit2)
sv$Visit2c[sv$status2=="Ongoing" & is.na(sv$Visit2) & sv$expected2<=last_fri]<-"missing"
sv$Visit2c[sv$status2=="Ongoing" & is.na(sv$Visit2) & sv$expected2>last_fri]<-"future"
sv$Visit2c[!is.na(sv$Visit2)]<-"on-schedule"

sv$Visit3c<-as.character(sv$Visit3)
sv$Visit3c[sv$status2=="Ongoing" & is.na(sv$Visit3) & sv$expected3<=last_fri]<-"missing"
sv$Visit3c[sv$status2=="Ongoing" & is.na(sv$Visit3) & sv$expected3>last_fri]<-"future"
sv$Visit3c[!is.na(sv$Visit3)]<-"on-schedule"


sv$Visit4c<-as.character(sv$Visit4)
sv$Visit4c[sv$status2=="Ongoing" & is.na(sv$Visit4) & sv$expected4<=last_fri]<-"missing"
sv$Visit4c[sv$status2=="Ongoing" & is.na(sv$Visit4) & sv$expected4>last_fri]<-"future"
sv$Visit4c[!is.na(sv$Visit4)]<-"on-schedule"

sv$Visit5c<-as.character(sv$Visit5)
sv$Visit5c[sv$status2=="Ongoing" & is.na(sv$Visit5) & sv$expected5<=last_fri]<-"missing"
sv$Visit5c[sv$status2=="Ongoing" & is.na(sv$Visit5) & sv$expected5>last_fri]<-"future"
sv$Visit5c[!is.na(sv$Visit5)]<-"on-schedule"

sv$Visit6c<-as.character(sv$Visit6)
sv$Visit6c[sv$status2=="Ongoing" & is.na(sv$Visit6) & sv$expected4<=last_fri]<-"missing"
sv$Visit6c[sv$status2=="Ongoing" & is.na(sv$Visit6) & sv$expected4>last_fri]<-"future"
sv$Visit6c[!is.na(sv$Visit6)]<-"on-schedule"


###############################################################################################
#Step 2-3: Defining subject visit progress as of today (on-schedule as EDC date/missing/future)
###############################################################################################
sv$visit<-rep(NA,nrow(sv))

sv$visit[sv$status2=="Ongoing" & sv$Visit6c!="future"]<-"v6"
sv$visit[sv$status2=="Ongoing" & is.na(sv$visit) & sv$Visit5c!="future"]<-"v5"
sv$visit[sv$status2=="Ongoing" & is.na(sv$visit) & sv$Visit4c!="future"]<-"v4"
sv$visit[sv$status2=="Ongoing" & is.na(sv$visit) & sv$Visit3c!="future"]<-"v3"
sv$visit[sv$status2=="Ongoing" & is.na(sv$visit) & sv$Visit2c!="future"]<-"v2"
sv$visit[sv$status2=="Ongoing" & is.na(sv$visit)]<-"v1"

sv$visit[sv$status2 %in% c("Ongoing-f","Withdrawal-f","Screening Failure")]<-"v1"
wit<-sv %>% 
  filter(status2 %in% c("Withdrawal","Screening Failure-w"))

for (i in 1:nrow(wit)){
  visit<-c(wit$Visit1[i],wit$Visit2[i],wit$Visit3[i],wit$Visit4[i],wit$Visit5[i],wit$Visit6[i])
  if (which.max(visit)==1) {wit$visit[i]<-"v1"}
  if (which.max(visit)==2) {wit$visit[i]<-"v2"}
  if (which.max(visit)==3) {wit$visit[i]<-"v3"}
  if (which.max(visit)==4) {wit$visit[i]<-"v4"}
  if (which.max(visit)==5) {wit$visit[i]<-"v5"}
  if (which.max(visit)==6) {wit$visit[i]<-"v6"}
  sv$visit[sv$Subject==wit$Subject[i]]<-wit$visit[i]
}

tabyl(sv,status2,visit)

####################
### Ongoing 
####################
ong<-sv %>% 
  filter(status2=="Ongoing") %>%
  select(Subject, visit)

crf_ong1<-crf %>% 
  filter(Subject %in% ong$Subject[ong$visit=="v1"]) %>% 
  filter(Visit %in% c("Subject Information", "Visit1")) %>% 
  filter(`Saving status`!="NA")

crf_ong2<-crf %>% 
  filter(Subject %in% ong$Subject[ong$visit=="v2"]) %>% 
  filter(Visit %in% c("Subject Information", "Visit1","Visit2")) %>% 
  filter(`Saving status`!="NA")

crf_ong3<-crf %>% 
  filter(Subject %in% ong$Subject[ong$visit=="v3"]) %>% 
  filter(Visit %in% c("Subject Information", "Visit1","Visit2","Visit3")) %>% 
  filter(`Saving status`!="NA")

crf_ong4<-crf %>% 
  filter(Subject %in% ong$Subject[ong$visit=="v4"]) %>% 
  filter(Visit %in% c("Subject Information", "Visit1","Visit2","Visit3","Visit4")) %>% 
  filter(`Saving status`!="NA")

crf_ong5<-crf %>% 
  filter(Subject %in% ong$Subject[ong$visit=="v5"]) %>% 
  filter(Visit %in% c("Subject Information", "Visit1","Visit2","Visit3","Visit4","Visit5")) %>% 
  filter(`Saving status`!="NA")

crf_ong6<-crf %>% 
  filter(Subject %in% ong$Subject[ong$visit=="v6"]) %>% 
  filter(Visit %in% c("Subject Information", "Visit1","Visit2","Visit3","Visit4","Visit5","Visit6")) %>% 
  filter(`Saving status`!="NA")

crf_ong<-bind_rows(crf_ong1,crf_ong2,crf_ong3,crf_ong4,crf_ong5,crf_ong6)

#N
#entered
crf_ong %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(`Saving status`,Visit)

#sdv
crf_ong %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(SDV, Visit)
#freezing
crf_ong %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(Status, Visit)

#V
#entered
crf_ong %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(`Saving status`,Visit)

#sdv
crf_ong %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(SDV, Visit)
#freezing
crf_ong %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(Status, Visit)


####################
### Withdrawal
####################

wit<-sv %>% 
  filter(status2 =="Withdrawal") %>% 
  select(Subject)

crf_wit<-crf %>% 
  filter(Subject %in% wit$Subject) %>% 
  filter(Visit %in% c("Subject Information", "Visit1","Visit2","Visit3","Visit4","Visit5","Visit6")) %>% 
  filter(`Saving status`!="NA")

#N
#entered
crf_wit %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(`Saving status`,Visit)
#sdv
crf_wit %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(SDV, Visit)
#freezing
crf_wit %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(Status, Visit)


#V
#entered
crf_wit %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(`Saving status`,Visit)

#sdv
crf_wit %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(SDV, Visit)
#freezing
crf_wit %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(Status, Visit)


####################
### Screening Fail
####################

sf<-sv %>% 
  filter(status2 =="Screening Failure") %>% 
  select(Subject)

rand_Sf<-crf %>% 
  filter(Subject %in% sf$Subject) %>% 
  filter(eCRF =="Randomization") %>%
  filter(`Saving status`=="Complete") %>% 
  select(Subject, Site)

crf_sfr<-crf %>% 
  filter(Subject %in% rand_Sf$Subject) %>% 
  filter(Visit %in% c("Subject Information", "Visit1","Visit2","Visit3","Visit4","Visit5","Visit6")) %>% 
  filter(`Saving status`!="NA")


#N
#entered
crf_sfr %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(`Saving status`,Visit)
#sdv
crf_sfr %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(SDV, Visit)
#freezing
crf_sfr %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(Status, Visit)


#V
#entered
crf_sfr %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(`Saving status`,Visit)

#sdv
crf_sfr %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(SDV, Visit)
#freezing
crf_sfr %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(Status, Visit)

rand_Sf %>% 
  tabyl(Site)

##############################################################################################################
### FINAL TABLE OUTPUT #######################################################################################
##############################################################################################################

###############
### Table A
###############
sv$status_this<-sv$status2
sv$status_this[sv$status2=="Withdrawal-f"|sv$status2=="Ongoing-f"]<-"Should be failure"
sv$status_this[sv$status2=="Screening Failure-w"]<-"Should be withdrawal"
sv<-sv %>% 
  mutate(site=case_when(
    Site=="Nguyen Dinh Chieu Hospital"~"N",
    Site=="Vinh Long City Health Center"~"V"
  ))

data<-tabyl(sv,status_this,site)
#data[nrow(data) + 1,] = c("Total",sum(data$N),sum(data$V))
aa<-tabyl(sv,status_this,site,visit)
la<-length(aa)

blank<-as.data.frame(matrix(nrow = 5, ncol=12))
dta<-cbind(data,blank)

colnames(dta)<-c("status_this","N","V",
                 "V1.N","V1.V",
                 "V2.N","V2.V",
                 "V3.N","V3.V",
                 "V4.N","V4.V",
                 "V5.N","V5.V",
                 "V6.N","V6.V")

#Visit 1
dta[,4:5]<-aa$v1[,2:3]
#Visit 2
dta[,6:7]<-aa$v2[,2:3]
#Visit 3
dta[,8:9]<-aa$v3[,2:3]
#Visit 4
dta[,10:11]<-aa$v4[,2:3]
#Visit 5
dta[,12:13]<-aa$v5[,2:3]
#Visit 6
dta[,14:15]<-aa$v6[,2:3]

### Table A output : dta
dta[nrow(data) + 1,] = c("Total",colSums(dta[,2:ncol(dta)],na.rm = T))


###############
### Table B
###############
ong_t<-sv %>% filter(status2=="Ongoing")

ong_t$Visit2c<-factor(ong_t$Visit2c,labels = c("on-schedule","missing","fututure"))
ong_t$Visit3c<-factor(ong_t$Visit3c,labels = c("on-schedule","missing","fututure"))
ong_t$Visit4c<-factor(ong_t$Visit4c,labels = c("on-schedule","missing","fututure"))
ong_t$Visit5c<-factor(ong_t$Visit5c,labels = c("on-schedule","missing","fututure"))
ong_t$Visit6c<-factor(ong_t$Visit6c,labels = c("on-schedule","missing","fututure"))

dtb<-as.data.frame(matrix(nrow = 1, ncol=36))
#1=on-schedule, 2=missing, 3=future
colnames(dtb)<-c("v1.1.n","v1.1.v","v1.2.n","v1.2.v","v1.3.n","v1.3.v",
                 "v2.1.n","v2.1.v","v2.2.n","v2.2.v","v2.3.n","v2.3.v",
                 "v3.1.n","v3.1.v","v3.2.n","v3.2.v","v3.3.n","v3.3.v",
                 "v4.1.n","v4.1.v","v4.2.n","v4.2.v","v4.3.n","v4.3.v",
                 "v5.1.n","v5.1.v","v5.2.n","v5.2.v","v5.3.n","v5.3.v",
                 "v6.1.n","v6.1.v","v6.2.n","v6.2.v","v6.3.n","v6.3.v")

#Visit 1
dtb[,1:2]<-dta[1,2:3]
dtb[,3:6]<-"-"

#Visit 2
v2<-tabyl(ong_t,site,Visit2c)
dtb[,7:12]<-cbind(v2$`on-schedule`,v2$missing,v2$fututure)

#Visit 3
v3<-tabyl(ong_t,site,Visit3c)
dtb[,13:18]<-cbind(v3$`on-schedule`,v3$missing,v3$fututure)

#Visit 4
v4<-tabyl(ong_t,site,Visit4c)
dtb[,19:24]<-cbind(v4$`on-schedule`,v4$missing,v4$fututure)

#Visit 5
v5<-tabyl(ong_t,site,Visit5c)
dtb[,25:30]<-cbind(v5$`on-schedule`,v5$missing,v5$fututure)

#Visit 6
v6<-tabyl(ong_t,site,Visit6c)
dtb[,31:36]<-cbind(v6$`on-schedule`,v6$missing,v6$fututure)


###############
### Table C
###############

######################
### Table C-1: Site N
######################

dtc1<-as.data.frame(matrix(nrow = 3, ncol=21))
rownames(dtc1)<-c("ongoing","withdrawal","failure-r")
colnames(dtc1)<-c("v0.ent","v1.ent","v2.ent","v3.ent","v4.ent","v5.ent","v6.ent",
                  "v0.sdv","v1.sdv","v2.sdv","v3.sdv","v4.sdv","v5.sdv","v6.sdv",
                  "v0.fre","v1.fre","v1.fre","v3.fre","v4.fre","v5.fre","v6.fre")
### Ongoing
ent<-crf_ong %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(`Saving status`,Visit)

ent[nrow(ent) + 1,] = c("all",colSums(ent[,2:ncol(ent)],na.rm = T))
visit<-ncol(ent)-1

#entry rate
dtc1[1,1:visit]<-as.character(percent(as.numeric(ent[1,2:ncol(ent)])/as.numeric(ent[3,2:ncol(ent)]),1))

#SDV rate
sdv<-crf_ong %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(SDV, Visit)

dtc1[1,8:(8+visit-1)]<-as.character(percent(as.numeric(sdv[1,2:ncol(sdv)])/as.numeric(ent[1,2:ncol(ent)]),1))

#Freezing rate
fre<-crf_ong %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(Status, Visit)

dtc1[1,15:(15+visit-1)]<-as.character(percent(as.numeric(fre[1,2:ncol(fre)])/as.numeric(sdv[1,2:ncol(sdv)]),1))

### Withdrawal
ent<-crf_wit %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(`Saving status`,Visit)

ent[nrow(ent) + 1,] = c("all",colSums(ent[,2:ncol(ent)],na.rm = T))
visit<-ncol(ent)-1

#entry rate
dtc1[2,1:visit]<-as.character(percent(as.numeric(ent[1,2:ncol(ent)])/as.numeric(ent[2,2:ncol(ent)]),1))

#SDV rate
sdv<-crf_wit %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(SDV, Visit)

dtc1[2,8:(8+visit-1)]<-as.character(percent(as.numeric(sdv[1,2:ncol(sdv)])/as.numeric(ent[1,2:ncol(sdv)]),1))

#Freezing rate
fre<-crf_wit %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(Status, Visit)

dtc1[2,15:(15+visit-1)]<-as.character(percent(as.numeric(fre[1,2:ncol(fre)])/as.numeric(sdv[1,2:ncol(sdv)]),1))


### Screening Failure (correct subject status: failure-r)

#entry rate
ent<-crf_sfr %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(`Saving status`,Visit)

ent[nrow(ent) + 1,] = c("all",colSums(ent[,2:ncol(ent)],na.rm = T))
visit<-ncol(ent)-1

dtc1[3,1:visit]<-as.character(percent(as.numeric(ent[1,2:ncol(ent)])/as.numeric(ent[2,2:ncol(ent)]),1))

#SDV rate
sdv<-crf_sfr %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(SDV, Visit)

dtc1[3,8:(8+visit-1)]<-as.character(percent(as.numeric(sdv[1,2:ncol(sdv)])/as.numeric(ent[1,2:ncol(sdv)]),1))

#Freezing rate
fre<-crf_sfr %>% 
  filter(Site=="N.Nguyen Dinh Chieu Hospital") %>% 
  tabyl(Status, Visit)

dtc1[3,15:(15+visit-1)]<-as.character(percent(as.numeric(fre[1,2:ncol(fre)])/as.numeric(sdv[1,2:ncol(sdv)]),1))

#re-order the table
order<-c("v0.ent","v0.sdv","v0.fre",
         "v1.ent","v1.sdv","v1.fre",
         "v2.ent","v2.sdv","v2.fre",
         "v3.ent","v3.sdv","v3.fre",
         "v4.ent","v4.sdv","v4.fre",
         "v5.ent","v5.sdv","v5.fre",
         "v6.ent","v6.sdv","v6.fre")
dtc10<-dtc1


######################
### Table C-2: Site V
######################

dtc2<-as.data.frame(matrix(nrow = 3, ncol=21))
rownames(dtc2)<-c("ongoing","withdrawal","failure-r")
colnames(dtc2)<-c("v0.ent","v1.ent","v2.ent","v3.ent","v4.ent","v5.ent","v6.ent",
                  "v0.sdv","v1.sdv","v2.sdv","v3.sdv","v4.sdv","v5.sdv","v6.sdv",
                  "v0.fre","v1.fre","v1.fre","v3.fre","v4.fre","v5.fre","v6.fre")
### Ongoing
ent<-crf_ong %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(`Saving status`,Visit)

ent[nrow(ent) + 1,] = c("all",colSums(ent[,2:ncol(ent)],na.rm = T))
visit<-ncol(ent)-1

#entry rate
dtc2[1,1:visit]<-as.character(percent(as.numeric(ent[1,2:ncol(ent)])/as.numeric(ent[3,2:ncol(ent)]),1))

#SDV rate
sdv<-crf_ong %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(SDV, Visit)

dtc2[1,8:(8+visit-1)]<-as.character(percent(as.numeric(sdv[1,2:ncol(sdv)])/as.numeric(ent[1,2:ncol(ent)]),1))

#Freezing rate
fre<-crf_ong %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(Status, Visit)

dtc2[1,15:(15+visit-1)]<-as.character(percent(as.numeric(fre[1,2:ncol(fre)])/as.numeric(sdv[1,2:ncol(sdv)]),1))

### Withdrawal
ent<-crf_wit %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(`Saving status`,Visit)

ent[nrow(ent) + 1,] = c("all",colSums(ent[,2:ncol(ent)],na.rm = T))
visit<-ncol(ent)-1

#entry rate
dtc2[2,1:visit]<-as.character(percent(as.numeric(ent[1,2:ncol(ent)])/as.numeric(ent[2,2:ncol(ent)]),1))

#SDV rate
sdv<-crf_wit %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(SDV, Visit)

dtc2[2,8:(8+visit-1)]<-as.character(percent(as.numeric(sdv[1,2:ncol(sdv)])/as.numeric(ent[1,2:ncol(sdv)]),1))

#Freezing rate
fre<-crf_wit %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(Status, Visit)

dtc2[2,15:(15+visit-1)]<-as.character(percent(as.numeric(fre[1,2:ncol(fre)])/as.numeric(sdv[1,2:ncol(sdv)]),1))


### Screening Failure (correct subject status: failure-r)

#entry rate
ent<-crf_sfr %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(`Saving status`,Visit)

ent[nrow(ent) + 1,] = c("all",colSums(ent[,2:ncol(ent)],na.rm = T))
visit<-ncol(ent)-1

dtc2[3,1:visit]<-as.character(percent(as.numeric(ent[1,2:ncol(ent)])/as.numeric(ent[2,2:ncol(ent)]),1))

#SDV rate
sdv<-crf_sfr %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(SDV, Visit)

dtc2[3,8:(8+visit-1)]<-as.character(percent(as.numeric(sdv[1,2:ncol(sdv)])/as.numeric(ent[1,2:ncol(sdv)]),1))

#Freezing rate
fre<-crf_sfr %>% 
  filter(Site=="V.Vinh Long City Health Center") %>% 
  tabyl(Status, Visit)

dtc2[3,15:(15+visit-1)]<-as.character(percent(as.numeric(fre[1,2:ncol(fre)])/as.numeric(sdv[1,2:ncol(sdv)]),1))

#re-order the table
dtc20<-dtc2

### footnote for reference
footnote<-as.data.frame(matrix(nrow = 7, ncol=1))
rownames(footnote)<-c("today date","last friday date","ongoing-f","withdrawal-f","failure-r(all)", "failure-r(n)","failure-r(v)")
footnote[1,]<-as.character(today)
footnote[2,]<-as.character(last_fri)
footnote[3,]<-tabyl(sv,status2)[2,2]
footnote[4,]<-tabyl(sv,status2)[6,2]
footnote[5,]<-nrow(rand_Sf)
footnote[6,]<-sum(rand_Sf$Site=="N.Nguyen Dinh Chieu Hospital")
footnote[7,]<-sum(rand_Sf$Site=="V.Vinh Long City Health Center")

##############################################################
### FINAL REPORT DATA EXPORT IN ONE EXCEL ####################
##############################################################

### the file will be save at the file path in the beginning!          
xlsx::write.xlsx(sv, "EV71_Subject_project_status.xlsx", sheetName = "Sheet1", 
           col.names = TRUE, row.names = TRUE, append = FALSE)

xlsx::write.xlsx(dta, "EV71_Subject_project_status.xlsx", sheetName = "table_a", 
                 col.names = TRUE, row.names = TRUE, append = TRUE)
xlsx::write.xlsx(dtb, "EV71_Subject_project_status.xlsx", sheetName = "table_b", 
                 col.names = TRUE, row.names = TRUE, append = TRUE)
xlsx::write.xlsx(dtc1, "EV71_Subject_project_status.xlsx", sheetName = "table_c1", 
                 col.names = TRUE, row.names = TRUE, append = TRUE)
xlsx::write.xlsx(dtc2, "EV71_Subject_project_status.xlsx", sheetName = "table_c2", 
                 col.names = TRUE, row.names = TRUE, append = TRUE)
xlsx::write.xlsx(footnote, "EV71_Subject_project_status.xlsx", sheetName = "footnote", 
                 col.names = TRUE, row.names = TRUE, append = TRUE)




