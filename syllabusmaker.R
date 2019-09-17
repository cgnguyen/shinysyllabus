####Packages and Basic Setup####
  #Packages
  library(needs)
  needs(tidyr)
  needs(lubridate)
  needs(officer)
  needs(ical)
  needs(dplyr)

####Define Basic Parameters####
<<<<<<< HEAD
startDate <- dmy("14-10-2019")
endDate <- dmy("15-02-2020")
=======
startDate <- dmy("17-Aug-2019")
endDate <- dmy("19-Feb-2020")
>>>>>>> fe6e1620ebd0aee0079efd3debfdb33ee620a0f8

#Weekdays select 2= Monday, 3=Tuesday etc. 
weekdays<-"Mon"



####Calculations####
#Find all days between the dates
myDates <-seq(from = startDate, to = endDate, by = "days")

#Identify Date of Weekday
<<<<<<< HEAD
select.vec<-which(wday(myDates) %in% c(2))
=======
select.vec<-which(wday(myDates) %in% c(2,4))
>>>>>>> fe6e1620ebd0aee0079efd3debfdb33ee620a0f8

#Subset Data by Weekday

date.vec<-myDates[select.vec]
<<<<<<< HEAD

=======
date.vec<-as.data.frame(date.vec)
date.vec$holiday<-""
>>>>>>> fe6e1620ebd0aee0079efd3debfdb33ee620a0f8


#Generate Number Vector
number.vec<-seq(1,length=length(date.vec))

#Generate Vector of Heading Names (German Format)

name.vec1<-paste("Sitzung",number.vec,sep=": ")
date.vec1<-paste(day(date.vec),month(date.vec), sep=".")

<<<<<<< HEAD
=======
####Include Breaks and Holidays####


####Check for Public Holidays####
#Read in Public Holidays for Berlin - find a better source than this and do this for each state?
ics_raw = readLines("http://i.cal.to/ical/58/berlin/feiertage/a6afca09.3e559e14-5d35ef26.ics")

holidays.ger = ical_parse_df(text=ics_raw) 

#Arrange Calender for easier Reading
holidays.ger<-holidays.ger %>%
  arrange(start) %>%
  select(start,end,summary) 

#Generate cleaned time vector
holidays.ger$date<-as.Date(ymd_hms(holidays.ger$start), "%a, %d %b %Y" , tz="GMT")


#Compare holiday vector to own dataset 
holiday.vec<-date.vec$date.vec  %in% holidays.ger$date


#Compare own dataset to holiday data frame
holiday.name.vec<-   holidays.ger[holidays.ger$date %in% date.vec$date.vec,]$summary

#Note if a day is a holiday in data.frame of dates
date.vec[holiday.vec,]$holiday<-as.character(holiday.name.vec)

>>>>>>> fe6e1620ebd0aee0079efd3debfdb33ee620a0f8




####Write Output to Word Document####
#Generate basic document
doc.full<-read_docx() %>%
  body_add_par("Seminarplan", style="Normal")


#Generate Date Section(s) that are right-oriented for each of the dates (German format)
list.dates<-list()
for(i in 1:length(date.vec1)){
<<<<<<< HEAD
  temp.date<-fpar(paste(day(date.vec[i]),
                        month(date.vec[i]),
                        year(date.vec[i]), 
=======
  temp.date<-fpar(paste(day(date.vec[i,]$date.vec),
                        month(date.vec[i,]$date.vec),
                        year(date.vec[i,]$date.vec), 
>>>>>>> fe6e1620ebd0aee0079efd3debfdb33ee620a0f8
                        sep="."), 
                  fp_p = fp_par(text.align = "right"))
  list.dates[[i]]<-temp.date
}

<<<<<<< HEAD
=======
#Generate Holiday Names
list.holidays<-list()

for(i in 1:length(date.vec1)){
  temp.holiday<-fpar(paste(date.vec[i,]$holiday), 
                     fp_p = fp_par(text.align = "center"))
  list.holidays[[i]]<-temp.holiday
}

>>>>>>> fe6e1620ebd0aee0079efd3debfdb33ee620a0f8


#Write seminar date in for loop
for (j in 1:length(date.vec1)){
  doc.full<-doc.full %>% 
    body_add_par(value = name.vec1[j], style =  "heading 1") %>% 
<<<<<<< HEAD
    body_add_fpar(list.dates[[j]])}
=======
    body_add_fpar(list.dates[[j]]) %>%
    body_add_fpar(list.holidays[[j]])
}
>>>>>>> fe6e1620ebd0aee0079efd3debfdb33ee620a0f8

#Write Output
doc.full%>% 
  print(target = "body_add_demo.docx")





<<<<<<< HEAD
# 
# ####Include Breaks and Holidays####
# 
# 
# ####Check for Public Holidays####
# #Read in Public Holidays for Berlin - find a better source than this and do this for each state?
# ics_raw = readLines("http://i.cal.to/ical/58/berlin/feiertage/a6afca09.3e559e14-5d35ef26.ics")
# 
# holidays.ger = ical_parse_df(text=ics_raw) 
# 
# #Arrange Calender for easier Reading
# holidays.ger<-holidays.ger %>%
#   arrange(start) %>%
#   select(start,end,summary) 
# 
# #Generate cleaned time vector
# holidays.ger$date<-as.Date(ymd_hms(holidays.ger$start), "%a, %d %b %Y" , tz="GMT")
# 
# 
# #Compare holiday vector to own dataset 
# holiday.vec<-date.vec$date.vec  %in% holidays.ger$date
# 
# 
# #Compare own dataset to holiday data frame
# holiday.name.vec<-   holidays.ger[holidays.ger$date %in% date.vec$date.vec,]$summary
# 
# #Note if a day is a holiday in data.frame of dates
# date.vec[holiday.vec,]$holiday<-as.character(holiday.name.vec)
# 


=======
>>>>>>> fe6e1620ebd0aee0079efd3debfdb33ee620a0f8
