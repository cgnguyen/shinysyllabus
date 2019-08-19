####Packages and Basic Setup####
  #Packages
  library(needs)
  needs(tidyr)
  needs(lubridate)
  needs(officer)
  needs(ical)
  needs(dplyr)
  
####Define Basic Parameters####
  startDate <- dmy("14-Oct-2019")
  endDate <- dmy("15-Feb-2020")
  
  #Weekdays select 2= Monday, 3=Tuesday etc. 
  weekdays<-"Mon"

  #Input Holidays and Semester Break
  

####Calculations for Days####
  #Find all days between the dates
  myDates <-seq(from = startDate, to = endDate, by = "days")

  #Identify Date of Weekday
    select.vec<-which(wday(myDates) %in% c(2))
    
  #Subset Data by Weekday
    
    date.vec<-myDates[select.vec]
    
    # # date.vec$holiday<-""
    # 

  #Generate Number Vector
    number.vec<-seq(1,length=length(date.vec))
  
  #Generate Vector of Heading Names (German Format)
    
    name.vec1<-paste("Sitzung",number.vec,sep=" ")
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
        
####Write Output to Word Document####
    #Generate basic document
    doc.full<-read_docx("template-deutsch.docx")
          
   
      # #Generate Holiday Names
      #   list.holidays<-list()
      #   
      #   for(i in 1:length(date.vec1)){
      #     temp.holiday<-fpar(paste(date.vec[i,]$holiday), 
      #                     fp_p = fp_par(text.align = "center"))
      #     list.holidays[[i]]<-temp.holiday
      #   }
      #   
    
   
    #Write seminar date in for loop
    for (i in 1:length(date.vec)){
      #Generate German output for date
      temp.date<-paste(day(date.vec[i]),
                       month(date.vec[i]),
                       year(date.vec[i]), 
                       sep=".")
    doc.full<-doc.full %>% 
      body_add_par(value = name.vec1[i], style =  "Sitzung") %>% 
      body_add_par(value = temp.date, style =  "Datum") 
    #%>%
      # body_add_fpar(list.holidays[[j]])
    }
   
    #Write Output
    doc.full%>% 
      print(target = "Seminarplan.docx")
    
    
    
    
                        
