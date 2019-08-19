library(shiny)  
  library(needs)
  library(tidyverse)
  library(lubridate)
  library(officer)
  library(ical)

  

ui <- fluidPage(
  titlePanel("Seminarplan-R"),
  
  sidebarLayout(
    sidebarPanel(
      #Input for Date Range
      dateInput("date_begin", label = h3("Beginn der Vorlesungszeit"), 
                language= "de",  
                format = "dd-mm-yyyy"),
      dateInput("date_end", label = h3("Ende der Vorlesungszeit"), 
                language= "de",
                format = "dd-mm-yyyy"),
      
      #Week Days 
      checkboxGroupInput("weekdays", label = h3("Sitzungstage"), 
      choices = list("Montag" = 2, 
                     "Dienstag" = 3, 
                     "Mittwoch" = 4, 
                     "Donnerstag" = 5, 
                     "Freitag" = 6, 
                     "Samstag" = 7, 
                     "Sontag" = 1), hr())
      ),
    
    mainPanel(
      textOutput("begin"),
      verbatimTextOutput("seminardays"),
      downloadButton('downloadData', 'Seminarplan Template Download')

    )
  )
)

server <- function(input, output) {
  
  ####Calculations####
  #Find all days between the dates


  
  
  
  output$begin <- renderText({ 
    paste("Das Seminar findet zwischen", 
          input$date_begin,  "und", 
          input$date_end, "statt")
  })
  
   output$downloadData <- downloadHandler(
     filename = function(){"seminarplanR.docx"},

     content = function(file) {
      ###Generate Date Vector###
       #Get Dates
       myDates <-seq(from = input$date_begin, to = input$date_end, by = "days")
       #Identify Date of Weekday
       select.vec<-which(wday(myDates) %in% c(input$weekdays))
       #Subset Data by Weekday
       date.vec<-myDates[select.vec]
       
       #Generate Number Vector
       number.vec<-seq(1,length=length(date.vec))

       #Generate Vector of Heading Names (German Format)
       name.vec1<-paste("Sitzung",number.vec,sep=" ")
       
      ###Generate File###
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
    }
   
    #Write Output
          doc.full%>% 
            print(target = file)
    }
   )
      
}



# Run the application 
shinyApp(ui = ui, server = server)

