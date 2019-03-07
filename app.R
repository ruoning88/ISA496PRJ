load("xgb_Set9.RData")
load("xgb_Run4.RData")
load("Explainer.RData")
load("ExplainerRun.RData")

library(shiny)
library(caret)
library(ggplot2)
library(shinyWidgets)
library(forecast)
library(dplyr)
library(xgboost)
library(devtools)
library(xgboostExplainer)
library(DiagrammeR)
library(readxl)
library(writexl)

options(shiny.sanitize.errors = FALSE)
ui<- fluidPage(
  titlePanel("Time Prediction Model"),
  tabsetPanel (
    tabPanel("SetupTime", fluid=TRUE, 
             setBackgroundColor("Ivory"),
             titlePanel(title="Set Up Time Predictions"),
             
             sidebarLayout(
               sidebarPanel(
                 fileInput('file1', 'Input your Data',
                           accept = c(".xlsx")),
                 
                 downloadButton('download1',"Download SetUp Time Predictions"),
                 actionButton("button", "Show Graph")
                 
               ),
               mainPanel(
                 plotOutput(outputId="xgbgraphset",height="500px",width="100%")
               )
             )),
    
    tabPanel("RunupTime", fluid=TRUE, 
             setBackgroundColor("Ivory"),
             titlePanel(title="Run Time Predictions"),
             
             sidebarLayout(
               sidebarPanel(
                 fileInput('file2', 'Input your Data',
                           accept = c(".xlsx")),
                 
                 downloadButton('download2',"Download Run Time Predictions"),
                 actionButton("button2", "Show Graph")
               ) 
               ,
               mainPanel(
                 plotOutput(outputId="xgbgraphrun",height="500px",width="100%")
               ))
    )
  )
)

server <- function (input,output,session) {
  downdata <- reactive ({
    if(is.null(input$file1)) return ()
    inFile <- input$file1
    setdata <- read_excel(inFile$datapath)
    colnames(setdata) <- c("Standard Setup Hours","Standard Run Hours","EaInkChange","HHAttachment","FullPartDC","FullDieCut","Machine4","100Coverage","RockTennMRA","DCAttachment",
                           "BoxWidth55","Michelman40E","PercInk","CuttingDieUN","CuttingDieOthers","WRADW","33MediumDW","26MediumDW",
                           "ExtGlueTab","26Medium","Ink50","GlueJointGlue","StyleRSC","StyleRSCL","NumInk","EaPlateChange","Grade40","Grade32","Grade200",
                           "40Medium","33Medium","WRAY","500BoxUnit","NUGHZ")
    
    validation.test <- setdata %>% select(EaInkChange,HHAttachment,FullPartDC,FullDieCut,Machine4,`100Coverage`,RockTennMRA,DCAttachment,
                                          BoxWidth55,Michelman40E,PercInk,CuttingDieUN,WRADW,`33MediumDW`,`26MediumDW`,
                                          ExtGlueTab,`26Medium`,Ink50,GlueJointGlue,StyleRSC,NumInk,StyleRSCL,EaPlateChange)
    validation.test <- as.data.frame(validation.test)
    validation.std <- setdata[,1]
    i <- sapply(validation.test, is.character)
    categoricalvl <- validation.test[,i]
    for(m in 1:ncol(categoricalvl)) {
      categoricalvl[,m] <- recode_factor(categoricalvl[,m],"Y"="1","N"="0")
      categoricalvl[,m] <- as.character(categoricalvl[,m])
      categoricalvl[,m] <- factor(categoricalvl[,m],levels=c("0","1"))
    }
    categoricalvl <- categoricalvl %>% model.matrix(~.-1,.)
    categoricalvl <- categoricalvl[,-1]
    colnames(categoricalvl) <- c("EaInkChange1","HHAttachment1","FullPartDC1","FullDieCut1","Machine41","`100Coverage`1","RockTennMRA1","DCAttachment1",
                                 "BoxWidth551","Michelman40E1","CuttingDieUN1","WRADW1","`33MediumDW`1","`26MediumDW`1",
                                 "ExtGlueTab1","`26Medium`1","Ink501","GlueJointGlue1","StyleRSC1","StyleRSCL1","EaPlateChange1")
    numericvl <- validation.test[,c(11,21)] %>% as.matrix()
    validation.tstx <- cbind(numericvl,categoricalvl)
    validation.tstx
    
    ptr.set.xgb <- predict(xgb_tune_Set9, newdata=validation.tstx)
    constant.Set <- 2.1671
    predictions.Set <- (10^ptr.set.xgb)-constant.Set ## This is the predicted diff between actual time and the time Batavia came up with
    validation.final <- cbind(validation.std,validation.test,predictions.Set)
    colnames(validation.final)[1] <- "Standard Setup Hours"
    validation.final$PredictedSetUpTime <- validation.final$predictions.Set + validation.final$`Standard Setup Hours`
    validation.final <- validation.final %>% select(-predictions.Set)
    validation.final
  })
  
  observeEvent(input$button, {
    output$xgbgraphset <- renderPlot({
      if(is.null(input$file1)) return ()
      inFile <- input$file1
      setdata <- read_excel(inFile$datapath)
      colnames(setdata) <- c("Standard Setup Hours","Standard Run Hours","EaInkChange","HHAttachment","FullPartDC","FullDieCut","Machine4","100Coverage","RockTennMRA","DCAttachment",
                             "BoxWidth55","Michelman40E","PercInk","CuttingDieUN","CuttingDieOthers","WRADW","33MediumDW","26MediumDW",
                             "ExtGlueTab","26Medium","Ink50","GlueJointGlue","StyleRSC","StyleRSCL","NumInk","EaPlateChange","Grade40","Grade32","Grade200",
                             "40Medium","33Medium","WRAY","500BoxUnit","NUGHZ")
      
      validation.test <- setdata %>% select(EaInkChange,HHAttachment,FullPartDC,FullDieCut,Machine4,`100Coverage`,RockTennMRA,DCAttachment,
                                            BoxWidth55,Michelman40E,PercInk,CuttingDieUN,WRADW,`33MediumDW`,`26MediumDW`,
                                            ExtGlueTab,`26Medium`,Ink50,GlueJointGlue,StyleRSC,NumInk,StyleRSCL,EaPlateChange)
      validation.test <- as.data.frame(validation.test)
      i <- sapply(validation.test, is.character)
      categoricalvl <- validation.test[,i]
      for(m in 1:ncol(categoricalvl)) {
        categoricalvl[,m] <- recode_factor(categoricalvl[,m],"Y"="1","N"="0")
        categoricalvl[,m] <- as.character(categoricalvl[,m])
        categoricalvl[,m] <- factor(categoricalvl[,m],levels=c("0","1"))
      }
      numericvl <- validation.test[,c(11,21)] 
      validation.tst <- cbind(numericvl,categoricalvl)
      
      xgbvalid <- xgb.DMatrix(data.matrix(validation.tst))
      colnames(xgbvalid) <- c("PercInk","NumInk","EaInkChange1","HHAttachment1","FullPartDC1","FullDieCut1",   
                              "Machine41","`100Coverage`1" ,"RockTennMRA1","DCAttachment1" , "BoxWidth551", "Michelman40E1", 
                              "CuttingDieUN1","WRADW1","`33MediumDW`1","`26MediumDW`1","ExtGlueTab1","`26Medium`1",   
                              "Ink501","GlueJointGlue1","StyleRSC1","StyleRSCL1","EaPlateChange1")
      
      model <- xgb_tune_Set9$finalModel
      m <- showWaterfall(model, explainer, xgbvalid, data.matrix(validation.tst),1, type = "regression") + 
        theme_bw()+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        ggtitle("Predictions for the 1st Observation")+
        xlab("Important Factors Used")+ylab("Values")+
        scale_y_continuous(limits=c(0,0.4),breaks=seq(-0.5,0.4,0.05))+
        theme(plot.title = element_text(size=10,face="bold",hjust = 0.5),panel.grid.major = element_blank(),
              axis.title=element_text(size=7,face="bold"))
      build <-ggplot_build(m)
      xlabl <- build$layout$panel_scales_x[[1]]$labels
      xlabl <- gsub("1 =","=",xlabl)
      xlabl[-grep("^Num|^Perc|2$",xlabl)] <- gsub("= 1","=0",xlabl[-grep("^Num|^Perc|2$",xlabl)])
      xlabl[-grep("^Num|^Perc|0$",xlabl)] <- gsub("= 2","=1",xlabl[-grep("^Num|^Perc|0$",xlabl)])
      m + scale_x_discrete(labels=xlabl)
    })
  })
  
  output$download1 <- downloadHandler(
    filename = function() {
      paste("SetupPredictions-", Sys.Date(), ".xlsx", sep="")
    },
    content = function(file) {
      write_xlsx(downdata(), file)
    }
  )
  
  
  
  ##------------------------------- Second Tab--Run Time -------------------------------------##
  downdata2 <- reactive ({
    if(is.null(input$file2)) return ()
    inFile2 <- input$file2
    rundata <- read_excel(inFile2$datapath)
    colnames(rundata) <- c("Standard Setup Hours","Standard Run Hours","EaInkChange","HHAttachment","FullPartDC","FullDieCut","Machine4","100Coverage","RockTennMRA","DCAttachment",
                           "BoxWidth55","Michelman40E","PercInk","CuttingDieUN","CuttingDieOthers","WRADW","33MediumDW","26MediumDW",
                           "ExtGlueTab","26Medium","Ink50","GlueJointGlue","StyleRSC","StyleRSCL","NumInk","EaPlateChange","Grade40","Grade32","Grade200",
                           "40Medium","33Medium","WRAY","500BoxUnit","NUGHZ")
    
    validation.test <-rundata %>% select(Ink50,`100Coverage`,FullDieCut,PercInk,DCAttachment,`33MediumDW`,
                                         Machine4,Grade40,BoxWidth55,Grade32,CuttingDieUN,`26Medium`,StyleRSCL,
                                         EaPlateChange,StyleRSC,`40Medium`,HHAttachment,Grade200,CuttingDieOthers,
                                         WRADW,`33Medium`, WRAY,ExtGlueTab,`500BoxUnit`,NUGHZ)
    validation.test <- as.data.frame(validation.test)
    validation.std <- rundata[,2]
    i <- sapply(validation.test, is.character)
    categoricalvl <- validation.test[,i]
    for(m in 1:ncol(categoricalvl)) {
      categoricalvl[,m] <- recode_factor(categoricalvl[,m],"Y"="1","N"="0")
      categoricalvl[,m] <- as.character(categoricalvl[,m])
      categoricalvl[,m] <- factor(categoricalvl[,m],levels=c("0","1"))
    }
    categoricalvl <- categoricalvl %>% model.matrix(~.-1,.)
    categoricalvl <- categoricalvl[,-1]
    colnames(categoricalvl) <- c("Ink501","`100Coverage`1","FullDieCut1","DCAttachment1","`33MediumDW`1",
                                 "Machine41","Grade401","BoxWidth551","Grade321","CuttingDieUN1","`26Medium`1","StyleRSCL1",
                                 "EaPlateChange1","StyleRSC1","`40Medium`1","HHAttachment1","Grade2001","CuttingDieOthers1",
                                 "WRADW1","`33Medium`1", "WRAY1","ExtGlueTab1","`500BoxUnit`1","NUGHZ1")
    numericvl <- validation.test[,c(4)] %>% as.matrix()
    validation.tstx <- cbind(numericvl,categoricalvl)
    colnames(validation.tstx)[1] <- "PercInk"
    validation.tstx
    
    ptr.run.xgb <- predict(xgb_tune_Run4, newdata=validation.tstx)
    constant.Run <- 2.626647
    predictions.Run <- (10^ptr.run.xgb)-constant.Run 
    validation.final <- cbind(validation.std,validation.test,predictions.Run)
    colnames(validation.final)[1] <- "Standard Run Hours"
    validation.final$PredictedRunTime <- validation.final$predictions.Run + validation.final$`Standard Run Hours`
    validation.final <- validation.final %>% select(-predictions.Run)
    validation.final
  })
  
  observeEvent(input$button2, {
    output$xgbgraphrun <- renderPlot({
      if(is.null(input$file2)) return ()
      inFile2 <- input$file2
      rundata <- read_excel(inFile2$datapath)
      colnames(rundata) <- c("Standard Setup Hours","Standard Run Hours","EaInkChange","HHAttachment","FullPartDC","FullDieCut","Machine4","100Coverage","RockTennMRA","DCAttachment",
                             "BoxWidth55","Michelman40E","PercInk","CuttingDieUN","CuttingDieOthers","WRADW","33MediumDW","26MediumDW",
                             "ExtGlueTab","26Medium","Ink50","GlueJointGlue","StyleRSC","StyleRSCL","NumInk","EaPlateChange","Grade40","Grade32","Grade200",
                             "40Medium","33Medium","WRAY","500BoxUnit","NUGHZ")
      
      validation.test <-rundata %>% select(Ink50,`100Coverage`,FullDieCut,PercInk,DCAttachment,`33MediumDW`,
                                           Machine4,Grade40,BoxWidth55,Grade32,CuttingDieUN,`26Medium`,StyleRSCL,
                                           EaPlateChange,StyleRSC,`40Medium`,HHAttachment,Grade200,CuttingDieOthers,
                                           WRADW,`33Medium`, WRAY,ExtGlueTab,`500BoxUnit`,NUGHZ)
      validation.test <- as.data.frame(validation.test)
      validation.std <- rundata[,2]
      i <- sapply(validation.test, is.character)
      categoricalvl <- validation.test[,i]
      for(m in 1:ncol(categoricalvl)) {
        categoricalvl[,m] <- recode_factor(categoricalvl[,m],"Y"="1","N"="0")
        categoricalvl[,m] <- as.character(categoricalvl[,m])
        categoricalvl[,m] <- factor(categoricalvl[,m],levels=c("0","1"))
      }
      numericvl <- validation.test[,c(4)] 
      validation.tst <- cbind(numericvl,categoricalvl)
      colnames(validation.tst)[1] <- "PercInk"
      validation.tst
      
      xgbvalid3 <- xgb.DMatrix(data.matrix(validation.tst))
      colnames(xgbvalid3) <- c("PercInk","Ink501","`100Coverage`1","FullDieCut1","DCAttachment1","`33MediumDW`1",
                               "Machine41","Grade401","BoxWidth551","Grade321","CuttingDieUN1","`26Medium`1","StyleRSCL1",
                               "EaPlateChange1","StyleRSC1","`40Medium`1","HHAttachment1","Grade2001","CuttingDieOthers1",
                               "WRADW1","`33Medium`1", "WRAY1","ExtGlueTab1","`500BoxUnit`1","NUGHZ1")
      
      model <- xgb_tune_Run4$finalModel
      m <- showWaterfall(model, explainerRun, xgbvalid3, data.matrix(validation.tst),1, type = "regression") + 
        theme_bw()+
        theme(axis.text.x = element_text(angle = 45, hjust = 1))+
        ggtitle("Predictions for the 1st Observation")+
        xlab("Important Factors Used")+ylab("Values")+
        scale_y_continuous(limits=c(0,0.7),breaks=seq(-0.5,0.7,0.05))+
        theme(plot.title = element_text(size=10,face="bold",hjust = 0.5),panel.grid.major = element_blank(),
              axis.title=element_text(size=7,face="bold"))
      build <-ggplot_build(m)
      xlabl <- build$layout$panel_scales_x[[1]]$labels
      xlabl <- gsub("1 =","=",xlabl)
      xlabl[-grep("^Perc|2$",xlabl)] <- gsub("= 1","=0",xlabl[-grep("^Perc|2$",xlabl)])
      xlabl[-grep("^Perc|0$",xlabl)] <- gsub("= 2","=1",xlabl[-grep("^Perc|0$",xlabl)])
      m + scale_x_discrete(labels=xlabl)
    })
  })
  
  output$download2 <- downloadHandler(
    filename = function() {
      paste("RunPredictions-", Sys.Date(), ".xlsx", sep="")
    },
    content = function(file) {
      write_xlsx(downdata2(), file)
    }
  )
}

shinyApp (ui=ui,server=server)
