library(officer)
library(nortest)
library(flextable)
library(broom)
library(tictoc)

tictoc::tic("Global")
# Inference Report
my.OS <- Sys.info()['sysname']



main <- function(path, var.type="all"){

#Read .csv File

if(path == ""){
	path <- file.choose(new = FALSE)
}
file_in <- file(path,"r")
file_out <- file(paste(path,"-REPORT-DESCRIPTIVE-",var.type,".docx",sep = ""),open="wt", blocking = FALSE, encoding ="UTF-8")

#defining the document name
if(grepl("linux",tolower(my.OS))){
  my.time <- format(Sys.time(), "%Y-%m-%d_%H%M%S")
  aux.file.path <- paste(path,"-REPORT-DESCRIPTIVE-ALL.docx",sep = "")
  aux.file.report <- file_out
  dir.create("/tmp/data/", showWarnings = FALSE)
  aux.file.rdata <- paste("/tmp/data/", my.time,"-TEMP.DATA",sep="")
}else{
  my.time <- format(Sys.time(), "%Y-%m-%d_%H%M%S")
  aux.file.path <- paste(path,"-REPORT-DESCRIPTIVE-ALL.docx",sep = "")
  dir.create("/tmp/data/", showWarnings = FALSE)
  aux.file.rdata <- paste("/tmp/data/", my.time,"-TEMP.DATA",sep="")
}
my_doc <- officer::read_docx()




file1 <- file_in

if(is.null(file1)){return(NULL)}

  a<-readLines(path,n=2)
  b<- nchar(gsub(";", "", a[1]))
  c<- nchar(gsub(",", "", a[1]))
  d<- nchar(gsub("\t", "", a[1]))
  if(b<c&b<d){
    #r <- utils::read.table(path,sep=";",stringsAsFactors = FALSE,nrows=1,na.strings = c("NA",""))
    r <- utils::read.table(path,sep=";",stringsAsFactors = FALSE,nrows=1)
    
    SEP=";"
  }else if(c<b&c<d){
    #r <- utils::read.table(path,sep=",",stringsAsFactors = FALSE,nrows=1,na.strings = c("NA",""))
    r <- utils::read.table(path,sep=",",stringsAsFactors = FALSE,nrows=1)
    SEP=","
  }else{
    #r <- utils::read.table(path,sep="\t",stringsAsFactors = FALSE,nrows=1,na.strings = c("NA",""))
    r <- utils::read.table(path,sep="\t",stringsAsFactors = FALSE,nrows=1)
    SEP="\t"
  }
  
  
  #r1<-utils::read.table(path,sep=SEP,stringsAsFactors = FALSE,nrows=1,skip = 1,na.strings = c("NA",""))
  r1<-utils::read.table(path,sep=SEP,stringsAsFactors = FALSE,nrows=1,skip = 1)
  aciertos <- 0
  for(i in 1: ncol(r)){
    if(class(r[1,i])==class(r1[1,i])){
      aciertos <- aciertos+1
    }
  }
  
  HEADER=T
  if(aciertos==ncol(r)){
    HEADER=F
  }
  
data <- utils::read.table(path,sep=SEP,header=HEADER)

#VARS   
as.class <- FALSE

#variables combinations
data.vars.vector <- names(data)
  
#type of selected vars
data.vars.vector <- data.vars.vector[which(class(data)==var.type)]
  
if (length(data.vars.vector) < 2){
	warning(paste("Not enough variables of type", var.type, "to proceed with Inference Reporting. Suggestion: use var.type input equal to all", sep=" "))
	return()
}
  


barPlot <- function(ip){
  src1 <- tempfile(fileext = ".png")
  png(filename = src1, width = 5, height = 6, units = 'in', res = 300)
  barplot(ip)
  dev.off()
  src1 <<- src1
}

boxPlot <- function(ip){
  src2 <- tempfile(fileext = ".png")
  png(filename = src2, width = 5, height = 6, units = 'in', res = 300)
  boxplot(ip)
  dev.off()
  src2 <<- src2
}

scatterPlot <- function(ip){
  src3 <- tempfile(fileext = ".png")
  png(filename = src3, width = 5, height = 6, units = 'in', res = 300)
  scatter.smooth(ip)
  dev.off()
  src3 <<- src3
}



writing <- function(what="text",vars.name.pass=NULL,vars.val.pass=NULL,analysis=NULL){
  
  
  
  
  if(what=="text"){
    
    if(analysis=="mean"){
      TextInfomation1 <<-  paste(vars.name.pass, " mean is:", mean(data[,vars.name.pass]))
      dyn.command <- paste(dyn.command,'body_add_par(TextInfomation1, style = "Normal") %>%') 
      dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
      
    }
    
    if(analysis=="sd"){
      
      TextInfomation2 <<-  paste(vars.name.pass, " sd is:", sd(data[,vars.name.pass]))
      dyn.command <- paste(dyn.command,'body_add_par(TextInfomation2, style = "Normal") %>%') 
      dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
      
    }
    
    if(analysis=="median"){
      
      TextInfomation3 <<-  paste(vars.name.pass, " median is:", median(data[,vars.name.pass]))
      dyn.command <- paste(dyn.command,'body_add_par(TextInfomation3, style = "Normal") %>%') 
      dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
      
    }
    
  }
  
  if(what=="scatter"){
    src3 <<- scatterPlot(data[,paste(vars.name.pass)])
    dyn.command <<- paste(dyn.command,'body_add_par("Scatter Plot is: ", style = "Normal") %>%') # blank paragraph
    dyn.command <<- paste(dyn.command,'body_add_img(src = src3, width = 5, height = 6, style = "centered") %>%')
    dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
  }
  
  if(what=="bar"){
    src1 <<- barPlot(data[,paste(vars.name.pass)])
    dyn.command <<- paste(dyn.command,'body_add_par("Bar Plot is: ", style = "Normal") %>%') # blank paragraph
    dyn.command <<- paste(dyn.command,'body_add_img(src = src1, width = 5, height = 6, style = "centered") %>%')
    dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
  }
  
  if(what=="box"){
    src2 <<- boxPlot(data[,paste(vars.name.pass)])
    dyn.command <<- paste(dyn.command,'body_add_par("Box Plot is: ", style = "Normal") %>%') # blank paragraph
    dyn.command <<- paste(dyn.command,'body_add_img(src = src2, width = 5, height = 6, style = "centered") %>%')
    dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
  }
  
  if(what=="table"){
    TableToSave <<- head(data)
    dyn.command <<- paste(dyn.command,'body_add_par("A short example of your data to analyze: ", style = "Normal") %>%') # blank paragraph
    dyn.command <<- paste(dyn.command, 'body_add_table(TableToSave, style = "table_template") %>%') 
    dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
  }
  
  
}


for(var in 1:(ncol(data)-1)){
  
  dyn.command <<- 'my_doc <- my_doc %>%'
  
  if(var ==1){
    writing(what = "table",vars.name.pass = NULL,analysis = NULL)
  }
  
  writing(what = "text",vars.name.pass = paste(names(data)[var]),analysis = "mean")
  writing(what = "text",vars.name.pass = paste(names(data)[var]),analysis = "sd")
  writing(what = "text",vars.name.pass = paste(names(data)[var]),analysis = "median")
  writing(what = "bar",vars.name.pass = paste(names(data)[var]),analysis = NULL)
  writing(what = "box",vars.name.pass = paste(names(data)[var]),analysis = NULL)
  writing(what = "scatter",vars.name.pass = paste(names(data)[var]),analysis = NULL)
  
  dyn.command <<- noquote(stri_replace_last_fixed(dyn.command, " %>%", ""))
  
  #browser()
  
  eval(parse(text=dyn.command))
  
  print(my_doc, target = "Temp.docx")
  
}

  
   
  tictoc::tic("Analysis")
    
  save.image(file = paste(aux.file.rdata))
  warning(Sys.time(),"Passed Descriptive\n\n")
  tictoc::toc()

tictoc::tic("Writing DOC")
print(my_doc, target = paste(aux.file.path))
tictoc::toc()


save.image(file = paste(aux.file.rdata))
warning(Sys.time(),"Passed Final Print\n\n")
tictoc::toc()

if(grepl("lin",tolower(my.OS))){
  close(aux.file.report)
}
warning("Passed Final Print")

sink(type = "message")
sink()
#end and print
}

#' Descriptive Report in .docx file
#'
#' This R Package asks for a .csv file with data and returns a report (.docx) with Descriptive Report concerning all possible variables (i.e. columns).
#' 
#' @import officer nortest flextable broom tictoc stats utils 
#'
#' @param path (Optional) A character vector with the path to data file. If empty character string (""), interface will appear to choose file. 
#' @param var.type (Optional) The type of variables to perform analysis, with possible values: "all", "numeric", "integer", "double", "factor", "character". 
#' @return The output
#'   will be a document in the same folder of the data file.
#' @examples
#' \dontrun{
#' data(iris)
#' write.csv(iris,file="iriscsvfile.csv")
#' docdescriptR(path="iriscsvfile.csv")
#' }
#' @references
#' \insertAllCited{}
docdescriptR <- function(path="", var.type="all"){
	main(path,var.type="all")
}
