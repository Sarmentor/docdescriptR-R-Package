library(officer)
library(nortest)
library(flextable)
library(broom)
library(tictoc)
library(magrittr) # needs to be run every time you start R and want to use %>%
library(dplyr)
library(plotrix)
library(DescTools)

`%>%` <- magrittr::`%>%`

tictoc::tic("Global")
# Inference Report
my.OS <- Sys.info()['sysname']



main <- function(path, var.type="all"){

#Read .csv File

if(path == ""){
	path <- file.choose(new = FALSE)
}
file_in <- file(path,"r")
file_out <- file(paste(tempdir(),"docdescriptR-REPORT-DESCRIPTIVE-",var.type,"-",format(Sys.time(), "%a-%b-%d-%X-%Y"),".docx",sep = ""),open="wt", blocking = FALSE, encoding ="UTF-8")

#defining the document name
if(grepl("linux",tolower(my.OS))){
  my.time <- format(Sys.time(), "%Y-%m-%d_%H%M%S")
  aux.file.path <- paste(path,"-REPORT-DESCRIPTIVE-",var.type,".docx",sep = "")
  aux.file.report <- file_out
  dir.create("/tmp/data/", showWarnings = FALSE)
  aux.file.rdata <- paste("/tmp/data/", my.time,"-TEMP.DATA",sep="")
}else{
  my.time <- format(Sys.time(), "%Y-%m-%d_%H%M%S")
  aux.file.path <- paste(path,"-REPORT-DESCRIPTIVE-",var.type,".docx",sep = "")
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


#variables combinations
data.vars.vector <- names(data)

#type of selected vars
if(var.type != "all"){
	data.vars.vector <- data.vars.vector[which(sapply(data,class)==var.type)]	

	if (length(data.vars.vector) < 1){
		warning(paste("Not enough variables of type", var.type, "to proceed with Descriptive Reporting. Suggestion: use var.type input equal to all", sep=" "))
		return()
	}

}
  


barplot <- function(ip,vars.name.pass){
  src1 <- tempfile(fileext = ".png")
  grDevices::png(filename = src1, width = 5, height = 6, units = 'in', res = 300)
  graphics::barplot(table(stats::na.omit(ip)),col=1:10, cex.names = 0.70,cex.axis = 0.70,
          main=paste("Bar Plot of",vars.name.pass, "Variable\n (with sample sizes)",sep=" "))
  grDevices::dev.off()
  src1 <- src1
}

boxplot <- function(ip,vars.name.pass){
  src2 <- tempfile(fileext = ".png")
  grDevices::png(filename = src2, width = 5, height = 6, units = 'in', res = 300)
  graphics::boxplot(ip)
  graphics::title(paste("Box plot of",vars.name.pass,"variable",sep = " "))
  grDevices::dev.off()
  src2 <- src2
}

scatterPlot <- function(ip,vars.name.pass){
  src3 <- tempfile(fileext = ".png")
  grDevices::png(filename = src3, width = 5, height = 6, units = 'in', res = 300)
  stats::scatter.smooth(ip)
  grDevices::dev.off()
  src3 <- src3
}

piePlot <- function(ip,vars.name.pass){
  src4 <- tempfile(fileext = ".png")
  grDevices::png(filename = src4, width = 5, height = 6, units = 'in', res = 300)
  lbls <- paste(names(table(ip)), "\n", table(ip), sep="")
  plotrix::pie3D(table(ip), labels = lbls,explode=0.1,labelcex=0.8,
        main=paste("Pie Chart of",vars.name.pass, "Variable\n (with sample sizes)",sep=" "))
  grDevices::dev.off()
  src4 <- src4
}



histPlot <- function(ip,vars.name.pass){
  src0 <- tempfile(fileext = ".png")
  grDevices::png(filename = src0, width = 5, height = 6, units = 'in', res = 300)
  my.breaks <- stats::quantile(ip,seq(0,1,by=0.1),na.rm = TRUE)
  labels = 1:(length(my.breaks)-1)
  # graphics::hist(ip,breaks = "Sturges", probability=TRUE ,xlim = extendrange(my.breaks,r=range(my.breaks,na.rm=TRUE),f=0.05), ylim=extendrange(ip,r=range(ip,na.rm=TRUE),f=0.05),
  #      col=c("green","darkgreen","blue","darkblue","purple","pink"),
  #      main=paste("graphics::histogram of",vars.name.pass,"variable",sep=" "))
  graphics::hist(stats::na.omit(ip),breaks = "Sturges", probability=FALSE ,
       col=c("green","darkgreen","blue","darkblue","purple","pink"),
       main=paste("histogram of",vars.name.pass,"variable",sep=" "),xlab = paste(vars.name.pass))
  grDevices::dev.off()
  src0 <- src0
}



writing <- function(what="descriptive",vars.name.pass=NULL,vars.val.pass=NULL,analysis=NULL){
  
  
  
  
  if(what=="descriptive"){ 
    
    if(analysis=="header1"){
      
      val.header1 <-  paste("Descriptive Analysis for All Variables:", sep = "")
      officer::body_add_par(my_doc, val.header1, style = "heading 1",pos = "after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
    }
    
    if(analysis=="header2"){
      
      val.header2 <-  paste('Variable "', vars.name.pass, '" - This is the Descriptive Analysis:', sep = "")
      officer::body_add_par(my_doc, val.header2, style = "heading 2",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
    }
    
    if(analysis=="summ"){
      
      val.summ <- summary(stats::na.omit(data[,vars.name.pass]))
      
      my.min <- as.numeric(val.summ["Min."])
      my.max <- as.numeric(val.summ["Max."])
      my.mean <- as.numeric(val.summ["Mean"])
      my.median <- as.numeric(val.summ["Median"])
      my.1stQu <- as.numeric(val.summ["1st Qu."])
      my.3rdQu <- as.numeric(val.summ["3rd Qu."])
      my.sd <- stats::sd(stats::na.omit(data[,vars.name.pass]))
      
      #output text with values for variable summary
      val.summ.text <- paste('"', vars.name.pass, '" variable varies between ',my.min ,' and ',my.max,'. The mean is ',round(my.mean,3),'. The standard deviation is ',round(my.sd,3),', that is, on average, the "',vars.name.pass,'" varies about ',round(my.sd,3),' of the mean.',' The first and third quartiles are ',my.1stQu,' and ',my.3rdQu,' respectively. This means that 50% (half) of elements of the sample have "', vars.name.pass,'" between ', my.1stQu,' and ',my.3rdQu,'.', sep="")
      officer::body_add_par(my_doc, val.summ.text, style = "Normal",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
    }
    
    if(analysis=="freqs"){
      data[,vars.name.pass] <- as.factor(data[,vars.name.pass])
      val.freqs <- as.numeric(freq <- table(stats::na.omit(data[,vars.name.pass])))
      
      #browser()
      #TODO
      my.Abs.Freq <- val.freqs 
      my.Rel.Freq <- as.numeric(prop.table(table(stats::na.omit(data[,vars.name.pass])))[vars.name.pass])
      my.CumAbs.Freq <- cumsum(val.freqs) 
      my.CumRel.Freq <- cumsum(my.Rel.Freq)
      
      ncats <- length(unique(data[,vars.name.pass]))
      
      if(ncats > 10){
        local.info.text <- paste('NOTE: Variable "',vars.name.pass,'"',' has more than 10 categories. Thus, it should be advised to divide this variable data in intervals or groups.\n','NONETHELESS, this software will continue with the analysis, at your own risk.', sep="")
        #TODO - call shiny dialog windows with message, with only a OK button, no close possibility.
      }
      
      val.freqs.text <- paste('Categorical variables are qualitative variables and cannot be represented by a number. Categorical variables could be nominal and ordinal categorical variables. The difference is that, regarding their presentation, while nominal variables may be presented randomly or in the preferred order of the analyst, the ordinal variables must be presented in the order that is more easily understood (lowest to high, for example). In case of categorical variables, the analysis that can be done is the frequency of each category.')
      officer::body_add_par(my_doc, val.freqs.text, style = "Normal",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
      for(cat in 1:ncats){
        
        #text with substitution of values for each category
        val.freqs.text <- paste('Analyses show that there are ',val.freqs[cat],' counts of category "', levels(data[,vars.name.pass])[cat] ,'", corresponding to ',round(val.freqs[cat]/sum(val.freqs),3)*100,'%. In total, there are ', length(stats::na.omit(data[,paste(vars.name.pass)])) ,' elements in the study.',sep="")
        officer::body_add_par(my_doc, val.freqs.text, style = "Normal",pos="after") 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
      }
      
      # obtem a tabela com frequencia das variaveis
      # obtem o nome da variavel
      name = names(freq)[freq == max(freq)]
      
      #se a contagem for igual para a variavel não faz
      if(length(unique(table(data[,vars.name.pass])))>1){
        
        val.freqs.text <- paste('The mode (most common element) of this variable is ', name,', with ',max(freq),' counts.',sep="")
        officer::body_add_par(my_doc, val.freqs.text, style = "Normal",pos="after") 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
        
      }
      
   }
    
    if(analysis == "na"){
      trigger <- is.na(data[,vars.name.pass])
      
      #if(any(trigger)){
      val <- length(which(trigger==TRUE))
      #}
      
      #output text with variable information about NA count
      val.na.text <-  paste(
        "Results show that there are ", val,' missing values in variable "',vars.name.pass,'"',sep="")
      officer::body_add_par(my_doc, val.na.text, style = "Normal",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
    }
    
    if(analysis=="mean"){
      val.mean <-  paste(vars.name.pass, " mean is:", mean(data[,vars.name.pass]),sep="")
      officer::body_add_par(my_doc, val.mean, style = "Normal",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
    }
    
    if(analysis=="sd"){
      
      val.sd <-  paste(vars.name.pass, " sd (standard deviation) is:", stats::sd(data[,vars.name.pass]))
      officer::body_add_par(my_doc, val.sd, style = "Normal",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
    }
    
    if(analysis=="median"){
      
      val.median <-  paste(vars.name.pass, " median is:", stats::median(data[,vars.name.pass]),sep="")
      officer::body_add_par(my_doc, val.median, style = "Normal",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      
    }
    
  }
  
  if(what=="scatter"){
    val.scatter <- scatterPlot(data[,paste(vars.name.pass)],vars.name.pass)
    officer::body_add_par(my_doc, "Scatter Plot is: ", style = "Normal",pos="after") # blank paragraph
    officer::body_add_img(my_doc, src = val.scatter, width = 5, height = 6, style = "centered",pos="after")
    officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
  }
  
  if(what=="bar"){
    val.bar <- barplot(data[,paste(vars.name.pass)],vars.name.pass)
    officer::body_add_par(my_doc, "Bar Plot is: ", style = "Normal",pos="after") # blank paragraph
    
    #to avoid dense bar plots
    if(length(unique(data[,paste(vars.name.pass)]))<=30){
      officer::body_add_img(my_doc, src = val.bar, width = 5, height = 6, style = "centered",pos="after")
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
    }else{
      officer::body_add_par(my_doc, "Analysis not possible due to data constraints! (unique values > 30)", style = "Normal",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
    }
    
  
  }
  
  if(what=="box"){
    val.box <- boxplot(data[,paste(vars.name.pass)],vars.name.pass)
    officer::body_add_par(my_doc, "Box Plot is: ", style = "Normal",pos="after") # blank paragraph
    officer::body_add_img(my_doc, src = val.box, width = 5, height = 6, style = "centered",pos="after")
    officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
  }
  
  if(what=="pie"){
    val.pie <- piePlot(data[,paste(vars.name.pass)],vars.name.pass)
    officer::body_add_par(my_doc, "Pie Plot is: ", style = "Normal",pos="after") # blank paragraph
   
     #to avoid dense pie charts
    if(length(unique(data[,paste(vars.name.pass)]))<=30){
      officer::body_add_img(my_doc, src = val.pie, width = 5, height = 6, style = "centered",pos="after")
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
    }else{
      officer::body_add_par(my_doc, "Analysis not possible due to data constraints! (unique values > 30)", style = "Normal",pos="after") 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
    }
  }
  
  if(what=="hist"){
    if(class(data[,paste(vars.name.pass)])=="Date"){
      return({0})
    }else{
      val.hist <- histPlot(data[,paste(vars.name.pass)],vars.name.pass)
      officer::body_add_par(my_doc, "histogram Plot is: ", style = "Normal",pos="after") # blank paragraph
      officer::body_add_img(my_doc, src = val.hist, width = 5, height = 6, style = "centered",pos="after")
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
    }
  }
  
  if(what=="table"){
    
    if(analysis=="head"){
      #browser()
      
      new.column <- data.frame("..." = rep("...",nrow(data)), stringsAsFactors = TRUE)
      new.table <- as.data.frame(utils::head(cbind(data[,1:ifelse(ncol(data<5),2,3)],new.column,data[,(ncol(data)-ifelse(ncol(data<5),1,3)):ncol(data)])))
      
      val.head <-flextable::flextable(new.table) %>% flextable::theme_booktabs() %>% flextable::fontsize() %>% flextable::autofit()
      officer::body_add_par(my_doc, "A short example of your data to analyze: ", style = "Normal",pos="after") # blank paragraph
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      flextable::body_add_flextable(my_doc, val.head, pos="after",split = TRUE) 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
    }
    if(analysis=="freqs"){
      val.freqs <-flextable::flextable(DescTools::Freq(data[,vars.name.pass]))
      officer::body_add_par(my_doc, "The table for frequencies is the following: ", style = "Normal",pos="after") # blank paragraph
      flextable::body_add_flextable(my_doc, flextable::autofit(flextable::fontsize(val.freqs, size = 6, part = "all")), pos="after",split = TRUE) 
      officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
    }
  }
}

tictoc::tic("Analysis")

#Descriptive Analysis
   
    for(var in 1:length(data.vars.vector)){
      
      
      if(var==1){
        writing(what = "table",vars.name.pass = NULL,analysis = "head")
        writing(what = "descriptive",vars.name.pass = paste(data.vars.vector[var]),analysis = "header1")
      }
      
      writing(what = "descriptive",vars.name.pass = paste(data.vars.vector[var]),analysis = "header2")
      #writing(what = "descriptive",vars.name.pass = paste(data.vars.vector[var]),analysis = "mean")
      #writing(what = "descriptive",vars.name.pass = paste(data.vars.vector[var]),analysis = "sd")
      #writing(what = "descriptive",vars.name.pass = paste(data.vars.vector[var]),analysis = "median")
      writing(what = "descriptive",vars.name.pass = paste(data.vars.vector[var]),analysis = "na")
      
      if((class(data[,paste(data.vars.vector[var])])=="character" || class(data[,paste(data.vars.vector[var])])=="factor") && length(unique(data[,paste(data.vars.vector[var])]))<30){
        writing(what = "descriptive",vars.name.pass = paste(data.vars.vector[var]),analysis = "freqs")
        writing(what = "bar",vars.name.pass = paste(data.vars.vector[var]),analysis = NULL)  
        writing(what = "pie",vars.name.pass = paste(data.vars.vector[var]),analysis = NULL)
      }else if((class(data[,paste(data.vars.vector[var])])!="character" && class(data[,paste(data.vars.vector[var])])!="factor") && length(unique(data[,paste(data.vars.vector[var])]))<=10){
        writing(what = "descriptive",vars.name.pass = paste(data.vars.vector[var]),analysis = "freqs")
        writing(what = "bar",vars.name.pass = paste(data.vars.vector[var]),analysis = NULL)  
        writing(what = "pie",vars.name.pass = paste(data.vars.vector[var]),analysis = NULL)
      }else if((class(data[,paste(data.vars.vector[var])])!="character" && class(data[,paste(data.vars.vector[var])])!="factor") && length(unique(data[,paste(data.vars.vector[var])]))>10){
        writing(what = "descriptive",vars.name.pass = paste(data.vars.vector[var]),analysis = "summ")
        writing(what = "hist",vars.name.pass = paste(data.vars.vector[var]),analysis = NULL)
        #integer but categorical
        x <- stats::na.omit(data[,var])
        if(any(as.integer(x) != x)){
          writing(what = "scatter",vars.name.pass = paste(data.vars.vector[var]),analysis = NULL)
          writing(what = "box",vars.name.pass = paste(data.vars.vector[var]),analysis = NULL)
        }
      }else{
        officer::body_add_par(my_doc, "Analysis not possible due to data constraints!", style = "Normal",pos="after") 
        officer::body_add_par(my_doc, "", style = "Normal",pos="after") # blank paragraph
      }  
        
       
      }
       
  
save.image(file = paste(aux.file.rdata))
warning(Sys.time(),"Passed Descriptive\n\n")

tictoc::toc()

tictoc::tic("Writing DOC")
print(my_doc, target = paste(aux.file.path))
tictoc::toc()

tictoc::tic("Writing RData")
save.image(file = paste(aux.file.rdata))
warning(Sys.time(),"Passed Final Tasks. Closing\n\n")
tictoc::toc()

if(grepl("lin",tolower(my.OS))){
  close(aux.file.report)
}
warning("Passed Final Close")

sink(type = "message")
sink()
#end and print
}

#' Descriptive Report in .docx file
#'
#' This R Package asks for a .csv file with data and returns a report (.docx) with Descriptive Report concerning all possible variables (i.e. columns).
#' 
#' @import officer nortest flextable broom tictoc utils grDevices graphics magrittr dplyr 
#'
#' @importFrom magrittr "%>%"
#' 
#' @param path (Optional) A character vector with the path to data file. If empty character string (""), interface will appear to choose file. 
#' @param var.type (Optional) The type of variables to perform analysis, with possible values: "all", "numeric", "integer", "double", "factor", "character". 
#' @return The output
#'   will be a document in the temp folder (tempdir()).
#' @examples
#' \donttest{
#' data(iris)
#' write.csv(iris,file="iriscsvfile.csv")
#' docdescriptR(path="iriscsvfile.csv")
#' }
#' @references
#' \insertAllCited{}
#' @export
docdescriptR <- function(path="", var.type="all"){
	main(path,var.type=var.type)
}
