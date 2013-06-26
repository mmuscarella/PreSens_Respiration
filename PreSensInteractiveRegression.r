### Analysis of PreSens Respiration Data
### Written By: Mario Muscarella
### Last Update: 20 June 2013

## This analysis uses the R package - rpanel which does interactive regression
## The goal is to use the interactive view to pick the area to analyze
## rpanel - Interactive Regression 
## Multiple trend chart version; Annual anomaly data

setwd("C:/Users/Mario Muscarella/Documents/PhD Files/PreSens/Analysis Script/")
rm(list=ls())
require("rpanel")||install.packages("rpanel")
require("gWidgets")||install.packages("gWidgets")
require("xlsx")||install.packages("xlsx")
#require("gWidgetsRGtk2")||install.packages("gWidgetsRGtk2")
#require("RGtk2")||install.packages("RGtk2")
#require("lattice")||install.packages("lattice")
#require("grid")||install.packages("grid")
#require("cairoDevice")||install.packages("cairoDevice")
#source("PreSensInteractiveRegressionFunctions.r")

#options("guiToolkit"="RGtk2")     
options(digits=6)
script <- "Mario's Bad Ass M-F'ing Interactive Regression Analysis"
user.name <- "M. Muscarella"

## Import Data #################################################################
infile <- gfile("Open PreSens Data...",type="open")
data.in <- read.xlsx(infile, "Sheet4", header=T, startRow=11) 
head(data.in)
colnames(data.in) <- c("Date", "Time", "A1", "A2", "A3", "A4", "A5", "A6", "B1",
   "B2", "B3", "B4", "B5", "B6", "C1", "C2", "C3", "C4", "C5", "C6", "D1", "D2",
   "D3", "D4", "D5", "D6", "Temp", "Error")
data.in$Date <- strptime(data.in[,1], format="%d.%m.%y %H:%M:%S")
data.in[data.in == "No Sensor"] <- NA
data.in$Time <- data.in$Time/3600 # Convert sec to Hrs
data.in[3:26] <- data.in[3:26] * 1000/32 # Convert mg O2/L to uM O2

## File for output
outfile <- gfile("File to Save Output",type="save", initialfilename="output.txt")
titles <- c("Sample", "Start", "End", "Rate (然 O2 Hr-1)", "R2", "P-value")
write.table(as.matrix(t(titles)), file=outfile, append=T, row.names=F, col.names=F, sep=",", quote=FALSE)

## Select Samples ##############################################################
samples <- as.factor(colnames(data.in)[3:26])

## Create Plotting Window
par(las=1)
 par(fig=c(0,1,0, 1), new = F)
 par(ps=9); par(cex.axis=c(0.9)); par(cex.lab=c(0.9)); par(oma=c(3,1,1,0.5)); par(mar=c(4,4,2,1))

## attach Data
attach(data.in)
   
### rpanel function ################
draw <- function(panel) {	
  if (panel$end > panel$start)
    {start <- panel$start
    end <- panel$end}
  else
    {start <- min(Time)
    end <- max(Time)}
                                                            
  data.in$samp <- panel$samp
  
# Text placement points    
  xaxis_pt <- min(Time) + 0.25*(max(Time)-min(Time)) 
  yaxis_pt <- 200
  
# Make data subset based on start & end yrs
  sub <- subset(data.in, Time >= start)
  sub <- subset(sub, Time <= end)
  
# Linear trend line lm & coefs / stats
  trend <- lm(sub$samp ~ sub$Time)
  a <- as.numeric(coef(trend)[1]);  b <- as.numeric(coef(trend)[2])
  r2 <- round(summary(trend)$r.squared, 3)
  p <- round(anova(trend)$'Pr(>F)'[1], 4)
  p <- ifelse (p == 0, "<0.001", p)
  start.2 <- signif(start, digits = 3)
  end.2 <- signif(end, digits = 3)
  rate <- signif(-b,3)
                     
# Make Basic plot
  plot(panel$samp ~ Time, type = "b", col = "darkgrey", xlab = "Time (Hrs)",
    ylab = expression(paste("Oxygen Concentration (然 O "[2],")")),
    par(bty="n"),xlim=c(0,max(Time)+10),ylim=c(150, 400), 
    xaxs = "i", yaxs = "i", axes = FALSE, cex.main = 0.95,
    main = "Interactive Regression of PreSens Respiration Data")
  axis(1, col = "grey"); axis(2, col = "grey")
  
# Calc vals for start-end regression line & add line  
  x_vals = c(panel$start, panel$end)
  y_vals = c(a+b*panel$start, a+b*panel$end)
	lines(x_vals, y_vals, col= "red")
  points(sub$Time, sub$samp, pch = 19, col = "red", type = "p")
  text(xaxis_pt, yaxis_pt, paste("Period: ", start.2, " to " , end.2, "Hrs"))  
  text(xaxis_pt, yaxis_pt - 10, bquote(Rate == .(rate) ~ 然 ~ O[2] ~ Hr^-1), cex=1) #expression(paste("Rate = ", rate, "然 O ", "Hr"^"-1")),cex=1)
  text(xaxis_pt, yaxis_pt - 20, bquote(R^2 == .(r2)), cex=1) #expression(paste("R"^"2"," = " , r2)), cex=1)
  text(xaxis_pt, yaxis_pt - 30, paste("P-value = ", p), cex=1)
  
### Outer Margin Annotation
	my_date <- format(Sys.time(), "%m/%d/%y")
	mtext(script, side = 1, line = .75, cex=0.8, outer = T, adj = 0)
  mtext(user.name, side = 1, line = 0.75, cex=0.8, outer=T, adj = 0.75)
	mtext(my_date, side = 1, line =.75, cex = 0.8, outer = T, adj = 1)
	data.out <- list(Start=start, End=end, Rate=rate, R2=r2, Pvalue=p)
	print(data.out)
	return(data.out)
	panel
  }
   
collect.data <- function(panel) {
  data.sample <- readline("What is the name of this sample?")
  data.start <- data.out$Start
  data.end <- data.out$End
  data.rate <- data.out$Rate
  data.r2 <- data.out$r2
  data.p <- data.out$Pvalue
  data.sample <- panel$samp
  data.out <- c(data.sample, data.start, data.end, data.rate, data.r2, data.p)
  write.table(as.matrix(t(data.out)), file=outfile, append=T, row.names=F, col.names=F, sep=",", quote=FALSE)
  panel
  } 
    
end.session <- function(panel) {
  dev.off(2)
  print.noquote("Good-Bye: Computer Will Now Self-Destruct")
  panel
  }
                            
## rpanel controls - enter start and end yrs for portion of full data set to be used
rpplot <- rp.control(title="Interactive Regression", start=0, end = max(Time), initval = samples[1])
rp.listbox(rpplot, variable = samp, vals = "samples", labels = samples, action = draw)
rp.slider(rpplot, start, action = draw, from = 0, to =  max(Time))
rp.slider(rpplot, end, action = draw, from = 0, to =  max(Time))
rp.doublebutton(rpplot, var = start, step = 1, title = "start fine adjustment", action = draw)
rp.doublebutton(rpplot, var = end, step = 1, title = "end fine adjustment", action = draw)
rp.button(rpplot, title = "save", action = collect.data)
rp.button(rpplot, title = "quit", action = end.session, quitbutton=T)
