## Survey User 2017
## User Alvaro A.

### Test Questions ##
pkgTest <- function(x)
{
  if (!require(x,character.only = TRUE))
  {
    install.packages(x,dep=TRUE)
    if(!require(x,character.only = TRUE)) stop("Package not found")
  }
}


pkgTest("svDialogs")

library(svDialogs)

## Ask something...

### Areas 
# - 1) Advanced Analytics
  # Project Assigned (Career Value, Enjoyment, )
  # Workload (Workload, Time of internship, Number of Projects)
  # Difficulty of project (Academic, Software, Business)
  # Advanced Analytics learnings (of Pfizer, PCH, within AA)
  # Enjoyment of project (Colleagues, Environment, Support Project)
  # What skills are valued in order to do this project optimally? (DataBase Administration, Business, Academic, Software)
  # 

# - 2) PCH Internship Activities
  # Value PCH Internship Program
  # PCH Activities (Speeches, outSession, Volunteer, oTher activities)
  # Workload of PCH activities 


# Q <- c("Q1",
#        "Q2",
#        "Q3",
#        "Q4",
#        "Q5",
#        "Q6",
#        "Q7",
#        "Q8",
#        "Q9",
#        "Q10",
#        "Overall value of Internship Activities",
#        "Overall enjoyment of the project",
#        "Overall difficulty of the project",
#        "Overall level of Satisfaction with the program"
# )
# 
# saveRDS(Q,"Q.rds")
u <- url("https://github.com/alvaroaguado3/Survey/blob/master/Q.rds?raw=true")
readRDS(gzcon(u)) -> Q

n <- rep(NA,length(Q))
for( i in 1:length(Q)){
  options(warn = -1)
  n[[i]] <- as.integer(dlgInput( paste(Q[i]," (0-10):  "))$res)
  while(is.na(n[[i]]) == TRUE){
    dlgMessage("Answer must be numeric from 0-10, Please answer the question again.")
    n[[i]] <- as.integer(dlgInput( paste(Q[i]," (0-10):  "))$res)
    next
  }
  options(warn = 0)
}


pkgTest("RDCOMClient")
library(RDCOMClient)
## init com api
OutApp <- COMCreate("Outlook.Application")
## create an email 
outMail = OutApp$CreateItem(0)
## configure  email parameter 
outMail[["To"]] = "alvaro.aguado@pfizer.com"
outMail[["subject"]] = "responses"
outMail[["body"]] = paste(n)
## send it                     
outMail$Send()
