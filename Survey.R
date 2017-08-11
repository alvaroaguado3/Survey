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

require(svDialogs)


u <- url("https://github.com/alvaroaguado3/Survey/blob/master/Q.rds?raw=true")
readRDS(gzcon(u)) -> Q


A <- function(Q){
  n <- rep(NA,1)
  type <- " (0-10)[NUM]:  "
  r <- 0
  options(warn = -1)
  while(r == 0){
  res <- dlgInput( paste(type,Q[1]))$res
    if(nchar(res)==0){
      res2 <- dlgMessage("Your response was empty, is that correct?", "yesno")$res
      if(res2 == "yes"){
        break
        }else{
      dlgMessage("You chose you redo the question. Please answer the question again.")
      next
        }
    }
  n[[1]] <- as.integer(res)
  if(is.na(n[[1]]) == TRUE | n[[1]] > 10 | n[[1]] < 0){
    dlgMessage("Answer must be numeric from 0-10, Please answer the question again.")
  next
  }else{
    break
  }
  }
  options(warn = 0)
  return(n)
}

B <- function(Q){
  n <- rep(NA,1)
  type <- " [TEXT]:  "
  r <- 0
  options(warn = -1)
  while(r == 0){
    res <- dlgInput( paste(type,Q[1]))$res
    if(nchar(res) < 6){
      res2 <- dlgMessage("Your response was empty or is very short, is that correct?", "yesno")$res
      if(res2 == "yes"){
        if(nchar(res)==0){
          res<-NA
          }else{
            n[[1]] <- as.character(res)  
          }
        break
      }else{
        dlgMessage("You chose you redo the question. Please answer the question again.")
        next
      }
    }else{
    n[[1]] <- as.character(res)
    break
    }
  }
  options(warn = 0)
  return(n)
}


C <- function(Q,c){
  n <- rep(NA,1)
  I <- paste0("1-",c)
  type <- paste0(" [",I," CHOICE]:  ")
  r <- 0
  options(warn = -1)
  while(r == 0){
    res <- dlgInput( paste(type,Q[1]))$res
    if(nchar(res)==0){
      res2 <- dlgMessage("Your response was empty, is that correct?", "yesno")$res
      if(res2 == "yes"){
        break
      }else{
        dlgMessage("You chose you redo the question. Please answer the question again.")
        next
      }
    }
    n[[1]] <- as.integer(res)
    if(is.na(n[[1]]) == TRUE | !(n[[1]] %in% (1:c)) ){
      dlgMessage(paste0("Answer must be a single number of the list [",I,"] Please answer the question again."))
      next
    }else{
      break
    }
  }
  options(warn = 0)
  return(n)
}

send <- function(responses){
  pkgTest("RDCOMClient")
  require(RDCOMClient)
  ## init com api
  OutApp <- COMCreate("Outlook.Application")
  ## create an email 
  outMail = OutApp$CreateItem(0)
  ## configure  email parameter 
  outMail[["To"]] = "alvaro.aguado@pfizer.com"
  outMail[["subject"]] = "responses"
  outMail[["body"]] = paste0(responses,collapse = "|")
  ## send it                     
  outMail$Send()
}

# D <- function(Q){
#   n <- rep(NA,length(Q))
#   for( i in 1:length(Q)){
#     options(warn = -1)
#     n[[i]] <- as.character(dlgInput( paste(Q[i]," (free Text):  "))$res)
#   }
# }

resp <- list()
Survey <- function(Q){
repeat{
  for(i in 1:nrow(Q)){
    if(i == 1){
      dlgMessage("Thank you for agreeing to take part in this survey measuring your experience here at Pfizer Consumer Healthcare")
      dlgMessage("This survey should only take 5 mins to complete. ")
      dlgMessage("Please answer with a number when the question has the code [NUM] (Only digits accepted)")
      dlgMessage("Please answer with text when the question has the code [TEXT]")
      dlgMessage("If you prefer not to answer some of the questions simply press Enter, to leave the question blank")
      
      dlgMessage("First tell us a little about yourself")
    }else{
    if(Q$block[i] != Q$block[i-1] & Q$block[i] == 2){
      dlgMessage("These Questions are about Advanced Analytics and the Project")
    }else if(Q$block[i] != Q$block[i-1] & Q$block[i] == 3){
      dlgMessage("These Questions are about Pfizer Environment")
    }else if(Q$block[i] != Q$block[i-1] & Q$block[i] == 4){
      dlgMessage("Finally Overall Experience of the Internship")
    }
    }

    if(Q[i,"type"]=="t"){
      resp[[i]] <- B(Q$Q[i])
    }else if(Q[i,"type"]=="n"){
      resp[[i]] <- A(Q$Q[i])
    }else if(Q[i,"type"]=="c"){
      resp[[i]] <- C(Q$Q[i],Q$Option[i])
    }
    
  }
  if(length(resp) == nrow(Q))
    dlgMessage("Congratulations you finished the Survey.")
    res2 <- dlgMessage("Do you want to transmit your responses?", "yesno")$res
    if(res2=="yes"){
      out <- send(unlist(resp))
      if(out == TRUE){
        dlgMessage("Your Responses were sent successfully. THANK YOU!")
        break}
    }else{
      res3 <- dlgMessage("Do you want to redo the Survey?", "yesno")$res
      if(res3=="yes"){
        res4 <- dlgMessage("Are you sure? All your previous answers will be lost", "yesno")$res
        if(res3=="yes"){
          resp <- list()
          next
        }else{
          out <- send(unlist(resp))
          if(out == TRUE){
            dlgMessage("Your Responses were sent successfully. THANK YOU!")
            break}
        }
      }else{
        out <- send(unlist(resp))
        if(out == TRUE){
          dlgMessage("Your Responses were sent successfully. THANK YOU!")
          break}
      }
      
    }
  }
}

Survey(Q)

