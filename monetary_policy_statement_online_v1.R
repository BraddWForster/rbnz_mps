##    Programme:  R
##
##    Objective:  Make a chart of previous Monetary Policy Statements
##
##    Plan of  : 
##    Attack   :  
##                1. Set up working directories
##                2. Import the data
##                3. Plot the Monetary Policy Statements
##
##    Author   :  Bradd Forster
##
##  Clear the decks and load up some functionality
    ##
    rm(list=ls(all=TRUE))

##  Core libraries
    ##
    library(dplyr)      ##  Using tibbles
    library(tidyr)      ##  Tidying data
    library(stringr)    ##  Dealing with strings
    library(lubridate)  ##  Dealing with dates
    library(readxl)      ##  Loading excel files
    library(writexl)    ##  Uploading excel files
    library(ggplot2)    ##  Making graphs
    library(plotly)     ## Interactive graph
    
################################################################################
## 1.                       Set up working directories                        ##                       
################################################################################
    
##  Set up paths for working directories
    ##
    userid <- 
    outputs_fn <- 
    raw_fn <- 
    scripts_fn <- 
    
##  Set the working directory
    ##
    setwd(paste0(raw_fn))

    
################################################################################
## 2.                           Import the data                               ##                       
################################################################################

##  Load in monetary policy statements
    ##
    ##  Locate the files that you want and same them
        ##
        file_names <- list.files(pattern = "^mps.*\\-data.xlsx")
        
    ##  Read the files into R
        ##
        ##  Make a blank list
            ##
            data_list <- list()
        
        ##  Load each file (identified in the locate step above) into a list element
            ##
            for(i in file_names){
              data_list[[i]] <- read_xlsx(path = i, sheet = 3)
              
              identifier_idx <- which(data_list[[i]][,1] == "Identifier")
              
              ocr_idx <- which(data_list[[i]][identifier_idx,] == "ocr")
              
              if (length(identifier_idx) > 0 && length(ocr_idx) > 0) {
                data_list[[i]] <- data_list[[i]][(identifier_idx + 1):nrow(data_list[[i]]), c(1, ocr_idx)]
              }
              
              colnames(data_list[[i]]) <- c("Date", "Ocr") 
              
              data_list[[i]]$Date <- as.Date(as.numeric(data_list[[i]]$Date) - 2, origin = "1899-12-30")
              
              data_list[[i]]$Mps <- as.Date(paste0("01",str_sub(i, start = 4, end = 8)), format = "%d%b%y")
            
              data_list[[i]]$Ocr <- as.numeric(data_list[[i]]$Ocr)
                
            }
        
            
            Data <- bind_rows(data_list)
        
            Data <- Data %>%
                      mutate(Forecast = ifelse(Date>Mps, TRUE, FALSE))
            
            Data <- Data %>%
              mutate(Forecast = factor(Forecast, levels = c(FALSE, TRUE), labels = c("Actual", "Forecast")),
                     Mps = factor(Mps)
                     )

            
################################################################################
## 3.                    Plot the Monetary Policy Statements                  ##
################################################################################
            
forecasts <- ggplot(data = Data %>% filter(Date>as.Date("2015-01-01")), aes(x= Date, y = Ocr, colour = Mps, linetype = Forecast)) +
  geom_line() +
  labs(y = "RBNZ Official Cash Rate (%)", x = "Date", colour = "MPS",linetype = "Type") +
  scale_linetype_manual(values = c("solid", "longdash"))
            
theme_mywebsite <- function(base_size = 9, base_family = "Montserrat") {
  theme_minimal(base_size = base_size, base_family = base_family) +
    theme(
      panel.grid.major = element_line(color = "#5A636A", size = 0.2),
      panel.grid.minor = element_blank(),
      axis.title = element_text(size = base_size + 2, face = "bold", color = "#AFB3B7"),
      axis.text = element_text(size = base_size, color = "#AFB3B7"),
      legend.title = element_text(size = base_size + 1, face = "bold", color = "#AFB3B7"),
      legend.text = element_text(size = base_size, color = "#AFB3B7"),
      plot.title = element_text(size = base_size + 4, face = "bold", hjust = 0.5, color = "#AFB3B7"),
      plot.subtitle = element_text(size = base_size + 1, color = "#AFB3B7", hjust = 0.5),
      plot.background = element_rect(fill = "#132E35", color = NA),
      legend.position = "bottom",
      legend.background = element_blank(),
      legend.key = element_blank()
    )
}
            
forecasts <- forecasts + theme_mywebsite()            
            
setwd(outputs_fn)

ggsave(filename = "forecasts_plot.jpg",
       plot = forecasts,
       width = 10, 
       height = 6, 
       dpi = 300)  

