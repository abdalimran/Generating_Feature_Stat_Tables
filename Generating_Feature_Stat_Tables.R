#-------------------------------------- Setting Working Directory-------------------------------

setwd("C:\\Users\\abdal\\Desktop\\Data & Design Lab DU\\Economic Census2013-20180301T134357Z-001\\Economic Census2013")

#-------------------------------------- Loading Libraries --------------------------------------

library(openxlsx)
library(tableHTML)

#------------------------------------- Loadting Data File --------------------------------------
data <- read.xlsx(xlsxFile = "Dhaka.xlsx",
                  sheet = 1,
                  colNames = TRUE)

#------------------------------------- Transforming Data Types ---------------------------------
# make sure all the ordinal, interval and ratio variable contains proper numeric data

data$Q3_UNIT_HEAD_AGE <- as.integer(data$Q3_UNIT_HEAD_AGE)

#-------------------------------------- Defining function to generate stats -------------------
describe_stats <- function(feature, feature_type){
  
  getmodefunc <- function(v) {
    uniqv <- na.omit(unique(v))
    uniqv[which.max(tabulate(match(v, uniqv)))]
  }
  
  if(feature_type == "ratio" || feature_type == "interval")
  {
    get_count <- round(length(feature))
    get_na <- round(sum(is.na(feature)))
    get_sum <- round(sum(feature, na.rm=TRUE),2)
    get_mean <- round(mean(feature, na.rm=TRUE),2)
    get_median <- round(median(feature, na.rm=TRUE),2)
    get_mode <- round(getmodefunc(feature),2)
    get_mdad <- round(mad(feature, na.rm=TRUE),2)
    get_mead <-round(mad(feature, center= mean(feature, na.rm=TRUE), na.rm=TRUE),2)
    get_sd <- round(sd(feature, na.rm=TRUE),2)
    get_var <- round(var(feature, na.rm=TRUE),2)
    get_range <- round(range(feature, na.rm=TRUE),2)
    get_quantile <- round(quantile(feature, na.rm=TRUE),2)
    
    des_table <- data.frame(get_count,
                            get_na,
                            get_sum,
                            get_mean,
                            get_median,
                            get_mode,
                            get_mdad,
                            get_mead,
                            get_sd,
                            get_var,
                            get_range[1],
                            get_range[2],
                            get_quantile[[1]],
                            get_quantile[[2]],
                            get_quantile[[3]],
                            get_quantile[[4]],
                            get_quantile[[5]])
    
    names(des_table) <- c("Count",
                          "Missing Values",
                          "Sum",
                          "Mean",
                          "Median",
                          "Mode",
                          "Median Absolute Deviation",
                          "Mean Absolute Deviation",
                          "Standard Deviation",
                          "Variance",
                          "Minimum",
                          "Maximum",
                          "0%",
                          "25%",
                          "50%",
                          "75%",
                          "100%")
    
    return(t(des_table))
  }else if(feature_type == "ordinal")
  {
    get_count <- round(length(feature))
    get_na <- round(sum(is.na(feature)))
    get_median <- round(median(feature),2)
    get_mode <- round(getmodefunc(feature),2)
    get_range <- round(range(feature),2)
    get_quantile <- round(quantile(feature, na.rm=TRUE),2)
    
    des_table <- data.frame(get_count,
                            get_na,
                            get_median,
                            get_mode,
                            get_range[1],
                            get_range[2],
                            get_quantile[[1]],
                            get_quantile[[2]],
                            get_quantile[[3]],
                            get_quantile[[4]],
                            get_quantile[[5]])
    
    names(des_table) <- c("Count",
                          "Missing Values",
                          "Median",
                          "Mode",
                          "Minimum",
                          "Maximum",
                          "0%",
                          "25%",
                          "50%",
                          "75%",
                          "100%")
    
    return(t(des_table))
  }else
  {
    get_count <- round(length(feature))
    get_na <- round(sum(is.na(feature)))
    des_table <- data.frame(get_count,
                            get_na)
    names(des_table) <- c("Count",
                          "Missing Values")
    
    return(t(des_table))
  }
  
}


#------------------------------- Loading feature list with variable type -----------------------

data2 <- read.xlsx(xlsxFile = "BBS Economic Census 2013 Dhaka.xlsx",
                  sheet = 1,
                  colNames = TRUE)

data2$Type<-tolower(data2$Type)



#------------------------------- Generating HTML Tables -------------------------------------

desc_tables <- NULL

for (row in 1:nrow(data2)) {
  
  desc_tables <- paste(desc_tables,
                      tableHTML(describe_stats(feature = data[[data2[row,1]]], feature_type = data2[row,3]),
                                headers = data2[row,1],
                                caption = paste(data2[row,2]," (",data2[row,3],")",sep = ''),
                                widths = c("200px","600px"),
                                theme = 'rshiny-blue') %>%
                        add_css_caption(css = list(c('color', 'font-size', 'background-color','text-align','padding-left'), c('white', '18px', '#728493','left','20px'))),
                      "<br>"
  )
  
}


#------------------------------- Exportin Report ------------------------------------------

write(desc_tables, file = "BBS_Features_Desc.html")
system("pandoc -V geometry:letterpaper -V geometry:margin=.5in -s BBS_Features_Desc.html -o BBS_Features_Desc.pdf")
