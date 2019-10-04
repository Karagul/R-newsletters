# Update all excel files that use Newsletters
#
# EO 11.04.2019
# 1. Read Newsletters from the excel file from Campaign Manager. OK
# 2. Calculate  _rate . OK
# 3. Date manipulations . OK
# 4. Read  the coded_newsletters and Match these with the Newsletters ok
# 5. Save a master Newsletters_list. OK.
# 6. update the sheets in various excel KPI workbooks.  newsletter_coded must include all cases.
#    newsletter_list should be without optins and tests and such.
# Developed at Eeros Thinkpad. Remember to check all file locations.
#
# EO 1.5.2019 moved from thinkpad to Eeros ThinkCentre

"remember to download 
https://veggienews.vebu.de/campaigns/reports/compareCampaigns.aspx

it can be left in the downloads folder. this script will move it proper location.

"

## House cleaning
# Remove all objects (= dropp all data and variables in memory)
detach(mydata)
rm(list = ls())

#### Create The Work Space 
## setwd("C:\\Users\\eero\\Google Drive\\ProVeg\\data")
setwd("//vebufiler01/KPIImpact/KPI-Impact/database/data")
# R uses \ as an escape character. Therefore you either need two \\ or / in a folder name.

# The most important data manipulation package that gives SQL like functions
# are installed with install_Eeros_favorite_packages.R
library(dplyr)
library(janitor)
library(readr)
library(stringr)
library(tidyr)
library(openxlsx)



# 1. READ the file

# Tab delimited text file
# mydata <- clean_names(read_delim("CompareCampaignsData.txt", "\t", escape_double = FALSE, trim_ws = TRUE))


file.copy("C:/Users/eolli/Downloads/comparecampaignsdata.csv", 
          "//vebufiler01/KPIImpact/KPI-Impact/database/data/comparecampaignsdata.csv",
          overwrite = TRUE,
          copy.mode = TRUE, 
          copy.date = TRUE)

# Comma Separeted Values
mydata <- clean_names(read_delim("CompareCampaignsData.csv", ",", escape_double = TRUE, trim_ws = TRUE), case="snake")
mydata <- rename(mydata, complaints = "spam_complaints")
mydata <- rename(mydata, date_sent_text = sent_date)



# 2. Calculate the rates

mydata$open_rate <- (opens/(total_recipients - bounces))
mydata$click_rate <- (clicks/opens)
mydata$bounce_rate <- (bounces/total_recipients)
mydata$unsubscribe_rate <- (unsubscriptions/(total_recipients - bounces))
mydata$complaint_rate <- (complaints/(total_recipients - bounces))
mydata <- remove_empty(mydata)
# these control the cell format type when the data is saved to excel.
class(mydata$open_rate) <- "percentage"
class(mydata$click_rate) <- "percentage"
class(mydata$unsubscribe_rate) <- "percentage"
class(mydata$complaint_rate) <- "percentage"

# Create number and format as percent rounded to one decimal place
# Dropped because Excel did not get the format type right, (good looking strings)
# mydata$open_percent     <- paste(round((opens/(total_recipients - bounces))*100,digits=1),"%",sep="")


View(mydata)



# 3. Date Manipulations

library(lubridate)

# calculate a real dates out of the string
mydata$sent_date <- as_date(parse_date_time(date_sent_text, "d!b!Y!", tz = "CET" ))
# check
class(mydata$sent_date)
min(mydata$sent_date)
max(mydata$sent_date)


# find the day of the week. 
mydata$sent_weekday <- wday(mydata$sent_date, label = TRUE, week_start = 1, locale = "UK")
View(mydata)
# why do I need the next line?
mydata <- rename(mydata, sent_weekday.x = "sent_weekday")

mydata <- select(mydata, campaign_name, sent_weekday, sent_date, date_sent_text, 
       total_recipients, 
       open_rate, opens, 
       click_rate, clicks, 
       bounce_rate, bounces, 
       unsubscribe_rate, unsubscriptions,
       complaint_rate , complaints)



mydata$type_by_script <- NA
mydata$newsletter_name_by_script <- NA


mydata <- mydata %>% 
   mutate(
      type_by_script = case_when(
         str_detect(str_to_lower(campaign_name), "opt-in"           )     ~ "Opt-in"     ,
         str_detect(str_to_lower(campaign_name), "fundraising"      )     ~ "Fundraising",
         str_detect(str_to_lower(campaign_name),"newsletter"        )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"vebu"              )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"proveg de"         )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"proveg \\d\\d"     )     ~ "Newsletter" ,  
         str_detect(str_to_lower(campaign_name),"testcommunity"     )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"gastro"            )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"food services"     )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"v-label"           )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"startup"           )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"incubator"         )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"aktiven-news"      )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"jobnews"           )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"food industry ger" )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"food industry int" )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"food industry"     )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"regionalgruppen"   )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"food services"     )     ~ "Newsletter" ,
         str_detect(str_to_lower(campaign_name),"test"              )     ~ "test"       ,
         str_detect(str_to_lower(campaign_name),"draft"             )     ~ "test"       ,            
         is.na(type_by_script)                                            ~ "missing"    
      )
   )

mydata <- mydata %>% 
   mutate(
      newsletter_name_by_script = case_when(
         str_detect(str_to_lower(campaign_name), "fundraising"      )     ~ "Fundraising"      ,
         str_detect(str_to_lower(campaign_name),"vebu"              )     ~ "VEBU"             ,
         str_detect(str_to_lower(campaign_name),"proveg de"         )     ~ "ProVeg DEU"       ,
         str_detect(str_to_lower(campaign_name),"proveg \\d\\d"     )     ~ "ProVeg DEU"       ,  
         str_detect(str_to_lower(campaign_name),"testcommunity"     )     ~ "Testcommunity"    ,
         str_detect(str_to_lower(campaign_name),"gastro"            )     ~ "Gastro"           ,
         str_detect(str_to_lower(campaign_name),"food services"     )     ~ "Food Services"    ,
         str_detect(str_to_lower(campaign_name),"v-label"           )     ~ "V-Label"          ,
         str_detect(str_to_lower(campaign_name),"startup"           )     ~ "Startup"          ,
         str_detect(str_to_lower(campaign_name),"incubator"         )     ~ "Incubator"        ,
         str_detect(str_to_lower(campaign_name),"aktiven-news"      )     ~ "Aktiven-News"     ,
         str_detect(str_to_lower(campaign_name),"jobnews"           )     ~ "Jobnews"          ,
         str_detect(str_to_lower(campaign_name),"food industry ger" )     ~ "Food Industry GER",
         str_detect(str_to_lower(campaign_name),"food industry int" )     ~ "Food Industry INT",
         str_detect(str_to_lower(campaign_name),"food industry"     )     ~ "Food Industry"    ,
         str_detect(str_to_lower(campaign_name),"regionalgruppen"   )     ~ "Regionalgruppen"  ,
         str_detect(str_to_lower(campaign_name),"food services"     )     ~ "Food Services"    ,
         is.na(newsletter_name_by_script)                                 ~ "missing"
      )
   )





#### COMBINE FILES.

library(readxl)
mydata_coded <- clean_names(read_excel("Newsletters_coded_2017-2019.xlsx", 
                             sheet = "Newsletter_coded", range = cell_cols(c("A:C","E")), 
                             col_types = c("text","text", "text","text","date")), case="snake")


#make sent_date a date.
mydata_coded$sent_date <-as_date(mydata_coded$sent_date)
#add the codes to the newsletter_list
mydata_coded <- left_join(mydata, mydata_coded, by = c("campaign_name", "sent_date"))
View(mydata_coded)

# Add time_variable dummy just to get the columns right 
mydata_coded$sent_000 <- hm("00:01")
#class(mydata_coded$sent_time)
#mydata_coded$sent_time


# Keep the manually made changes and just juse scripting to add the new missing ones.
mydata_coded <- mydata_coded %>%
   mutate(newsletter_name = case_when(
            !is.na(newsletter_name) ~ newsletter_name,
            is.na(newsletter_name)  ~ newsletter_name_by_script)
          )
mydata_coded <- mydata_coded %>%
   mutate(type = case_when(
      !is.na(type) ~ type,
       is.na(type)  ~ type_by_script))




# I want to replace missing values of type, when campaign_name contains opt-in.
# but this results in a shorter tibble. I want to keep the rest of the dataframe.
#mydata_coded %>%
#   dplyr::filter(is.na(type)) %>%
#   dplyr::filter(stringr::str_detect(str_to_lower(campaign_name), "opt-in")) %>% 
#   dplyr::mutate(type = "Opt-In")


#newsletter_name_list <- c("newsletter" , "vebu" , "proveg de" , "proveg-testcommunity" , "gastro" , "food services" , 
#   "v-label" ,"startup" ,"incubator" , "aktiven-news" , "jobnews	" ,
# "food industry ger" , "food industry int" , "food industry	" , "regionalgruppen " , "food services")


# organize the columns to the same orders as in the excel file.
newsletter_coded <- select(mydata_coded, type, newsletter_name, campaign_name, 
                 sent_weekday, sent_date, date_sent_text, sent_000,
                 total_recipients, 
                 open_rate, opens, 
                 click_rate, clicks, 
                 bounce_rate, bounces, 
                 unsubscribe_rate, unsubscriptions,
                 complaint_rate , complaints)
View(newsletter_coded)

##  TEST of coding ##
newsletter_list <- dplyr::filter(newsletter_coded, grepl('Newsletter|Fundraising', type))
newsletter_list <- dplyr::filter(newsletter_list, newsletter_list$sent_date>ymd("2016-12-31"))
View(newsletter_list)

# I need still to
# - IF recipients for one mailing < 0.3*average number of newsletter recipients
#   it is probably not a newsletter but a opt in.
# - until this is done one must manually check and update the newsletter_list
#   before it can be copied to all destinations.





### 6. Create a excel file with the information

library(openxlsx)
wb_newsletter <- openxlsx::createWorkbook()
openxlsx::addWorksheet(wb_newsletter, "Newsletter_list")
openxlsx::writeDataTable(wb_newsletter,"Newsletter_list", x = newsletter_list, 
                colNames = TRUE, 
                withFilter = TRUE)
openxlsx::saveWorkbook(wb_newsletter, file = "Newsletter_list.xlsx", overwrite = TRUE)
openxlsx::saveWorkbook(wb_newsletter, file = "\\\\VEBUFILER01/KPIImpact/KPI-Impact/Public Information/Newsletter_list.xlsx", overwrite = TRUE)

wb_all <- createWorkbook()
addWorksheet(wb_all, "Newsletter_coded")
writeDataTable(wb_all,"Newsletter_coded", x = newsletter_coded, 
               colNames = TRUE, 
               withFilter = TRUE)
saveWorkbook(wb_all, file = "Newsletters_coded_2017-2019.xlsx", overwrite = TRUE)
saveWorkbook(wb_all, file = "\\\\VEBUFILER01/KPIImpact/KPI-Impact/Public Information/Newsletters_coded_2017-2019.xlsx", overwrite = TRUE)

rm(wb_all)


#THIS SHould be ALSO SAVED IN mEDIAOUTREACH v-LABEL AND ANY OTHER FOLDER THAT LINKS TO THE NEWSLETTER_LIST.XLSX
"\\VEBUFILER01\KPIImpact\KPI-Impact\Projects\Media Outreach\PVDE_MediaOutreach.xlsx"
"\\VEBUFILER01\KPIImpact\KPI-Impact\Projects\Media Outreach\PVINT_MediaOutreach.xlsx"
"\\VEBUFILER01\KPIImpact\KPI-Impact\Projects\VLabel\VLabel.xlsx"


### Replace the sheet in an existing file ####
# unfortunately the resulting file does not open well = corrupt
#wb_overview <- loadWorkbook("\\\\VEBUFILER01/KPIImpact/KPI-Impact/Public Information/Overview of Newsletters.xlsx")
# show the names of sheets in use
#if(any(names(wb_overview)=="Newsletter_list")) {
#removeWorksheet(wb_overview, "Newsletter_list")
#addWorksheet(wb_overview, "Newsletter_list")
#writeDataTable(wb_overview,"Newsletter_list", x = newsletter_list, 
#               colNames = TRUE, 
#              withFilter = TRUE)
#saveWorkbook(wb_overview, file = "\\\\VEBUFILER01/KPIImpact/KPI-Impact/Public Information/Overview of Newsletters.xlsx", overwrite = TRUE)
#} else {
##   print("The workbook does not contain the correct sheet")}



## House cleaning
# Remove all objects (= dropp all data and variables in memory)
detach(mydata)
rm(list = ls())

