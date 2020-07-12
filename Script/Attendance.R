library(tidyverse)
library(lubridate)
library(bizdays)
library(readxl)
library(googledrive)
drive_auth(email = "3pl.hr.pnd@gmail.com")
create.calendar(name = "mycal", weekdays = c("friday","saturday"))

# 2nd Quarter
drive_download("2020 - 2nd Quarter Monthly Attendance Sheet", path = file.path(getwd(),"DATA","2020 - 2nd Quarter Monthly Attendance Sheet.xlsx"),overwrite = TRUE)

## March
Attendance <- read_excel(file.path(getwd(),"DATA","2020 - 2nd Quarter Monthly Attendance Sheet.xlsx"),sheet = "March",skip = 10) %>% 
  gather(3:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),3,2020,sep = "/")) %>% relocate(Comment, .after = last_col()) %>%
  rename_at(vars(colnames(.)), ~c("ID","Name","Date","Attendance","Comment"))
## April
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 2nd Quarter Monthly Attendance Sheet.xlsx"),sheet = "April",skip = 10) %>% 
                                      gather(3:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),4,2020,sep = "/")) %>% 
                                      relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Date","Attendance","Comment")))
## May
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 2nd Quarter Monthly Attendance Sheet.xlsx"),sheet = "May",skip = 10) %>% 
                                      gather(3:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),5,2020,sep = "/")) %>% relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Date","Attendance","Comment")))
## June
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 2nd Quarter Monthly Attendance Sheet.xlsx"),sheet = "June",skip = 10) %>% 
                                      gather(3:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),6,2020,sep = "/")) %>% relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Date","Attendance","Comment")))

# 3rd Quarter
drive_download("2020 - 3rd Quarter Monthly Attendance Sheet", path = file.path(getwd(),"DATA","2020 - 3rd Quarter Monthly Attendance Sheet.xlsx"),overwrite = TRUE)

## July
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 3rd Quarter Monthly Attendance Sheet.xlsx"),sheet = "July",skip = 10) %>% 
                                      gather(3:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),7,2020,sep = "/")) %>% 
                                      relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Date","Attendance","Comment")))
## August
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 3rd Quarter Monthly Attendance Sheet.xlsx"),sheet = "August",skip = 10) %>% 
                                      gather(3:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),8,2020,sep = "/")) %>% 
                                      relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Date","Attendance","Comment")))
## September
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 3rd Quarter Monthly Attendance Sheet.xlsx"),sheet = "September",skip = 10) %>% 
                                      gather(3:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),9,2020,sep = "/")) %>% 
                                      relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Date","Attendance","Comment")))
