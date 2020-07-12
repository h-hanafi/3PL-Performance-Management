library(tidyverse)
library(lubridate)
library(bizdays)
library(readxl)
library(googledrive)
drive_auth(email = "3pl.hr.pnd@gmail.com")
create.calendar(name = "mycal", weekdays = c("friday","saturday"))

drive_download("2020 - 2nd Quarter Monthly Attendance Sheet", path = file.path(getwd(),"DATA","2020 - 2nd Quarter Monthly Attendance Sheet.xlsx"),overwrite = TRUE)
read_excel(file.path(getwd(),"DATA","2020 - 2nd Quarter Monthly Attendance Sheet.xlsx"),sheet = "March",skip = 10) %>% 
  gather(3:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Day),3,2020,sep = "/")) %>%  %>% 
  rename(c("ID","Name",))

                                                                 