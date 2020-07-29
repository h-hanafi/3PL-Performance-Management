library(tidyverse)
library(lubridate)
library(bizdays)
library(readxl)
library(googledrive)
library(grDevices)
drive_auth(email = "3pl.hr.pnd@gmail.com")

# Variable Creation
create.calendar(name = "mycal", weekdays = c("friday","saturday"))

Att_Legend <- tibble(Status = c("Working Days","Late","Authorised Sick Leave","Unauthorised Sick Leave","Maternity Leave","Marriage Leave","Annual Leave","Compassionate Leave","Authorized Leave Without Pay","Unauthorized Leave Without Pay","Authorized Leave With Pay","Other Leave","Dependents Leave","Not Employed","Overtime"),
                     Attendance = c("W","L","SA","SU","ML","MAL","ANL","CL","ALW","ULW","ALP","OL","DPL","X","OVT"))

# 2nd Quarter
drive_download("2020 - 2nd Quarter Monthly Attendance Sheet", path = file.path(getwd(),"DATA","2020 - 2nd Quarter Monthly Attendance Sheet.xlsx"),overwrite = TRUE)

## March
Attendance <- read_excel(file.path(getwd(),"DATA","2020 - 2nd Quarter Monthly Attendance Sheet.xlsx"),sheet = "March",skip = 10) %>% 
  gather(4:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),3,2020,sep = "/")) %>% relocate(Comment, .after = last_col()) %>%
  rename_at(vars(colnames(.)), ~c("ID","Name","Status","Date","Attendance","Comment"))
## April
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 2nd Quarter Monthly Attendance Sheet.xlsx"),sheet = "April",skip = 10) %>% 
                                      gather(4:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),4,2020,sep = "/")) %>% 
                                      relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Status","Date","Attendance","Comment")))
## May
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 2nd Quarter Monthly Attendance Sheet.xlsx"),sheet = "May",skip = 10) %>% 
                                      gather(4:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),5,2020,sep = "/")) %>% 
                                      relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Status","Date","Attendance","Comment")))
## June
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 2nd Quarter Monthly Attendance Sheet.xlsx"),sheet = "June",skip = 10) %>% 
                                      gather(4:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),6,2020,sep = "/")) %>% 
                                      relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Status","Date","Attendance","Comment")))

# 3rd Quarter
drive_download("2020 - 3rd Quarter Monthly Attendance Sheet", path = file.path(getwd(),"DATA","2020 - 3rd Quarter Monthly Attendance Sheet.xlsx"),overwrite = TRUE)

## July
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 3rd Quarter Monthly Attendance Sheet.xlsx"),sheet = "July",skip = 10) %>% 
                                      gather(4:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),7,2020,sep = "/")) %>% 
                                      relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Status","Date","Attendance","Comment")))
## August
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 3rd Quarter Monthly Attendance Sheet.xlsx"),sheet = "August",skip = 10) %>% 
                                      gather(4:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),8,2020,sep = "/")) %>% 
                                      relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Status","Date","Attendance","Comment")))
## September
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 3rd Quarter Monthly Attendance Sheet.xlsx"),sheet = "September",skip = 10) %>% 
                                      gather(4:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),9,2020,sep = "/")) %>% 
                                      relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Status","Date","Attendance","Comment")))



# 4th Quarter
drive_download("2020 - 4th Quarter Monthly Attendance Sheet", path = file.path(getwd(),"DATA","2020 - 4th Quarter Monthly Attendance Sheet.xlsx"),overwrite = TRUE)

## October
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 4th Quarter Monthly Attendance Sheet.xlsx"),sheet = "October",skip = 10) %>% 
                                      gather(4:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),10,2020,sep = "/")) %>% 
                                      relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Status","Date","Attendance","Comment")))

## November
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 4th Quarter Monthly Attendance Sheet.xlsx"),sheet = "November",skip = 10) %>% 
                                      gather(4:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),11,2020,sep = "/")) %>% 
                                      relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Status","Date","Attendance","Comment")))

## December
Attendance <-  Attendance %>% rbind(read_excel(file.path(getwd(),"DATA","2020 - 4th Quarter Monthly Attendance Sheet.xlsx"),sheet = "December",skip = 10) %>% 
                                      gather(4:(ncol(.)-1),key = "Date", value = "Attendance") %>% mutate(Date = paste(as.numeric(Date),12,2020,sep = "/")) %>% 
                                      relocate(Comment, .after = last_col()) %>%
                                      rename_at(vars(colnames(.)), ~c("ID","Name","Status","Date","Attendance","Comment")))

# Final Data Wrangling
Attendance <- Attendance %>% filter(!is.na(ID)) %>% mutate(Date = dmy(Date)) %>% arrange(desc(ID)) %>%  mutate(Name = factor(Name, levels = unique(Name)),
                                                                                                               ID = factor(ID, levels = unique(ID)),
                                                                                                               Attendance = factor(Attendance,Att_Legend$Attendance))
# Figures
## March
Attendance %>% filter(month(Date) == 3) %>% ggplot(aes(Date,Name, fill = Attendance)) + geom_tile(color = "black") + 
  scale_x_date(expand = c(0,0), date_breaks = "1 week", date_labels = "%d") + scale_y_discrete(expand = c(0,0)) + 
  scale_fill_brewer(palette = "Set2", na.value = "grey50") + 
  labs(title = "March Attendance", y = "Staff Name") + theme(plot.title = element_text(hjust = 0.5)) +
ggsave("03-2020 Attendance.png", device = "png",path = file.path(getwd(),"Reports"))

## April
Attendance %>% filter(month(Date) == 4) %>% ggplot(aes(Date,Name, fill = Attendance)) + geom_tile(color = "black") + 
  scale_x_date(expand = c(0,0), date_breaks = "1 week", date_labels = "%d") + scale_y_discrete(expand = c(0,0)) + 
  scale_fill_brewer(palette = "Set2", na.value = "grey50") + 
  labs(title = "April Attendance", y = "Staff Name") + theme(plot.title = element_text(hjust = 0.5))
ggsave("04-2020 Attendance.png", device = "png",path = file.path(getwd(),"Reports"))

## May
Attendance %>% filter(month(Date) == 5) %>% ggplot(aes(Date,Name, fill = Attendance)) + geom_tile(color = "black") + 
  scale_x_date(expand = c(0,0), date_breaks = "1 week", date_labels = "%d") + scale_y_discrete(expand = c(0,0)) + 
  scale_fill_brewer(palette = "Set2", na.value = "grey50") + 
  labs(title = "May Attendance", y = "Staff Name") + theme(plot.title = element_text(hjust = 0.5))
ggsave("05-2020 Attendance.png", device = "png",path = file.path(getwd(),"Reports"))

## June
Attendance %>% filter(month(Date) == 6) %>% ggplot(aes(Date,Name, fill = Attendance)) + geom_tile(color = "black") + 
  scale_x_date(expand = c(0,0), date_breaks = "1 week", date_labels = "%d") + scale_y_discrete(expand = c(0,0)) + 
  scale_fill_brewer(palette = "Set2", na.value = "grey50") + 
  labs(title = "June Attendance", y = "Staff Name") + theme(plot.title = element_text(hjust = 0.5))
ggsave("06-2020 Attendance.png", device = "png",path = file.path(getwd(),"Reports"))

## July
Attendance %>% filter(month(Date) == 7) %>% ggplot(aes(Date,Name, fill = Attendance)) + geom_tile(color = "black") + 
  scale_x_date(expand = c(0,0), date_breaks = "1 week", date_labels = "%d") + scale_y_discrete(expand = c(0,0)) + 
  scale_fill_brewer(palette = "Set2", na.value = "grey50") + 
  labs(title = "July Attendance", y = "Staff Name") + theme(plot.title = element_text(hjust = 0.5))
ggsave("07-2020 Attendance.png", device = "png",path = file.path(getwd(),"Reports"))

## August
Attendance %>% filter(month(Date) == 8) %>% ggplot(aes(Date,Name, fill = Attendance)) + geom_tile(color = "black") + 
  scale_x_date(expand = c(0,0), date_breaks = "1 week", date_labels = "%d") + scale_y_discrete(expand = c(0,0)) + 
  scale_fill_brewer(palette = "Set2", na.value = "grey50") + 
  labs(title = "August Attendance", y = "Staff Name") + theme(plot.title = element_text(hjust = 0.5))
ggsave("08-2020 Attendance.png", device = "png",path = file.path(getwd(),"Reports"))

## September
Attendance %>% filter(month(Date) == 9) %>% ggplot(aes(Date,Name, fill = Attendance)) + geom_tile(color = "black") + 
  scale_x_date(expand = c(0,0), date_breaks = "1 week", date_labels = "%d") + scale_y_discrete(expand = c(0,0)) + 
  scale_fill_brewer(palette = "Set2", na.value = "grey50") + 
  labs(title = "September Attendance", y = "Staff Name") + theme(plot.title = element_text(hjust = 0.5))
ggsave("09-2020 Attendance.png", device = "png",path = file.path(getwd(),"Reports"))

## October
Attendance %>% filter(month(Date) == 10) %>% ggplot(aes(Date,Name, fill = Attendance)) + geom_tile(color = "black") + 
  scale_x_date(expand = c(0,0), date_breaks = "1 week", date_labels = "%d") + scale_y_discrete(expand = c(0,0)) + 
  scale_fill_brewer(palette = "Set2", na.value = "grey50") + 
  labs(title = "September Attendance", y = "Staff Name") + theme(plot.title = element_text(hjust = 0.5))
ggsave("10-2020 Attendance.png", device = "png",path = file.path(getwd(),"Reports"))

## November
Attendance %>% filter(month(Date) == 11) %>% ggplot(aes(Date,Name, fill = Attendance)) + geom_tile(color = "black") + 
  scale_x_date(expand = c(0,0), date_breaks = "1 week", date_labels = "%d") + scale_y_discrete(expand = c(0,0)) + 
  scale_fill_brewer(palette = "Set2", na.value = "grey50") + 
  labs(title = "November Attendance", y = "Staff Name") + theme(plot.title = element_text(hjust = 0.5))
ggsave("11-2020 Attendance.png", device = "png",path = file.path(getwd(),"Reports"))

## December
Attendance %>% filter(month(Date) == 12) %>% ggplot(aes(Date,Name, fill = Attendance)) + geom_tile(color = "black") + 
  scale_x_date(expand = c(0,0), date_breaks = "1 week", date_labels = "%d") + scale_y_discrete(expand = c(0,0)) + 
  scale_fill_brewer(palette = "Set2", na.value = "grey50") + 
  labs(title = "December Attendance", y = "Staff Name") + theme(plot.title = element_text(hjust = 0.5))
ggsave("12-2020 Attendance.png", device = "png",path = file.path(getwd(),"Reports"))

# Summary
Attendance %>% group_by(ID,Attendance,.drop = F) %>%  filter(Date <= Sys.Date(), Status %in%  c("Permanent","Probation")) %>%
  summarise(Count = n()) %>% ungroup() %>%  left_join(Attendance %>% distinct(ID, .keep_all = T) %>% select(ID,Name)) %>% relocate(Name, .after = ID) %>%
  left_join(Att_Legend) %>% relocate(Status, .after = Attendance) %>% arrange(desc(ID)) %>%  write_excel_csv(file.path(getwd(),"Reports","Attendance Summary.csv"))


