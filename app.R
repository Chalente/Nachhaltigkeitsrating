library(tidyverse)
library(dplyr)
library(readxl)
library(treemap)
library(fmsb)
library(stringr)
library(gridExtra)

#set working directory
setwd("~/Documents/R/Team Nachhaltigkeit/Ergebnisse")


#CLEANING

BASE_Nachhaltigkeit_Umfrage <- read_excel("Documents/R/Realace Analytics Tool/BASE_Nachhaltigkeit_Umfrage.xlsx", 
                                          col_names = FALSE, col_types = c("text", 
                                                                           "numeric", "skip", "text", "text"), 
                                          skip = 5)

colnames(BASE_Nachhaltigkeit_Umfrage)[1]  <- "Parameter"   # change column name for x column
colnames(BASE_Nachhaltigkeit_Umfrage)[2]  <- "Gewichtung"  # change column name for x column
colnames(BASE_Nachhaltigkeit_Umfrage)[3]  <- "Ist"  # change column name for x column
colnames(BASE_Nachhaltigkeit_Umfrage)[4]  <- "Soll"   # change column name for x column

BASE_Nachhaltigkeit_Umfrage <-  BASE_Nachhaltigkeit_Umfrage[-c(31, 30, 29, 28, 1),]
BASE_Nachhaltigkeit_Umfrage <-  BASE_Nachhaltigkeit_Umfrage[-c(26:27),]

BASE_Nachhaltigkeit_Umfrage <-  BASE_Nachhaltigkeit_Umfrage[-c(8, 14, 19, 24),]

BASE_Nachhaltigkeit_Umfrage$Parameter <- BASE_Nachhaltigkeit_Umfrage$Parameter %>% replace_na(replace = "")

BASE_Nachhaltigkeit_Umfrage["Gruppe"] <- ""
BASE_Nachhaltigkeit_Umfrage$Gruppe[2:7] <- "Ökologisch"
BASE_Nachhaltigkeit_Umfrage$Gruppe[9:12] <- "Sozial"
BASE_Nachhaltigkeit_Umfrage$Gruppe[14:16] <- "Ökonomisch"
BASE_Nachhaltigkeit_Umfrage$Gruppe[18:20] <- "Basal"

BASE_Nachhaltigkeit_Umfrage <-  BASE_Nachhaltigkeit_Umfrage[-c(1, 8, 13, 17),]

BASE_Nachhaltigkeit_Umfrage[6,1] <- "Zwischenergebnis Ökologie"
BASE_Nachhaltigkeit_Umfrage[10,1] <- "Zwischenergebnis Sozial"
BASE_Nachhaltigkeit_Umfrage[13,1] <- "Zwischenergebnis Ökomisch"
BASE_Nachhaltigkeit_Umfrage[16,1] <- "Zwischenergebnis Basal"

#Umfragen laden

HO <- c("Nachhaltigkeit_Umfrage_hauke.xlsx")
CM <- c("BASE_Nachhaltigkeit_Umfrage_CM.xlsx")
AK <- c("BASE_Nachhaltigkeit_Umfrage_AK.xlsx")
SV <- c("Nachhaltigkeit_Umfrage_SV.xlsx")
PK <- c("BASE_Nachhaltigkeit_Umfrage_PK.xlsx")


####CLEANING FUNCTION

Cleaning <- function(x){
  
BASE_Nachhaltigkeit_Umfrage <- read_excel(x, 
                                          col_names = FALSE, col_types = c("text", 
                                                                           "numeric", "skip", "text", "text"), 
                                          skip = 5)

colnames(BASE_Nachhaltigkeit_Umfrage)[1]  <- "Parameter"   # change column name for x column
colnames(BASE_Nachhaltigkeit_Umfrage)[2]  <- "Gewichtung"  # change column name for x column
colnames(BASE_Nachhaltigkeit_Umfrage)[3]  <- "Ist"  # change column name for x column
colnames(BASE_Nachhaltigkeit_Umfrage)[4]  <- "Soll"   # change column name for x column

BASE_Nachhaltigkeit_Umfrage <-  BASE_Nachhaltigkeit_Umfrage[-c(31, 30, 29, 28, 1),]
BASE_Nachhaltigkeit_Umfrage <-  BASE_Nachhaltigkeit_Umfrage[-c(26:27),]

BASE_Nachhaltigkeit_Umfrage <-  BASE_Nachhaltigkeit_Umfrage[-c(8, 14, 19, 24),]

BASE_Nachhaltigkeit_Umfrage$Parameter <- BASE_Nachhaltigkeit_Umfrage$Parameter %>% replace_na(replace = "")

BASE_Nachhaltigkeit_Umfrage["Gruppe"] <- ""
BASE_Nachhaltigkeit_Umfrage$Gruppe[2:7] <- "Ökologisch"
BASE_Nachhaltigkeit_Umfrage$Gruppe[9:12] <- "Sozial"
BASE_Nachhaltigkeit_Umfrage$Gruppe[14:16] <- "Ökonomisch"
BASE_Nachhaltigkeit_Umfrage$Gruppe[18:20] <- "Basal"

BASE_Nachhaltigkeit_Umfrage <-  BASE_Nachhaltigkeit_Umfrage[-c(1, 8, 13, 17),]

BASE_Nachhaltigkeit_Umfrage[6,1] <- "Zwischenergebnis Ökologie"
BASE_Nachhaltigkeit_Umfrage[10,1] <- "Zwischenergebnis Sozial"
BASE_Nachhaltigkeit_Umfrage[13,1] <- "Zwischenergebnis Ökomisch"
BASE_Nachhaltigkeit_Umfrage[16,1] <- "Zwischenergebnis Basal"

Sammeldatei <- rbind(as.data.frame(BASE_Nachhaltigkeit_Umfrage))

return(Sammeldatei)

}


Hauke <- Cleaning(HO) %>% mutate(Person = "HO")
Antje <- Cleaning(AK) %>% mutate(Person = "AK")
Cornelius <- Cleaning(CM) %>% mutate(Person = "CM")
Antje <- Cleaning(AK) %>% mutate(Person = "AK")
Phillippa <- Cleaning(PK) %>% mutate(Person = "PK")
Sebastian <- Cleaning(SV) %>% mutate(Person = "SV")

Sammling_Excels <- as.data.frame(rbind(Hauke, Antje, Cornelius, Phillippa, Sebastian ))

#Ist-Zustand
Ist_Zustand <- Sammling_Excels %>% select(-c(Soll, Person)) 
Ist_Zustand$Ist = as.integer(Ist_Zustand$Ist)
Ist_Zustand <- Ist_Zustand %>% group_by(Parameter) %>% summarize(Mean_Bewertung = mean(Ist))

#Soll-Zustand
Soll_Zustand <- Sammling_Excels %>% select(-c(Ist, Person)) 
Soll_Zustand$Soll = as.integer(Soll_Zustand$Soll)
Soll_Zustand <- Soll_Zustand %>% group_by(Parameter) %>% summarize(Mean_Bewertung = mean(Soll))

#VISUALISIERUNG

#Vorbereitung TreeMap
Sammlung_Zusammenfassung <- Sammling_Excels %>% group_by(Parameter, Gewichtung, Gruppe) %>% summarize(Durchschnittliche_Bewertung = mean(Gewichtung))

#Farben-Palette aus Elementetafel übernehmen
#palette_treemap <- c("#00000", "#92865")
#Sammlung_Zusammenfassung$Farbe <- ""
#Farbzuordnung nach Zeile 26 

#Zwischenenprozente und Gesmtprozente raus
BASE_Nachhaltigkeit_Zwischenergebnisse <- subset(Sammlung_Zusammenfassung, grepl("^Zwischenergebnis", Parameter)) 

# treemap nur Obergruppen
Treemap <- treemap(BASE_Nachhaltigkeit_Zwischenergebnisse,
        index= "Gruppe",
        vSize="Gewichtung",
        type="index",
        border.col = "white")

#treemap mit Untergruppen aka Parametern
white_palette <- c("#RRGGBB")
Treemap_alle <-  Sammlung_Zusammenfassung[-c(5),] %>% filter(!str_detect(Parameter, "Zwischenergebnis.*"))


treemap(Treemap_alle,
                   index= c("Gruppe", "Parameter"),
                   vSize="Gewichtung",
                   type="index",
                   fontcolor.labels = "black",
                   border.col = "white",
                   overlap.labels = 1,
                  align.labels = c("center", "center")) 

#SpiderRadar

#Firmenfarben setzen
realace_blau <- c("#20205F")

#Rows zu Columns
modify_for_rc <- function(y){
BASE_Nachhaltigkeit_Spider <- as.data.frame(t(y))
colnames(BASE_Nachhaltigkeit_Spider) <- BASE_Nachhaltigkeit_Spider[1,]
BASE_Nachhaltigkeit_Spider <- BASE_Nachhaltigkeit_Spider[-1,]
BASE_Nachhaltigkeit_Spider <- BASE_Nachhaltigkeit_Spider[-3,]
#BASE_Nachhaltigkeit_Spider$Dekarbonisierung <-  replace_na(BASE_Nachhaltigkeit_Spider$Dekarbonisierung, replace = "0")
#BASE_Nachhaltigkeit_Spider <- BASE_Nachhaltigkeit_Spider[-3,]
#BASE_Nachhaltigkeit_Spider <- BASE_Nachhaltigkeit_Spider[-1,]

BASE_Nachhaltigkeit_Spider <- BASE_Nachhaltigkeit_Spider %>% select(-c("Zwischenergebnis Ökologie", "Zwischenergebnis Sozial", "Zwischenergebnis Ökomisch", "Zwischenergebnis Basal", "Gesamtnutzwert"))
BASE_Nachhaltigkeit_Spider[,1:12] <- lapply(BASE_Nachhaltigkeit_Spider[,1:12], as.numeric)

#adding max und min
BASE_Nachhaltigkeit_Spider <- rbind(rep(2,10) , rep(-2,10) , BASE_Nachhaltigkeit_Spider)

#Reorder Columns
#BASE_Nachhaltigkeit_Spider <- BASE_Nachhaltigkeit_Spider %>% select(Dekarbonisierung, Bauweise, Standort, Betrieb, Prozess, Versorgung, Wohlfühlen, Empowerment, "Life Cycle Value", Governance, Nutzen, Smartness )
BASE_Nachhaltigkeit_Spider <- BASE_Nachhaltigkeit_Spider %>% select(Smartness, Nutzen, Governance, "Life Cycle Value", Standort, Bauweise, Dekarbonisierung, Betrieb, Prozess, Wohlfühlen, Empowerment, Versorgung )

}

RC_Ist <- modify_for_rc(Ist_Zustand)
RC_Soll <- modify_for_rc(Soll_Zustand)


#RadarChart Ist

Ist_Chart <- radarchart(RC_Ist,axistype=6, title= "Ist", pcol= realace_blau , pfcol=scales::alpha(realace_blau, 0.5), plwd=3, plty= "solid", pdensity = 32,
           
           cglcol="grey", cglty=3, axislabcol="grey", caxislabels=seq(0,20,5), cglwd=1.1,
           
           )

#RadarChart Soll


Soll_Chart <- radarchart(RC_Soll,axistype=6, title= "Soll", pcol= realace_blau , pfcol=scales::alpha(realace_blau, 0.5), plwd=3, plty= "solid", pdensity = 32,
           
           cglcol="grey", cglty=3, axislabcol="grey", caxislabels=seq(0,20,5), cglwd=1.1
           
)


