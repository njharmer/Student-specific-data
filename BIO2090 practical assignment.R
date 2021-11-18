# Script to generate unique assignments for BIO2090 practical
# Should provide each student with unique data and markers with an answer sheet.

# Load packages
library(ggplot2)
library(officer)
library(magrittr)
setwd("C:/Users/njh209/OneDrive - University of Exeter/Nic's documents/R/resources")

# Load student numbers
Students = c("C123456", "C234567")

# Set up the protein datasets
MW = c(66000, 60000, 17, 12, 13, 14, 20, 11)
names <- c("BSA", "catalase", "myoglobin", "cytochrome c (horse)",  "ribonuclease A",
           "a-lactalbumin", "GrpE", "cytochrome c (C. jejuni)")
A410 <- c(0,1,3,3,0,0,0,3)
protset <- matrix(data = c(1,3,6,2,3,6,1,4,6,1,4,7,1,3,8,1,5,8,2,3,8,2,5,7), nrow = 3, ncol = 8)

# Set up table variables
headers <- c("Fractions", "A280 nm", "A410 nm")

#Actions for each individual dataset

for (j in 1:length(Students)) {
  
  # Generate randomised values for student
  set <- sample(1:8,1)
  sets <- protset[,set]
  prots <- data.frame(names = c(names[sets[1]],names[sets[2]],names[sets[3]]),
                      MW = c(MW[sets[1]],MW[sets[2]],MW[sets[3]]),
                      A410 = c(A410[sets[1]],A410[sets[2]],A410[sets[3]]))
  if(set == 1 || set == 5) {
    doc <- paste0("Dataset15-",sample(1:6,1));
  } else if(set == 2) {
    doc <- paste0("Dataset2-",sample(1:5,1));
  } else if(set == 3 || set == 6) {
    doc <- paste0("Dataset36-",sample(1:4,1));
  } else if(set == 4) {
    doc <- "Dataset4";
  } else if(set == 7) {
    doc <- paste0("Dataset7-",sample(1:4,1));
  } else if(set == 8) {
    doc <- paste0("Dataset8-",sample(1:4,1));
  }
  
  collen <- sample(450:550,1)
  colrad <- 5
  colvol <- pi*(collen/10)*((colrad/10)^2)
  colvoid <- colvol/3
  theoplat <- sample (400:600,1)
  m <- -1/log10(150000)
  c <- 1
  IEX <- c(sample(13:17,1),2.5,(sample(30:40,1)/100),0,sample(43:51,1),2.5,(sample(30:40,1)/100),0)
  IEX[4] <- IEX[3]*prots$A410[2]; IEX[8] <- IEX[7]*prots$A410[3]
  back280 <- sample(400:450,1)/10000
  back410 <- sample(350:400,1)/10000
  
  # Run SEC column
  prots$Kav <- c + (m*log10(prots$MW))
  prots$Ve <- colvoid + (prots$Kav*(colvol-colvoid))
  prots$W <- (5.54*(prots$Ve^2)*1000/(collen*theoplat))^0.5
  SEC <- data.frame(Fractions = c(0:35),
                    A280nm = c(0),
                    A410nm = c(0))
  SEC2 <- data.frame(Fractions = c(0:35,0:35),
                     Abs = c(0),
                     Wavelength = c(0))
  
  for (i in 1:36) {
    SEC$A280nm[i] <- round((back280 + (sample(20:100,1)/10000) +(((0.3*exp(-(((i+colvoid-2)-prots$Ve[1])^2)/
                                                                           (0.5*(prots$W[1]^2)))) + (0.3*exp(-(((i+colvoid-2)-prots$Ve[2])^2)/(0.5*(prots$W[2]^2))))+
                                                                  (0.3*exp(-(((i+colvoid-2)-prots$Ve[3])^2)/(0.5*(prots$W[3]^2)))))*sample(90:110,1)/100)),3)
    SEC$A410nm[i] <- round((back410 + (sample(20:100,1)/10000) +
                            (((0.3*prots$A410[1]*exp(-(((i+colvoid-2)-prots$Ve[1])^2)/(0.5*(prots$W[1]^2)))) + 
                                (0.3*prots$A410[2]*exp(-(((i+colvoid-2)-prots$Ve[2])^2)/(0.5*(prots$W[2]^2)))) + 
                                (0.3*prots$A410[3]*exp(-(((i+colvoid-2)-prots$Ve[3])^2)/(0.5*(prots$W[3]^2)))))*
                               sample(90:110,1)/100)),3)
    SEC2$Abs[i] <- SEC$A280nm[i]; SEC2$Abs[i+36] <- SEC$A410nm[i]
    SEC2$Wavelength[i] <- "280 nm"; SEC2$Wavelength[i+36] <- "410 nm"
  }
  
  #Run IEX column
  IEC <- data.frame(Fractions = c(0:60),
                    A280nm = c(0),
                    A410nm = c(0))
  IEC2 <- data.frame(Fractions = c(0:60,0:60),
                     Abs = c(0),
                     Wavelength = c(0))
  
  for (i in 1:61) {
    IEC$A280nm[i] <- round((back280 + (sample(20:100,1)/10000) + (((IEX[3]*exp(-(((i-1)-IEX[1])^2)/
                                                                               (0.5*(IEX[2]^2)))) + (IEX[7]*exp(-(((i-1)-IEX[5])^2)/(0.5*(IEX[6]^2)))))*sample(90:110,1)/
                                                                  100)),3)
    IEC$A410nm[i] <- round((back410 + (sample(20:100,1)/10000) + (((IEX[4]*exp(-(((i-1)-IEX[1])^2)/
                                                                               (0.5*(IEX[2]^2)))) + (IEX[8]*exp(-(((i-1)-
                                                                                                                     IEX[5])^2)/(0.5*(IEX[6]^2)))))*sample(90:110,1)/100)),3)
    IEC2$Abs[i] <- IEC$A280nm[i]; IEC2$Abs[i+61] <- IEC$A410nm[i]
    IEC2$Wavelength[i] <- "280 nm"; IEC2$Wavelength[i+61] <- "410 nm"
  }
  
  # Generate student files
  readout <- read_docx(paste0("Start_files/",doc, ".docx"))
  readout <- readout %>%
    cursor_begin %>%
    body_add_par(paste0("Data for student: ", Students[j])) %>%
    body_add_par("") %>%
    body_add_par("Data for size exclusion chromatography:", style = "Normal") %>%
    body_add_par("") %>% 
    body_add_par(paste0("Column void volume: ", round(colvoid), " mL"), style = "Normal") %>%
    body_add_par("") %>% 
    body_add_table(value = SEC, style="Normal Table") %>%
    body_add_par("") %>%
    body_add_par("") %>%
    body_add_par("Data for ion exchange chromatography:") %>%
    body_add_par("") %>%
    body_add_table(value = IEC, style="Normal Table") %>%
    body_add_par("") %>%
    body_add_par("") %>%
    
    print(readout, target = paste0("Student_files/", Students[j], ".docx"))
  
  # Prepare figures
  se <- ggplot(data = SEC2, aes(x = Fractions, y=Abs)) 
  se <- se + geom_path() + geom_point(mapping = aes(size = 2, color = Wavelength)) +
    labs(y="Absorbance / AU", x="Fraction") 
  
  ie <- ggplot(data = IEC2, aes(x = Fractions, y=Abs)) 
  ie <- ie + geom_path() + geom_point(mapping = aes(size = 2, color = Wavelength)) + 
    labs(y="Absorbance / AU", x="Fraction")
  
  # Generate marker files
  
  readout <- read_docx()
  readout <- readout %>%
    body_add_par("BIO2090 Laboratory report assignment", style = 'heading 1') %>%
    body_add_par("") %>%
    body_add_par(paste0("Results for student: ", Students[j])) %>%
    body_add_par("") %>%
    body_add_par(paste0("Correct proteins: ", names[sets[1]], ", ", names[sets[2]], 
                        ", ", names[sets[3]]), style = "Normal") %>%
    body_add_par("") %>%
    body_add_par("Graph for size exclusion chromatography") %>%
    body_add_par("") %>%
    body_add_gg(value = se, style = "centered") %>%
    body_add_par("") %>%
    body_add_break() %>%
    body_add_par("Graph for ion exchange chromatography:") %>%
    body_add_par("") %>%
    body_add_gg(value = ie, style="centered") %>%
    
    print(readout, target = paste0("Lecturer_files/", Students[j], "_answers.docx"))
  
}

