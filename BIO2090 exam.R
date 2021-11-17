# Script to generate unique assignments for BIO2090 exam
# Should provide each student with unique data and markers with an answer sheet.

# Load packages
library(ggplot2)
library(officer)
library(magrittr)
setwd("C:/Users/njh209/OneDrive - University of Exeter/Nic's documents/R/resources")

# Load student numbers
Students = c("C12345", "C23456")

# Set errors
error <- c(15, 7)

#Actions for each individual dataset

for (j in 1:length(Students)) {
  
  # Generate Bradford data
  Bradford <- data.frame(Tube = c(1:6),
                         Volume_of_BSA_solution_uL = c(0,4,8,12,16,20),
                         Volume_of_water_uL = c(1000,996,992,988,984,980),
                         Absorbance_595nm = c(0),
                         concentration_ug_mL = c(0),
                         Corrected_Absorbance_595nm = c(0)) 
  grad <- 0.0025+(sample(0:100,1)/100000)
  back <- 0.9+(sample(0:200,1)/1000)
  
  for (i in 1:6) {
    Bradford[i,4] <- round(((100-error[2]+((sample(0:(error[2]*200),1))/100))/100)*(back+(Bradford[i,2]*10
                                                                                          *grad*((100-error[1]+((sample(0:(error[1]*200),1))/100))/100))), digits = 3)
    Bradford[i,5] <- Bradford[i,2]*10
    Bradford[i,6] <- Bradford[i,4]-Bradford[1,4]
  }  
  
  # Calculate linear regression of Bradford
  Brlr <- lm(Bradford[,6]~Bradford[,5])
  slope <- as.numeric(Brlr$coefficients[2])
  yint <- as.numeric(Brlr$coefficients[1])
  
  # Generate table of results for unknowns
  Fractions <- data.frame(Fraction = c("Homogenate, H", "Cytosol, C", "Mitochondria, M", "Nuclear, N",
                                       "30% ammonium sulfate", "50% ammonium sulfate", "80% ammonium sulfate"),
                          Volume_of_sample_uL = c(5),
                          Volume_of_water_uL = c(996),
                          Absorbance_595nm = c((back+0.35+sample(0:200,1)/1000), 
                                               (back+0.5+sample(0:300,1)/1000),
                                               (back+0.4+sample(0:400,1)/1000),
                                               (back+0.25+sample(0:300,1)/1000),
                                               (back+0.3+sample(0:50,1)/1000),
                                               (back+0.5+sample(0:200,1)/1000),
                                               (back+0.3+sample(0:100,1)/1000)),
                          Corrected_Absorbance = c(0),
                          x_value = c(0),
                          concentration = c(0))
  
  for (i in 1:7) {
    Fractions[i,5] <- Fractions[i,4]-Bradford[1,4]
    Fractions[i,6] <- signif((Fractions[i,5]-yint)/slope,4)
    Fractions[i,7] <- signif(Fractions[i,6]/5,3)
  }
  
  # Generate PGK data
  PGK <- data.frame(Fraction = c("Homogenate, H", "Cytosol, C", "Mitochondria, M", "Nuclear, N",
                                 "30% ammonium sulfate", "50% ammonium sulfate", "80% ammonium sulfate"),
                    Gradient_of_slope_AU_per_min = c((-0.03-sample(0:200,1)/1000), 
                                                     (-0.03-sample(0:200,1)/1000),
                                                     (sample(0:100,1)/-1000),
                                                     (-0.1-sample(0:100,1)/1000),
                                                     (-0.03-sample(0:20,1)/1000),
                                                     (-0.03-sample(0:30,1)/1000),
                                                     (-0.15-sample(0:200,1)/1000)),
                    Product_concentration = c(0),
                    Activity = c(0),
                    Mass_protein = c(0),
                    Specific_activity = c(0))
  
  for (i in 1:7) {
    PGK[i,3] <- signif(PGK[i,2]*1000000/-6220,5)
    PGK[i,4] <- PGK[i,3]/1000
    PGK[i,5] <- Fractions[i,7]*0.005
    PGK[i,6] <- signif(PGK[i,4]/PGK[i,5],3)
  }
  
  # Generate student files
  
  readout <- read_docx()
  readout <- readout %>%
    body_add_par("BIO2090 2020/21 January Examination", style = 'heading 1') %>%
    body_add_par("") %>%
    body_add_par(paste0("Individual data for student: ", Students[j]," for Question 1")) %>%
    body_add_par("") %>%
    body_add_par("Bradford standard curve data:", style = "Normal") %>%
    body_add_par("") %>% 
    body_add_table(value = Bradford[,1:4], style="table_template") %>%
    body_add_par("") %>%
    body_add_par("") %>%
    body_add_par("Sample Bradford data:") %>%
    body_add_par("") %>%
    body_add_table(value = Fractions[,1:4], style="table_template") %>%
    body_add_par("") %>%
    body_add_par("") %>%
    body_add_par("PGK assay data:") %>%
    body_add_par("") %>%
    body_add_table(value = PGK[,1:2], style="table_template") %>%
    
    
    print(readout, target = paste0("Student_files/", Students[j], ".docx"))
  
  # Prepare figures
  se <- ggplot(data = Bradford, aes(x = concentration_ug_mL, y=Corrected_Absorbance_595nm)) + 
    stat_smooth(method = "lm",  formula = y~x, level = 0.75)
  se <- se + geom_point(mapping = aes(size = 2)) + labs(y="Absorbance / AU", x="concentration / ug/mL")
  
  # Generate lecturer files
  
  readout <- read_docx()
  readout <- readout %>%
    body_add_par("BIO2090 2020/21 Exam Dataset", style = 'heading 1') %>%
    body_add_par("") %>%
    body_add_par(paste0("Data and answers for student: ", Students[j])) %>%
    body_add_par("") %>%
    body_add_par("") %>%
    body_add_par("Bradford standard curve data:", style = "Normal") %>%
    body_add_par("") %>% 
    body_add_table(value = Bradford, style="table_template") %>%
    body_add_par("") %>%
    body_add_par("") %>%
    body_add_par(paste0("Bradford graph: formula is y = ", signif(slope,3),"x + ",signif(yint,3))) %>%
    body_add_par("") %>%
    body_add_gg(value = se, style = "centered") %>%
    body_add_par("") %>%
    body_add_par("") %>%
    body_add_par("Sample Bradford data:") %>%
    body_add_par("") %>%
    body_add_table(value = Fractions, style="table_template") %>%
    body_add_par("") %>%
    body_add_par("") %>%
    body_add_par("PGK assay data:") %>%
    body_add_par("") %>%
    body_add_table(value = PGK, style="table_template") %>%
    
    
    print(readout, target = paste0("Lecturer_files/", Students[j], "_answers.docx"))
  
}

