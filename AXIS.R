library(readxl)
library(tidyverse)
library(officer)
library(rvg)
library(ggpubr)
library(cowplot)
library(svMisc)

#Import data

dat <- read_excel("rawdata send(Dung).xlsx", 
                 range = cell_rows(3:151), col_names = F)
names(dat) <- read_excel("rawdata send(Dung).xlsx", col_names = F)[1,]
name_dat <- read_excel("rawdata send(Dung).xlsx")[1,]

#process data
dat <- dat %>% mutate_if(is.character,as.factor)
levels(dat$I2) <- c("Single","Married","Single","Single")
dat <- dat %>% mutate_if(is.character,as.factor)
levels(dat$D4) <- c("Food Preparer", "Influencer", "Influencer")
levels(dat$A1) <- c("Central", "Southern", "Northern", "Southern",
                    "Northern", "Northern")
levels(dat$I3) <- c("Middle school/below", "High school", "High school",
                    "Middle school/below", "College and above", "College and above")
levels(dat$I5) <- c("High", rep("Medium",12),rep("Low",4))
levels(dat$I2) <- c("Single","Married","Single","Single")
levels(dat$D5) <- c("Non_Primary Caregiver", "Primary Caregiver")

#plotting
name_plot <- data.frame(D4 = "Cooking Responsibility", A1 = "Area",
                        D2C = "Age", I3 = "Education",
                        I2 = "Marital Status", D5 = "Parental Status",
                        I5 = "Income")

n <- 1
myplot <- vector(mode = "list", length = length(189:222))
for (i in names(dat)[189:222]) {
      progress(n,progress.bar = TRUE)
      ##donut
      donut <- dat %>% count_(i) %>%
            mutate(Percent = n/sum(n)*100) %>%
            ggplot(aes_string(x=2, y='Percent', fill=i)) + 
            theme_void()+
            theme(plot.title = element_text(hjust = 0.5),
                  legend.position = "bottom")+
            geom_col(position = "stack", width = 1)+
            geom_text(aes(label=round(Percent)), 
                      position = position_stack(vjust=0.5))+
            xlim(0.5,2.5)+
            coord_polar(theta = "y")+
            labs(title = "TOTAL", fill="")
      
      ##bar
      plot <- lapply(names(name_plot), function(x){
            test <- dat %>% group_by_at(x) %>% 
                  count_(i) %>% 
                  mutate(Percent = n/sum(n)*100)
            ggplot(test, aes_string(x = x, y = "Percent", fill = i)) + 
                  theme_void()+
                  theme(plot.title = element_text(hjust = 0.5))+
                  geom_bar(stat="identity") +
                  geom_text(aes(label=round(Percent), y = Percent),
                            position=position_stack(vjust=0.5)) +
                  labs(title = name_plot[,which(names(name_plot)==x)],
                       fill = "")
      })
      
      myplot[[n]] <- ggarrange(donut,
                             ggarrange(ggarrange(plotlist = plot[1:3], ncol = 3,
                                                 widths = c(2,3,5), legend = "none"),
                                       ggarrange(plotlist = plot[4:7], ncol = 4, 
                                                 widths = c(3,2,2,3), legend = "none"),
                                       nrow = 2, align = "h"),
                             ncol = 2, widths = c(1,2))
      n <- n+1
}

#create Powerpoint
path <- "test.pptx"
if(!file.exists(path)) {
      out <- read_pptx()
} else {
      out <- read_pptx(path)
}

for (k in 1:length(myplot)) {
      progress(k,progress.bar = TRUE)
      out %>%
            add_slide(layout = "Title and Content", master = "Office Theme") %>%
            ph_with(myplot[[k]], location = ph_location_type(type = "body")) %>%
            ph_with(names(name_dat)[k], location = ph_location(left = 0.5, bottom =0.5)) %>%
            print(target = path)
}