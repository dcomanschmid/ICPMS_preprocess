#############################################################
# ICPMS data formating                                      #
#  - read in data from comma separated :)) txt files        #            
#  - format data                                            #                                          
#     *for each text file create an excel workbook(WB)      #
#     *each WB will have a worksheet(WS) per chemical       #
#     *each WS will have a color :))) specific to the       #
#      chemical class                                       #
#     *each WS will have a column with the sample names     #
#      and columns for mean, sd etc.                        #
#  - save output in excel workbooks	                    #                                                                                            #                                                           #                                       
#                                                           #
# Diana Coman Schmid                                        #      
# diana.comanschmid@eawag.ch                                #
#############################################################

#clear the R workspace
rm(list = ls())

#install (Tools tab -> Insatll Packages) andload libraries
library("XLConnect")

#set the working directory (where your text files are)
#replace the text between the quotes with the file path of the folder where your input files are stored (only store files with this specific ICPMS format in this folder)
wd <- "..."


# show the files available in the working directory
list.files(wd)

list.files(wd,pattern = "*.txt")

file.list = list.files(wd, pattern = "*.TXT")# check if file extention is always .TXT
names(file.list) <- file.list
file.list
str(file.list)

#read in the text files 

dat.in <- list()
for (f in names(file.list)){  
  dat.in[[f]] <- read.table(file.path(wd,f),header=TRUE,sep=",",fill = TRUE,na.strings = "-",blank.lines.skip = TRUE)#,encoding="iso-8859-1"
}
# summary(dat.in)
length(dat.in)

head(dat.in[[1]])


# dat.in <- dat.in[c(2,6,7)]
# summary(dat.in)
# head(dat.in[[1]])

dat.infile <- list()
for (l in names(dat.in)){ 
  colnames(dat.in[[l]])[1] <- "Sample.t" 
  dat.in[[l]]$Sample.t <- gsub(" ","",dat.in[[l]]$Sample.t)#remove space char
  dat.in[[l]]$Isotope <- gsub(" ","",dat.in[[l]]$Isotope)#remove space char
  iso.t <- unique(dat.in[[l]]$Isotope) 
  iso <- iso.t[nchar(iso.t) > 0]
  sample.t <- unique(dat.in[[l]]$Sample.t)
  sample <- sample.t[nchar(sample.t) > 0 & nchar(sample.t) <25]# !!! check if nchar 25 makes sense
  infile <- dat.in[[l]][1:length(rep(sample,each=length(iso))),]
  infile$RepSample <- rep(sample,each=length(iso))
  pat="Ag|Cu|Zn|Mn|Ni|Cd|Pb|Ca|Fe|Mg|Na|K"
  chem <- colnames(infile)[grep(pat,colnames(infile))]
  sam.l <- list()
  for (s in chem){
    sam.df <- as.data.frame(infile[,c("Sample.t","RepSample","Isotope",s)])
    vals <- unique(sam.df$Isotope)
    vals.df <- as.data.frame(unique(sam.df$RepSample))
    for (v in vals){
      vals.df <- cbind(vals.df,sam.df[which(sam.df$Isotope == v),s])
    }
    colnames(vals.df) <- c("Sample",vals)
    sam.l[[s]] <- vals.df
  }
  dat.infile[[l]] <- sam.l
}
summary(dat.infile)
length(dat.infile)

head(dat.infile[[2]])


for (l in names(dat.infile)){
  for (ll in names(dat.infile[[l]])){
    chem.col <- sample(c(28,13,51,40,46,12),1,replace = FALSE)
    wb.res <- loadWorkbook(file.path(paste(wd,l,"formatted.xlsx",sep="")), create = TRUE)
    createSheet(wb.res, name=ll)
    writeWorksheet(wb.res, dat.infile[[l]][[ll]],ll,header=TRUE)
#     setSheetColor(wb.res,ll,ifelse(grepl("Ag",ll),23,ifelse(grepl("Cu",ll),57,28)))
    setSheetColor(wb.res,ll,chem.col)
    saveWorkbook(wb.res)
  }
}






