#GO29006_data.clean.R
#Reads in and cleans data from GO29006 Preliminary Data Sample Tracker for QC and final transfer 29JUN2015_GNE edits.xlsx
#TRACKER file:  
#GO29006 Preliminary Data Sample Tracker for QC and final transfer 29JUN2015_GNE edits.xlsx.
#Test File
#Updated on 25AUG2015 for Product file.
#New data tracker: GO29006 Preliminary Data Sample Tracker for QC and final transfer 29JUN2015.xlsx


rm(list=ls())
library(openxlsx)
library(car)

#read-in data file
in.file<- "C:/Projects/TDxDataTransfers/GO29006/GO29006 Preliminary Data Tracker_07JAN2016.xlsx" 
password<- "15-GO29006"
library(RDCOMClient)
eApp<- COMCreate("Excel.Application")
wk<- eApp$Workbooks()$Open(Filename=in.file, Password=password)
tf<- tempfile()
wk$Sheets(1)$SaveAs(tf, 3)
input.dat<- read.delim(sprintf("%s.txt", tf), header = TRUE, sep = "\t")

time<-gsub(":","-",gsub(" ","-",Sys.time()))  #attach time to related output files.
filename0<-paste0("C:/Projects/TDxDataTransfers/GO29006/GO29006-TRACKER-",time,".RData")


source("C:/payam Rscripts/bye.empty.R")
input.dat<- bye.empty(input.dat)
save(input.dat,file=filename0)

#Remove extra column 
#input.dat<-input.dat[,1:41]

#Remove extra row 
input.dat<-input.dat[1:32,]


################################################################################################################################
################################################################################################################################
################################################################################################################################
#Remove carriage returns, line breaks, special character; double quotes, commas, and empty lines.
################################################################################################################################
################################################################################################################################
################################################################################################################################

COLnames0<- c("COMMENTS","Comments","Cytoplasm.Comments","Membrane.Comments","Primary.Specimen.Identifier.....Ventana.Accessioning..","PHARMA_ID","SPECIMEN_SOURCE") 
tmp.input.dat<- apply(input.dat[COLnames0],2,as.character)
tmp.input.dat<- apply(tmp.input.dat,2,function(y){toupper(y)})
tmp.input.dat<- apply(tmp.input.dat,2,function(y){gsub(pattern="[\r\n]",replacement=". ",y)})
tmp.input.dat<- apply(tmp.input.dat,2,function(y){gsub(pattern="[\"]",replacement="",y)})
tmp.input.dat<- apply(tmp.input.dat,2,function(y){gsub(pattern=",",replacement="",y)})
input.dat[COLnames0]<- tmp.input.dat


# You can add similar code lines for other possible changes. The only change you need is to define new COLnames.
COLnames1<- c("COMMENTS","Comments","Cytoplasm.Comments","Membrane.Comments","X..tumor","X..viable","X..necrotic","Cytoplasm.0","Cytoplasm.1.","Cytoplasm.2.",
              "Cytoplasm.3.","Cytoplasm.H.score","Membrane.0","Membrane.1.","Membrane.2.","Membrane.3.","Membrane.H.score","NaPi2b.Scoring.Assignment")  #For NA cleaning
input.dat[COLnames1][!is.na(input.dat[COLnames1]) & (input.dat[COLnames1]==". " | input.dat[COLnames1]=="" | input.dat[COLnames1]=="N/A" | input.dat[COLnames1]=="n/a"
                     | input.dat[COLnames1]=="na" | input.dat[COLnames1]=="NA")]<-NA
#
#COLnames2<- c("SPECIMEN_SOURCE")  #For "Not Provided" cleaning
#input.dat[COLnames2][!is.na(input.dat[COLnames2]) & (input.dat[COLnames2]==". " | input.dat[COLnames2]=="" | input.dat[COLnames2]=="N/A")]<-"NOT PROVIDED"

################################################################################################################################
################################################################################################################################
################################################################################################################################
#Wide data formatting.
################################################################################################################################
################################################################################################################################
################################################################################################################################

################################################################################################################################
# FFS Table Variables #
################################################################################################################################
input.dat$STUDYID<- "GO29006"
################################################################################################################################
input.dat$ZBNAM<- "VENT"
################################################################################################################################
input.dat$PATNUM<- as.numeric(as.character(input.dat$PATIENT_Enrollment.NUMBER))
################################################################################################################################
input.dat$VISIT<- "SCRN"
################################################################################################################################
input.dat$BOMTPT<- matrix(NA,dim(input.dat)[1],1)
################################################################################################################################
input.dat$SAMPLE.COLLECTION_DATE<- as.character(input.dat$SAMPLE.COLLECTION_DATE)

input.dat$BOMD<- as.Date(input.dat$SAMPLE.COLLECTION_DATE,format="%d-%b-%y")
input.dat$BOMD<- format(input.dat$BOMD,"%Y%m%d")
input.dat$BOMD[is.na(input.dat$BOMD)]<- 99999999
################################################################################################################################
input.dat$BOMTM<- matrix(NA,dim(input.dat)[1],1)
################################################################################################################################
input.dat$ACCSNM<- toupper(as.character(input.dat$Primary.Specimen.Identifier.....Ventana.Accessioning..))
input.dat$ACCSNM[input.dat$ACCSNM=="LS-13-21-3572, L-S 1A"]<- "LS-13-21-3572,L-S 1A"

input.dat$ACCSNM[is.na(input.dat$ACCSNM)]<-"Not Provided"
################################################################################################################################
input.dat$BOMREFID<-  toupper(as.character(input.dat$PHARMA_ID))
input.dat$BOMREFID<-ifelse(test=nchar(input.dat$BOMREFID)>8,
                           yes=apply(matrix(input.dat$BOMREFID,ncol = length(input.dat$BOMREFID),nrow = 1),2,
                                     function(x)strsplit(x," ")[[1]][1]),
                           no=input.dat$BOMREFID
                           )

################################################################################################################################
input.dat$BOMLOC<- toupper(as.character(input.dat$SPECIMEN_SOURCE))
################################################################################################################################
input.dat$ZBSPEC[grepl("slide",input.dat$SAMPLE_RECEIVED,ignore.case = TRUE)]<- "SLIDE"
input.dat$ZBSPEC[grepl("block",input.dat$SAMPLE_RECEIVED,ignore.case = TRUE)]<- "BLOCK"

################################################################################################################################
input.dat$ZBSPCCND[grepl("nbf",input.dat$Fixative,ignore.case = TRUE) | 
                     grepl("formalin",input.dat$Fixative,ignore.case = TRUE)]<- "FFPE"
input.dat$ZBSPCCND[grepl("unknown",input.dat$Fixative,ignore.case = TRUE)]<- "UNKNOWN"

################################################################################################################################
## FFS Table Variables (Appendix B) ##
################################################################################################################################

#####H&E########################################################################################################################

################################################################################################################################
input.dat$TUMINVP<- as.character(input.dat$X..tumor)
################################################################################################################################
input.dat$VIATUP<- as.numeric(as.character(input.dat$X..viable))
################################################################################################################################
input.dat$NECRPC<- as.numeric(as.character(input.dat$X..necrotic))
################################################################################################################################
input.dat$COMMENT<- toupper(as.character(input.dat$Comments))
################################################################################################################################

#####CYTOPLASM##################################################################################################################
input.dat$CYTINT0<- as.numeric(as.character(input.dat$Cytoplasm.0))
input.dat$CYTINT1<- as.numeric(as.character(input.dat$Cytoplasm.1.))
input.dat$CYTINT2<- as.numeric(as.character(input.dat$Cytoplasm.2.))
input.dat$CYTINT3<- as.numeric(as.character(input.dat$Cytoplasm.3.))
input.dat$CYSTOT<- as.numeric(as.character(input.dat$Cytoplasm.H.score))
input.dat$CYTSTC<- toupper(as.character(input.dat$Cytoplasm.Comments))
################################################################################################################################

#####MEMBRANE###################################################################################################################
input.dat$MEMINT0<- as.numeric(gsub("[\r\n]","",as.character(input.dat$Membrane.0)))
input.dat$MEMINT1<- as.numeric(gsub("[\r\n]","",as.character(input.dat$Membrane.1.)))
input.dat$MEMINT2<- as.numeric(gsub("[\r\n]","",as.character(input.dat$Membrane.2.)))
input.dat$MEMINT3<- as.numeric(gsub("[\r\n]","",as.character(input.dat$Membrane.3.)))
input.dat$MEMSTOT<- as.character(input.dat$Membrane.H.score)
input.dat$IHCSCORE<- gsub("[\r\n]","",as.character(input.dat$NaPi2b.Scoring.Assignment))
input.dat$IHCSCORE<- gsub(" ","",input.dat$IHCSCORE)
input.dat$IHCSCORE[input.dat$IHCSCORE=="3"]<- "3+"
input.dat$MEMSTAIN<- toupper(as.character(input.dat$Membrane.Comments))

################################################################################################################################
# FFS Table Variables (Continued) #
################################################################################################################################

input.dat$ZBASYID<- matrix(NA,dim(input.dat)[1],1)
################################################################################################################################

input.dat$BOMTD<- format(as.Date(as.character(input.dat$Date.Scored),format="%d-%b-%y"),format="%Y%m%d")
input.dat$BOMTD[is.na(input.dat$BOMTD)]<- "99999999"

################################################################################################################################
input.dat$BOMTTM<- matrix(NA,dim(input.dat)[1],1)
################################################################################################################################
input.dat$ZBDIL<- matrix(NA,dim(input.dat)[1],1)
################################################################################################################################

######ZBTESTCD##################################################################################################################

zbtestcd.list=c("TUMINVP","VIATUP","NECRPC","COMMENT","CYTINT0","CYTINT1","CYTINT2","CYTINT3","CYSTOT","CYTSTC",
                "MEMINT0","MEMINT1","MEMINT2","MEMINT3","MEMSTOT","IHCSCORE","MEMSTAIN")

ZBTESTCD<- matrix(zbtestcd.list,dim(input.dat)[1]*length(zbtestcd.list),1)

COLnames.IHC.CYT<- (ZBTESTCD=="CYTINT0" | ZBTESTCD=="CYTINT1" | ZBTESTCD=="CYTINT2" | ZBTESTCD=="CYTINT3" | ZBTESTCD=="CYSTOT")
COLnames.IHC.MEM<- (ZBTESTCD=="MEMINT0" | ZBTESTCD=="MEMINT1" | ZBTESTCD=="MEMINT2" | ZBTESTCD=="MEMINT3" | ZBTESTCD=="MEMSTOT" | ZBTESTCD=="IHCSCORE")
input.dat$HE.BOMREAS<- tolower(as.character(input.dat$COMMENT))
input.dat$IHC.CYT.BOMREAS<- tolower(as.character(input.dat$CYTSTC))
input.dat$IHC.MEM.BOMREAS<- tolower(as.character(input.dat$MEMSTAIN))

zbtestcd.HE<- ZBTESTCD=="TUMINVP" | ZBTESTCD=="VIATUP" | ZBTESTCD=="NECRPC" | ZBTESTCD=="COMMENT"

zbtestcd.IHC<- ZBTESTCD=="CYTINT0" | ZBTESTCD=="CYTINT1" | ZBTESTCD=="CYTINT2" | ZBTESTCD=="CYTINT3" | ZBTESTCD=="CYSTOT" | ZBTESTCD=="CYTSTC" | 
               ZBTESTCD=="MEMINT0" | ZBTESTCD=="MEMINT1" | ZBTESTCD=="MEMINT2" | ZBTESTCD=="MEMINT3" | ZBTESTCD=="MEMSTOT" | ZBTESTCD=="IHCSCORE" | 
               ZBTESTCD=="MEMSTAIN"


####HE.BOMREAS###################################################################################################################

input.dat$HE.BOMREAS<- ifelse(
  test=(grepl("no",input.dat$HE.BOMREAS, ignore.case = TRUE) & grepl("tissue",input.dat$HE.BOMREAS, ignore.case = TRUE) & !grepl("tumor",input.dat$HE.BOMREAS, ignore.case = TRUE)),
  yes="NOE",
  no=ifelse(
    test=(grepl("no",input.dat$HE.BOMREAS, ignore.case = TRUE) & grepl("tumor",input.dat$HE.BOMREAS, ignore.case = TRUE)) | 
          grepl("rejected",input.dat$HE.BOMREAS, ignore.case = TRUE) | 
         (grepl("stain",input.dat$HE.BOMREAS, ignore.case = TRUE) & grepl("slide",input.dat$HE.BOMREAS, ignore.case = TRUE)) | 
          grepl("h&e",input.dat$HE.BOMREAS, ignore.case = TRUE) | grepl("fell",input.dat$HE.BOMREAS, ignore.case = TRUE) |
      grepl("falling",input.dat$HE.BOMREAS, ignore.case = TRUE) | grepl("off",input.dat$HE.BOMREAS, ignore.case = TRUE) , 
    yes="NSR" ,
    no=ifelse(
      test=is.na(input.dat$HE.BOMREAS),
      yes="NOE",
      no="NSR"
    )
  ) 
)

#input.dat$HE.BOMREAS<- "NOE"

######IHC.BOMREAS##############################################################################################################

#input.dat$IHC.BOMREAS<- "NOE"
input.dat$IHC.CYT.BOMREAS<- ifelse(
  test= (grepl("no",input.dat$IHC.CYT.BOMREAS, ignore.case = TRUE) & grepl("tissue",input.dat$IHC.CYT.BOMREAS, ignore.case = TRUE)),
  yes="NOE",
  no=ifelse(
    test=(grepl("no",input.dat$IHC.CYT.BOMREAS, ignore.case = TRUE) & grepl("tumor",input.dat$IHC.CYT.BOMREAS, ignore.case = TRUE)) | 
      grepl("rejected",input.dat$IHC.CYT.BOMREAS, ignore.case = TRUE) | 
      (grepl("stain",input.dat$IHC.CYT.BOMREAS, ignore.case = TRUE) & grepl("slide",input.dat$IHC.CYT.BOMREAS, ignore.case = TRUE)) | 
      grepl("h&e",input.dat$IHC.CYT.BOMREAS, ignore.case = TRUE) | grepl("fell",input.dat$IHC.CYT.BOMREAS, ignore.case = TRUE) |
      grepl("falling",input.dat$HE.BOMREAS, ignore.case = TRUE) | grepl("off",input.dat$IHC.CYT.BOMREAS, ignore.case = TRUE) , 
    yes="NSR" ,
    no=ifelse(
      test=is.na(input.dat$IHC.CYT.BOMREAS),
      yes="NOE",
      no="NSR"
    )
  ) 
)


input.dat$IHC.MEM.BOMREAS<- ifelse(
  test=(grepl("no",input.dat$IHC.MEM.BOMREAS, ignore.case = TRUE) & grepl("tissue",input.dat$IHC.MEM.BOMREAS, ignore.case = TRUE)),
  yes="NOE",
  no=ifelse(
    test=(grepl("no",input.dat$IHC.MEM.BOMREAS, ignore.case = TRUE) & grepl("tumor",input.dat$IHC.MEM.BOMREAS, ignore.case = TRUE)) | 
      grepl("rejected",input.dat$IHC.MEM.BOMREAS, ignore.case = TRUE) | 
      (grepl("stain",input.dat$IHC.MEM.BOMREAS, ignore.case = TRUE) & grepl("slide",input.dat$IHC.MEM.BOMREAS, ignore.case = TRUE)) | 
      grepl("h&e",input.dat$IHC.MEM.BOMREAS, ignore.case = TRUE) | grepl("fell",input.dat$IHC.MEM.BOMREAS, ignore.case = TRUE) |
      grepl("falling",input.dat$HE.BOMREAS, ignore.case = TRUE) | grepl("off",input.dat$IHC.MEM.BOMREAS, ignore.case = TRUE) , 
    yes="NSR" ,
    no=ifelse(
      test=is.na(input.dat$IHC.MEM.BOMREAS),
      yes="NOE",
      no="NSR"
    )
  ) 
)


################################################################################################################################
################################################################################################################################
################################################################################################################################
#Wide to long data conversion.
################################################################################################################################
################################################################################################################################
################################################################################################################################

ID2<- matrix(NA,dim(input.dat)[1]*length(zbtestcd.list),1)
ZBMETHOD<- matrix(NA,dim(input.dat)[1]*length(zbtestcd.list),1)
ZBTARGET<- matrix(NA,dim(input.dat)[1]*length(zbtestcd.list),1)
ZBSTARG<- matrix(NA,dim(input.dat)[1]*length(zbtestcd.list),1)
BOMORESU<- matrix(NA,dim(input.dat)[1]*length(zbtestcd.list),1)
BOMRESN<- matrix(NA,dim(input.dat)[1]*length(zbtestcd.list),1)
BOMRESC<- matrix(NA,dim(input.dat)[1]*length(zbtestcd.list),1)
BOMSTAT<- matrix(NA,dim(input.dat)[1]*length(zbtestcd.list),1)
BOMREAS<- matrix(NA,dim(input.dat)[1]*length(zbtestcd.list),1)

####ID2#########################################################################################################################

ID2<- rep(input.dat$ACCSNM,each=length(zbtestcd.list))

####ZBMETHOD######################################################################################################################

ZBMETHOD[zbtestcd.HE]<- "HEMATOXYLIN & EOSIN STAIN"
ZBMETHOD[zbtestcd.IHC]<- "IHC"

####ZBTARGET######################################################################################################################

ZBTARGET[zbtestcd.HE]<- "NOT APPLICABLE"
ZBTARGET[zbtestcd.IHC]<- "SLC34A2"

####ZBSTARG######################################################################################################################

ZBSTARG[zbtestcd.HE]<- "NOT APPLICABLE"
ZBSTARG[zbtestcd.IHC]<- "CLONE 10H1"

####BOMORESU######################################################################################################################

bomoresu.percent<- (ZBTESTCD=="TUMINVP" | ZBTESTCD=="VIATUP" | ZBTESTCD=="NECRPC" | ZBTESTCD=="CYTINT0" | ZBTESTCD=="CYTINT1" | ZBTESTCD=="CYTINT2" | 
                    ZBTESTCD=="CYTINT3" | ZBTESTCD=="MEMINT0" | ZBTESTCD=="MEMINT1" | ZBTESTCD=="MEMINT2" | ZBTESTCD=="MEMINT3")

bomoresu.na<- ZBTESTCD=="COMMENT" | (ZBTESTCD=="CYTSTC") | (ZBTESTCD=="MEMSTAIN")

bomoresu.unitless<- (ZBTESTCD=="CYSTOT") | (ZBTESTCD=="MEMSTOT") |  (ZBTESTCD=="IHCSCORE") 

BOMORESU[bomoresu.percent]<- "%"
BOMORESU[bomoresu.unitless]<- NA
BOMORESU[bomoresu.na]<- NA

####BOMRESC and BOMRESN######################################################################################################################

BOMRESC[(ZBTESTCD=="TUMINVP") | (ZBTESTCD=="VIATUP") | (ZBTESTCD=="NECRPC")]<- NA

BOMRESN[(ZBTESTCD=="VIATUP")]<- as.numeric(input.dat$VIATUP)

BOMRESN[(ZBTESTCD=="NECRPC")]<- as.numeric(input.dat$NECRPC)

BOMRESN[(ZBTESTCD=="COMMENT")]<- NA
BOMRESC[(ZBTESTCD=="COMMENT")]<- input.dat$COMMENT

############################################

BOMRESN[(ZBTESTCD=="TUMINVP")]<- ifelse (test= !is.na(input.dat$TUMINVP),
                                         yes=ifelse( test=input.dat$TUMINVP=="<1",
                                                     yes= NA,
                                                     no= ifelse( test=grepl("[^0-9]",input.dat$TUMINVP),
                                                                 yes=as.numeric(gsub("[^0-9]","",input.dat$TUMINVP)),
                                                                 no=as.numeric(gsub("[^0-9]","",input.dat$TUMINVP))
                                                     )
                                         ),
                                         no=BOMRESN[(ZBTESTCD=="TUMINVP")]
)

############################################

BOMRESC[(ZBTESTCD=="TUMINVP")]<- ifelse (test= !is.na(input.dat$TUMINVP),
                                         yes=ifelse( test=input.dat$TUMINVP=="<1",
                                                     yes= input.dat$TUMINVP,
                                                     no= ifelse( test=grepl("[^0-9]",input.dat$TUMINVP),
                                                                 yes=input.dat$TUMINVP,
                                                                 no=NA
                                                     )
                                         ),
                                         no=BOMRESC[(ZBTESTCD=="TUMINVP")]
)

####################################################################################################################################

BOMRESC[(ZBTESTCD=="CYTINT0") | (ZBTESTCD=="CYTINT1") | (ZBTESTCD=="CYTINT2") | (ZBTESTCD=="CYTINT3") | (ZBTESTCD=="CYSTOT") |
          (ZBTESTCD=="MEMINT0") | (ZBTESTCD=="MEMINT1") | (ZBTESTCD=="MEMINT2") | (ZBTESTCD=="MEMINT3") | (ZBTESTCD=="MEMSTOT")]<- NA

############################################

BOMRESN[(ZBTESTCD=="CYTINT0")]<- ifelse (test= !is.na(input.dat$CYTINT0),
                                          yes=ifelse( test=input.dat$CYTINT0=="<1",
                                                      yes= NA,
                                                      no= ifelse( test=grepl("[^0-9]",input.dat$CYTINT0),
                                                                  yes=as.numeric(gsub("[^0-9]","",input.dat$CYTINT0)),
                                                                  no=as.numeric(gsub("[^0-9]","",input.dat$CYTINT0))
                                                      )
                                          ),
                                          no=BOMRESN[(ZBTESTCD=="CYTINT0")]
)

############################################

BOMRESC[(ZBTESTCD=="CYTINT0")]<- ifelse (test= !is.na(input.dat$CYTINT0),
                                          yes=ifelse( test=input.dat$CYTINT0=="<1",
                                                      yes= input.dat$CYTINT0,
                                                      no= ifelse( test=grepl("[^0-9]",input.dat$CYTINT0),
                                                                  yes=input.dat$CYTINT0,
                                                                  no=NA
                                                      )
                                          ),
                                          no=BOMRESC[(ZBTESTCD=="CYTINT0")]
)

############################################

BOMRESN[(ZBTESTCD=="CYTINT1")]<- ifelse (test= !is.na(input.dat$CYTINT1),
                                         yes=ifelse( test=input.dat$CYTINT1=="<1",
                                                     yes= NA,
                                                     no= ifelse( test=grepl("[^0-9]",input.dat$CYTINT1),
                                                                 yes=as.numeric(gsub("[^0-9]","",input.dat$CYTINT1)),
                                                                 no=as.numeric(gsub("[^0-9]","",input.dat$CYTINT1))
                                                     )
                                         ),
                                         no=BOMRESN[(ZBTESTCD=="CYTINT1")]
)

############################################

BOMRESC[(ZBTESTCD=="CYTINT1")]<- ifelse (test= !is.na(input.dat$CYTINT1),
                                         yes=ifelse( test=input.dat$CYTINT1=="<1",
                                                     yes= input.dat$CYTINT1,
                                                     no= ifelse( test=grepl("[^0-9]",input.dat$CYTINT1),
                                                                 yes=input.dat$CYTINT1,
                                                                 no=NA
                                                     )
                                         ),
                                         no=BOMRESC[(ZBTESTCD=="CYTINT1")]
)

############################################

BOMRESN[(ZBTESTCD=="CYTINT2")]<- ifelse (test= !is.na(input.dat$CYTINT2),
                                         yes=ifelse( test=input.dat$CYTINT2=="<1",
                                                     yes= NA,
                                                     no= ifelse( test=grepl("[^0-9]",input.dat$CYTINT2),
                                                                 yes=as.numeric(gsub("[^0-9]","",input.dat$CYTINT2)),
                                                                 no=as.numeric(gsub("[^0-9]","",input.dat$CYTINT2))
                                                     )
                                         ),
                                         no=BOMRESN[(ZBTESTCD=="CYTINT2")]
)

############################################

BOMRESC[(ZBTESTCD=="CYTINT2")]<- ifelse (test= !is.na(input.dat$CYTINT2),
                                         yes=ifelse( test=input.dat$CYTINT2=="<1",
                                                     yes= input.dat$CYTINT2,
                                                     no= ifelse( test=grepl("[^0-9]",input.dat$CYTINT2),
                                                                 yes=input.dat$CYTINT2,
                                                                 no=NA
                                                     )
                                         ),
                                         no=BOMRESC[(ZBTESTCD=="CYTINT2")]
)

############################################

BOMRESN[(ZBTESTCD=="CYTINT3")]<- ifelse (test= !is.na(input.dat$CYTINT3),
                                         yes=ifelse( test=input.dat$CYTINT3=="<1",
                                                     yes= NA,
                                                     no= ifelse( test=grepl("[^0-9]",input.dat$CYTINT3),
                                                                 yes=as.numeric(gsub("[^0-9]","",input.dat$CYTINT3)),
                                                                 no=as.numeric(gsub("[^0-9]","",input.dat$CYTINT3))
                                                     )
                                         ),
                                         no=BOMRESN[(ZBTESTCD=="CYTINT3")]
)

############################################

BOMRESC[(ZBTESTCD=="CYTINT3")]<- ifelse (test= !is.na(input.dat$CYTINT3),
                                         yes=ifelse( test=input.dat$CYTINT3=="<1",
                                                     yes= input.dat$CYTINT3,
                                                     no= ifelse( test=grepl("[^0-9]",input.dat$CYTINT3),
                                                                 yes=input.dat$CYTINT3,
                                                                 no=NA
                                                     )
                                         ),
                                         no=BOMRESC[(ZBTESTCD=="CYTINT3")]
)

############################################

BOMRESN[(ZBTESTCD=="CYSTOT")]<- ifelse (test= !is.na(input.dat$CYSTOT),
                                         yes=ifelse( test=input.dat$CYSTOT=="<1",
                                                     yes= NA,
                                                     no= ifelse( test=grepl("[^0-9]",input.dat$CYSTOT),
                                                                 yes=as.numeric(gsub("[^0-9]","",input.dat$CYSTOT)),
                                                                 no=as.numeric(gsub("[^0-9]","",input.dat$CYSTOT))
                                                     )
                                         ),
                                         no=BOMRESN[(ZBTESTCD=="CYSTOT")]
)

############################################

BOMRESC[(ZBTESTCD=="CYSTOT")]<- ifelse (test= !is.na(input.dat$CYSTOT),
                                         yes=ifelse( test=input.dat$CYSTOT=="<1",
                                                     yes= input.dat$CYSTOT,
                                                     no= ifelse( test=grepl("[^0-9]",input.dat$CYSTOT),
                                                                 yes=input.dat$CYSTOT,
                                                                 no=NA
                                                     )
                                         ),
                                         no=BOMRESC[(ZBTESTCD=="CYSTOT")]
)

#############################################

BOMRESC[(ZBTESTCD=="CYTSTC")]<- input.dat$CYTSTC
BOMRESN[(ZBTESTCD=="CYTSTC")]<- NA

####################################################################################################################################

BOMRESN[(ZBTESTCD=="MEMINT0")]<- input.dat$MEMINT0

BOMRESN[(ZBTESTCD=="MEMINT1")]<- input.dat$MEMINT1

BOMRESN[(ZBTESTCD=="MEMINT2")]<- input.dat$MEMINT2

BOMRESN[(ZBTESTCD=="MEMINT3")]<- input.dat$MEMINT3

BOMRESN[(ZBTESTCD=="MEMSTOT")]<- ifelse (test= !is.na(input.dat$MEMSTOT),
                                         yes=ifelse( test=input.dat$MEMSTOT=="<1",
                                                     yes= NA,
                                                     no= ifelse( test=grepl("[^0-9]",input.dat$MEMSTOT),
                                                                 yes=as.numeric(gsub("[^0-9]","",input.dat$MEMSTOT)),
                                                                 no=as.numeric(gsub("[^0-9]","",input.dat$MEMSTOT))
                                                     )
                                         ),
                                         no=BOMRESN[(ZBTESTCD=="MEMSTOT")]
)

############################################

BOMRESC[(ZBTESTCD=="MEMSTOT")]<- ifelse (test= !is.na(input.dat$MEMSTOT),
                                         yes=ifelse( test=input.dat$MEMSTOT=="<1",
                                                     yes= input.dat$MEMSTOT,
                                                     no= ifelse( test=grepl("[^0-9]",input.dat$MEMSTOT),
                                                                 yes=input.dat$MEMSTOT,
                                                                 no=NA
                                                     )
                                         ),
                                         no=BOMRESC[(ZBTESTCD=="MEMSTOT")]
)

####################################################################################################################################

BOMRESC[(ZBTESTCD=="IHCSCORE")]<- input.dat$IHCSCORE
BOMRESN[(ZBTESTCD=="IHCSCORE")]<- NA

BOMRESC[(ZBTESTCD=="MEMSTAIN")]<- input.dat$MEMSTAIN
BOMRESN[(ZBTESTCD=="MEMSTAIN")]<- NA



#####BOMSTAT and BOMREAS##########################################################################################################################

bomstat.nd<- (ZBTESTCD=="COMMENT") | (ZBTESTCD=="CYTSTC") | (ZBTESTCD=="MEMSTAIN") |(ZBTESTCD=="TUMINVP") | (ZBTESTCD=="VIATUP") | (ZBTESTCD=="NECRPC") |
  (ZBTESTCD=="CYTINT0") | (ZBTESTCD=="CYTINT1") | (ZBTESTCD=="CYTINT2") | (ZBTESTCD=="CYTINT3") | (ZBTESTCD=="CYSTOT") | 
  (ZBTESTCD=="MEMINT0") | (ZBTESTCD=="MEMINT1") | (ZBTESTCD=="MEMINT2") | (ZBTESTCD=="MEMINT3") | (ZBTESTCD=="MEMSTOT") | (ZBTESTCD=="IHCSCORE") 


bomres.HE.nd<- (ZBTESTCD=="COMMENT") | (ZBTESTCD=="TUMINVP") | (ZBTESTCD=="VIATUP") | (ZBTESTCD=="NECRPC") 

bomres.IHC.mem.nd<-  (ZBTESTCD=="MEMSTAIN") | (ZBTESTCD=="MEMINT0") | (ZBTESTCD=="MEMINT1") | (ZBTESTCD=="MEMINT2") | (ZBTESTCD=="MEMINT3") | (ZBTESTCD=="MEMSTOT") | 
                 (ZBTESTCD=="IHCSCORE") 

bomres.IHC.cyt.nd<- (ZBTESTCD=="CYTSTC") | (ZBTESTCD=="CYTINT0") | (ZBTESTCD=="CYTINT1") | (ZBTESTCD=="CYTINT2") | (ZBTESTCD=="CYTINT3") | (ZBTESTCD=="CYSTOT") 

####################################################################

BOMSTAT[bomstat.nd]<- ifelse(
  test=(is.na(BOMRESC[bomstat.nd]) & is.na(BOMRESN[bomstat.nd])),
  yes="ND",
  no=NA
)

BOMREAS[bomres.HE.nd]<- ifelse(
  test=(is.na(BOMRESC[bomres.HE.nd]) & is.na(BOMRESN[bomres.HE.nd])),
  yes=rep(input.dat$HE.BOMREAS,each=length(bomres.HE.nd[bomres.HE.nd==TRUE])/dim(input.dat)[1]), 
  no=NA
)

BOMREAS[bomres.IHC.cyt.nd]<- ifelse(
  test=(is.na(BOMRESC[bomres.IHC.cyt.nd]) & is.na(BOMRESN[bomres.IHC.cyt.nd])),
  yes=rep(input.dat$IHC.CYT.BOMREAS,each=length(bomres.IHC.cyt.nd[bomres.IHC.cyt.nd==TRUE])/dim(input.dat)[1]),
  no=NA
)

BOMREAS[bomres.IHC.mem.nd]<- ifelse(
  test=(is.na(BOMRESC[bomres.IHC.mem.nd]) & is.na(BOMRESN[bomres.IHC.mem.nd])),
  yes=rep(input.dat$IHC.MEM.BOMREAS,each=length(bomres.IHC.mem.nd[bomres.IHC.mem.nd==TRUE])/dim(input.dat)[1]),
  no=NA
)

##################################################################################

match.index=match(ID2,input.dat$ACCSNM)
first=rbind(input.dat[match.index,])

attach(first)
dat=data.frame(cbind(STUDYID,ZBNAM,PATNUM,VISIT,BOMTPT,BOMD,BOMTM,ACCSNM,BOMREFID,BOMLOC,ZBSPEC,ZBSPCCND,ZBMETHOD,ZBTARGET,ZBSTARG,ZBTESTCD,
                     BOMRESN,BOMORESU,BOMRESC,BOMSTAT,BOMREAS,ZBASYID,BOMTD,BOMTTM,ZBDIL))
detach(first)
names(dat)=c("STUDYID","ZBNAM","PATNUM","VISIT","BOMTPT","BOMD","BOMTM","ACCSNM","BOMREFID","BOMLOC","ZBSPEC","ZBSPCCND","ZBMETHOD",
             "ZBTARGET","ZBSTARG","ZBTESTCD","BOMRESN","BOMORESU","BOMRESC","BOMSTAT","BOMREAS","ZBASYID","BOMTD","BOMTTM","ZBDIL")

################################################################################################
dat.long<- dat
################################################################################################
#Save and write.
################################################################################################

# attach time to output file
time<-gsub(":","-",gsub(" ","-",Sys.time())) 

filename1<-paste0("C:/Projects/TDxDataTransfers/GO29006/GO29006-PRODUCTION-",time,".RData")
save(dat.long,file=filename1)

filename2<-paste0("C:/Projects/TDxDataTransfers/GO29006/GO29006-PRODUCTION-",time,".xlsx")
write.xlsx(dat.long, file=filename2,showNA=FALSE,col.names=TRUE,row.names=FALSE)

