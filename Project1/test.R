library(readxl)
mydata <- read_excel("U:/Projects/FOLR1/Data/Raw/14JAN2016 FOLR1 RPA and IUO Stained Cases.xlsx", 
                     sheet = "Sheet3",
                     skip = 5)

mydata <- mydata[-1,]
colnames(mydata) <- c('id', 'iuo.percent', 'iuo.intensity', 'iuo.bkgd', 'iuo.statu50', 'iuo.status75', 'rpa.percent', 
                      'rpa.intensity', 'rpa.status', 'rpa.status50', 'rpa.status75', 'vendor', 'Tumor.type')

Get_score <- function(data){
    if (grepl('-', data) > 0) {
        result <- floor(mean(as.numeric(unlist(strsplit(data,'-')))))
    }
    else{
        result <- as.numeric(data)
    }
    return(result)
}

