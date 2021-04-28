# imposto la cartella dove sono i file scaricati
cartella <- "METTI IL PERCORSO QUI"
setwd(cartella)

# scarico lo script magico
source("https://raw.githubusercontent.com/pqstri/covaximx_converter/main/importer.R")

# faccio la magia
dati <- do_magic()

# salvo i dati scaricati in spss
haven::write_sav(dati, format(Sys.time(), "CovaxiMS_%d%b%Y_%He%M.sav"))

# o in Excel
openxlsx::write.xlsx(dati, format(Sys.time(), "CovaxiMS_%d%b%Y_%He%M.xlsx"))