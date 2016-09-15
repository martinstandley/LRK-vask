## Dette programmet er utviklet i Difis prosjekt for bedre anskaffelsesstatistikk
## Kontakt: Martin Standley, 91383954 martin.standley@difi.no
## Programmet leser inn ett uttrekk av regnskapsinformasjon i csv format, vasker og skrive ut vasket data i
## et nytt csv fil med "-vasket" tilført til filnavnet.
## Filnavnet spesifiseres i koden nedenfor - variabel: dataSettNavn

require ("dplyr")
require ("stringr")
require ("readxl")

library(dplyr)
library(stringr)
library(readxl)

## spesifiser datasettnavn her (uten .csv - det legges på automatisk)
dataSettNavn <- "Analyserapport-kopi-2016-01-20"

## forventede colNames: 
## Bilagsnr, BA(T), Bilagsdato, Koststed, Koststed(T), Saknr, Saknr.(T), Kap/post,Kap/post(T),
## Leverandør,Leverandør(T),Organisasjonsnr,Giro,Konto,Konto(T),Val,Kontantbeløp

dataFilInn <- paste(dataSettNavn, ".csv", sep="")
if (file.exists(dataFilInn)) {
    # Les in .csv fil hvis den finnes
    LRK <- read.table(dataFilInn, header=TRUE, dec=",", sep=";", stringsAsFactors = FALSE)
    ## rett opp kolonnenavn som ikke leses riktig av read.csv
    names(LRK) <- sub(".T.", "(T)", names(LRK)) 
    # fjern tusenskille blank i hoveddata - legges ut uten blank
    LRK$Kontantbeløp <- gsub(" ","", LRK$Kontantbeløp)
    # konverter fra comma desimal til punkt
    LRK$Kontantbeløp <- gsub(",",".", LRK$Kontantbeløp)
} else { 
    # prøv å lese en .xlsx fil med samme navn (Agresso produserer bare xlsx)
    colTypes <- c("text", "text","date","text","text","text","text","text","text","text","text","text","text","text","text","text","numeric")
    dataFilInn <- paste(dataSettNavn, ".xlsx", sep="")
    LRK <- read_excel(dataFilInn, col_types = colTypes, na=" ")
    # konverter til text med desimal punkt
    LRK$Kontantbeløp <- as.character(LRK$Kontantbeløp)
}
# NB: Kontantbeløp skal alltid være character med decimal punkt her

## siste linje er ofte summen - ta bort hvis bilagsnr er ugyldig
if (is.na(LRK[dim(LRK)[1], "Bilagsnr"] )) {LRK <- slice(LRK, 1:(dim(LRK)[1]-1))}

## Sett opp loggen og legg inn antall konteringer og beløpssum
vaskLog <- data.frame(postNr=0, Hendelse = "Vask start", Info = as.character(Sys.time()), stringsAsFactors = FALSE)
vaskLog <- rbind(vaskLog, c(0, "Antall konteringer", dim(LRK)[1]))
vaskLog <- rbind(vaskLog, c(0, "Sum av alle beløp", sub("\\.", ",", sum(as.numeric(LRK$Kontantbeløp)))))

## Fjern evt. "MVA" i orgnr
LRK$Organisasjonsnr <- sub("MVA", "", LRK$Organisasjonsnr)

# legg inn en ny kolonne: VaskeStatus
LRK$vaskeStatus <- "Ikke-vasket"

# spesifiser filteret for organisasjonsnavn (brukes til å skille fra personnavn). NB: må ikke forekomme i navn!
legalOrgStrings <- c("AS", "A/S", "BA", "ApS", "Ltd", "LTD", "ltd", "Agency", "AISBL", "irektorat", "verket", "styrels")
legalOrgStrings <- c(legalOrgStrings, "AB", "NUF", "SARL", "mbH", "kommune", "nstitut", "School", "kademie", "cademy")
legalOrgStrings <- c(legalOrgStrings, "epartment", "epartement", "Hotel", "inisterie", "inistry", "ylke")

# function returns true if name contains one or more of the legal strings or is only upper case
## (and is therefore assumed to be an organisation)
chkOrgName <- function (name) {
    res <- FALSE
    for (i in 1: length(legalOrgStrings)) {res <- res | grepl(legalOrgStrings[i], name)}
    res <- res | (toupper(name) == name)
    res
    }

## Gå gjennom datasett linje for linje og  rensk
vaskTeller <- 0
konkhensynTeller <- 0
DTeller <- 1


for (i in 1:dim(LRK)[1]) { 
    # rensk organisasjonsnummer og leverandør navn for personinfo
    orgNr <- LRK[i,"Organisasjonsnr"]
    orgName <- LRK[i, "Leverandør(T)"]
    # test for gyldig norsk orgnr (lengde =9, starter med 8 elelr 9 og resten er tall)
    if ((str_length(orgNr) != 9) | (!str_detect(orgNr,"[8-9][0-9]{8}"))) {
        ## hvis ikke norsk orgnr så må vi passe for personinfo
        ## hvis både leverandørnavn og orgnr er allerede tomme gjør ingenting
        if ((str_length(orgNr) !=0) | (str_length(LRK[i, "Leverandør(T)"]) != 0)) {
            ## sjekk om leverandørnavnser ut som bedrift
            if (chkOrgName(orgName)) {
                ## gjenkjent som orgnavn, men ikke gyldig orgnr - lag D-orgnr.
                ## legges inn senere ***
            } else {
                ## set orgnr til null og slett leverandør navn (kan inneholde personnavn)
                ## kunne senere legg inn oppslag i foretaksregisteret her senere
                vaskTeller <- vaskTeller+1
                vaskLog <- rbind(vaskLog, c(i, "Vasket ut", LRK[i, "Leverandør(T)"]))
                LRK[i,"Organisasjonsnr"] <- "00000000"
                LRK[i, "Leverandør(T)"] <- "*V* Slettet- person *V*"
                LRK[i,"vaskeStatus"] <- "Vasket"
            }
        }
    }
    ## Varsle konkurransehensyn hvis beløp>10000 og mindre enn 4 utbetalinger ila året    
    if ((as.numeric(LRK[i, "Kontantbeløp"]) > 10000) & (sum (LRK[,"Organisasjonsnr"] == LRK[i,"Organisasjonsnr"])< 4)) {
        LRK[i,"vaskeStatus"] <- "Sjekk konkurransehensyn"
        konkhensynTeller <- konkhensynTeller + 1
        }
} 
 
## antall vasket orgnr/navn og varsling om kokurransehensyn til log og avslutt
vaskLog <- rbind(vaskLog, c(i, "Antall vasket ut", as.character(vaskTeller)))
vaskLog <- rbind(vaskLog, c(i, "Antall konkurransehensyn", as.character(konkhensynTeller)))
vaskLog <- rbind(vaskLog, c(i, "Vask avsluttet", as.character(Sys.time())))

# legg tilbake desimal komma - merk dobbel escape på punktum (en escape for \)
LRK$Kontantbeløp <- gsub("\\.",",", LRK$Kontantbeløp)

## skriv ut vasket fil og logg som csv med inputfilens navn + "-vasket" og "-logg"
outputFileCSV <- paste (dataSettNavn, "-vasket.csv", sep="")
write.table(LRK, outputFileCSV, row.names=FALSE, sep=";")
write.table(vaskLog, paste (dataSettNavn, "-logg.csv", sep=""), row.names=FALSE, sep=";")