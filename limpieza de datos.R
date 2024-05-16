if (!require("dplyr")) {
  install.packages("dplyr")
  library("dplyr")}

if (!require("xlsx")) {
  install.packages("xlsx")
  library("xlsx")}

setwd("C:/Users/Usuario/Documentos/GITHUB/Data/Data procesada")
consultoras = read.csv(file.choose())

#DATA ADJUDICACIONES

directorio = "C:/Users/Usuario/Documentos/GITHUB/Data/Adjudicaciones"
archivos = list.files(directorio, pattern="*.xls", full.names=TRUE)
listas = lapply(archivos, read_excel)
dataadjudicaciones = as.data.frame(listas[1]) %>% mutate(EMPRESA = consultoras$EMPRESA[1])

for (i in 2:14){ 
  datos = as.data.frame(listas[i]) %>% mutate(EMPRESA=consultoras$EMPRESA[i])
  dataadjudicaciones = rbind(dataadjudicaciones,datos)
}

dataadjudicaciones = dataadjudicaciones %>% mutate_if(is.numeric,~replace(.,is.na(.),0)) %>% 
  mutate(MESBUENAPRO = format(as.Date(FECHABUENAPRO),"%B")) %>%
  mutate(FACTORMES = as.numeric(format(as.Date(FECHABUENAPRO),"%m"))) %>%
  mutate(AÑOBUENAPRO = as.numeric(format(as.Date(FECHABUENAPRO),"%Y"))) %>% 
  mutate(MONTOFINAL=MONTOADJUDICADOSOLES+MONTO_ADICIONAL_S2-MONTO_REDUCCION_S2) %>% 
  select(-c(1:3,5:8,11:15,19:32)) %>% mutate_all(.,~replace(.,is.na(.),"NO APLICA"))

write.xlsx(dataadjudicaciones,file="data procesada.xlsx",sheetName="ADJUDICACIONES",row.names = FALSE,append=FALSE)

#DATA ÓRDENES DE COMPRA

directorio = "C:/Users/Usuario/Documentos/GITHUB/Data/Órdenes de compra"
archivos = list.files(directorio, pattern="*.xls", full.names=TRUE)
listas = lapply(archivos, read_excel)
dataordenes = as.data.frame(listas[1]) %>% mutate(EMPRESA = consultoras$EMPRESA[1])

for (j in 2:14){ 
  datos = as.data.frame(listas[j]) %>% mutate(EMPRESA=consultoras$EMPRESA[j])
  dataordenes = rbind(dataordenes,datos)
}

dataordenes = dataordenes %>% mutate(MES = format(as.Date(FECHA_EMISION),"%B")) %>% 
  mutate(FACTORMES = as.numeric(format(as.Date(FECHA_EMISION),"%m"))) %>%
  mutate(AÑO = as.numeric(format(as.Date(FECHA_EMISION),"%Y"))) %>% select(-c(1,2,4,7,8))

write.xlsx(dataordenes,file="data procesada.xlsx",sheetName="ÓRDENES DE COMPRA",row.names = FALSE,append=TRUE)