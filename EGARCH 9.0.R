#ATIVANDO AS BIBLIOTECAS *Caso necessite instalar utilize o comando install.packages()*
library(Metrics)
library(xlsx)
library(fpp2)
library(rugarch)
library(erer)

#--------------------------------------------------EXPORTACAODE DADOS-------------------------------------------------------------------------------------------
#Expota e separa os dados tipo date dos dados para formação das series temporais
Ex = read.xlsx(file.choose(), header = T,sheetIndex = 1)
Data= Ex[,1]
Dados= Ex[,-1]

#--------------------------------------------------FORMACAO DA SERIE-------------------------------------------------------------------------------------------
#Cria uma arquivo no fprmato ts (serie temporal)
serie = ts(Dados[,1], start = c(2000,1), end = c(2020,2),frequency = 12) %>% window(end = c(2012,12))

#---------------------------------------------------------EGARCH----------------------------------------------------------------------------------------------
EGARCH= function (x,Data_IN,Data_OUT,Dados_reais_OUT){
#Faz a modelagem e previsão
egarchspec = ugarchspec(variance.model = list(model = "eGARCH", garchOrder = c(1,1)), mean.model = list(armaOrder = c (1,1)))
EGARCH= ugarchfit(egarchspec, x, solver ='hybrid')
prevegarch = ugarchforecast(EGARCH, n.ahead= 24)

#cria dois arquivos do tipo data.frame para armazenar as informacoes de previsao dentro e fora da amostra
Previsão_D = data.frame(Data_IN, Dados_reais_IN = (prevegarch@model[["modeldata"]][["data"]]), Previsão_IN = (EGARCH@fit[["fitted.values"]]))
Previsão_F= data.frame(Data_OUT, Dados_reais_OUT, Previsão_OUT = prevegarch@forecast[["seriesFor"]])
names(Previsão_F)[3]=c("Previsão_OUT")

#calcula os erros de previsao dentro e fora da amostra pelas metricas de erros e amarzena em um data.frame
Previsão = prevegarch@forecast[["seriesFor"]]
Erro_IN = data.frame(MAPE_IN = mape(Previsão_D$Dados_reais_IN, Previsão_D$Previsão_IN)*100,
                     MAE_IN = mae(Previsão_D$Dados_reais_IN, Previsão_D$Previsão_IN),
                     RMSE_IN =rmse(Previsão_D$Dados_reais_IN, Previsão_D$Previsão_IN))
                     print(Erro_IN)
Erro_OUT = data.frame(MAPE_OUT = mape(Previsão_F$Dados_reais_OUT, Previsão)*100,
                      MAE_OUT = mae(Previsão_F$Dados_reais_OUT,Previsão),
                      RMSE_OUT = rmse(Previsão_F$Dados_reais_OUT,Previsão))
                      print(Erro_OUT)

#Armazena os parâmetros da momdelagem em um data.frame
Parametros = data.frame(EGARCH@fit[["coef"]])

#o planilhador cria e exporta os dados obtidos para um planilha de excel
##PLANILHA MODELO

wb = createWorkbook()

sheet = createSheet(wb, "Prev.in")
cs1 <- CellStyle(wb) + Font(wb, isBold=TRUE) + Border()  # header
addDataFrame(Previsão_D, sheet=sheet, startColumn=1, row.names= FALSE, colnamesStyle=cs1)
addDataFrame(Erro_IN, sheet=sheet, startColumn=4, row.names= FALSE, colnamesStyle=cs1)

sheet = createSheet(wb, "Prev.out")
addDataFrame(Previsão_F, sheet=sheet, startColumn=1, row.names= FALSE, colnamesStyle=cs1)
addDataFrame(Erro_OUT, sheet=sheet, startColumn=4, row.names= FALSE, colnamesStyle=cs1)

sheet = createSheet(wb, "Parâmetros")
addDataFrame(Parametros, sheet=sheet, startColumn=1, col.names= FALSE, colnamesStyle=cs1)

saveWorkbook(wb, "EGARCH.xlsx")

return(Previsão)

}

EGARCH(serie,Data[1:156],Data[157:180],Dados[157:180,1])
