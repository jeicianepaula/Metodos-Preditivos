#ATIVANDO AS BIBLIOTECAS *Caso necessite instalar utilize o comando install.packages()*
library(Metrics)
library(xlsx)
library(fpp2)
library(rugarch)
library(erer)

#--------------------------------------------------EXPORTACAODE DADOS-------------------------------------------------------------------------------------------
Ex = read.xlsx(file.choose(), header = T,sheetIndex = 1)
Data= Ex[,1]
Dados= Ex[,-1]
#--------------------------------------------------FORMACAO DA SERIE-------------------------------------------------------------------------------------------

serie = ts(Dados[,1], start = c(2000,1), end = c(2020,2),frequency = 12) %>% window(end = c(2012,12))

#---------------------------------------------------------HOLT----------------------------------------------------------------------------------------------

HOLT = function (x,Data_IN,Data_OUT,Dados_reais_OUT){
  
  HOLT =holt(x, h=24)
  print(HOLT[["model"]])
  
Previsão_D= data.frame(Data_IN, Dados_reais_IN = (HOLT[["x"]]), Previsão_IN = (HOLT[["fitted"]]))
Previsão_F= data.frame(Data_OUT, Dados_reais_OUT, Previsão_OUT = (HOLT[["mean"]]))

#Erro_IN
Erro_IN = data.frame(MAPE_IN = mape(Previsão_D$Dados_reais_IN, Previsão_D$Previsão_IN)*100,
                     MAE_IN = mae(Previsão_D$Dados_reais_IN, Previsão_D$Previsão_IN),
                     RMSE_IN =rmse(Previsão_D$Dados_reais_IN, Previsão_D$Previsão_IN))

#Erro_OUT                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         
Erro_OUT = data.frame(MAPE_OUT = mape(Previsão_F$Dados_reais_OUT, Previsão_F$Previsão_OUT)*100,
                      MAE_OUT = mae(Previsão_F$Dados_reais_OUT, Previsão_F$Previsão_OUT),
                      RMSE_OUT = rmse(Previsão_F$Dados_reais_OUT, Previsão_F$Previsão_OUT))

Parametros= c(HOLT[["model"]][["loglik"]], HOLT[["model"]][["aic"]], HOLT[["model"]][["bic"]], HOLT[["model"]][["aicc"]],
                       HOLT[["model"]][["par"]][["alpha"]], HOLT[["model"]][["par"]][["beta"]],
                       HOLT[["model"]][["par"]][["l"]], HOLT[["model"]][["par"]][["b"]],
                       HOLT[["model"]][["sigma2"]])

Parametros= matrix(Parametros, ncol = 2, nrow = 9)
Parametros[,1] = c("loglik", "aic","aicc","bic", "alpha", "beta", "l", "beta","sigma2" )

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

saveWorkbook(wb, "Holt sem amortecimento.xlsx")

}
  
HOLT(serie, Data[1:156], Data[157:180],Dados[157:180,1])


