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

#---------------------------------------------------------HOLT-W MULTIPLICATIVO----------------------------------------------------------------------------------------------
HOLTWM = function (x,Data_IN,Data_OUT,Dados_reais_OUT,h){
  
#Faz a modelagem e previsão  
HOLTWM = hw(x,seasonal = "multiplicative", h=h)
print(HOLTWM[["model"]])

#cria dois arquivos do tipo data.frame para armazenar as informacoes de previsao dentro e fora da amostra
Previsão_D = data.frame(Data_IN, Dados_reais_IN = (HOLTWM[["x"]]), Previsão_IN = (HOLTWM[["fitted"]])) 
Previsão_F= data.frame(Data_OUT, Dados_reais_OUT, Previsão_OUT = (HOLTWM[["mean"]]))

#calcula os erros de previsao dentro e fora da amostra pelas metricas de erros e amarzena em um data.frame
Erro_IN = data.frame(MAPE_IN = mape(Previsão_D$Dados_reais_IN, Previsão_D$Previsão_IN)*100,
                     MAE_IN = mae(Previsão_D$Dados_reais_IN, Previsão_D$Previsão_IN),
                     RMSE_IN =rmse(Previsão_D$Dados_reais_IN, Previsão_D$Previsão_IN))
                     print(Erro_IN)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               
Erro_OUT = data.frame(MAPE_OUT = mape(Previsão_F$Dados_reais_OUT, Previsão_F$Previsão_OUT)*100,
                      MAE_OUT = mae(Previsão_F$Dados_reais_OUT, Previsão_F$Previsão_OUT),
                      RMSE_OUT = rmse(Previsão_F$Dados_reais_OUT, Previsão_F$Previsão_OUT))
                      print(Erro_OUT)
#Armazena os parâmetros da momdelagem em um data.frame
Parametros_modelos = c(HOLTWM[["model"]][["loglik"]], HOLTWM[["model"]][["aic"]], HOLTWM[["model"]][["bic"]], HOLTWM[["model"]][["aicc"]],
                       HOLTWM[["model"]][["par"]][["alpha"]], HOLTWM[["model"]][["par"]][["beta"]],
                       HOLTWM[["model"]][["par"]][["l"]], HOLTWM[["model"]][["par"]][["b"]],
                       HOLTWM[["model"]][["sigma2"]])

Parametros = matrix(Parametros, ncol = 2, nrow = 9)
Parametros[,1] = c("loglik", "aic","aicc","bic", "alpha", "beta", "l", "beta","sigma2" )

#o planilhador cria e exporta os dados obtidos para um planilha de excel
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

saveWorkbook(wb, "Holt Winters Multiplicativo.xlsx")

return(HOLTWM[["mean"]])

}

HOLTWM(serie,Data[1:156],Data[157:180],Dados[157:180,1], 24)
