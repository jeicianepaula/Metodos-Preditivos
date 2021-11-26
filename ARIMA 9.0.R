#ATIVANDO AS BIBLIOTECAS *Caso necessite instalar utilize o comando install.packages()*
library(Metrics) #METRICAS DE ERRO
library(xlsx) #IMPOTAÇÃO DO DADOS 
library(fpp2) # MODELOS DE PREVISÃO LINEARES
library(rugarch) #MODELOS DE PREVISÕ NÃO LINEARES
library(erer) #econometria

#--------------------------------------------------EXPORTACAO  DE DADOS-------------------------------------------------------------------------------------------
#Expota e separa os dados tipo date dos dados para formação das series temporais
Ex = read.xlsx(file.choose(), header = T,sheetIndex = 1)
Data= Ex[,1]
Dados= Ex[,-1]

#--------------------------------------------------FORMACAO DA SERIE-------------------------------------------------------------------------------------------
#Cria uma arquivo no fprmato ts (serie temporal)
serie = ts(Dados[,1], start = c(2000,1), end = c(2020,2),frequency = 12) %>% window(end = c(2012,12))

#---------------------------------------------------------ARIMA/SARIMA----------------------------------------------------------------------------------------------
ARIMA = function (x,Data_IN,Data_OUT,Dados_reais_OUT){
  
    #Faz a modelagem e previsão  
    ARIMA = serie %>% auto.arima(trace=T, stepwise = F, approximation = F) %>% forecast(h=24,level = c(85,95))
    print(ARIMA[["model"]])
    
    #cria dois arquivos do tipo data.frame para armazenar as informacoes de previsao dentro e fora da amostra
    Previsão_D = data.frame(Data_IN, Dados_reais_IN = ARIMA[["model"]][["x"]], Previsão_IN = ARIMA[["fitted"]])
    Previsão_F= data.frame(Data_OUT, Dados_reais_OUT, Previsão_OUT = (ARIMA[["mean"]])) 
    
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
    Parametros_modelos= c(ARIMA[["method"]],ARIMA[["model"]][["sigma2"]], 
                          ARIMA[["model"]][["aic"]],ARIMA[["model"]][["aicc"]], 
                          ARIMA[["model"]][["bic"]])
                          
    Parametros= matrix(Parametros_modelos, ncol = 2, nrow = 5)
    Parametros[,1] = c("method", "aic","aicc","bic", "sigma2")
    
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
    saveWorkbook(wb, "ARIMA.xlsx")
    
    return(ARIMA[["mean"]])
}

ARIMA(serie,Data[1:156],Data[157:180],Dados[157:180,1])

