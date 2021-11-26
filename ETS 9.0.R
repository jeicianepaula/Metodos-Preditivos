#ATIVANDO AS BIBLIOTECAS *Caso necessite instalar utilize o comando install.packages()*
library(Metrics) 
library(xlsx)
library(fpp2)

#--------------------------------------------------EXPORTACAODE DADOS-------------------------------------------------------------------------------------------
#Expota e separa os dados tipo date dos dados para formação das series temporais
Dados= read.xlsx(file.choose(), header = F,sheetIndex = 1)
Data= Dados[,1]
Dados1= Dados[,-1]
Dados1= na.interp(Dados1) 
print(Dados1)

#--------------------------------------------------FORMACAO DA SERIE-------------------------------------------------------------------------------------------
#Cria uma arquivo no formato ts (serie temporal)
serie = ts(Dados1, start = c(2015,1), end = c(2019,4),frequency = 4) %>% window(end = c(2018,4))
print(serie)
plot(serie)
autoplot(serie)+ ylim(c(17.5, 30)) +labs( title = "Serie Temporal") + theme_minimal() +
        theme(legend.title = element_blank(), text = element_text(family = "", colour = "gray20")) +
        theme(plot.title = element_text(hjust = 0.5)) + geom_line(size = 1)

#---------------------------------------------------------ETS----------------------------------------------------------------------------------------------
ETS = function (x,Data_IN,Data_OUT,Dados_reais_OUT){
 
  #Faz a modelagem e previsão
  ETS = ets(serie) %>% forecast(h = 4)
  print(ETS[["model"]])
  
  #cria dois arquivos do tipo data.frame para armazenar as informacoes de previsao dentro e fora da amostra
  Previsão_D = data.frame(Data_IN  ,Dados_reais_IN = (ETS[["x"]]),  Previsão_IN = (ETS[["fitted"]]))
  Previsão_F= data.frame(Data_OUT, Dados_reais_OUT, Previsão_OUT = (ETS[["mean"]]))
 
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
  Parametros = c(ETS[["model"]][["method"]],ETS[["model"]][["loglik"]], ETS[["model"]][["aic"]], ETS[["model"]][["bic"]], ETS[["model"]][["aicc"]],
                 ETS[["model"]][["par"]][["alpha"]],
                 ETS[["model"]][["par"]][["l"]])
  Parametros = matrix(Parametros, ncol = 2, nrow = 7)
  Parametros[,1] = c("Modelo","loglik", "aic","aicc","bic", "alpha","l")
  
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
  
  saveWorkbook(wb, "ETS.xlsx")
  return(ETS[["mean"]])

}
  
ETS(serie,Data[1:16],Data[17:20],Dados[17:20,2])

