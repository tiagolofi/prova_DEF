
#### TRATAR ARQUIVO PARA A PROVA DEF ####

packs = c("readxl", "tidyr", "magrittr", "lubridate")

for(i in packs){
  cat("Instalando packages...\n")
  install.packages(i)
}

# setwd("C:/Users/usuario/Documents/Ambientes/DEF/prova_DEF")

require(magrittr)

corrige_data <- function(data){
  
  lubridate::floor_date(
    as.Date("1900-01-01") + as.numeric(data),
    "month"
  )
  
}

cria_arquivo_tratado <- function(dir = NULL){
  
  ## 'dir' = local onde o arquivo ficará salvo 
  
  choices = c("CUSTOS - Contrato de Gestão", "CUSTOS - Outras Despesas",
              "PAGAMENTOS - Contratos de Gestão", "PAGAMENTOS - Outras Despesas")
  
  filename = menu(
    choices = choices, 
    title = "Qual informação você deseja exportar?",
    graphics = TRUE
  )
  
  print(filename)
  
  if(filename == 1){
    range = "B5:K11"
  } else 
  if(filename == 2){
    range = "B14:K19"
  } else
  if(filename == 3){
    range = "B25:K31"
  } else
  if(filename == 4){
    range = "B35:K40"
  }
  
  file <- readxl::read_excel(
    path = "Planilha para prova do DEF.xlsx",
    range = range
  )
  
  cat("Arrumando arquivo e corrigindo datas...\n")
  file %<>% 
    tidyr::gather(
      key = "Mês",
      value = "Valor", 
      4:10
    )
  
  file$`Mês` %<>% corrige_data
  
  print(file)
  
  if(!is.null(dir)){
    cat(paste0("Exportando arquivo para ", dir, "...\n"))
    write.csv2(
      x = file,
      file = paste0(dir, "/", choices[filename], ".csv"), 
      row.names = F, 
      na = ""
    )
  } else {
    cat("Exportando arquivo...\n")
    write.csv2(
      x = file,
      file = paste0(choices[filename], ".csv"), 
      row.names = F, 
      na = ""
    )
  } 
  cat("Arquivo exportado com sucesso!")
}

i = 1

repeat {

  cat(paste0("Salvando arquivo número ", i))
  
  cria_arquivo_tratado(
    dir = NULL
  )
  
  i=i+1
  
  if(i == 5){
    break()
  }
  
}

rm(i, choices, corrige_data, cria_arquivo_tratado)
