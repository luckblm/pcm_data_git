#ECONOMIA----
bases$subtematica <- "-"

##01 - AGRICULTURA----
###01.1 - Quantidade Produzida e Valor da Produção de Abacaxi (em toneladas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Agricultura"
d1$indicador <- "Quantidade Produzida e Valor da Produção de Abacaxi (em mil frutos), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Abacaxi (Mil frutos)"
d1$categoria2 <- "Quantidade produzida (Toneladas)	"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Agricultura"
d2$indicador <- "Quantidade Produzida e Valor da Produção de Abacaxi (em mil frutos), segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Abacaxi (Mil frutos)"
d2$categoria2 <- "Quantidade produzida (Toneladas)	"
d3$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Agricultura"
d3$indicador <- "Quantidade Produzida e Valor da Produção de Abacaxi (em mil frutos), segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Abacaxi (Mil frutos)"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Agricultura"
d4$indicador <- "Quantidade Produzida e Valor da Produção de Abacaxi (em mil frutos), segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Abacaxi (Mil frutos)"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/agricultura/01.1 - Quantidade Produzida e Valor da Produção de Abacaxi.xlsx",
  overwrite = TRUE
)
###01.2 - Quantidade Produzida e Valor da Produção de Arroz (em casca) (em toneladas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Agricultura"
d1$indicador <- "Quantidade Produzida e Valor da Produção de Arroz (em casca) (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Arroz (em casca)"
d1$categoria2 <- "Quantidade produzida (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Agricultura"
d2$indicador <- "Quantidade Produzida e Valor da Produção de Arroz (em casca) (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Arroz (em casca)"
d2$categoria2 <- "Quantidade produzida (Toneladas)"
d3$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Agricultura"
d3$indicador <- "Quantidade Produzida e Valor da Produção de Arroz (em casca) (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Arroz (em casca)"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Agricultura"
d4$indicador <- "Quantidade Produzida e Valor da Produção de Arroz (em casca) (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Arroz (em casca)"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"



# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/agricultura/01.2 - Quantidade Produzida e Valor da Produção de Arroz.xlsx",
  overwrite = TRUE
)
###01.3 - Quantidade Produzida e Valor da Produção de Feijão (em grão)(em toneladas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Agricultura"
d1$indicador <- "Quantidade Produzida e Valor da Produção de Feijão (em grão)(em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Feijão (em grão)"
d1$categoria2 <- "Quantidade produzida (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Agricultura"
d2$indicador <- "Quantidade Produzida e Valor da Produção de Feijão (em grão)(em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Feijão (em grão)"
d2$categoria2 <- "Quantidade produzida (Toneladas)"
d3$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Agricultura"
d3$indicador <- "Quantidade Produzida e Valor da Produção de Feijão (em grão)(em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Feijão (em grão)"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Agricultura"
d4$indicador <- "Quantidade Produzida e Valor da Produção de Feijão (em grão)(em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Feijão (em grão)"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"



# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/agricultura/01.3 - Quantidade Produzida e Valor da Produção de Feijão.xlsx",
  overwrite = TRUE
)
###01.4 - Quantidade Produzida e Valor da Produção de Mandioca (em toneladas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Agricultura"
d1$indicador <- "Quantidade Produzida e Valor da Produção de Mandioca (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Mandioca"
d1$categoria2 <- "Quantidade produzida (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Agricultura"
d2$indicador <- "Quantidade Produzida e Valor da Produção de Mandioca (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Mandioca"
d2$categoria2 <- "Quantidade produzida (Toneladas)"
d3$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Agricultura"
d3$indicador <- "Quantidade Produzida e Valor da Produção de Mandioca (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Mandioca"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Agricultura"
d4$indicador <- "Quantidade Produzida e Valor da Produção de Mandioca (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Mandioca"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"



# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/agricultura/01.4 - Quantidade Produzida e Valor da Produção de Mandioca.xlsx",
  overwrite = TRUE
)
###01.5 - Quantidade Produzida e Valor da Produção de Milho (em grão) (em toneladas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Agricultura"
d1$indicador <- "Quantidade Produzida e Valor da Produção de Milho (em grão) (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Milho (em grão)"
d1$categoria2 <- "Quantidade produzida (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Agricultura"
d2$indicador <- "Quantidade Produzida e Valor da Produção de Milho (em grão) (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Milho (em grão)"
d2$categoria2 <- "Quantidade produzida (Toneladas)"
d3$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Agricultura"
d3$indicador <- "Quantidade Produzida e Valor da Produção de Milho (em grão) (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Milho (em grão)"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Agricultura"
d4$indicador <- "Quantidade Produzida e Valor da Produção de Milho (em grão) (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Milho (em grão)"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"



# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/agricultura/01.5 - Quantidade Produzida e Valor da ###Produção de Milho.xlsx",
  overwrite = TRUE
)
###01.6 - Quantidade Produzida e Valor da Produção de Soja (em grão) (em toneladas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Agricultura"
d1$indicador <- "Quantidade Produzida e Valor da Produção de Soja (em grão)  (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Soja (em grão)"
d1$categoria2 <- "Quantidade produzida (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Agricultura"
d2$indicador <- "Quantidade Produzida e Valor da Produção de Soja (em grão)  (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Soja (em grão)"
d2$categoria2 <- "Quantidade produzida (Toneladas)"
d3$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Agricultura"
d3$indicador <- "Quantidade Produzida e Valor da Produção de Soja (em grão)  (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Soja (em grão)"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Agricultura"
d4$indicador <- "Quantidade Produzida e Valor da Produção de Soja (em grão)  (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Soja (em grão)"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"



# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/agricultura/01.6 - Quantidade Produzida e Valor da Produção de Soja.xlsx",
  overwrite = TRUE
)
###01.7 - Quantidade Produzida e Valor da Produção de Açaí----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Agricultura"
d1$indicador <- "Quantidade Produzida e Valor da Produção de Açaí (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Açaí (toneladas)"
d1$categoria2 <- "Quantidade produzida (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Agricultura"
d2$indicador <- "Quantidade Produzida e Valor da Produção de Açaí (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Açaí (toneladas)"
d2$categoria2 <- "Quantidade produzida (Toneladas)"
d3$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Agricultura"
d3$indicador <- "Quantidade Produzida e Valor da Produção de Açaí (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Açaí (toneladas)"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Agricultura"
d4$indicador <- "Quantidade Produzida e Valor da Produção de Açaí (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Açaí (toneladas)"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"



# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/agricultura/01.7 - Quantidade Produzida e Valor da Produção de Açaí.xlsx",
  overwrite = TRUE
)
###01.8 - Quantidade Produzida e Valor da Produção de Banana----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Agricultura"
d1$indicador <- "Quantidade Produzida e Valor da Produção de Banana (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Banana (Toneladas)"
d1$categoria2 <- "Quantidade produzida (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Agricultura"
d2$indicador <- "Quantidade Produzida e Valor da Produção de Banana (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Banana (Toneladas)"
d2$categoria2 <- "Quantidade produzida (Toneladas)"
d3$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Agricultura"
d3$indicador <- "Quantidade Produzida e Valor da Produção de Banana (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Banana (Toneladas)"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Agricultura"
d4$indicador <- "Quantidade Produzida e Valor da Produção de Banana (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Banana (Toneladas)"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"



# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/agricultura/01.8 - Quantidade Produzida e Valor da Produção de Banana.xlsx",
  overwrite = TRUE
)
###01.9 - Quantidade Produzida e Valor da Produção de Cacau----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Agricultura"
d1$indicador <- "Quantidade Produzida e Valor da Produção de Cacau (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Cacau (Toneladas)"
d1$categoria2 <- "Quantidade produzida (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Agricultura"
d2$indicador <- "Quantidade Produzida e Valor da Produção de Cacau (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Cacau (Toneladas)"
d2$categoria2 <- "Quantidade produzida (Toneladas)"
d3$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Agricultura"
d3$indicador <- "Quantidade Produzida e Valor da Produção de Cacau (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Cacau (Toneladas)"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Agricultura"
d4$indicador <- "Quantidade Produzida e Valor da Produção de Cacau (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Cacau (Toneladas)"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"



# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/agricultura/01.9 - Quantidade Produzida e Valor da Produção de Cacau.xlsx",
  overwrite = TRUE
)
###01.10 - Quantidade Produzida e Valor da Produção de Coco-da-baia----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Agricultura"
d1$indicador <- "Quantidade Produzida e Valor da Produção de Coco-da-baia (em mil frutos), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Coca-da-baia (Mil frutos)"
d1$categoria2 <- "Quantidade produzida (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Agricultura"
d2$indicador <- "Quantidade Produzida e Valor da Produção de Coco-da-baia (em mil frutos), segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Coca-da-baia (Mil frutos)"
d2$categoria2 <- "Quantidade produzida (Toneladas)"
d3$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Agricultura"
d3$indicador <- "Quantidade Produzida e Valor da Produção de Coco-da-baia (em mil frutos), segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Coca-da-baia (Mil frutos)"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Agricultura"
d4$indicador <- "Quantidade Produzida e Valor da Produção de Coco-da-baia (em mil frutos), segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Coca-da-baia (Mil frutos)"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"



# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/agricultura/01.10 - Quantidade Produzida e Valor da Produção de Coco-da-baia.xlsx",
  overwrite = TRUE
)
###01.11 - Quantidade Produzida e Valor da ###Produção de Dendê----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Agricultura"
d1$indicador <- "Quantidade Produzida e Valor da Produção de Dendê (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Dendê (toneladas)"
d1$categoria2 <- "Quantidade produzida (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Agricultura"
d2$indicador <- "Quantidade Produzida e Valor da Produção de Dendê (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Dendê (toneladas)"
d2$categoria2 <- "Quantidade produzida (Toneladas)"
d3$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Agricultura"
d3$indicador <- "Quantidade Produzida e Valor da Produção de Dendê (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Dendê (toneladas)"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Agricultura"
d4$indicador <- "Quantidade Produzida e Valor da Produção de Dendê (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Dendê (toneladas)"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/agricultura/.xlsx",
  overwrite = TRUE
)
###01.12 - Quantidade Produzida e Valor da Produção de Pimenta-do-reino----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Agricultura"
d1$indicador <- "Quantidade Produzida e Valor da Produção de Pimenta-do-reino (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Pimenta-do-reino (toneladas)"
d1$categoria2 <- "Quantidade produzida (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Agricultura"
d2$indicador <- "Quantidade Produzida e Valor da Produção de Pimenta-do-reino (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Pimenta-do-reino (toneladas)"
d2$categoria2 <- "Quantidade produzida (Toneladas)"
d3$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Agricultura"
d3$indicador <- "Quantidade Produzida e Valor da Produção de Pimenta-do-reino (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Pimenta-do-reino (toneladas)"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Agricultura"
d4$indicador <- "Quantidade Produzida e Valor da Produção de Pimenta-do-reino (em toneladas), segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Pimenta-do-reino (toneladas)"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"



# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/agricultura/01.12 - Quantidade Produzida e Valor da Produção de Pimenta-do-reino.xlsx",
  overwrite = TRUE
)

##02 - EXTRATIVISMO VEGETAL E SILVICULTURA----
###02.1 - Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) do Total de Produtos Alimentícios da Extração Vegetal----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Extrativismo Vegetal e Silvicultura"
d1$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) do Total de Produtos Alimentícios da Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Total de Produtos Alimentícios"
d1$categoria2 <- "Quantidade produzida na extração vegetal (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Extrativismo Vegetal e Silvicultura"
d2$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) do Total de Produtos Alimentícios da Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Total de Produtos Alimentícios"
d2$categoria2 <- "Quantidade produzida na extração vegetal (Toneladas)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Extrativismo Vegetal e Silvicultura"
d3$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) do Total de Produtos Alimentícios da Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Total de Produtos Alimentícios"
d3$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Extrativismo Vegetal e Silvicultura"
d4$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) do Total de Produtos Alimentícios da Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Total de Produtos Alimentícios"
d4$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d4$categoria3 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/extrativismo_vegetal_e_silvicultura/02.1 - Quantidade Produzida e Valor da Produção do Total de Produtos Alimentícios da Extração Vegetal.xlsx",
  overwrite = TRUE
)
###02.2 - Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Açaí na Extração Vegetal----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Extrativismo Vegetal e Silvicultura"
d1$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Açaí na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Açaí (fruto)"
d1$categoria2 <- "Quantidade produzida na extração vegetal (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Extrativismo Vegetal e Silvicultura"
d2$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Açaí na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Açaí (fruto)"
d2$categoria2 <- "Quantidade produzida na extração vegetal (Toneladas)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Extrativismo Vegetal e Silvicultura"
d3$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Açaí na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Açaí (fruto)"
d3$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Extrativismo Vegetal e Silvicultura"
d4$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Açaí na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Açaí (fruto)"
d4$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d4$categoria3 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/extrativismo_vegetal_e_silvicultura/02.2 - Quantidade Produzida e Valor da ###Produção (Mil Reais) de Açaí na Extração Vegetal.xlsx",
  overwrite = TRUE
)
###02.3 - Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Castanha-do-Pará na Extração Vegetal----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Extrativismo Vegetal e Silvicultura"
d1$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Castanha-do-Pará na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Castanha-do-Pará"
d1$categoria2 <- "Quantidade produzida na extração vegetal (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Extrativismo Vegetal e Silvicultura"
d2$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Castanha-do-Pará na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Castanha-do-Pará"
d2$categoria2 <- "Quantidade produzida na extração vegetal (Toneladas)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Extrativismo Vegetal e Silvicultura"
d3$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Castanha-do-Pará na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Castanha-do-Pará"
d3$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Extrativismo Vegetal e Silvicultura"
d4$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Castanha-do-Pará na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Castanha-do-Pará"
d4$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d4$categoria3 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/extrativismo_vegetal_e_silvicultura/02.3 - Quantidade Produzida e Valor da ###Produção de Castanha-do-Pará na Extração Vegetal.xlsx",
  overwrite = TRUE
)
###02.4 - Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Palmito na Extração Vegetal----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Extrativismo Vegetal e Silvicultura"
d1$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Palmito na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Palmito"
d1$categoria2 <- "Quantidade produzida na extração vegetal (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Extrativismo Vegetal e Silvicultura"
d2$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Palmito na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Palmito"
d2$categoria2 <- "Quantidade produzida na extração vegetal (Toneladas)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Extrativismo Vegetal e Silvicultura"
d3$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Palmito na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Palmito"
d3$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Extrativismo Vegetal e Silvicultura"
d4$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Palmito na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Palmito"
d4$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d4$categoria3 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/extrativismo_vegetal_e_silvicultura/02.4 - Quantidade Produzida e Valor da ###Produção de Palmito na Extração Vegetal.xlsx",
  overwrite = TRUE
)
###02.5 - Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Fibras na Extração Vegetal----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Extrativismo Vegetal e Silvicultura"
d1$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Fibras na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Fibras"
d1$categoria2 <- "Quantidade produzida na extração vegetal (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Extrativismo Vegetal e Silvicultura"
d2$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Fibras na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Fibras"
d2$categoria2 <- "Quantidade produzida na extração vegetal (Toneladas)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Extrativismo Vegetal e Silvicultura"
d3$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Fibras na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Fibras"
d3$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Extrativismo Vegetal e Silvicultura"
d4$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Fibras na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Fibras"
d4$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d4$categoria3 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/extrativismo_vegetal_e_silvicultura/02.5 - Quantidade Produzida e Valor da ###Produção de Fibras na Extração Vegetal.xlsx",
  overwrite = TRUE
)
###02.6 - Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Carvão Vegetal na Extração Vegetal----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Extrativismo Vegetal e Silvicultura"
d1$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Carvão Vegetal na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Carvão vegetal"
d1$categoria2 <- "Quantidade produzida na extração vegetal (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Extrativismo Vegetal e Silvicultura"
d2$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Carvão Vegetal na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Carvão vegetal"
d2$categoria2 <- "Quantidade produzida na extração vegetal (Toneladas)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Extrativismo Vegetal e Silvicultura"
d3$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Carvão Vegetal na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Carvão vegetal"
d3$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Extrativismo Vegetal e Silvicultura"
d4$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Carvão Vegetal na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Carvão vegetal"
d4$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d4$categoria3 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/extrativismo_vegetal_e_silvicultura/02.6 - Quantidade Produzida e Valor da ###Produção de Carvão Vegetal na Extração Vegetal.xlsx",
  overwrite = TRUE
)
###02.7 - Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Lenha na Extração Vegetal----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Extrativismo Vegetal e Silvicultura"
d1$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Lenha na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Lenha"
d1$categoria2 <- "Quantidade produzida na extração vegetal (Metros cúbicos)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Extrativismo Vegetal e Silvicultura"
d2$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Lenha na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Lenha"
d2$categoria2 <- "Quantidade produzida na extração vegetal (Metros cúbicos)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Extrativismo Vegetal e Silvicultura"
d3$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Lenha na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Lenha"
d3$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Extrativismo Vegetal e Silvicultura"
d4$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Lenha na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Lenha"
d4$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d4$categoria3 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/extrativismo_vegetal_e_silvicultura/02.7 - Quantidade Produzida (Toneladas) e Valor da ###Produção de Lenha na Extração Vegetal.xlsx",
  overwrite = TRUE
)
###02.8 - Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Madeira em Tora na Extração Vegetal----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Extrativismo Vegetal e Silvicultura"
d1$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Madeira em Tora na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Madeira em tora"
d1$categoria2 <- "Quantidade produzida na extração vegetal (Metros cúbicos)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Extrativismo Vegetal e Silvicultura"
d2$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Madeira em Tora na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Madeira em tora"
d2$categoria2 <- "Quantidade produzida na extração vegetal (Metros cúbicos)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Extrativismo Vegetal e Silvicultura"
d3$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Madeira em Tora na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Madeira em tora"
d3$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Extrativismo Vegetal e Silvicultura"
d4$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Madeira em Tora na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Madeira em tora"
d4$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d4$categoria3 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/extrativismo_vegetal_e_silvicultura/02.8 - Quantidade Produzida e Valor da ###Produção (Mil Reais) de Madeira em Tora na Extração Vegetal.xlsx",
  overwrite = TRUE
)
###02.9 - Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Oleaginosos na Extração Vegetal----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Extrativismo Vegetal e Silvicultura"
d1$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Oleaginosos na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Oleaginosos"
d1$categoria2 <- "Quantidade produzida na extração vegetal (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Extrativismo Vegetal e Silvicultura"
d2$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Oleaginosos na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Oleaginosos"
d2$categoria2 <- "Quantidade produzida na extração vegetal (Toneladas)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Extrativismo Vegetal e Silvicultura"
d3$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Oleaginosos na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Oleaginosos"
d3$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Extrativismo Vegetal e Silvicultura"
d4$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Oleaginosos na Extração Vegetal, segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Oleaginosos"
d4$categoria2 <- "Valor da produção na extração vegetal (Mil Reais)"
d4$categoria3 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/extrativismo_vegetal_e_silvicultura/02.9 - Quantidade Produzida e Valor da ###Produção (Mil Reais) de Oleaginosos na Extração Vegetal.xlsx",
  overwrite = TRUE
)
###02.10 - Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Carvão Vegetal da Silvicultura----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Extrativismo Vegetal e Silvicultura"
d1$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Carvão Vegetal da Silvicultura, segundo Brasil,  Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Carvão vegetal"
d1$categoria2 <- "Quantidade produzida na silvicultura (Toneladas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Extrativismo Vegetal e Silvicultura"
d2$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Carvão Vegetal da Silvicultura, segundo Brasil,  Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Carvão vegetal"
d2$categoria2 <- "Quantidade produzida na silvicultura (Toneladas)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Extrativismo Vegetal e Silvicultura"
d3$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Carvão Vegetal da Silvicultura, segundo Brasil,  Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Carvão vegetal"
d3$categoria2 <- "Valor da produção na silvicultura (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Extrativismo Vegetal e Silvicultura"
d4$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Carvão Vegetal da Silvicultura, segundo Brasil,  Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Carvão vegetal"
d4$categoria2 <- "Valor da produção na silvicultura (Mil Reais)"
d4$categoria3 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/extrativismo_vegetal_e_silvicultura/02.10 - Quantidade Produzida e Valor da ###Produção (Mil Reais) de Carvão Vegetal da Silvicultura.xlsx",
  overwrite = TRUE
)
###02.11 - Quantidade Produzida (Toneladas) e Valor da ###Produção (Mil Reais) de Lenha da Silvicultura----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Extrativismo Vegetal e Silvicultura"
d1$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Lenha da Silvicultura, segundo Brasil,  Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Lenha"
d1$categoria2 <- "Quantidade produzida na silvicultura (Metros cúbicos)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Extrativismo Vegetal e Silvicultura"
d2$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Lenha da Silvicultura, segundo Brasil,  Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Lenha"
d2$categoria2 <- "Quantidade produzida na silvicultura (Metros cúbicos)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Extrativismo Vegetal e Silvicultura"
d3$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Lenha da Silvicultura, segundo Brasil,  Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Lenha"
d3$categoria2 <- "Valor da produção na silvicultura (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Extrativismo Vegetal e Silvicultura"
d4$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Lenha da Silvicultura, segundo Brasil,  Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Lenha"
d4$categoria2 <- "Valor da produção na silvicultura (Mil Reais)"
d4$categoria3 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/extrativismo_vegetal_e_silvicultura/02.11 - Quantidade Produzida (Toneladas) e Valor da ###Produção de Lenha da Silvicultura.xlsx",
  overwrite = TRUE
)
###02.12 - Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Madeira em Tora da Silvicultura----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Extrativismo Vegetal e Silvicultura"
d1$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Madeira em Tora da Silvicultura, segundo Brasil,  Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Madeira em tora"
d1$categoria2 <- "Quantidade produzida na silvicultura (Metros cúbicos)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Extrativismo Vegetal e Silvicultura"
d2$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Madeira em Tora da Silvicultura, segundo Brasil,  Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Madeira em tora"
d2$categoria2 <- "Quantidade produzida na silvicultura (Metros cúbicos)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Extrativismo Vegetal e Silvicultura"
d3$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Madeira em Tora da Silvicultura, segundo Brasil,  Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Madeira em tora"
d3$categoria2 <- "Valor da produção na silvicultura (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Extrativismo Vegetal e Silvicultura"
d4$indicador <- "Quantidade Produzida (Toneladas) e Valor da Produção (Mil Reais) de Madeira em Tora da Silvicultura, segundo Brasil,  Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Madeira em tora"
d4$categoria2 <- "Valor da produção na silvicultura (Mil Reais)"
d4$categoria3 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/extrativismo_vegetal_e_silvicultura/02.12 - Quantidade Produzida (Toneladas) e Valor da Produção de Madeira em Tora da Silvicultura.xlsx",
  overwrite = TRUE
)

##03 - PECUÁRIA----
###03.1 - Efetivo de Rebanho Bovino----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Pecuária"
d1$indicador <- "Efetivo de Rebanho Bovino, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Rebanho Bovino"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Pecuária"
d2$indicador <- "Efetivo de Rebanho Bovino, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/pecuaria/03.1 - Efetivo de Rebanho Bovino.xlsx",
  overwrite = TRUE
)
###03.2 - Efetivo de Rebanho Bubalino----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Pecuária"
d1$indicador <- "Efetivo de Rebanho Bubalino, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Rebanho Bubalino"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Pecuária"
d2$indicador <- "Efetivo de Rebanho Bubalino, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/pecuaria/03.2 - Efetivo de Rebanho Bubalino.xlsx",
  overwrite = TRUE
)
###03.3 - Efetivo de Rebanho Suíno Total----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Pecuária"
d1$indicador <- "Efetivo de Rebanho Suíno Total, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Rebanho Suíno - total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Pecuária"
d2$indicador <- "Efetivo de Rebanho Suíno Total, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/pecuaria/03.3 - Efetivo de Rebanho Suíno Total.xlsx",
  overwrite = TRUE
)
###03.4 - Efetivo de Rebanho de Galináceos Total----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Pecuária"
d1$indicador <- "Efetivo de Rebanho de Galináceos Total (Cabeças), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Rebanho Galináceos - Total (Cabeças)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Pecuária"
d2$indicador <- "Efetivo de Rebanho de Galináceos Total (Cabeças), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/pecuaria/03.4 - Efetivo de Rebanho de Galináceos Total.xlsx",
  overwrite = TRUE
)
###03.5 - Abate de Frangos (Cabeças)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Pecuária"
d1$indicador <- "Abate de Frangos (Cabeças), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Abate de Frangos (Cabeças)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Pecuária"
d2$indicador <- "Abate de Frangos (Cabeças), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/pecuaria/03.5 - Abate de Frangos_Cabeças.xlsx",
  overwrite = TRUE
)
###03.6 - Abate de Rebanho Bovino (Cabeças)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Pecuária"
d1$indicador <- "Abate de Rebanho Bovino (Cabeças), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Abate de Rebanho Bovino (Cabeças)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Pecuária"
d2$indicador <- "Abate de Rebanho Bovino (Cabeças), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/pecuaria/03.6 - Abate de Rebanho Bovino_Cabeças.xlsx",
  overwrite = TRUE
)
###03.7 - Abate de Rebanho Suíno (Cabeças)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Pecuária"
d1$indicador <- "Abate de Rebanho Suíno (Cabeças), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Abate de Rebanho Suíno (Cabeças)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Pecuária"
d2$indicador <- "Abate de Rebanho Suíno (Cabeças), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/pecuaria/03.7 - Abate de Rebanho Suíno_Cabeças.xlsx",
  overwrite = TRUE
)
###03.8 - Produção de Leite (Mil Litros) e Valor da Produção (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Pecuária"
d1$indicador <- "Produção (Mil Litros) e Valor da Produção (Mil Reais) de Leite, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Leite"
d1$categoria2 <- "Produção (Mil litros)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Pecuária"
d2$indicador <- "Produção (Mil Litros) e Valor da Produção (Mil Reais) de Leite, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Leite"
d2$categoria2 <- "Produção (Mil litros)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Pecuária"
d3$indicador <- "Produção (Mil Litros) e Valor da Produção (Mil Reais) de Leite, segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Leite"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Pecuária"
d4$indicador <- "Produção (Mil Litros) e Valor da Produção (Mil Reais) de Leite, segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Leite"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/pecuaria/03.8 - Produção de Leite e Valor da Produção.xlsx",
  overwrite = TRUE
)
###03.9 - Produção de Ovos de Galinha (Mil Dúzias) e Valor da Produção (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Pecuária"
d1$indicador <- "Produção (Mil Dúzias) e Valor da Produção (Mil Reais)  de Ovos de Galinha, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Ovos de galinha"
d1$categoria2 <- "Produção (Mil litros)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Pecuária"
d2$indicador <- "Produção (Mil Dúzias) e Valor da Produção (Mil Reais)  de Ovos de Galinha, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ovos de galinha"
d2$categoria2 <- "Produção (Mil litros)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Pecuária"
d3$indicador <- "Produção (Mil Dúzias) e Valor da Produção (Mil Reais)  de Ovos de Galinha, segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Ovos de galinha"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Pecuária"
d4$indicador <- "Produção (Mil Dúzias) e Valor da Produção (Mil Reais)  de Ovos de Galinha, segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Ovos de galinha"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/pecuaria/03.9 - Produção de Ovos de Galinha e Valor da ###Produção.xlsx",
  overwrite = TRUE
)
###03.10 - Produção de Ovos de Codorna (Mil Dúzias) e Valor da Produção (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Pecuária"
d1$indicador <- "Produção de Ovos de Codorna (Mil Dúzias) e Valor da Produção (Mil Reais), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Ovos de codorna"
d1$categoria2 <- "Produção de origem animal (Mil dúzias)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Pecuária"
d2$indicador <- "Produção de Ovos de Codorna (Mil Dúzias) e Valor da Produção (Mil Reais), segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ovos de codorna"
d2$categoria2 <- "Produção de origem animal (Mil dúzias)"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Pecuária"
d3$indicador <- "Produção de Ovos de Codorna (Mil Dúzias) e Valor da Produção (Mil Reais), segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Ovos de codorna"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Pecuária"
d4$indicador <- "Produção de Ovos de Codorna (Mil Dúzias) e Valor da Produção (Mil Reais), segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Ovos de codorna"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria3 <- "Ranking"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/pecuaria/03.10 - Produção de Ovos de Codorna e Valor da ###Produção.xlsx",
  overwrite = TRUE
)
###03.11 - Produção de Mel de Abelha (Quilogramas) e Valor da ###Produção (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Pecuária"
d1$indicador <- "Produção (Quilogramas) e Valor da Produção (Mil Reais) de Mel de Abelha, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Mel de abelha"
d1$categoria2 <- "Produção de origem animal (Quilogramas)"
d1$categoria3 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Pecuária"
d2$indicador <- "Produção (Quilogramas) e Valor da Produção (Mil Reais) de Mel de Abelha, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Mel de abelha"
d2$categoria2 <- "Produção de origem animal (Quilogramas)"
d2$categoria2 <- "Ranking"

d3 <- bases
d3$tematica <- "Economia"
d3$subtematica <- "Pecuária"
d3$indicador <- "Produção (Quilogramas) e Valor da Produção (Mil Reais) de Mel de Abelha, segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Mel de abelha"
d3$categoria2 <- "Valor da produção (Mil Reais)"
d3$categoria3 <- "-"

d4 <- bases
d4$tematica <- "Economia"
d4$subtematica <- "Pecuária"
d4$indicador <- "Produção (Quilogramas) e Valor da Produção (Mil Reais) de Mel de Abelha, segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Mel de abelha"
d4$categoria2 <- "Valor da produção (Mil Reais)"
d4$categoria2 <- "Ranking"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/pecuaria/03.11 - Produção de Mel de Abelha e Valor da ###Produção.xlsx",
  overwrite = TRUE
)

##04 - INDUSTRIA----
###04.1 - Total da Indústria - Número de unidades locais (unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Total da Indústria - Número de unidades locais (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Total da Indústria - Número de unidades locais (Unidades)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Total da Indústria - Número de unidades locais (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.1 - Total da Indústria - Número de unidades locais.xlsx",
  overwrite = TRUE
)
###04.2 - Total da Indústria - Pessoal ocupado em 31.12 (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Total da Indústria - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- ""
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Total da Indústria - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.2 - Total da Indústria - Pessoal ocupado em 31_12.xlsx",
  overwrite = TRUE
)
###04.3 - Total da Indústria - Salários, retiradas e outras remunerações (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Total da Indústria - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Total da Indústria - Salários, retiradas e outras remunerações (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Total da Indústria - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.2 - Total da Indústria - Pessoal ocupado em 31_12.xlsx",
  overwrite = TRUE
)
###04.4 - Total da Indústria - Receitas líquidas de vendas (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Total da Indústria - Receitas líquidas de vendas (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Total da Indústria - Receitas líquidas de vendas (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Total da Indústria - Receitas líquidas de vendas (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.4 - Total da Indústria - Receitas líquidas de vendas.xlsx",
  overwrite = TRUE
)
###04.5 - Total da Indústria - Custos das Operações Industriais (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Total da Indústria - Custos das Operações Industriais (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Total da Indústria - Custos das Operações Industriais (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Total da Indústria - Custos das Operações Industriais (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.5 - Total da Indústria - Custos das Operações Industriais.xlsx",
  overwrite = TRUE
)
###04.6 - Total da Indústria - Valor bruto da Produção industrial (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Total da Indústria - Valor Bruto da Produção Industrial (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Total da Indústria - Valor bruto da produção industrial (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Total da Indústria - Valor Bruto da Produção Industrial (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.6 - Total da Indústria - Valor bruto da Produção industrial.xlsx",
  overwrite = TRUE
)
###04.7 - Indústrias Extrativas - Número de unidades locais (unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Indústrias Extrativas - Número de unidades locais (unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Indústrias Extrativas - Número de unidades locais (Unidades)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Indústrias Extrativas - Número de unidades locais (unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.7 - Indústrias Extrativas - Número de unidades locais.xlsx",
  overwrite = TRUE
)
###04.8 - Indústrias Extrativas - Pessoal ocupado em 31.12 (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Indústrias Extrativas - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Indústrias Extrativas - Pessoal ocupado em 31/12 (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Indústrias Extrativas - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.8 - Indústrias Extrativas - Pessoal ocupado em 31_12_Pessoas.xlsx",
  overwrite = TRUE
)
###04.9 - Indústrias Extrativas - Salários, retiradas e outras remunerações (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Indústrias Extrativas - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Indústrias Extrativas - Salários, retiradas e outras remunerações (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Indústrias Extrativas - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.9 - Indústrias Extrativas - Salários, retiradas e outras remunerações.xlsx",
  overwrite = TRUE
)
###04.10 - Indústrias Extrativas - Receitas líquidas de vendas (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Indústrias Extrativas - Total de receitas líquidas de vendas (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Indústrias Extrativas - Total de receitas líquidas de vendas (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Indústrias Extrativas - Total de receitas líquidas de vendas (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.10 - Indústrias Extrativas - Receitas líquidas de vendas (Mil Reais).xlsx",
  overwrite = TRUE
)
###04.11 - Indústrias Extrativas - Custos das operações industriais (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Indústrias Extrativas - Custos das operações industriais (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Indústrias Extrativas - Custos das operações industriais (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Indústrias Extrativas - Custos das operações industriais (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.11 - Indústrias Extrativas - Custos das operações industriais.xlsx",
  overwrite = TRUE
)
###04.12 - Indústrias Extrativas - Valor bruto da Produção industrial (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Indústrias Extrativas - Valor bruto da produção industrial (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Indústrias Extrativas - Valor bruto da produção industrial (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Indústrias Extrativas - Valor bruto da produção industrial (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.12 - Indústrias Extrativas - Valor bruto da Produção industrial.xlsx",
  overwrite = TRUE
)
###04.13 - Indústrias de Transformação - Número de unidades locais (unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Indústrias de Transformação - Número de unidades locais (unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Indústrias de Transformação - Número de unidades locais (Unidades)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Indústrias de Transformação - Número de unidades locais (unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.13 - Indústrias de Transformação - Número de unidades locais.xlsx",
  overwrite = TRUE
)
###04.14 - Indústrias de Transformação - Pessoal ocupado em 31.12 (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Indústrias de Transformação - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Indústrias de Transformação - Pessoal ocupado em 31/12 (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Indústrias de Transformação - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.14 - Indústrias de Transformação - Pessoal ocupado em 31_12.xlsx",
  overwrite = TRUE
)
###04.15 - Indústrias de Transformação - Salários, retiradas e outras remunerações (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Indústrias de Transformação - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Indústrias de Transformação - Salários, retiradas e outras remunerações (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Indústrias de Transformação - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.15 - Indústrias de Transformação - Salários, retiradas e outras remunerações.xlsx",
  overwrite = TRUE
)
###04.16 - Indústrias de Transformação - Receitas líquidas de vendas (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Indústrias de Transformação - Total de receitas líquidas de vendas (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Indústrias de Transformação - Total de receitas líquidas de vendas (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Indústrias de Transformação - Total de receitas líquidas de vendas (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.16 - Indústrias de Transformação - Receitas líquidas de vendas.xlsx",
  overwrite = TRUE
)
###04.17 - Indústrias de Transformação - Custos das operações industriais----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Indústrias de Transformação - Custos das operações industriais (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Indústrias de Transformação - Total de custos das operações industriais (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Indústrias de Transformação - Custos das operações industriais (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.17 - Indústrias de Transformação - Custos das operações industriais.xlsx",
  overwrite = TRUE
)
###04.18 - Indústrias de Transformação - Valor bruto da Produção industrial (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Industria"
d1$indicador <- "Indústrias de Transformação - Valor bruto da produção industrial (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Indústrias de Transformação - Valor bruto da produção industrial (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Industria"
d2$indicador <- "Indústrias de Transformação - Valor bruto da produção industrial (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/industria/04.18 - Indústrias de Transformação - Valor bruto da Produção industrial.xlsx",
  overwrite = TRUE
)

##05 - SERVIÇOS----
###05.1 - Total de Serviço - Número de empresas (Unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Total de Serviço - Número de empresas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Número de empresas (Unidades)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Total de Serviço - Número de empresas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.1 - Total de Serviço - Número de empresas.xlsx",
  overwrite = TRUE
)
###05.2 - Total de Serviço - Pessoal ocupado em 31/12 (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Total de Serviço - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Pessoal ocupado em 31/12 (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Total de Serviço - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.2 - Total de Serviço - Pessoal ocupado em 31_12_Pessoas.xlsx",
  overwrite = TRUE
)
###05.3 - Total de Serviço - Receita bruta de Serviços (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Total de Serviço - Receita bruta de serviços (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Receita bruta de serviços (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Total de Serviço - Receita bruta de serviços (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.3 - Total de Serviço - Receita bruta de Serviços.xlsx",
  overwrite = TRUE
)
###05.4 - Total de Serviço - Salários, retiradas e outras remunerações (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Total de Serviço - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Receita bruta de serviços (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Total de Serviço - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/.xlsx",
  overwrite = TRUE
)
###05.5 - Serviços à Família - Número de empresas (Unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços prestados à família - Número de empresas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços prestados à família - Número de empresas (Unidades)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços prestados à família - Número de empresas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.5 - Serviços à Família - Número de empresas.xlsx",
  overwrite = TRUE
)
###05.6 - Serviços à Família - Pessoal ocupado em 31/12 (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços prestados à família - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços prestados à família - Pessoal ocupado em 31/12 (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços prestados à família - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.6 - Serviços à Família - Pessoal ocupado em 31_12.xlsx",
  overwrite = TRUE
)
###05.7 - Serviços à Família - Receita bruta de Serviços (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços prestados à família - Receita bruta de serviços (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços prestados à família - Receita bruta de serviços (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços prestados à família - Receita bruta de serviços (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.7 - Serviços à Família - Receita bruta de Serviços.xlsx",
  overwrite = TRUE
)
###05.8 - Serviços à Família - Salários, retiradas e outras remunerações (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços prestados à família - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços prestados à família - Salários, retiradas e outras remunerações (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços prestados à família - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.8 - Serviços à Família - Salários, retiradas e outras remunerações.xlsx",
  overwrite = TRUE
)
###05.9 - Serviços de informação e comunicação - Número de empresas (Unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços de informação e comunicação - Número de empresas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços de informação e comunicação - Número de empresas (Unidades)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços de informação e comunicação - Número de empresas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.9 - Serviços de informação e comunicação - Número de empresas.xlsx",
  overwrite = TRUE
)
###05.10 - Serviços de informação e comunicação - Pessoal ocupado em 31/12 (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços de informação e comunicação - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços de informação e comunicação - Pessoal ocupado em 31/12 (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços de informação e comunicação - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.10 - Serviços de informação e comunicação - Pessoal ocupado em 31_12.xlsx",
  overwrite = TRUE
)
###05.11 - Serviços de informação e comunicação - Receita bruta de Serviços (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços de informação e comunicação - Receita bruta de serviços (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços de informação e comunicação - Receita bruta de serviços (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços de informação e comunicação - Receita bruta de serviços (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.11 - Serviços de informação e comunicação - Receita bruta de ###Serviços.xlsx",
  overwrite = TRUE
)
###05.12 - Serviços de informação e comunicação - Salários, retiradas e outras remunerações (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços de informação e comunicação - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços de informação e comunicação - Salários, retiradas e outras remunerações (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços de informação e comunicação - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.12 - Serviços de informação e comunicação - Salários, retiradas e outras remunerações.xlsx",
  overwrite = TRUE
)
###05.13 - Serviços Profissionais, administrativos e complementares - Número de empresas (Unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços profissionais, administrativos e complementares - Número de empresas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços profissionais, administrativos e complementares - Número de empresas (Unidades)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços profissionais, administrativos e complementares - Número de empresas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.13 - Serviços Profissionais, administrativos e complementares - Número de empresas.xlsx",
  overwrite = TRUE
)
###05.14 - Serviços Profissionais, administrativos e complementares - Pessoal ocupado em 31/12 (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços de informação e comunicação - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços de informação e comunicação - Pessoal ocupado em 31/12 (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços de informação e comunicação - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.14 - Serviços Profissionais, administrativos e complementares - Pessoal ocupado em 31_12.xlsx",
  overwrite = TRUE
)
###05.15 - Serviços Profissionais, administrativos e complementares - Receita bruta de Serviços (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços profissionais, administrativos e complementares - Receita bruta de serviços (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços profissionais, administrativos e complementares - Receita bruta de serviços (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços profissionais, administrativos e complementares - Receita bruta de serviços (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.15 - Serviços Profissionais, administrativos e complementares - Receita bruta de ###Serviços.xlsx",
  overwrite = TRUE
)
###05.16 - Serviços Profissionais, administrativos e complementares - Salários, retiradas e outras remunerações (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços de informação e comunicação - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços de informação e comunicação - Salários, retiradas e outras remunerações (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços de informação e comunicação - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.16 - Serviços Profissionais, administrativos e complementares - Salários, retiradas e outras remunerações.xlsx",
  overwrite = TRUE
)
###05.17 - Transportes, Serviços auxiliares aos Transportes e correio - Número de empresas (Unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Transportes, serviços auxiliares aos transportes e correio - Número de empresas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Transportes, serviços auxiliares aos transportes e correio - Número de empresas (Unidades)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Transportes, serviços auxiliares aos transportes e correio - Número de empresas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.17 - Transportes, Serviços auxiliares aos Transportes e correio - Número de empresas.xlsx",
  overwrite = TRUE
)
###05.18 - Transportes, serviços auxiliares aos Transportes e correio - Pessoal ocupado em 31/12 (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Transportes, serviços auxiliares aos transportes e correio - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Transportes, serviços auxiliares aos transportes e correio - Pessoal ocupado em 31/12 (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Transportes, serviços auxiliares aos transportes e correio - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.18 - Transportes, serviços auxiliares aos Transportes e correio - Pessoal ocupado em 31_12.xlsx",
  overwrite = TRUE
)
###05.19 - Transportes, Serviços auxiliares aos Transportes e correio - Receita bruta de ###Serviços (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Transportes, serviços auxiliares aos transportes e correio - Receita bruta de serviços (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Transportes, serviços auxiliares aos transportes e correio - Receita bruta de serviços (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Transportes, serviços auxiliares aos transportes e correio - Receita bruta de serviços (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.19 - Transportes, Serviços auxiliares aos Transportes e correio - Receita bruta de ###Serviços.xlsx",
  overwrite = TRUE
)
###05.20 - Transportes, Serviços auxiliares aos ###Transportes e correio - Salários, retiradas e outras remunerações (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Transportes, serviços auxiliares aos transportes e correio - Salários, retiradas e outras remunerações, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Transportes, serviços auxiliares aos transportes e correio - Salários, retiradas e outras remunerações (Mil R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Transportes, serviços auxiliares aos transportes e correio - Salários, retiradas e outras remunerações, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.20 - Transportes_Serviços auxiliares aos Transportes e correio - Salários retiradas e outras remunerações.xlsx",
  overwrite = TRUE
)
###05.21 - Atividades Imobiliárias - Número de empresas (Unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Atividades Imobiliárias - Número de empresas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Atividades Imobiliárias - Número de empresas (Unidades)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Atividades Imobiliárias - Número de empresas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.21 - Atividades Imobiliárias - Número de empresas.xlsx",
  overwrite = TRUE
)
###05.22 - Atividades Imobiliárias - Pessoal ocupado em 31/12 (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Atividades imobiliárias - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Atividades imobiliárias - Pessoal ocupado em 31/12 (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Atividades imobiliárias - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.22 - Atividades Imobiliárias - Pessoal ocupado em 31_12.xlsx",
  overwrite = TRUE
)
###05.23 - Atividades Imobiliárias - Receita bruta de Serviços (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Atividades imobiliárias - Receita bruta de serviços (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Atividades imobiliárias - Receita bruta de serviços (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Atividades imobiliárias - Receita bruta de serviços (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.23 - Atividades Imobiliárias - Receita bruta de Serviços.xlsx",
  overwrite = TRUE
)
###05.24 - Atividades Imobiliárias - Salários, retiradas e outras remunerações (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Atividades imobiliárias - Salários, retiradas e outras remunerações, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Atividades imobiliárias - Salários, retiradas e outras remunerações (Mil R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Atividades imobiliárias - Salários, retiradas e outras remunerações, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.24 - Atividades Imobiliárias - Salários, retiradas e outras remunerações.xlsx",
  overwrite = TRUE
)
###05.25 - Serviços de Manutenção e reparação - Número de empresas (Unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços de manutenção e reparação - Número de empresas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços de manutenção e reparação - Número de empresas (Unidades)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços de manutenção e reparação - Número de empresas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.25 - Serviços de Manutenção e reparação - Número de empresas.xlsx",
  overwrite = TRUE
)
###05.26 - Serviços de Manutenção e reparação - Pessoal ocupado em 31/12 (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços de manutenção e reparação - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços de manutenção e reparação - Pessoal ocupado em 31/12 (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços de manutenção e reparação - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.26 - Serviços de Manutenção e reparação - Pessoal ocupado em 31_12.xlsx",
  overwrite = TRUE
)
###05.27 - Serviços de Manutenção e reparação - Receita bruta de ###Serviços (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços de manutenção e reparação - Receita bruta de serviços (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços de manutenção e reparação - Receita bruta de serviços (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços de manutenção e reparação - Receita bruta de serviços (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.27 - Serviços de Manutenção e reparação - Receita bruta de ###Serviços.xlsx",
  overwrite = TRUE
)
###05.28 - Serviços de Manutenção e reparação - Salários, retiradas e outras remunerações (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Serviços"
d1$indicador <- "Serviços de manutenção e reparação - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Serviços de manutenção e reparação - Salários, retiradas e outras remunerações (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Serviços"
d2$indicador <- "Serviços de manutenção e reparação - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/servicos/05.28 - Serviços de Manutenção e reparação - Salários, retiradas e outras remunerações.xlsx",
  overwrite = TRUE
)

##06 - CONSTRUÇÃO CIVIL----
###06.1 - Empresas de construção - Número de empresas ativas (Unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Construção Civil"
d1$indicador <- "Empresas de construção - Número de empresas ativas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Número de empresas ativas (Unidades)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Construção Civil"
d2$indicador <- "Empresas de construção - Número de empresas ativas (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/construcao_civil/06.1 - Empresas de construção - Número de empresas ativas.xlsx",
  overwrite = TRUE
)
###06.2 - Empresas de construção - Pessoal ocupado em 31/12 (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Construção Civil"
d1$indicador <- "Empresas de construção - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Empresas de construção - Pessoal ocupado em 31/12 (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Construção Civil"
d2$indicador <- "Empresas de construção - Pessoal ocupado em 31/12 (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/construcao_civil/06.2 - Empresas de construção - Pessoal ocupado em 31_12.xlsx",
  overwrite = TRUE
)
###06.3 - Empresas de construção - Salários, retiradas e outras remunerações (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Construção Civil"
d1$indicador <- "Empresas de construção - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Empresas de construção - Salários, retiradas e outras remunerações (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Construção Civil"
d2$indicador <- "Empresas de construção - Salários, retiradas e outras remunerações (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/construcao_civil/06.3 - Empresas de construção - Salários, retiradas e outras remunerações.xlsx",
  overwrite = TRUE
)
###06.4 - Empresas de construção - Custos das obras e/ou ###Serviços da construção (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Construção Civil"
d1$indicador <- "Empresas de construção - Custos das obras e/ou serviços da construção (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Empresas de construção - Custos das obras e/ou serviços da construção (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Construção Civil"
d2$indicador <- "Empresas de construção - Custos das obras e/ou serviços da construção (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/construcao_civil/06.4 - Empresas de construção - Custos das obras e/ou ###Serviços da construção.xlsx",
  overwrite = TRUE
)
###06.5 - Empresas de construção - Receita líquida (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Construção Civil"
d1$indicador <- "Empresas de construção - Receita líquida (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Empresas de construção - Receita líquida (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Construção Civil"
d2$indicador <- "Empresas de construção - Receita líquida (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/construcao_civil/06.5 - Empresas de construção - Receita líquida.xlsx",
  overwrite = TRUE
)
###06.6 - Empresas de construção - Valor bruto da Produção (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Construção Civil"
d1$indicador <- "Empresas de construção - Valor bruto da produção (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Empresas de construção - Valor bruto da produção (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Construção Civil"
d2$indicador <- "Empresas de construção - Valor bruto da produção (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/construcao_civil/06.6 - Empresas de construção - Valor bruto da Produção.xlsx",
  overwrite = TRUE
)

##07 - COMÉRCIO----
###07.1 - Total de Comércio - Número de unidades locais com receita de revenda (Unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Total de Comércio - Número de unidades locais com receita de revenda (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Número de unidades locais com receita de revenda (Unidades)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Total de Comércio - Número de unidades locais com receita de revenda (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.1 - Total de Comércio - Número de unidades locais com receita de revenda.xlsx",
  overwrite = TRUE
)
###07.2 - Total de Comércio - Pessoal ocupado em 3112 em empresas comerciais (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Total de Comércio - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas),  Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Pessoal ocupado em 31/12 em empresas comerciais (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Total de Comércio - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas),  Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.2 - Total de Comércio - Pessoal ocupado em 3112 em empresas comerciais.xlsx",
  overwrite = TRUE
)
###07.3 - Total de Comércio - Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Total de Comércio - Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Total de Comércio - Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.3 - Total de Comércio - Gastos com salários, retiradas e outras remunerações em empresas comerciais.xlsx",
  overwrite = TRUE
)
###07.4 - Total de Comércio - Receita bruta de revenda de mercadorias (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Total de Comércio - Receita bruta de revenda de mercadorias (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Receita bruta de revenda de mercadorias (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Total de Comércio - Receita bruta de revenda de mercadorias (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.4 - Total de Comércio - Receita bruta de revenda de mercadorias.xlsx",
  overwrite = TRUE
)
###07.5 - Comércio de veículos, peças e motocicletas - Número de unidades locais com receita de revenda (Unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Comércio de veículos, peças e motocicletas - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Comércio de veículos, peças e motocicletas - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Comércio de veículos, peças e motocicletas - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.5 - Comércio de veículos peças e motocicletas - Número de unidades locais com receita de revenda.xlsx",
  overwrite = TRUE
)
###07.6 - Comércio de veículos, peças e motocicletas - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Comércio de veículos, peças e motocicletas - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Comércio de veículos, peças e motocicletas - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Comércio de veículos, peças e motocicletas - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.6 - Comércio de veículos peças e motocicletas - Pessoal ocupado em 31_12 em empresas comerciais.xlsx",
  overwrite = TRUE
)
###07.7 - Comércio de veículos, peças e motocicletas - Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Comércio de veículos, peças e motocicletas - Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Comércio de veículos, peças e motocicletas - Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.7 - Comércio de veículos, peças e motocicletas - Gastos com salários, retiradas e outras remunerações em empresas comerciais.xlsx",
  overwrite = TRUE
)
###07.8 - Comércio de veículos, peças e motocicletas - Receita bruta de revenda de mercadorias (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Comércio de veículos, peças e motocicletas - Receita bruta de revenda de mercadorias (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Comércio de veículos, peças e motocicletas - Receita bruta de revenda de mercadorias (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Comércio de veículos, peças e motocicletas - Receita bruta de revenda de mercadorias (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.8 - Comércio de veículos peças e motocicletas - Receita bruta de revenda de mercadorias.xlsx",
  overwrite = TRUE
)
###07.9 - Comércio por atacado - Número de unidades locais com receita de revenda (Unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Comércio por atacado - Número de unidades locais com receita de revenda (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Comércio por atacado - Número de unidades locais com receita de revenda (Unidades)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Comércio por atacado - Número de unidades locais com receita de revenda (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.9 - Comércio por atacado - Número de unidades locais com receita de revenda.xlsx",
  overwrite = TRUE
)
###07.10 - Comércio por atacado - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Comércio por atacado - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Comércio por atacado - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Comércio por atacado - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.10 - Comércio por atacado - Pessoal ocupado em 31/12 em empresas comerciais.xlsx",
  overwrite = TRUE
)
###07.11 - Comércio por atacado - Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Comércio por atacado - Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Comércio por atacado - Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Comércio por atacado - Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.11 - Comércio por atacado - Gastos com salários, retiradas e outras remunerações em empresas comerciais.xlsx",
  overwrite = TRUE
)
###07.12 - Comércio por atacado - Receita bruta de revenda de mercadorias (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Comércio por atacado - Receita bruta de revenda de mercadorias (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Comércio por atacado - Receita bruta de revenda de mercadorias (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Comércio por atacado - Receita bruta de revenda de mercadorias (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.12 - Comércio por atacado - Receita bruta de revenda de mercadorias.xlsx",
  overwrite = TRUE
)
###07.13 - Comércio varejista - Número de unidades locais com receita de revenda (Unidades)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Comércio varejista - Número de unidades locais com receita de revenda (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Comércio varejista - Número de unidades locais com receita de revenda (Unidades)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Comércio varejista - Número de unidades locais com receita de revenda (Unidades), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.13 - Comércio varejista - Número de unidades locais com receita de revenda.xlsx",
  overwrite = TRUE
)
###07.14 - Comércio varejista - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Comércio varejista - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Comércio varejista - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Comércio varejista - Pessoal ocupado em 31/12 em empresas comerciais (Pessoas), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.14 - Comércio varejista - Pessoal ocupado em 31_12 em empresas comerciais.xlsx",
  overwrite = TRUE
)
###07.15 - Comércio varejista - Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Comércio varejista - Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Comércio varejista - Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Comércio varejista - Gastos com salários, retiradas e outras remunerações em empresas comerciais (Mil Reais), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.15 - Comércio varejista - Gastos com salários, retiradas e outras remunerações em empresas comerciais.xlsx",
  overwrite = TRUE
)
###07.16 - Comércio varejista - Receita bruta de revenda de mercadorias (Mil Reais)----
d1 <- bases
d1$tematica <- "Economia"
d1$subtematica <- "Comércio"
d1$indicador <- "Comércio varejista - Receita bruta de revenda de mercadorias (Mil Reais),  Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Comércio varejista - Receita bruta de revenda de mercadorias (Mil Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Economia"
d2$subtematica <- "Comércio"
d2$indicador <- "Comércio varejista - Receita bruta de revenda de mercadorias (Mil Reais),  Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2
)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(
    wb,
    sheet = nome,
    cols = 1:ncol(dados_lista[[nome]]),
    widths = "auto"
  )
}
# Salvar o arquivo Excel
saveWorkbook(
  wb,
  "economia/comercio/07.16 - Comércio varejista - Receita bruta de revenda de mercadorias.xlsx",
  overwrite = TRUE
)
