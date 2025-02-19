#TERRITÓRIO----

##01 - Quantidade de Municípios e Respectivas Áreas Territoriais----
bases$subtematica <- "-"

d1 <- bases
d1$tematica <- "Território"
d1$indicador <- "Quatidade de Municípios e Respectivas Áreas Territoriais, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Municípios"
d1$categoria2 <- "Quantidade"

d2 <- bases
d2$tematica <- "Território"
d2$indicador <- "Quatidade de Municípios e Respectivas Áreas Territoriais, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Municípios"
d2$categoria2 <- "Ranking"

d3 <- bases
d3$tematica <- "Território"
d3$indicador <- "Quatidade de Municípios e Respectivas Áreas Territoriais, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Área Territorial (km²)"
d3$categoria2 <- "Área"

d4 <- bases
d4$tematica <- "Território"
d4$indicador <- "Quatidade de Municípios e Respectivas Áreas Territoriais, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Área Territorial (km²)"
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
  "territorio/01 - Quantidade de Municípios e Respectivas Áreas Territoriais.xlsx",
  overwrite = TRUE
)
##02 - Quantidade de Alguns tipos de Unidades Territoriais----
bases$subtematica <- "-"

d1 <- bases
d1$tematica <- "Território"
d1$indicador <- "Quantidade de alguns tipos de Unidades Territoriais exixtentes, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Unidades Territoriais"
d1$categoria2 <- "Distritos"
d1$categoria3 <- "Quantidade"

d2 <- bases
d2$tematica <- "Território"
d2$indicador <- "Quantidade de alguns tipos de Unidades Territoriais exixtentes, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Unidades Territoriais"
d2$categoria2 <- "Distritos"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Território"
d3$indicador <- "Quantidade de alguns tipos de Unidades Territoriais exixtentes, segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "Unidades Territoriais"
d3$categoria2 <- "Vilas"
d3$categoria3 <- "Quantidade"

d4 <- bases
d4$tematica <- "Território"
d4$indicador <- "Quantidade de alguns tipos de Unidades Territoriais exixtentes, segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "Unidades Territoriais"
d4$categoria2 <- "Vilas"
d4$categoria3 <- "Ranking"

d5 <- bases
d5$tematica <- "Território"
d5$indicador <- "Quantidade de alguns tipos de Unidades Territoriais exixtentes, segundo Brasil, Grandes Regiões e Unidades da Federação"
d5$categoria1 <- "Unidades Territoriais"
d5$categoria2 <- "Terra Índígena"
d5$categoria3 <- "Quantidade "

d6 <- bases
d6$tematica <- "Território"
d6$indicador <- "Quantidade de alguns tipos de Unidades Territoriais exixtentes, segundo Brasil, Grandes Regiões e Unidades da Federação"
d6$categoria1 <- "Unidades Territoriais"
d6$categoria2 <- "Terra Índígena"
d6$categoria3 <- "Ranking"

d7 <- bases
d75$tematica <- "Território"
d7$indicador <- "Quantidade de alguns tipos de Unidades Territoriais exixtentes, segundo Brasil, Grandes Regiões e Unidades da Federação"
d7$categoria1 <- "Unidades Territoriais"
d7$categoria2 <- "UCA¹"
d7$categoria3 <- "Quantidade "

d8 <- bases
d8$tematica <- "Território"
d8$indicador <- "Quantidade de alguns tipos de Unidades Territoriais exixtentes, segundo Brasil, Grandes Regiões e Unidades da Federação"
d8$categoria1 <- "Unidades Territoriais"
d8$categoria2 <- "UCA¹"
d8$categoria3 <- "Ranking"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6,
  "d7" = d7,
  "d8" = d8
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
  "territorio/02 - Quantidade de Alguns tipos de Unidades Territoriais.xlsx",
  overwrite = TRUE
)