#PREVIDÊNCIA SOCIAL----
bases$subtematica <- "-"

##01 - Número de Benefícios Previdenciários Emitidos----
d1 <- bases
d1$tematica <- "Previdência Social"
d1$indicador <- "Número de Benefícios Previdenciários Emitidos"
d1$categoria1 <- "Número de Benefícios"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Previdência Social"
d2$indicador <- "Número de Benefícios Previdenciários Emitidos"
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
  "previdencia_social/01 - Número de Benefícios Previdenciários Emitidos.xlsx",
  overwrite = TRUE
)
##02 - Número de Benefícios previdenciários Emitidos Urbano----
d1 <- bases
d1$tematica <- "Previdência Social"
d1$indicador <- "Número de Benefícios previdenciários Emitidos Urbano"
d1$categoria1 <- "Número de Benefícios"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Previdência Social"
d2$indicador <- "Número de Benefícios previdenciários Emitidos Urbano"
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
  "previdencia_social/02 - Número de Benefícios previdenciários Emitidos Urbano.xlsx",
  overwrite = TRUE
)
##03 - Número de Benefícios previdenciários Emitidos Rural----
d1 <- bases
d1$tematica <- "Previdência Social"
d1$indicador <- "Número de Benefícios previdenciários Emitidos Rural"
d1$categoria1 <- "Número de Benefícios"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Previdência Social"
d2$indicador <- "Número de Benefícios previdenciários Emitidos Rural"
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
  "previdencia_social/03 - Número de Benefícios previdenciários Emitidos Rural.xlsx",
  overwrite = TRUE
)
##04 - Valor Aplicado em Benefícios previdenciários Emitidos Total (R$ Mil)----
d1 <- bases
d1$tematica <- "Previdência Social"
d1$indicador <- "Valor Aplicado em Benefícios previdenciários Emitidos Total (R$ Mil)"
d1$categoria1 <- "Valor Aplicado Total (R$ Mil)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Previdência Social"
d2$indicador <- "Valor Aplicado em Benefícios previdenciários Emitidos Total (R$ Mil)"
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
  "previdencia_social/Valor Aplicado em Benefícios previdenciários Emitidos Total.xlsx",
  overwrite = TRUE
)
##05 - Valor Aplicado em Benefícios previdenciários Emitidos Urbano (R$ Mil)----
d1 <- bases
d1$tematica <- "Previdência Social"
d1$indicador <- "Valor Aplicado em Benefícios previdenciários Emitidos Urbano (R$ Mil)"
d1$categoria1 <- "Valor Aplicado Total (R$ Mil)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Previdência Social"
d2$indicador <- "Valor Aplicado em Benefícios previdenciários Emitidos Urbano (R$ Mil)"
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
  "previdencia_social/05 - Valor Aplicado em Benefícios previdenciários Emitidos Urbano.xlsx",
  overwrite = TRUE
)
##06 - Valor Aplicado em Benefícios previdenciários Emitidos Rural (R$ Mil)----
d1 <- bases
d1$tematica <- "Previdência Social"
d1$indicador <- "Valor Aplicado em Benefícios previdenciários Emitidos Rural (R$ Mil)"
d1$categoria1 <- "Valor Aplicado Total (R$ Mil)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Previdência Social"
d2$indicador <- "Valor Aplicado em Benefícios previdenciários Emitidos Rural (R$ Mil)"
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
  "previdencia_social/06 - Valor Aplicado em Benefícios previdenciários Emitidos Rural.xlsx",
  overwrite = TRUE
)
##07 - Valor Arrecado pela Previdência Social (R$ Mil)----
d1 <- bases
d1$tematica <- "Previdência Social"
d1$indicador <- "Valor Arrecado pela Previdência Social (R$ Mil)"
d1$categoria1 <- "Valor Arrecado (R$ Mil)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Previdência Social"
d2$indicador <- "Valor Arrecado pela Previdência Social (R$ Mil)"
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
  "previdencia_social/07 - Valor Arrecado pela Previdência Social.xlsx",
  overwrite = TRUE
)