#SEGURANÇA----
bases$subtematica <- "-"

##01 - Taxa de Homicídios Total por 100.000 habitantes----
d1 <- bases
d1$tematica <- "Segurança"
d1$indicador <- "Taxa de Homicídios Total por 100.000 habitantes"
d1$categoria1 <- "Taxa"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Segurança"
d2$indicador <- "Taxa de Homicídios Total por 100.000 habitantes"
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
  "seguranca/01 - Taxa de Homicídios Total por 100.000 habitantes.xlsx",
  overwrite = TRUE
)
##02 - Taxa de Homicídios de Jovens (15 a 29 anos) por 100.000 Jovens----
d1 <- bases
d1$tematica <- "Segurança"
d1$indicador <- "Taxa de Homicídios de Jovens (15 a 29 anos) por 100.000 Jovens"
d1$categoria1 <- "Taxa"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Segurança"
d2$indicador <- "Taxa de Homicídios de Jovens (15 a 29 anos) por 100.000 Jovens"
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
  "seguranca/02 - Taxa de Homicídios de Jovens (15 a 29 anos) por 100.000 Jovens.xlsx",
  overwrite = TRUE
)
##03 - Taxa de Morte no Trânsito por 100.000 habitantes----
d1 <- bases
d1$tematica <- "Segurança"
d1$indicador <- "Taxa de Morte no Trânsito por 100.000 habitantes"
d1$categoria1 <- "Taxa"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Segurança"
d2$indicador <- "Taxa de Morte no Trânsito por 100.000 habitantes"
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
  "seguranca/03 - Taxa de Morte no Trânsito por 100.000 habitantes.xlsx",
  overwrite = TRUE
)