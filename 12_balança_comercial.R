#BALANÇA COMERCIAL----

bases$subtematica <- "-"

##01 - Exportação (US$)----

d1 <- bases
d1$tematica <- "Balança Comercial"
d1$indicador <- "Exportação (US$), Segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- "Exportação (US$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Balança Comercial"
d2$indicador <- "Exportação (US$), Segundo Brasil, Grandes regiões e Unidades da Federação"
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
  "balanca_comercial/01 - Exportação.xlsx",
  overwrite = TRUE
)
##02 - Importação (US$)----
d1 <- bases
d1$tematica <- "Balança Comercial"
d1$indicador <- "Importação (US$), Segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- "Importação (US$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Balança Comercial"
d2$indicador <- "Importação (US$), Segundo Brasil, Grandes regiões e Unidades da Federação"
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
  "balanca_comercial/02 - Importação.xlsx",
  overwrite = TRUE
)
##03 - Saldo da Balança Comercial (US$)----
d1 <- bases
d1$tematica <- "Balança Comercial"
d1$indicador <- "Saldo da Balança Comercial (US$), Segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- "Saldo (US$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Balança Comercial"
d2$indicador <- "Saldo da Balança Comercial (US$), Segundo Brasil, Grandes regiões e Unidades da Federação"
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
  "balanca_comercial/03 - Saldo da Balança Comercial.xlsx",
  overwrite = TRUE
)