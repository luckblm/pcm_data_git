#MEIO AMBIENTE----
bases$subtematica <- "-"

##01 - Taxa de Desflorestamento Anual----

d1 <- bases
d1$tematica <- "Meio Ambiente"
d1$indicador <- "Desflorestamento Anual (km2/ano), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Desflorestamento (KM²)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Meio Ambiente"
d2$indicador <- "Desflorestamento Anual (km2/ano), Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "meio_ambiente/01 - Taxa de Desflorestamento Anual.xlsx",
  overwrite = TRUE
)
##02 - Degradação Florestal na Amazônia Legal----
d1 <- bases
d1$tematica <- "Meio Ambiente"
d1$indicador <- "Degradação Florestal do Bioma Amazônicol  (km2 /ano), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Degradação Florestal Bioma Amazônico  (km 2 /ano)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Meio Ambiente"
d2$indicador <- "Degradação Florestal do Bioma Amazônicol  (km2 /ano), Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "meio_ambiente/02 - Degradação Florestal na Amazônia Legal.xlsx",
  overwrite = TRUE
)
##03 - Área de Florestas (km²) e Hidrografia (km²), Amazônia Legal----
d1 <- bases
d1$tematica <- "Meio Ambiente"
d1$indicador <- "Área de Floresta e Hidrografia - segundo Amazônia Legal"
d1$categoria1 <- "Área de Floresta (km²)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Meio Ambiente"
d2$indicador <- "Área de Floresta e Hidrografia - segundo Amazônia Legal"
d2$categoria1 <- "Área de Floresta (km²)"
d2$categoria2 <- "Ranking"

d3 <- bases
d3$tematica <- "Meio Ambiente"
d3$indicador <- "Área de Floresta e Hidrografia - segundo Amazônia Legal"
d3$categoria1 <- "Hidrografia (km²)"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Meio Ambiente"
d4$indicador <- "Área de Floresta e Hidrografia - segundo Amazônia Legal"
d4$categoria1 <- "Hidrografia (km²)"
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
  "meio_ambiente/03 - Área de Florestas e Hidrografia_Amazônia Legal.xlsx",
  overwrite = TRUE
)
##04 - Focos de Calor----
d <- bases
d$tematica <- "Meio Ambiente"
d$indicador <- "Focos de Calor, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d$categoria1 <- "Focos de Calor"
d$categoria2 <- "-"

d <- bases
d$tematica <- "Meio Ambiente"
d$indicador <- "Focos de Calor, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d$categoria1 <- "Ranking"
d$categoria2 <- "-"


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
  "meio_ambiente/04 - Focos de Calor.xlsx",
  overwrite = TRUE
)