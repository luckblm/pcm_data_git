#INFRAESTRUTURA----
bases$subtematica <- "-"

##01 - Frota de veículos Emplacados----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Frota de Veículos Emplacados, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Frota de Veículos Emplacados"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Frota de Veículos Emplacados, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "infraestrutura/01 - Frota de veículos Emplacados.xlsx",
  overwrite = TRUE
)
##02 - Energia Elétrica (cobertura domiciliar)----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Energia Elétrica (cobertura domiciliar), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Energia Elétrica (cobertura domiciliar)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Energia Elétrica (cobertura domiciliar), segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "infraestrutura/02 - Energia Elétrica_cobertura domiciliar.xlsx",
  overwrite = TRUE
)
##03 - Consumo Total de Energia Elétrica (GWh)----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Consumo Total de Energia Elétrica, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Consumo (GWh)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Consumo Total de Energia Elétrica, segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "infraestrutura/03 - Consumo Total de Energia Elétrica_GWh.xlsx",
  overwrite = TRUE
)
##04 - Consumo Residencial de Eletricidade (GWh)----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Consumo Residencial de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Consumo Residencial de Energia Elétrica (GWh)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Consumo Residencial de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "infraestrutura/04 - Consumo Residencial de Eletricidade_GWh.xlsx",
  overwrite = TRUE
)
##05 - Consumo Comercial de Eletricidade (GWh)----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Consumo Comercial de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Consumo Comercial de Energia Elétrica (GWh)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Consumo Comercial de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "infraestrutura/05 - Consumo Comercial de Eletricidade (GWh).xlsx",
  overwrite = TRUE
)
##06 - Consumo Industrial de Eletricidade (GWh)----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Consumo Industrial de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Consumo Industrial de Energia Elétrica (GWh)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Consumo Industrial de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "infraestrutura/06 - Consumo Industrial de Eletricidade_GWh.xlsx",
  overwrite = TRUE
)
##07 - Outros Consumos de Eletricidade (GWh)----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Outros Consumos de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Outros Consumo de Energia Elétrica (GWh)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Outros Consumos de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "infraestrutura/07 - Outros Consumos de Eletricidade_GWh.xlsx",
  overwrite = TRUE
)
##08 - Consumidores Totais de Energia Elétrica----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Consumidores Totais de Energia Elétrica, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Consumidores"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Consumidores Totais de Energia Elétrica, segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "infraestrutura/08 - Consumidores Totais de Energia Elétrica.xlsx",
  overwrite = TRUE
)
##09 - Consumidores Residenciais de Eletricidade (GWh)----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Consumidores Residenciais de Energia Eletrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Consumidores Residenciais de Energia Elétrica"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Consumidores Residenciais de Energia Eletrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "infraestrutura/09 - Consumidores Residenciais de Eletricidade_=GWh.xlsx",
  overwrite = TRUE
)
##10 - Consumidores Comerciais de Eletricidade (GWh)----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Consumidores Comerciais de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Consumidores Comerciais de Energia Elétrica"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Consumidores Comerciais de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "infraestrutura/.xlsx",
  overwrite = TRUE
)
##11 - Consumidores Industriais de Eletricidade (GWh)----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Consumidores Industriais de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Consumidores Industriais de Energia Elétrica"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Consumidores Industriais de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "infraestrutura/11 - Consumidores Industriais de Eletricidade_GWh.xlsx",
  overwrite = TRUE
)
##12 - Outros Consumidores de Eletricidade (GWh)----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Outros Consumidores de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Outros Consumidores de Energia Elétrica"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Outros Consumidores de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "infraestrutura/12 - Outros Consumidores de Eletricidade_GWh.xlsx",
  overwrite = TRUE
)
##13 - Geração de Energia Elétrica (GWh)----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Geração de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Geração de Energia Elétrica (GWh)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Geração de Energia Elétrica, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "infraestrutura/13 - Geração de Energia Elétrica_GWh.xlsx",
  overwrite = TRUE
)
##14 - Percentual de Domicílios com Canalização Interna----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Percentual de Domicílios com Canalização Interna, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Percentual de Domicílios com Canalização Interna"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Percentual de Domicílios com Canalização Interna, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "infraestrutura/.xlsx",
  overwrite = TRUE
)
##15 - Percentual de domicílios com existência de utilização de Internet----

d1 <- bases
d1$tematica <- "Infraestrutura"
d1$indicador <- "Percentual de domicílios com existência de utilização de Internet, segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- ""
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Infraestrutura"
d2$indicador <- "Percentual de domicílios com existência de utilização de Internet, segundo Brasil, Grandes regiões e Unidades da Federação"
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
  "infraestrutura/15 - Percentual de domicílios com existência de utilização de Internet.xlsx",
  overwrite = TRUE
)