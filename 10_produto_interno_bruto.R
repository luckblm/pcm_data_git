#PRODUTO INTERNO BRUTO----

bases$subtematica <- "-"

##01 - Produto Interno Bruto (valores correntes)----

d1 <- bases
d1$tematica <- "Produto Interno Bruto"
d1$indicador <- "Produto Interno Bruto (valores correntes), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "valores correntes"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Produto Interno Bruto"
d2$indicador <- "Produto Interno Bruto (valores correntes), segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "produto_interno_bruto/01 - Produto Interno Bruto_valores correntes.xlsx",
  overwrite = TRUE
)
##02 - Participação no Produto Interno Bruto----
d1 <- bases
d1$tematica <- "Produto Interno Bruto"
d1$indicador <- "Participação das Grandes Regiões e Unidades da Federação no Produto Interno Bruto"
d1$categoria1 <- "Participação no Produto Interno Bruto (%)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Produto Interno Bruto"
d2$indicador <- "Participação das Grandes Regiões e Unidades da Federação no Produto Interno Bruto"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d1" = d2
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
  "produto_interno_bruto/02 - Participação no Produto Interno Bruto.xlsx",
  overwrite = TRUE
)
##03 - Crescimento Real do Produto Interno Bruto----
d1 <- bases
d1$tematica <- "Produto Interno Bruto"
d1$indicador <- "Crescimento Real do Produto Interno Bruto, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Crescimento Real do Produto Interno Bruto (%)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Produto Interno Bruto"
d2$indicador <- "Crescimento Real do Produto Interno Bruto, segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "produto_interno_bruto/03 - Crescimento Real do Produto Interno Bruto.xlsx",
  overwrite = TRUE
)
##04 - Valor Adicionado Bruto Total----
d1 <- bases
d1$tematica <- "Produto Interno Bruto"
d1$indicador <- "Valor Adicionado Bruto Total, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Valor Adicionado Bruto Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Produto Interno Bruto"
d2$indicador <- "Valor Adicionado Bruto Total, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d1" = d2
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
  "produto_interno_bruto/04 - Valor Adicionado Bruto Total.xlsx",
  overwrite = TRUE
)
##05 - Impostos sobre produtos, líquidos de subsídios----
d1 <- bases
d1$tematica <- "Produto Interno Bruto"
d1$indicador <- "Impostos sobre produtos, líquidos de subsídios, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Impostos sobre produtos, líquidos de subsídios (Milhão Reais)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Produto Interno Bruto"
d2$indicador <- "Impostos sobre produtos, líquidos de subsídios, segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "produto_interno_bruto/05 - Impostos sobre produtos, líquidos de subsídios.xlsx",
  overwrite = TRUE
)
##06 - Produto Interno Bruto Per Capita (valores correntes)----
d1 <- bases
d1$tematica <- "Produto Interno Bruto"
d1$indicador <- "Produto Interno Bruto Per Capita (valores correntes), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "PIB Per Capita (1,00 R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Produto Interno Bruto"
d2$indicador <- "Produto Interno Bruto Per Capita (valores correntes), segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d1" = d2
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
  "produto_interno_bruto/06 - Produto Interno Bruto Per Capita_valores correntes.xlsx",
  overwrite = TRUE
)
##07 - Valor Adicionado Bruto do Setor Agropecuário----
d1 <- bases
d1$tematica <- "Produto Interno Bruto"
d1$indicador <- "Valor Adicionado Bruto do Setor Agropecuário, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Valor Adicionado Bruto"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Produto Interno Bruto"
d2$indicador <- "Valor Adicionado Bruto do Setor Agropecuário, segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "produto_interno_bruto/07 - Valor Adicionado Bruto do Setor Agropecuário.xlsx",
  overwrite = TRUE
)
##08 - Valor Adicionado Bruto do Setor Industrial----
d1 <- bases
d1$tematica <- "Produto Interno Bruto"
d1$indicador <- "Valor Adicionado Bruto do Setor Industrial, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Valor Adicionado Bruto"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Produto Interno Bruto"
d2$indicador <- "Valor Adicionado Bruto do Setor Industrial, segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "produto_interno_bruto/08 - Valor Adicionado Bruto do Setor Industrial.xlsx",
  overwrite = TRUE
)
##09 - Valor Adicionado Bruto do Setor de Serviços----
d1 <- bases
d1$tematica <- "Produto Interno Bruto"
d1$indicador <- "Valor Adicionado Bruto do Setor de Serviços, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Valor Adicionado Bruto"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Produto Interno Bruto"
d2$indicador <- "Valor Adicionado Bruto do Setor de Serviços, segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "produto_interno_bruto/09 - Valor Adicionado Bruto do Setor de Serviços.xlsx",
  overwrite = TRUE
)
##10 - Valor Adicionado Bruto da Administração Pública----
d1 <- bases
d1$tematica <- "Produto Interno Bruto"
d1$indicador <- "Valor Adicionado Bruto da Administração Pública, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Valor Adicionado Bruto"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Produto Interno Bruto"
d2$indicador <- "Valor Adicionado Bruto da Administração Pública, segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d1" = d2
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
  "produto_interno_bruto/10 - Valor Adicionado Bruto da Administração Pública.xlsx",
  overwrite = TRUE
)