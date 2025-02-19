#HABITAÇÃO E SANEAMENTO----
bases$subtematica <- "-"

##01 - Percentual de Domicílios com Rede Geral de Distribuição de Água----
d1 <- bases
d1$tematica <- "Habitação e Saneamento"
d1$indicador <- "Percentual de Domicílios com Rede Geral de Distribuição de Água"
d1$categoria1 <- "Percentual"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Habitação e Saneamento"
d2$indicador <- "Percentual de Domicílios com Rede Geral de Distribuição de Água"
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
  "habitacao_e_saneamento/01 - Percentual de Domicílios com Rede Geral de Distribuição de Água.xlsx",
  overwrite = TRUE
)
##02 - Percentual de Domicílios com Esgotamento Sanitário por Rede Geral ou Fossa Ligada à Rede----
d1 <- bases
d1$tematica <- "Habitação e Saneamento"
d1$indicador <- "Percentual de Domicílios com Esgotamento Sanitário por Rede Geral ou Fossa Ligada à Rede"
d1$categoria1 <- "Percentual"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Habitação e Saneamento"
d2$indicador <- "Percentual de Domicílios com Esgotamento Sanitário por Rede Geral ou Fossa Ligada à Rede"
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
  "habitacao_e_saneamento/02 - Percentual de Domicílios com Esgotamento Sanitário por Rede Geral ou Fossa Ligada à Rede.xlsx",
  overwrite = TRUE
)
##03 - Percentual de Domicílios com Coleta Direta e Indireta de Lixo----
d1 <- bases
d1$tematica <- "Habitação e Saneamento"
d1$indicador <- "Percentual de Domicílios com Coleta Direta e Indireta de Lixo"
d1$categoria1 <- "Percentual"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Habitação e Saneamento"
d2$indicador <- "Percentual de Domicílios com Coleta Direta e Indireta de Lixo"
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
  "habitacao_e_saneamento/03 - Percentual de Domicílios com Coleta Direta e Indireta de Lixo.xlsx",
  overwrite = TRUE
)
##04 - Percentual de Domicílios com Materiais Duráveis----
d1 <- bases
d1$tematica <- "Habitação e Saneamento"
d1$indicador <- "Percentual de Domicílios com Materiais Duráveis"
d1$categoria1 <- "Percentual"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Habitação e Saneamento"
d2$indicador <- "Percentual de Domicílios com Materiais Duráveis"
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
  "habitacao_e_saneamento/04 - Percentual de Domicílios com Materiais Duráveis.xlsx",
  overwrite = TRUE
)
##05 - Percentual de Domicílios Próprios já Quitados----
d1 <- bases
d1$tematica <- "Habitação e Saneamento"
d1$indicador <- "Percentual de Domicílios Próprios já Quitados"
d1$categoria1 <- "Percentual"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Habitação e Saneamento"
d2$indicador <- "Percentual de Domicílios Próprios já Quitados"
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
  "habitacao_e_saneamento/05 - Percentual de Domicílios Próprios já Quitados.xlsx",
  overwrite = TRUE
)
##06 - Percentual de domicílios que possuem banheiro de uso exclusivo----
d1 <- bases
d1$tematica <- "Habitação e Saneamento"
d1$indicador <- "Percentual de domicílios que possuem banheiro de uso exclusivo"
d1$categoria1 <- "Percentual"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Habitação e Saneamento"
d2$indicador <- "Percentual de domicílios que possuem banheiro de uso exclusivo"
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
  "habitacao_e_saneamento/06 - Percentual de domicílios que possuem banheiro de uso exclusivo.xlsx",
  overwrite = TRUE
)