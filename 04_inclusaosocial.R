#INCLUSÃO SOCIAL----

bases$subtematica <- "-"

##01 - Famílias Cadastradas no Sistema CAD Único----
d1 <- bases
d1$tematica <- "Inclusão social"
d1$indicador <- "Famílias Cadastradas no Sistema CAD Único"
d1$categoria1 <- "Total de famílias cadastradas"
d1$categoria2 <- "Total"

d2 <- bases
d2$tematica <- "Inclusão social"
d2$indicador <- "Famílias Cadastradas no Sistema CAD Único"
d2$categoria1 <- "Total de famílias cadastradas"
d2$categoria2 <- "Ranking"

d3 <- bases
d3$tematica <- "Inclusão social"
d3$indicador <- "Famílias Cadastradas no Sistema CAD Único"
d3$categoria1 <- "Famílias com renda per capita abaixo de ½ Salário Mínino"
d3$categoria2 <- "Total"

d4 <- bases
d4$tematica <- "Inclusão social"
d4$indicador <- "Famílias Cadastradas no Sistema CAD Único"
d4$categoria1 <- "Famílias com renda per capita abaixo de ½ Salário Mínino"
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
  "inclusao_social/01 - Famílias Cadastradas no Sistema CAD Único.xlsx",
  overwrite = TRUE
)
##02 - Pessoas Cadastradas no Sistema CAD Único----
d1 <- bases
d1$tematica <- "Inclusão Social"
d1$indicador <- "Pessoas Cadastradas no Sistema CAD Único"
d1$categoria1 <- "Total de pessoas inscritas"
d1$categoria2 <- "Total"

d2 <- bases
d2$tematica <- "Inclusão Social"
d2$indicador <- "Pessoas Cadastradas no Sistema CAD Único"
d2$categoria1 <- "Total de pessoas inscritas"
d2$categoria2 <- "Ranking"

d3 <- bases
d3$tematica <- "Inclusão Social"
d3$indicador <- "Pessoas Cadastradas no Sistema CAD Único"
d3$categoria1 <- "Pessoas inscritas no Cadastro Único com renda mensal de até meio salário mínimo"
d3$categoria2 <- "Total"

d4 <- bases
d4$tematica <- "Inclusão Social"
d4$indicador <- "Pessoas Cadastradas no Sistema CAD Único"
d4$categoria1 <- "Pessoas inscritas no Cadastro Único com renda mensal de até meio salário mínimo"
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
  "inclusao_social/02 - Pessoas Cadastradas no Sistema CAD Único.xlsx",
  overwrite = TRUE
)
##03 - Famílias Beneficiadas do Programa Bolsa Família----
d1 <- bases
d1$tematica <- "Inclusão Social"
d1$indicador <- "Famílias Beneficiadas do Programa Bolsa Família"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Inclusão Social"
d2$indicador <- "Famílias Beneficiadas do Programa Bolsa Família"
d2$categoria1 <- "Raking"
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
  "inclusao_social/03 - Famílias Beneficiadas do Programa Bolsa Família.xlsx",
  overwrite = TRUE
)
##04 - Valor Aplicado no Programa Bolsa Família (R$)----
d1 <- bases
d1$tematica <- "Inclusão Social"
d1$indicador <- "Valor Aplicado no Programa Bolsa Família (R$)"
d1$categoria1 <- "Valor total (R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Inclusão Social"
d2$indicador <- "Valor Aplicado no Programa Bolsa Família (R$)"
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
  "inclusao_social/04 - Valor Aplicado no Programa Bolsa Família.xlsx",
  overwrite = TRUE
)
##05 - Famílias Cadastradas no Programa Auxílio Brasil----
d1 <- bases
d1$tematica <- "Inclusão Social"
d1$indicador <- "Famílias Cadastradas no Programa Auxílio Brasil"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Inclusão Social"
d2$indicador <- "Famílias Cadastradas no Programa Auxílio Brasil"
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
  "inclusao_social/05 - Famílias Cadastradas no Programa Auxílio Brasil.xlsx",
  overwrite = TRUE
)
##06 - Valor Aplicado no Programa Auxílio Brasil----
d1 <- bases
d1$tematica <- "Inclusão Social"
d1$indicador <- "Valor Aplicado no Programa Auxílio Brasil"
d1$categoria1 <- "Valor total (R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Inclusão Social"
d2$indicador <- "Valor Aplicado no Programa Auxílio Brasil"
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
  "inclusao_social/06 - Valor Aplicado no Programa Auxílio Brasil.xlsx",
  overwrite = TRUE
)

