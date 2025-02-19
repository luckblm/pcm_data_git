#MERCADO DE TRABALHO----

bases$subtematica <- "-"
##01 - Número de Vínculos Total----

d1 <- bases
d1$tematica <- "Mercado de Trabalho"
d1$indicador <- "Número de Vínculos Empregatícios Total, Segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- "Número de Vínculos Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Mercado de Trabalho"
d2$indicador <- "Número de Vínculos Empregatícios Total, Segundo Brasil, Grandes regiões e Unidades da Federação"
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
  "mercado_de_trabalho/01 - Número de Vínculos Total.xlsx",
  overwrite = TRUE
)
##02 - Número de Vínculos por Sexo----
d1 <- bases
d1$tematica <- "Mercado de Trabalho"
d1$indicador <- "Número de Vínculos por Sexo, Segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- "Masculino"
d1$categoria2 <- "Número de Vínculos"

d2 <- bases
d2$tematica <- "Mercado de Trabalho"
d2$indicador <- "Número de Vínculos por Sexo, Segundo Brasil, Grandes regiões e Unidades da Federação"
d2$categoria1 <- "Masculino"
d2$categoria2 <- "Ranking"

d3 <- bases
d3$tematica <- "Mercado de Trabalho"
d3$indicador <- "Número de Vínculos por Sexo, Segundo Brasil, Grandes regiões e Unidades da Federação"
d3$categoria1 <- "Feminino"
d3$categoria2 <- "Número de Vínculos"

d4 <- bases
d4$tematica <- "Mercado de Trabalho"
d4$indicador <- "Número de Vínculos por Sexo, Segundo Brasil, Grandes regiões e Unidades da Federação"
d4$categoria1 <- "Feminino"
d4$categoria2 <- "Ranking"

d5 <- bases
d5$tematica <- "Mercado de Trabalho"
d5$indicador <- "Número de Vínculos por Sexo, Segundo Brasil, Grandes regiões e Unidades da Federação"
d5$categoria1 <- "Total"
d5$categoria2 <- "Número de Vínculos"

d6 <- bases
d6$tematica <- "Mercado de Trabalho"
d6$indicador <- "Número de Vínculos por Sexo, Segundo Brasil, Grandes regiões e Unidades da Federação"
d6$categoria1 <- "Total"
d6$categoria2 <- "Ranking"


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6
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
  "mercado_de_trabalho/02 - Número de Vínculos por Sexo.xlsx",
  overwrite = TRUE
)
##03 - Número de Vínculos por Grande Setor----
d1 <- bases
d1$tematica <- "Mercado de Trabalho"
d1$indicador <- "Número de Vínculos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- "Indústria"
d1$categoria2 <- "Número de Vínculos"

d2 <- bases
d2$tematica <- "Mercado de Trabalho"
d2$indicador <- "Número de Vínculos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d2$categoria1 <- "Indústria"
d2$categoria2 <- "Ranking"

d3 <- bases
d3$tematica <- "Mercado de Trabalho"
d3$indicador <- "Número de Vínculos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d3$categoria1 <- "Construção Civil	"
d3$categoria2 <- "Número de Vínculos"

d4 <- bases
d4$tematica <- "Mercado de Trabalho"
d4$indicador <- "Número de Vínculos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d4$categoria1 <- "Construção Civil	"
d4$categoria2 <- "Ranking"

d5 <- bases
d5$tematica <- "Mercado de Trabalho"
d5$indicador <- "Número de Vínculos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d5$categoria1 <- "Comércio"
d5$categoria2 <- "Número de Vínculos"

d6 <- bases
d6$tematica <- "Mercado de Trabalho"
d6$indicador <- "Número de Vínculos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d6$categoria1 <- "Comércio"
d6$categoria2 <- "Ranking"

d7 <- bases
d7$tematica <- "Mercado de Trabalho"
d7$indicador <- "Número de Vínculos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d7$categoria1 <- "Serviços"
d7$categoria2 <- "Número de Vínculos"

d8 <- bases
d8$tematica <- "Mercado de Trabalho"
d8$indicador <- "Número de Vínculos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d8$categoria1 <- "Serviços"
d8$categoria2 <- "Ranking"

d9 <- bases
d9$tematica <- "Mercado de Trabalho"
d9$indicador <- "Número de Vínculos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d9$categoria1 <- "Agropecuária"
d9$categoria2 <- "Número de Vínculos"

d10 <- bases
d10$tematica <- "Mercado de Trabalho"
d10$indicador <- "Número de Vínculos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d10$categoria1 <- "Agropecuária"
d10$categoria2 <- "Ranking"

d11 <- bases
d10$tematica <- "Mercado de Trabalho"
d10$indicador <- "Número de Vínculos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d10$categoria1 <- "Total"
d10$categoria2 <- "Número de Vínculos"

d12 <- bases
d11$tematica <- "Mercado de Trabalho"
d11$indicador <- "Número de Vínculos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d11$categoria1 <- "Total"
d11$categoria2 <- "Ranking"

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
  "d8" = d8,
  "d9" = d9,
  "d10" = d10,
  "d11" = d11,
  "d12" = d12
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
  "mercado_de_trabalho/03 - Número de Vínculos por Grande Setor.xlsx",
  overwrite = TRUE
)
##04 - Remuneração Média do Trabalhador----
d1 <- bases
d1$tematica <- "Mercado de Trabalho"
d1$indicador <- "Remuneração Média do Trabalhador, Segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- "Remuneração Média"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Mercado de Trabalho"
d2$indicador <- "Remuneração Média do Trabalhador, Segundo Brasil, Grandes regiões e Unidades da Federação"
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
  "mercado_de_trabalho/04 - Remuneração Média do Trabalhador.xlsx",
  overwrite = TRUE
)
##05 - Remuneração Média do Trabalhador por Sexo----
d1 <- bases
d1$tematica <- "Mercado de Trabalho"
d1$indicador <- "Remuneração Média do Trabalhador por Sexo, Segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- "Masculino"
d1$categoria2 <- "Remuneração Média"

d2 <- bases
d2$tematica <- "Mercado de Trabalho"
d2$indicador <- "Remuneração Média do Trabalhador por Sexo, Segundo Brasil, Grandes regiões e Unidades da Federação"
d2$categoria1 <- "Masculino"
d2$categoria2 <- "Ranking"

d3 <- bases
d3$tematica <- "Mercado de Trabalho"
d3$indicador <- "Remuneração Média do Trabalhador por Sexo, Segundo Brasil, Grandes regiões e Unidades da Federação"
d3$categoria1 <- "Feminino"
d3$categoria2 <- "Remuneração Média"

d4 <- bases
d4$tematica <- "Mercado de Trabalho"
d4$indicador <- "Remuneração Média do Trabalhador por Sexo, Segundo Brasil, Grandes regiões e Unidades da Federação"
d4$categoria1 <- "Feminino"
d4$categoria2 <- "Ranking"

d5 <- bases
d5$tematica <- "Mercado de Trabalho"
d5$indicador <- "Remuneração Média do Trabalhador por Sexo, Segundo Brasil, Grandes regiões e Unidades da Federação"
d5$categoria1 <- "Total"
d5$categoria2 <- "Remuneração Média"

d6 <- bases
d6$tematica <- "Mercado de Trabalho"
d6$indicador <- "Remuneração Média do Trabalhador por Sexo, Segundo Brasil, Grandes regiões e Unidades da Federação"
d6$categoria1 <- "Total"
d6$categoria2 <- "Ranking"

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d2" = d2,
  "d3" = d3,
  "d4" = d4,
  "d5" = d5,
  "d6" = d6
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
  "mercado_de_trabalho/05 - Remuneração Média do Trabalhador por Sexo.xlsx",
  overwrite = TRUE
)
##06 - Número de Estabelecimentos----
d1 <- bases
d1$tematica <- "Mercado de Trabalho"
d1$indicador <- "Número de Estabelecimentos Total, Segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- "Número de Estabelecimentos Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Mercado de Trabalho"
d2$indicador <- "Número de Estabelecimentos Total, Segundo Brasil, Grandes regiões e Unidades da Federação"
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
  "mercado_de_trabalho/06 - Número de Estabelecimentos.xlsx",
  overwrite = TRUE
)
##07 - Número de Estabelecimentos por Grande Setor----
d1 <- bases
d1$tematica <- "Mercado de Trabalho"
d1$indicador <- "Número de Estabelecimentos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- "Indústria"
d1$categoria2 <- "Número de Vínculos por Grande Setor"

d2 <- bases
d2$tematica <- "Mercado de Trabalho"
d2$indicador <- "Número de Estabelecimentos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d2$categoria1 <- "Indústria"
d2$categoria2 <- "Ranking"

d3 <- bases
d3$tematica <- "Mercado de Trabalho"
d3$indicador <- "Número de Estabelecimentos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d3$categoria1 <- "Construção Civil"
d3$categoria2 <- "Número de Vínculos por Grande Setor"

d4 <- bases
d4$tematica <- "Mercado de Trabalho"
d4$indicador <- "Número de Estabelecimentos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d4$categoria1 <- "Construção Civil"
d4$categoria2 <- "Ranking"

d5 <- bases
d5$tematica <- "Mercado de Trabalho"
d5$indicador <- "Número de Estabelecimentos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d5$categoria1 <- "Comércio"
d5$categoria2 <- "Número de Vínculos por Grande Setor"

d6 <- bases
d6$tematica <- "Mercado de Trabalho"
d6$indicador <- "Número de Estabelecimentos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d6$categoria1 <- "Comércio"
d6$categoria2 <- "Ranking"

d7 <- bases
d7$tematica <- "Mercado de Trabalho"
d7$indicador <- "Número de Estabelecimentos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d7$categoria1 <- "Serviços"
d7$categoria2 <- "Número de Vínculos por Grande Setor"

d8 <- bases
d8$tematica <- "Mercado de Trabalho"
d8$indicador <- "Número de Estabelecimentos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d8$categoria1 <- "Serviços"
d8$categoria2 <- "Ranking"

d9 <- bases
d9$tematica <- "Mercado de Trabalho"
d9$indicador <- "Número de Estabelecimentos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d9$categoria1 <- "Agropecuária"
d9$categoria2 <- "Número de Vínculos por Grande Setor"

d10 <- bases
d10$tematica <- "Mercado de Trabalho"
d10$indicador <- "Número de Estabelecimentos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d10$categoria1 <- "Agropecuária"
d10$categoria2 <- "Ranking"

d11 <- bases
d11$tematica <- "Mercado de Trabalho"
d11$indicador <- "Número de Estabelecimentos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d11$categoria1 <- "Total"
d11$categoria2 <- "Número de Vínculos por Grande Setor"

d12 <- bases
d12$tematica <- "Mercado de Trabalho"
d12$indicador <- "Número de Estabelecimentos por Grande Setor, Segundo Brasil, Grandes regiões e Unidades da Federação"
d12$categoria1 <- "Total"
d12$categoria2 <- "Ranking"
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
  "d8" = d8,
  "d9" = d9,
  "d10" = d10,
  "d11" = d11,
  "d12" = d12
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
  "mercado_de_trabalho/07 - Número de Estabelecimentos por Grande Setor.xlsx",
  overwrite = TRUE
)
##08 - Condição em relação à força de trabalho----
d1 <- bases
d1$tematica <- "Mercado de Trabalho"
d1$indicador <- "Condição em relação à força de trabalho, segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- "Pessoas de 14 anos ou mais de idade (Mil pessoas), por condição em relação à força de trabalho"
d1$categoria2 <- "Força de trabalho"
d1$categoria3 <- "Quantidade"

d2 <- bases
d2$tematica <- "Mercado de Trabalho"
d2$indicador <- "Condição em relação à força de trabalho, segundo Brasil, Grandes regiões e Unidades da Federação"
d2$categoria1 <- "Pessoas de 14 anos ou mais de idade (Mil pessoas), por condição em relação à força de trabalho"
d2$categoria2 <- "Força de trabalho"
d2$categoria3 <- "Ranking"

d3$tematica <- "Mercado de Trabalho"
d3$indicador <- "Condição em relação à força de trabalho, segundo Brasil, Grandes regiões e Unidades da Federação"
d3$categoria1 <- "Pessoas de 14 anos ou mais de idade (Mil pessoas), por condição em relação à força de trabalho"
d3$categoria2 <- "Fora da força de trabalho"
d3$categoria3 <- "Quantidade"

d4 <- bases
d4$tematica <- "Mercado de Trabalho"
d4$indicador <- "Condição em relação à força de trabalho, segundo Brasil, Grandes regiões e Unidades da Federação"
d4$categoria1 <- "Pessoas de 14 anos ou mais de idade (Mil pessoas), por condição em relação à força de trabalho"
d4$categoria2 <- "Fora da força de trabalho"
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
  "mercado_de_trabalho/08 - Condição em relação à força de trabalho.xlsx",
  overwrite = TRUE
)
##09 - População Ocupada e Desocupada----
d <- bases
d1$tematica <- "Mercado de Trabalho"
d1$indicador <- "População Ocupada e Desocupada na Força de Trabalho, Segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- "Pessoas de 14 anos ou mais de idade (Mil pessoas), Ocupada e Desocupada na Força de Trabalho"
d1$categoria2 <- "Ocupado"
d1$categoria3 <- "Quantidade"

d2 <- bases
d2$tematica <- "Mercado de Trabalho"
d2$indicador <- "População Ocupada e Desocupada na Força de Trabalho, Segundo Brasil, Grandes regiões e Unidades da Federação"
d2$categoria1 <- "Pessoas de 14 anos ou mais de idade (Mil pessoas), Ocupada e Desocupada na Força de Trabalho"
d2$categoria2 <- "Ocupado"
d2$categoria3 <- "Ranking"

d3 <- bases
d3$tematica <- "Mercado de Trabalho"
d3$indicador <- "População Ocupada e Desocupada na Força de Trabalho, Segundo Brasil, Grandes regiões e Unidades da Federação"
d3$categoria1 <- "Pessoas de 14 anos ou mais de idade (Mil pessoas), Ocupada e Desocupada na Força de Trabalho"
d3$categoria2 <- "Desocupado"
d3$categoria3 <- "Quantidade"

d4 <- bases
d4$tematica <- "Mercado de Trabalho"
d4$indicador <- "População Ocupada e Desocupada na Força de Trabalho, Segundo Brasil, Grandes regiões e Unidades da Federação"
d4$categoria1 <- "Pessoas de 14 anos ou mais de idade (Mil pessoas), Ocupada e Desocupada na Força de Trabalho"
d4$categoria2 <- "Desocupado"
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
  "mercado_de_trabalho/09 - População Ocupada e Desocupada.xlsx",
  overwrite = TRUE
)
##10 - Taxa de Ocupação----
d1 <- bases
d1$tematica <- "Mercado de Trabalho"
d1$indicador <- "Taxa de Ocupação, Segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- "Taxa de Ocupação"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Mercado de Trabalho"
d2$indicador <- "Taxa de Ocupação, Segundo Brasil, Grandes regiões e Unidades da Federação"
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
  "mercado_de_trabalho/10 - Taxa de Ocupação.xlsx",
  overwrite = TRUE
)
##11 - Taxa de Desocupação----
d1 <- bases
d1$tematica <- "Mercado de Trabalho"
d1$indicador <- "Taxa de Desocupação, Segundo Brasil, Grandes regiões e Unidades da Federação"
d1$categoria1 <- "Taxa de Desocupação"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Mercado de Trabalho"
d2$indicador <- "Taxa de Desocupação, Segundo Brasil, Grandes regiões e Unidades da Federação"
d2$categoria1 <- "Taxa de Desocupação"
d2$categoria2 <- "Ranking"


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
  "mercado_de_trabalho/11 - Taxa de Desocupação.xlsx",
  overwrite = TRUE
)