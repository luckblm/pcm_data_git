#EDUCAÇÃO----
##EDUCAÇÃO BÁSICA----

bases$subtematica <- "Educação Básica"

###01 - Número de Matrículas na Creche - Ensino Regular por Dependência Administrativa----


d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Matrículas na Creche - Ensino Regular por Dependência Administrativa"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Matrículas na Creche - Ensino Regular por Dependência Administrativa"
d2$categoria1 <- "Federal"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Matrículas na Creche - Ensino Regular por Dependência Administrativa"
d3$categoria1 <- "Estadual"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Matrículas na Creche - Ensino Regular por Dependência Administrativa"
d4$categoria1 <- "Municipal"
d4$categoria2 <- "-"

d5 <- bases
d5$tematica <- "Educação"
d5$indicador <- "Número de Matrículas na Creche - Ensino Regular por Dependência Administrativa"
d5$categoria1 <- "Privada"
d5$categoria2 <- "-"

d6 <- bases
d6$tematica <- "Educação"
d6$indicador <- "Número de Matrículas na Creche - Ensino Regular por Dependência Administrativa"
d6$categoria1 <- "Classificação"
d6$categoria2 <- "-"


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
  "educacao/educacao_basica/01 - Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa.xlsx",
  overwrite = TRUE
)

###02 - Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa----

d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa"
d2$categoria1 <- "Federal"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa"
d3$categoria1 <- "Estadual"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa"
d4$categoria1 <- "Municipal"
d4$categoria2 <- "-"

d5 <- bases
d5$tematica <- "Educação"
d5$indicador <- "Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa"
d5$categoria1 <- "Privada"
d5$categoria2 <- "-"

d6 <- bases
d6$tematica <- "Educação"
d6$indicador <- "Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa"
d6$categoria1 <- "Classificação"
d6$categoria2 <- "-"


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
  "educacao/educacao_basica/02 - Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa.xlsx",
  overwrite = TRUE
)

###03 - Número de Matrículas no Ensino Fundamental - Ensino Regular por Dependência Administrativa----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Matrículas no Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Matrículas no Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d2$categoria1 <- "Federal"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Matrículas no Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d3$categoria1 <- "Estadual"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Matrículas no Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d4$categoria1 <- "Municipal"
d4$categoria2 <- "-"

d5 <- bases
d5$tematica <- "Educação"
d5$indicador <- "Número de Matrículas no Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d5$categoria1 <- "Privada"
d5$categoria2 <- "-"

d6 <- bases
d6$tematica <- "Educação"
d6$indicador <- "Número de Matrículas no Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d6$categoria1 <- "Classificação"
d6$categoria2 <- "-"


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
  "educacao/educacao_basica/03 - Número de Matrículas no Ensino Fundamental - Ensino Regular por Dependência Administrativa.xlsx",
  overwrite = TRUE
)

###04 - Número de Matrículas no Ensino Médio - Ensino Regular por Dependência Administrativa----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Matrículas no Ensino Médio - Ensino Regular por Dependência Administrativa"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Matrículas no Ensino Médio - Ensino Regular por Dependência Administrativa"
d2$categoria1 <- "Federal"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Matrículas no Ensino Médio - Ensino Regular por Dependência Administrativa"
d3$categoria1 <- "Estadual"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Matrículas no Ensino Médio - Ensino Regular por Dependência Administrativa"
d4$categoria1 <- "Municipal"
d4$categoria2 <- "-"

d5 <- bases
d5$tematica <- "Educação"
d5$indicador <- "Número de Matrículas no Ensino Médio - Ensino Regular por Dependência Administrativa"
d5$categoria1 <- "Privada"
d5$categoria2 <- "-"

d6 <- bases
d6$tematica <- "Educação"
d6$indicador <- "Número de Matrículas no Ensino Médio - Ensino Regular por Dependência Administrativa"
d6$categoria1 <- "Classificação"
d6$categoria2 <- "-"


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
  "educacao/educacao_basica/04 - Número de Matrículas no Ensino Médio - Ensino Regular por Dependência Administrativa.xlsx",
  overwrite = TRUE
)
###05 - Número de Matrículas na Educação Profissional por Dependência Administrativa----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Matrículas na Educação Profissional por Dependência Administrativa"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Matrículas na Educação Profissional por Dependência Administrativa"
d2$categoria1 <- "Federal"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Matrículas na Educação Profissional por Dependência Administrativa"
d3$categoria1 <- "Estadual"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Matrículas na Educação Profissional por Dependência Administrativa"
d4$categoria1 <- "Municipal"
d4$categoria2 <- "-"

d5 <- bases
d5$tematica <- "Educação"
d5$indicador <- "Número de Matrículas na Educação Profissional por Dependência Administrativa"
d5$categoria1 <- "Privada"
d5$categoria2 <- "-"

d6 <- bases
d6$tematica <- "Educação"
d6$indicador <- "Número de Matrículas na Educação Profissional por Dependência Administrativa"
d6$categoria1 <- "Ranking"
d6$categoria2 <- "-"


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
  "educacao/educacao_basica/05 - Número de Matrículas na Educação Profissional por Dependência Administrativa.xlsx",
  overwrite = TRUE
)
###06 - Número de Estabelecimentos de Creche - Ensino Regular por Dependência Administrativa----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Estabelecimentos de Creche - Ensino Regular por Dependência Administrativa"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Estabelecimentos de Creche - Ensino Regular por Dependência Administrativa"
d2$categoria1 <- "Federal"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Estabelecimentos de Creche - Ensino Regular por Dependência Administrativa"
d3$categoria1 <- "Estadual"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Estabelecimentos de Creche - Ensino Regular por Dependência Administrativa"
d4$categoria1 <- "Municipal"
d4$categoria2 <- "-"

d5 <- bases
d5$tematica <- "Educação"
d5$indicador <- "Número de Estabelecimentos de Creche - Ensino Regular por Dependência Administrativa"
d5$categoria1 <- "Privada"
d5$categoria2 <- "-"

d6 <- bases
d6$tematica <- "Educação"
d6$indicador <- "Número de Estabelecimentos de Creche - Ensino Regular por Dependência Administrativa"
d6$categoria1 <- "Ranking"
d6$categoria2 <- "-"


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
  "educacao/educacao_basica/06 - Número de Estabelecimentos de Creche - Ensino Regular por Dependência Administrativa.xlsx",
  overwrite = TRUE
)
###07 - Número de Estabelecimentos de Pré-Escola - Ensino Regular por Dependência Administrativa----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Estabelecimentos de Pré-Escola - Ensino Regular por Dependência Administrativa"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Estabelecimentos de Pré-Escola - Ensino Regular por Dependência Administrativa"
d2$categoria1 <- "Federal"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Estabelecimentos de Pré-Escola - Ensino Regular por Dependência Administrativa"
d3$categoria1 <- "Estadual"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Estabelecimentos de Pré-Escola - Ensino Regular por Dependência Administrativa"
d4$categoria1 <- "Municipal"
d4$categoria2 <- "-"

d5 <- bases
d5$tematica <- "Educação"
d5$indicador <- "Número de Estabelecimentos de Pré-Escola - Ensino Regular por Dependência Administrativa"
d5$categoria1 <- "Privada"
d5$categoria2 <- "-"

d6 <- bases
d6$tematica <- "Educação"
d6$indicador <- "Número de Estabelecimentos de Pré-Escola - Ensino Regular por Dependência Administrativa"
d6$categoria1 <- "Ranking"
d6$categoria2 <- "-"


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
  "educacao/educacao_basica/ - .xlsx",
  overwrite = TRUE
)
###08 - Número de Estabelecimentos de Ensino Fundamental - Ensino Regular por Dependência Administrativa----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Estabelecimentos de Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Estabelecimentos de Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d2$categoria1 <- "Federal"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Estabelecimentos de Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d3$categoria1 <- "Estadual"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Estabelecimentos de Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d4$categoria1 <- "Municipal"
d4$categoria2 <- "-"

d5 <- bases
d5$tematica <- "Educação"
d5$indicador <- "Número de Estabelecimentos de Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d5$categoria1 <- "Privada"
d5$categoria2 <- "-"

d6 <- bases
d6$tematica <- "Educação"
d6$indicador <- "Número de Estabelecimentos de Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d6$categoria1 <- "Ranking"
d6$categoria2 <- "-"


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
  "educacao/educacao_basica/08 - Número de Estabelecimentos de Ensino Fundamental - Ensino Regular por Dependência Administrativa.xlsx",
  overwrite = TRUE
)
###09 - Número de Estabelecimentos de Ensino Médio - Ensino Regular por Dependência Administrativa----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Estabelecimentos de Ensino Médio - Ensino Regular por Dependência Administrativa"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Estabelecimentos de Ensino Médio - Ensino Regular por Dependência Administrativa"
d2$categoria1 <- "Federal"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Estabelecimentos de Ensino Médio - Ensino Regular por Dependência Administrativa"
d3$categoria1 <- "Estadual"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Estabelecimentos de Ensino Médio - Ensino Regular por Dependência Administrativa"
d4$categoria1 <- "Municipal"
d4$categoria2 <- "-"

d5 <- bases
d5$tematica <- "Educação"
d5$indicador <- "Número de Estabelecimentos de Ensino Médio - Ensino Regular por Dependência Administrativa"
d5$categoria1 <- "Privada"
d5$categoria2 <- "-"

d6 <- bases
d6$tematica <- "Educação"
d6$indicador <- "Número de Estabelecimentos de Ensino Médio - Ensino Regular por Dependência Administrativa"
d6$categoria1 <- "Ranking"
d6$categoria2 <- "-"


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
  "educacao/educacao_basica/09 - Número de Estabelecimentos de Ensino Médio - Ensino Regular por Dependência Administrativa.xlsx",
  overwrite = TRUE
)
###10 - Número de Estabelecimentos de Educação Profissional - Ensino Regular por Dependência Administrativa----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Estabelecimentos de Educação Profissional - Ensino Regular por Dependência Administrativa"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Estabelecimentos de Educação Profissional - Ensino Regular por Dependência Administrativa"
d2$categoria1 <- "Federal"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Estabelecimentos de Educação Profissional - Ensino Regular por Dependência Administrativa"
d3$categoria1 <- "Estadual"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Estabelecimentos de Educação Profissional - Ensino Regular por Dependência Administrativa"
d4$categoria1 <- "Municipal"
d4$categoria2 <- "-"

d5 <- bases
d5$tematica <- "Educação"
d5$indicador <- "Número de Estabelecimentos de Educação Profissional - Ensino Regular por Dependência Administrativa"
d5$categoria1 <- "Privada"
d5$categoria2 <- "-"

d6 <- bases
d6$tematica <- "Educação"
d6$indicador <- "Número de Estabelecimentos de Educação Profissional - Ensino Regular por Dependência Administrativa"
d6$categoria1 <- "Ranking"
d6$categoria2 <- "-"


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
  "educacao/educacao_basica/10 - Número de Estabelecimentos de Educação Profissional - Ensino Regular por Dependência Administrativa.xlsx",
  overwrite = TRUE
)
###11 - Taxa de Aprovação no Ensino Fundamental----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Taxa de Aprovação no Ensino Fundamental"
d1$categoria1 <- "Taxa de Aprovação"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Taxa de Aprovação no Ensino Fundamental"
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
  "educacao/educacao_basica/11 - Taxa de Aprovação no Ensino Fundamental.xlsx",
  overwrite = TRUE
)
###12 - Taxa de Reprovação no Ensino Fundamental----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Taxa de Reprovação no Ensino Fundamental"
d1$categoria1 <- "Taxa de Reprovação"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Taxa de Reprovação no Ensino Fundamental"
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
  "educacao/educacao_basica/12 - Taxa de Reprovação no Ensino Fundamental.xlsx",
  overwrite = TRUE
)
###13 - Taxa de Abandono no Ensino Fundamental----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Taxa de Abandono no Ensino Fundamental"
d1$categoria1 <- "Taxa de Abandono"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Taxa de Abandono no Ensino Fundamental"
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
  "educacao/educacao_basica/13 - Taxa de Abandono no Ensino Fundamental.xlsx",
  overwrite = TRUE
)
###14 - Taxa de Aprovação no Ensino Médio----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Taxa de Aprovação no Ensino Médio"
d1$categoria1 <- "Taxa de Aprovação"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Taxa de Aprovação no Ensino Médio"
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
  "educacao/educacao_basica/14 - Taxa de Aprovação no Ensino Médio.xlsx",
  overwrite = TRUE
)
###15 - Taxa de Reprovação no Ensino Médio----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Taxa de Reprovação no Ensino Médio"
d1$categoria1 <- "Taxa de Reprovação"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Taxa de Reprovação no Ensino Médio"
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
  "educacao/educacao_basica/15 - Taxa de Reprovação no Ensino Médio.xlsx",
  overwrite = TRUE
)
###16 - Taxa de Abandono no Ensino Médio----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Taxa de Abandono no Ensino Médio "
d1$categoria1 <- "Taxa de Abandono"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Taxa de Abandono no Ensino Médio"
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
  "educacao/educacao_basica/16 - Taxa de Abandono no Ensino Médio.xlsx",
  overwrite = TRUE
)
###17 - Nota IDEB na Rede de Ensino Pública - Séries Iniciais----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Nota IDEB na Rede de Ensino Pública - Séries Iniciais"
d1$categoria1 <- "Nota"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Nota IDEB na Rede de Ensino Pública - Séries Iniciais"
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
  "educacao/educacao_basica/17 - Nota IDEB na Rede de Ensino Pública - Séries Iniciais.xlsx",
  overwrite = TRUE
)
###18 - Nota IDEB na Rede de Ensino Pública - Séries Finais----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Nota IDEB na Rede de Ensino Pública - Séries Finais"
d1$categoria1 <- "Nota"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Nota IDEB na Rede de Ensino Pública - Séries Finais"
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
  "educacao/educacao_basica/18 - Nota IDEB na Rede de Ensino Pública - Séries Finais.xlsx",
  overwrite = TRUE
)
###19 - Nota IDEB na Rede de Ensino Estadual - Séries Iniciais----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Nota IDEB na Rede de Ensino Estadual - Séries Iniciais"
d1$categoria1 <- "Nota"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Nota IDEB na Rede de Ensino Estadual - Séries Iniciais"
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
  "educacao/educacao_basica/19 - Nota IDEB na Rede de Ensino Estadual - Séries Iniciais.xlsx",
  overwrite = TRUE
)
###20 - Nota IDEB na Rede de Ensino Estadual Séries Finais----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Nota IDEB na Rede de Ensino Estadual Séries Finais"
d1$categoria1 <- "Nota"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Nota IDEB na Rede de Ensino Estadual Séries Finais"
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
  "educacao/educacao_basica/20 - Nota IDEB na Rede de Ensino Estadual Séries Finais.xlsx",
  overwrite = TRUE
)

##EDUCAÇÃO SUPERIOR----

bases$subtematica <- "Educação Superior"

###01 - Número de Instituições de Ensino Superior----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Instituições de Ensino Superior"
d1$categoria1 <- "Pública"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Instituições de Ensino Superior"
d2$categoria1 <- "Privada"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Instituições de Ensino Superior"
d3$categoria1 <- "Total"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Instituições de Ensino Superior"
d4$categoria1 <- "Ranking"
d4$categoria2 <- "-"


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
  "educacao/educacao_superior/01 - Número de Instituições de Ensino Superior.xlsx",
  overwrite = TRUE
)
###02 - Número de Docentes em Instituições de Ensino Superior----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Docentes em Instituições de Ensino Superior"
d1$categoria1 <- "Pública"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Docentes em Instituições de Ensino Superior"
d2$categoria1 <- "Privada"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Docentes em Instituições de Ensino Superior"
d3$categoria1 <- "Total"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Docentes em Instituições de Ensino Superior"
d4$categoria1 <- "Ranking"
d4$categoria2 <- "-"


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
  "educacao/educacao_superior/02 - Número de Docentes em Instituições de Ensino Superior.xlsx",
  overwrite = TRUE
)
###03 - Número de Cursos de Graduação Presenciais----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Cursos de Graduação Presenciais"
d1$categoria1 <- "Pública"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Cursos de Graduação Presenciais"
d2$categoria1 <- "Privada"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Cursos de Graduação Presenciais"
d3$categoria1 <- "Total"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Cursos de Graduação Presenciais"
d4$categoria1 <- "Ranking"
d4$categoria2 <- "-"


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
  "educacao/educacao_superior/03 - Número de Cursos de Graduação Presenciais.xlsx",
  overwrite = TRUE
)
###04 - Número de Matrículas em Cursos de Graduação Presenciais em Instituições de Ensino Superior----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Matrículas em Cursos de Graduação Presenciais em Instituições de Ensino Superior"
d1$categoria1 <- "Pública"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Matrículas em Cursos de Graduação Presenciais em Instituições de Ensino Superior"
d2$categoria1 <- "Privada"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Matrículas em Cursos de Graduação Presenciais em Instituições de Ensino Superior"
d3$categoria1 <- "Total"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Matrículas em Cursos de Graduação Presenciais em Instituições de Ensino Superior"
d4$categoria1 <- "Ranking"
d4$categoria2 <- "-"


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
  "educacao/educacao_superior/04 - Número de Matrículas em Cursos de Graduação Presenciais em Instituições de Ensino Superior.xlsx",
  overwrite = TRUE
)
###05 - Número de Concluintes em Cursos de Graduação Presenciais em Instituições de Ensino Superior----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Concluintes em Cursos de Graduação Presenciais em Instituições de Ensino Superior"
d1$categoria1 <- "Pública"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Concluintes em Cursos de Graduação Presenciais em Instituições de Ensino Superior"
d2$categoria1 <- "Privada"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Concluintes em Cursos de Graduação Presenciais em Instituições de Ensino Superior"
d3$categoria1 <- "Total"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Concluintes em Cursos de Graduação Presenciais em Instituições de Ensino Superior"
d4$categoria1 <- "Ranking"
d4$categoria2 <- "-"


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
  "educacao/educacao_superior/05 - Número de Concluintes em Cursos de Graduação Presenciais em Instituições de Ensino Superior.xlsx",
  overwrite = TRUE
)
###06 - Número de Matrículas em Cursos de Graduação EAD----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Matrículas em Cursos de Graduação EAD"
d1$categoria1 <- "Pública"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Matrículas em Cursos de Graduação EAD"
d2$categoria1 <- "Privada"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Matrículas em Cursos de Graduação EAD"
d3$categoria1 <- "Total"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Matrículas em Cursos de Graduação EAD"
d4$categoria1 <- "Ranking"
d4$categoria2 <- "-"


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
  "educacao/educacao_superior/06 - Número de Matrículas em Cursos de Graduação EAD.xlsx",
  overwrite = TRUE
)
###07 - Número de Concluintes em Cursos de Graduação à Distancia (Públicas e Privadas)----
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Concluintes em Cursos de Graduação à Distancia (Públicas e Privadas)"
d1$categoria1 <- "Pública"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Concluintes em Cursos de Graduação à Distancia (Públicas e Privadas)"
d2$categoria1 <- "Privada"
d2$categoria2 <- "-"

d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Concluintes em Cursos de Graduação à Distancia (Públicas e Privadas)"
d3$categoria1 <- "Total"
d3$categoria2 <- "-"

d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Concluintes em Cursos de Graduação à Distancia (Públicas e Privadas)"
d4$categoria1 <- "Ranking"
d4$categoria2 <- "-"


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
  "educacao/educacao_superior/07 - Número de Concluintes em Cursos de Graduação à Distancia_Públicas e Privadas - .xlsx",
  overwrite = TRUE
)
