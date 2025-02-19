#SAÚDE----
bases$subtematica <- "-"
##01 - Taxa de Mortalidade Infantil----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Taxa de Mortalidade Infantil"
d1$categoria1 <- "Taxa"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Taxa de Mortalidade Infantil"
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
  "saude/01 - Taxa de Mortalidade Infantil.xlsx",
  overwrite = TRUE
)
##02 - Taxa de Mortalidade na Infância----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Taxa de Mortalidade na Infância"
d1$categoria1 <- "Taxa"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Taxa de Mortalidade na Infância"
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
  "saude/02 - Taxa de Mortalidade na Infância.xlsx",
  overwrite = TRUE
)
##03 - Taxa de Mortalidade Materna----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Taxa de Mortalidade Materna"
d1$categoria1 <- "Taxa"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Taxa de Mortalidade Materna"
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
  "saude/03 - Taxa de Mortalidade Materna.xlsx",
  overwrite = TRUE
)
##04 - Taxa de Natalidade----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Taxa de Natalidade"
d1$categoria1 <- "Taxa"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Taxa de Natalidade"
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
  "saude/04 - Taxa de Natalidade.xlsx",
  overwrite = TRUE
)
##05 - Taxa de Mortalidade Geral----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Taxa de Mortalidade Geral"
d1$categoria1 <- "Taxa"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Taxa de Mortalidade Geral"
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
  "saude/05 - Taxa de Mortalidade Geral.xlsx",
  overwrite = TRUE
)
##06 - Proporção de Mortes por Sexo----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Proporção de Mortes por Sexo"
d1$categoria1 <- "Proporção"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Proporção de Mortes por Sexo"
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
  "saude/06 - Proporção de Mortes por Sexo.xlsx",
  overwrite = TRUE
)

##07 - Hospitais por 10 Mil Habitantes----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Hospitais por 10 Mil Habitantes"
d1$categoria1 <- "Quantidade"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Hospitais por 10 Mil Habitantes"
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
  "saude/07 - Hospitais por 10 Mil Habitantes.xlsx",
  overwrite = TRUE
)

##08 - Postos e Centros de Saúde por 10 Mil Habitantes----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Postos e Centros de Saúde por 10 Mil Habitantes"
d1$categoria1 <- "Quantidade"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Postos e Centros de Saúde por 10 Mil Habitantes"
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
  "saude/08 - Postos e Centros de Saúde por 10 Mil Habitantes.xlsx",
  overwrite = TRUE
)

##09 - Médicos por 10 Mil Habitantes----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Médicos por 10 Mil Habitantes"
d1$categoria1 <- "Quantidade"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Médicos por 10 Mil Habitantes"
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
  "saude/09 - Médicos por 10 Mil Habitantes.xlsx",
  overwrite = TRUE
)
##10 - Leitos Hospitalares por Mil Habitantes----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Leitos Hospitalares por Mil Habitantes"
d1$categoria1 <- "Quantidade"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Leitos Hospitalares por Mil Habitantes"
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
  "saude/10 - Leitos Hospitalares por Mil Habitantes.xlsx",
  overwrite = TRUE
)
##11 - Percentual de Nascidos Vivos com 7 ou mais Consultas Pré-Natal----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Percentual de Nascidos Vivos com 7 ou mais Consultas Pré-Natal"
d1$categoria1 <- "Percentual"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Percentual de Nascidos Vivos com 7 ou mais Consultas Pré-Natal"
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
  "saude/11 - Percentual de Nascidos Vivos com 7 ou mais Consultas Pré-Natal.xlsx",
  overwrite = TRUE
)
##12 - Percentual de Nascidos Vivos por Tipo de Parto----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Percentual de Nascidos Vivos por Tipo de Parto"
d1$categoria1 <- "Percentual"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Percentual de Nascidos Vivos por Tipo de Parto"
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
  "saude/12 - Percentual de Nascidos Vivos por Tipo de Parto.xlsx",
  overwrite = TRUE
)
##13 - Percentual de Nascidos Vivos de Mães Adolescentes na Faixa Etária de 10 a 19 anos----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Percentual de Nascidos Vivos de Mães Adolescentes na Faixa Etária de 10 a 19 anos"
d1$categoria1 <- "Percentual"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Percentual de Nascidos Vivos de Mães Adolescentes na Faixa Etária de 10 a 19 anos"
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
  "saude/13 - Percentual de Nascidos Vivos de Mães Adolescentes na Faixa Etária de 10 a 19 anos.xlsx",
  overwrite = TRUE
)
##14 - Óbitos por Neoplasias, segundo Brasil----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Óbitos por Neoplasias, segundo Brasil"
d1$categoria1 <- "Quantidade"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Óbitos por Neoplasias, segundo Brasil"
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
  "saude/14 - Óbitos por Neoplasias, segundo Brasil.xlsx",
  overwrite = TRUE
)
##15 - Óbitos por Doenças Infecciosas e Parasitárias----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Óbitos por Doenças Infecciosas e Parasitárias"
d1$categoria1 <- "Quantidade"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Óbitos por Doenças Infecciosas e Parasitárias"
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
  "saude/15 - Óbitos por Doenças Infecciosas e Parasitárias.xlsx",
  overwrite = TRUE
)
##16 - Óbitos por Doenças do Aparelho Circulatório----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Óbitos por Doenças do Aparelho Circulatório"
d1$categoria1 <- "Quantidade"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Óbitos por Doenças do Aparelho Circulatório"
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
  "saude/16 - Óbitos por Doenças do Aparelho Circulatório.xlsx",
  overwrite = TRUE
)
##17 - Óbitos por Doenças do Aparelho Digestivo----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Óbitos por Doenças do Aparelho Digestivo"
d1$categoria1 <- "Quantidade"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Óbitos por Doenças do Aparelho Digestivo"
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
  "saude/17 - Óbitos por Doenças do Aparelho Digestivo.xlsx",
  overwrite = TRUE
)
##18 - Óbitos por Doenças do Aparelho Respiratório----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Óbitos por Doenças do Aparelho Respiratório"
d1$categoria1 <- "Quantidade"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Óbitos por Doenças do Aparelho Respiratório"
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
  "saude/18 - Óbitos por Doenças do Aparelho Respiratório.xlsx",
  overwrite = TRUE
)
##19 - Óbitos por Algumas Afecções Originadas no Período Perinatal----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Óbitos por Algumas Afecções Originadas no Período Perinatal"
d1$categoria1 <- "Quantidade"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Óbitos por Algumas Afecções Originadas no Período Perinatal"
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
  "saude/19 - Óbitos por Algumas Afecções Originadas no Período Perinatal.xlsx",
  overwrite = TRUE
)
##20 - Taxa de Cobertura Populacional dos Agentes Comunitários de Saúde(SÉRIE ENCERRADA)----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Taxa de Cobertura Populacional dos Agentes Comunitários de Saúde(SÉRIE ENCERRADA)"
d1$categoria1 <- "Taxa"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Taxa de Cobertura Populacional dos Agentes Comunitários de Saúde(SÉRIE ENCERRADA)"
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
  "saude/20 - Taxa de Cobertura Populacional dos Agentes Comunitários de Saúde(SÉRIE ENCERRADA).xlsx",
  overwrite = TRUE
)
##21 - Taxa de Cobertura Populacional das Equipes de Saúde da Família(SÉRIE ENCERRADA)----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Taxa de Cobertura Populacional das Equipes de Saúde da Família(SÉRIE ENCERRADA)"
d1$categoria1 <- "Taxa"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Taxa de Cobertura Populacional das Equipes de Saúde da Família(SÉRIE ENCERRADA)"
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
  "saude/21 - Taxa de Cobertura Populacional das Equipes de Saúde da Família(SÉRIE ENCERRADA).xlsx",
  overwrite = TRUE
)
##22 - Cobertura Populacional (%) da Atenção Básica----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Cobertura Populacional (%) da Atenção Básica"
d1$categoria1 <- "Percentual"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Cobertura Populacional (%) da Atenção Básica"
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
  "saude/22 - Cobertura Populacional (%) da Atenção Básica.xlsx",
  overwrite = TRUE
)

