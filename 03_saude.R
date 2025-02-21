#SAÚDE----
bases$subtematica <- "-"
##01 - Taxa de Mortalidade Infantil----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Taxa de Mortalidade Infantil"
d1$categoria1 <- "Taxa"
d1$categoria2 <- "-"

# Criando os vetores dos anos
d1$`2019` <- c(12.39, 15.12, 11.51, 15.91, 15.99, 18.06, 15.14, 18.82, 11.70, 
            13.70, 14.36, 14.65, 12.23, 12.42, 13.03, 12.25, 13.23, 17.28, 
            15.06, 11.52, 11.45, 10.65, 13.16, 11.05, 10.24, 10.31, 9.61, 
            10.62, 11.84, 11.10, 12.69, 13.11, 8.53)

d1$`2020` <- c(11.52, 14.54, 13.02, 16.05, 13.90, 19.19, 14.86, 18.04, 10.62, 
            12.97, 13.74, 13.88, 11.62, 11.30, 12.68, 11.62, 11.98, 15.89, 
            14.34, 10.52, 10.45, 9.76, 12.60, 9.88, 9.08, 9.30, 9.32, 8.64, 
            11.19, 10.92, 12.08, 11.36, 9.76)

d1$`2021` <- c(11.90, 15.01, 12.34, 17.84, 14.80, 19.35, 14.73, 19.94, 12.59, 
            13.12, 13.53, 13.75, 10.70, 12.23, 12.63, 12.42, 13.36, 14.04, 
            14.93, 10.94, 10.69, 11.24, 12.74, 10.38, 9.45, 9.46, 9.23, 9.59, 
            11.73, 10.69, 12.67, 12.10, 10.57)

d1$`2022` <- c(12.59, 15.11, 13.41, 17.19, 15.69, 18.79, 14.66, 18.07, 12.33, 
            14.09, 15.32, 15.76, 11.73, 11.06, 14.72, 13.27, 12.83, 17.63, 
            15.35, 11.64, 11.37, 10.79, 13.16, 11.31, 10.23, 10.32, 9.79, 
            10.49, 12.56, 12.35, 14.08, 12.66, 10.08)

d1$`2023` <- c(12.62, 15.90, 12.30, 17.03, 17.08, 23.85, 15.04, 21.01, 12.70, 
            13.80, 14.85, 14.99, 11.72, 11.17, 12.98, 13.24, 13.52, 18.51, 
            14.79, 11.74, 11.29, 11.59, 13.45, 11.37, 9.95, 10.79, 9.07, 
            9.67, 12.80, 13.55, 14.04, 12.44, 10.83)


d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Taxa de Mortalidade Infantil"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"
d2$`2023` <- c(NA, 1, 18, 5, 4, 1, 6, 2, 16, 
               2, 8, 7, 19, 23, 15, 14, 12, 3, 
               9, 4, 22, 20, 13, 21, 5, 25, 27, 26, 
               3, 11, 10, 17, 24)



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

# Criando os vetores dos anos com d1$
d1$`2019` <- c(14.43, 18.11, 13.58, 19.41, 19.17, 22.44, 18.01, 22.60, 14.11, 
               15.81, 16.94, 16.75, 14.21, 14.33, 15.13, 14.10, 15.68, 19.21, 
               17.15, 13.40, 13.40, 12.62, 15.29, 12.80, 11.81, 12.15, 10.89, 
               12.08, 13.94, 13.11, 15.12, 15.20, 10.28)

d1$`2020` <- c(13.20, 17.19, 14.81, 18.76, 16.63, 22.46, 17.57, 20.77, 13.15, 
               14.80, 16.19, 16.03, 13.33, 12.75, 14.10, 13.43, 13.65, 17.87, 
               16.05, 11.95, 11.99, 11.40, 14.34, 11.12, 10.37, 10.74, 10.51, 
               9.84, 12.91, 12.83, 14.39, 12.96, 10.72)

d1$`2021` <- c(13.77, 17.80, 14.51, 20.96, 17.97, 25.68, 17.02, 23.08, 15.29, 
               15.01, 15.97, 15.62, 12.46, 13.45, 14.49, 14.18, 15.45, 16.31, 
               16.71, 12.57, 12.18, 13.26, 14.59, 11.95, 11.02, 10.95, 10.57, 
               11.46, 13.87, 12.64, 15.42, 14.25, 11.99)

d1$`2022` <- c(15.04, 18.52, 16.18, 20.64, 19.46, 24.60, 17.76, 21.81, 15.48, 
               16.69, 18.16, 18.58, 13.86, 13.96, 17.57, 15.78, 15.70, 20.37, 
               17.89, 13.83, 13.57, 13.51, 15.71, 13.33, 12.12, 12.45, 11.54, 
               12.20, 15.33, 15.12, 17.48, 15.30, 12.19)

d1$`2023` <- c(14.96, 19.51, 14.89, 20.77, 21.18, 33.25, 18.15, 26.11, 14.43, 
               16.21, 17.75, 17.10, 13.54, 13.15, 15.33, 15.84, 15.82, 21.44, 
               17.30, 13.85, 13.46, 13.87, 15.98, 13.28, 11.66, 12.53, 10.67, 
               11.45, 15.30, 16.26, 17.32, 14.61, 12.69)


d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Taxa de Mortalidade na Infância"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"

d2$`2023` <- c(NA, 1, 16, 5, 4, 1, 6, 2, 18, 
                       2, 7, 10, 20, 23, 15, 13, 14, 3, 
                       9, 4, 21, 19, 12, 22, 5, 25, 27, 26, 
                       3, 11, 8, 17, 24)

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

