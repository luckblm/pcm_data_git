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

# Criando o vetor de ranking para 2023
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

# Criando o vetor de ranking para 2023
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

# Criando os vetores dos anos com d1$
d1$`2019` <- c(60.58, 79.69, 51.80, 49.14, 86.32, 75.24, 93.97, 32.56, 61.35, 
               67.31, 75.01, 95.97, 83.60, 72.68, 77.99, 54.74, 58.23, 36.70, 
               56.78, 56.85, 45.54, 60.08, 87.99, 50.41, 41.18, 46.92, 31.62, 
               41.61, 61.39, 54.93, 74.76, 73.87, 21.22)

d1$`2020` <- c(78.64, 102.77, 73.65, 33.02, 103.13, 145.35, 116.60, 102.51, 75.86, 
               90.70, 95.21, 103.92, 108.28, 78.11, 99.33, 77.05, 82.75, 103.83, 
               83.09, 72.88, 54.61, 79.97, 118.52, 63.91, 50.14, 57.42, 37.79, 
               51.25, 79.40, 38.73, 96.43, 97.02, 55.89)

d1$`2021` <- c(127.30, 154.83, 180.82, 95.55, 161.88, 287.75, 135.64, 126.73, 193.71, 
               124.92, 152.76, 147.90, 121.40, 163.48, 140.95, 94.29, 118.83, 105.75, 
               117.01, 119.54, 102.42, 121.92, 194.35, 100.14, 119.59, 142.98, 106.74, 
               102.86, 144.54, 149.38, 150.41, 153.91, 107.80)

d1$`2022` <- c(64.01, 80.92, 44.17, 34.52, 75.84, 160.42, 85.16, 88.14, 93.13, 
               74.19, 86.77, 125.45, 70.34, 87.42, 47.16, 62.16, 76.52, 115.69, 
               62.71, 57.37, 53.18, 67.66, 89.26, 47.02, 49.20, 49.06, 41.75, 
               55.40, 62.85, 71.63, 58.45, 63.51, 58.45)

d1$`2023` <- c(60.32, 75.72, 46.00, 41.55, 85.24, 122.29, 72.99, 92.69, 77.78, 
               64.61, 92.64, 61.75, 51.32, 68.52, 67.91, 62.15, 70.92, 41.37, 
               59.97, 58.39, 50.47, 44.07, 95.41, 50.61, 44.19, 49.37, 37.20, 
               43.81, 61.47, 64.63, 75.14, 52.28, 59.07)

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Taxa de Mortalidade Materna"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"

# Criando o vetor de ranking para 2023
d2$`2023` <- c(NA, 1, 22, 25, 5, 1, 8, 3, 6, 
               2, 4, 14, 18, 10, 11, 13, 9, 26, 
               15, 4, 20, 23, 2, 19, 5, 21, 27, 24, 
               3, 12, 7, 17, 16)


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

# Criando os vetores para os anos dentro de d1$
d1$`2019` <- c(13.56, 17.02, 15.21, 18.46, 18.73, 24.13, 16.08, 18.16, 15.54,
               14.11, 16.02, 14.64, 14.15, 12.56, 14.36, 13.95, 14.92, 14.22, 13.26,
               12.48, 12.14, 13.67, 12.05, 12.70, 12.88, 13.42, 13.68, 11.83,
               14.79, 15.72, 16.89, 13.69, 14.07)

d1$`2020` <- c(12.89, 16.15, 14.36, 16.93, 17.98, 21.80, 15.30, 16.98, 14.92,
               13.43, 14.91, 13.78, 13.27, 12.32, 13.96, 13.36, 14.42, 13.71, 12.66,
               11.82, 11.61, 13.23, 11.47, 11.93, 12.42, 12.70, 13.50, 11.45,
               13.96, 14.70, 16.18, 13.04, 12.88)

d1$`2021` <- c(12.55, 16.36, 14.01, 17.31, 18.37, 21.30, 15.62, 17.08, 14.77,
               13.28, 15.19, 13.98, 13.01, 12.20, 13.81, 13.05, 14.50, 13.34, 12.38,
               11.27, 11.31, 12.78, 10.87, 11.26, 11.94, 12.24, 13.15, 10.85,
               13.71, 14.85, 16.21, 12.62, 12.29)

d1$`2022` <- c(12.62, 16.66, 15.75, 17.45, 18.40, 20.56, 15.76, 18.56, 14.92,
               12.97, 14.46, 12.91, 12.77, 12.12, 12.80, 12.96, 14.62, 12.91, 12.29,
               11.55, 11.44, 13.49, 11.23, 11.54, 12.02, 12.29, 12.90, 11.11,
               13.77, 14.68, 15.90, 12.72, 12.75)

d1$`2023` <- c(11.98, 15.32, 13.74, 16.47, 16.60, 18.82, 14.63, 16.20, 14.77,
               12.33, 13.87, 12.51, 12.08, 11.47, 12.50, 12.18, 14.46, 12.71, 11.47,
               10.93, 11.00, 12.80, 10.23, 10.99, 11.57, 11.89, 12.21, 10.78,
               13.38, 13.98, 15.50, 12.62, 11.98)


d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Taxa de Natalidade"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"

# Criando o vetor de ranking para 2023
d2$`2023` <- c(NA, 1, 11, 3, 2, 1, 7, 4, 6,
               3, 10, 15, 19, 23, 16, 18, 8, 13, 22,
               5, 24, 12, 27, 25, 4, 21, 17, 26,
               2, 9, 5, 14, 20)

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

# Criando os vetores para os anos dentro de d1$
d1$`2019` <- c(6.42, 4.65, 4.69, 4.65, 4.42, 4.59, 4.72, 4.17, 5.10,
               6.18, 4.96, 6.27, 6.20, 6.21, 6.81, 6.73, 6.08, 5.86, 6.28,
               6.97, 6.66, 6.08, 8.38, 6.67, 6.88, 6.52, 5.90, 7.84,
               5.46, 6.05, 5.26, 5.85, 4.25)

d1$`2020` <- c(7.35, 5.84, 5.72, 5.43, 5.89, 5.67, 5.94, 5.36, 5.83,
               7.25, 6.08, 7.21, 7.57, 6.98, 7.70, 7.96, 7.21, 6.81, 7.18,
               7.90, 7.14, 7.16, 9.91, 7.55, 7.35, 7.17, 6.40, 8.12,
               6.48, 6.78, 6.64, 6.80, 5.31)

d1$`2021` <- c(8.59, 6.42, 7.74, 6.06, 6.81, 6.60, 5.94, 5.41, 7.16,
               7.70, 6.24, 7.97, 7.97, 7.52, 8.54, 8.34, 7.46, 7.13, 7.70,
               9.41, 8.88, 7.98, 10.84, 9.25, 9.55, 9.71, 8.16, 10.27,
               7.99, 8.82, 8.02, 8.42, 6.17)

d1$`2022` <- c(7.60, 5.56, 6.50, 5.01, 5.11, 5.10, 5.60, 5.24, 6.20,
               7.38, 5.89, 7.44, 7.35, 7.33, 8.15, 7.95, 7.39, 6.69, 7.61,
               8.20, 7.92, 7.30, 9.39, 7.97, 8.19, 7.85, 6.74, 9.57,
               6.38, 7.41, 5.94, 6.70, 5.13)

d1$`2023` <- c(6.91, 5.09, 5.73, 4.78, 4.80, 4.74, 5.08, 4.91, 5.68,
               6.67, 5.52, 6.68, 6.59, 6.46, 6.97, 7.18, 6.81, 6.17, 6.94,
               7.51, 7.39, 6.70, 8.43, 7.29, 7.29, 7.09, 6.12, 8.32,
               5.94, 6.82, 5.76, 6.18, 4.74)


d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Taxa de Mortalidade Geral"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"

# Criando o vetor de ranking para 2023
d2$`2023` <- c(NA, 5, 19, 25, 24, 26, 22, 23, 20,
               3, 21, 12, 13, 14, 7, 5, 10, 16, 8,
               1, 3, 11, 1, 4, 2, 6, 17, 2,
               4, 9, 18, 15, 27)

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
d1$categoria2 <- "Masculino"

# Criando os vetores para Masculino por ano
d1$`2019` <- c(55.23, 60.93, 61.37, 61.81, 60.95, 62.25, 60.60, 60.39, 61.38,
                 55.95, 58.19, 55.95, 55.05, 55.80, 53.94, 54.49, 54.63, 56.75, 57.44,
                 53.74, 55.00, 56.70, 51.95, 53.78, 54.74, 56.54, 55.37, 52.93,
                 58.36, 57.52, 61.69, 58.30, 54.90)

d1$`2020` <- c(56.15, 61.29, 61.58, 61.93, 61.04, 61.40, 60.96, 62.70, 62.44,
                 57.01, 59.77, 57.54, 57.06, 56.57, 54.98, 54.91, 56.30, 57.75, 58.00,
                 54.51, 55.50, 57.25, 52.80, 54.70, 55.75, 57.26, 56.80, 53.89,
                 59.16, 58.65, 62.04, 58.85, 56.49)

d1$`2021` <- c(55.40, 59.66, 61.15, 58.48, 58.79, 61.33, 59.36, 60.84, 60.84,
                 56.03, 58.95, 56.92, 55.72, 55.22, 54.42, 53.96, 55.17, 56.89, 57.08,
                 54.07, 54.96, 55.79, 51.73, 54.57, 55.29, 56.87, 55.88, 53.46,
                 58.14, 57.50, 60.77, 57.95, 55.62)

d1$`2022` <- c(54.71, 60.05, 61.57, 58.48, 60.42, 62.51, 59.31, 59.62, 61.15,
                 55.65, 57.47, 56.30, 55.14, 55.41, 54.27, 54.08, 55.08, 56.34, 56.69,
                 53.17, 54.24, 55.74, 50.97, 53.41, 54.16, 55.92, 54.48, 52.49,
                 57.75, 57.90, 61.52, 57.24, 53.55)

d1$`2023` <- c(54.80, 59.99, 61.23, 58.12, 60.03, 60.29, 59.21, 64.40, 61.18,
                 55.77, 58.25, 56.95, 55.20, 54.78, 53.88, 54.19, 55.19, 56.47, 56.73,
                 53.20, 54.58, 56.22, 51.19, 53.17, 54.37, 55.94, 54.67, 52.82,
                 57.78, 57.45, 61.53, 57.29, 54.03)


d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Proporção de Mortes por Sexo"
d2$categoria1 <- "Proporção"
d2$categoria2 <- "Feminino"

# Criando os vetores para Feminino por ano
d2$`2019` <- c(44.73, 38.96, 38.52, 37.97, 38.99, 37.68, 39.25, 39.59, 38.47,
                44.00, 41.74, 43.98, 44.93, 44.14, 46.02, 45.46, 45.34, 43.18, 42.49,
                46.22, 44.99, 43.26, 47.97, 46.20, 45.24, 43.44, 44.61, 47.05,
                41.60, 42.48, 38.24, 41.65, 45.09)

d2$`2020` <- c(43.81, 38.62, 38.38, 37.70, 38.94, 38.44, 38.95, 37.30, 37.45,
                42.95, 40.20, 42.35, 42.92, 43.41, 44.99, 45.03, 43.67, 42.20, 41.96,
                45.45, 44.48, 42.64, 47.12, 45.28, 44.23, 42.72, 43.18, 46.10,
                40.81, 41.34, 37.93, 41.09, 43.50)

d2$`2021` <- c(44.56, 40.27, 38.83, 41.32, 41.18, 38.62, 40.54, 39.14, 39.10,
                43.93, 41.01, 42.97, 44.26, 44.75, 45.54, 45.99, 44.80, 43.04, 42.89,
                45.90, 45.02, 44.14, 48.20, 45.40, 44.68, 43.09, 44.11, 46.49,
                41.84, 42.48, 39.19, 42.03, 44.35)

d2$`2022` <- c(45.25, 39.88, 38.42, 41.40, 39.55, 37.49, 40.57, 40.33, 38.80,
                44.30, 42.44, 43.63, 44.83, 44.53, 45.72, 45.87, 44.88, 43.58, 43.26,
                46.80, 45.75, 44.17, 48.95, 46.57, 45.82, 44.05, 45.51, 47.49,
                42.19, 42.08, 38.45, 42.71, 46.30)

d2$`2023` <- c(45.17, 39.94, 38.74, 41.71, 39.94, 39.65, 40.69, 35.55, 38.82,
                46.77, 41.68, 42.96, 44.76, 45.14, 46.10, 45.77, 44.78, 43.47, 43.23,
                42.18, 45.40, 43.76, 48.74, 46.81, 45.60, 44.03, 45.32, 47.15,
                44.18, 42.54, 38.43, 42.66, 45.95)


d3 <- bases
d3$tematica <- "Saúde"
d3$indicador <- "Proporção de Mortes por Sexo"
d3$categoria1 <- "Proporção"
d3$categoria2 <- "Ignorado"

# Criando os vetores para Ignorado por ano
d3$`2019` <- c(0.04, 0.12, 0.11, 0.22, 0.05, 0.07, 0.14, 0.03, 0.15,
                     0.05, 0.07, 0.07, 0.02, 0.06, 0.04, 0.05, 0.03, 0.07, 0.07,
                     0.03, 0.01, 0.04, 0.08, 0.02, 0.02, 0.02, 0.02, 0.01,
                     0.04, 0.00, 0.07, 0.05, 0.01)

d3$`2020` <- c(0.04, 0.08, 0.04, 0.37, 0.02, 0.17, 0.09, 0.00, 0.11,
                     0.04, 0.03, 0.11, 0.02, 0.02, 0.02, 0.06, 0.03, 0.06, 0.05,
                     0.04, 0.02, 0.11, 0.08, 0.02, 0.02, 0.02, 0.02, 0.02,
                     0.04, 0.02, 0.03, 0.06, 0.01)

d3$`2021` <- c(0.04, 0.07, 0.01, 0.20, 0.03, 0.05, 0.10, 0.02, 0.06,
                     0.04, 0.05, 0.11, 0.02, 0.03, 0.03, 0.05, 0.02, 0.07, 0.03,
                     0.03, 0.02, 0.06, 0.07, 0.03, 0.03, 0.04, 0.01, 0.04,
                     0.02, 0.01, 0.05, 0.02, 0.03)

d3$`2022` <- c(0.04, 0.08, 0.01, 0.12, 0.03, 0.00, 0.12, 0.05, 0.04,
                     0.05, 0.09, 0.08, 0.03, 0.06, 0.02, 0.05, 0.03, 0.08, 0.05,
                     0.03, 0.01, 0.09, 0.08, 0.02, 0.02, 0.03, 0.01, 0.02,
                     0.05, 0.02, 0.03, 0.05, 0.15)

d3$`2023` <- c(0.04, 0.05, 0.03, 0.17, 0.03, 0.06, 0.10, 0.05, 0.00,
                     0.07, 0.07, 0.09, 0.04, 0.07, 0.02, 0.04, 0.03, 0.06, 0.05,
                     0.02, 0.02, 0.02, 0.07, 0.02, 0.03, 0.02, 0.01, 0.03,
                     0.03, 0.01, 0.04, 0.05, 0.01)

d4 <- bases
d4$tematica <- "Saúde"
d4$indicador <- "Proporção de Mortes por Sexo"
d4$categoria1 <- "Masculino"
d4$categoria2 <- "Ranking"

# Criando o vetor de ranking para Ranking em 2023
d4$`2023` <- c(NA, 1, 3, 9, 6, 5, 7, 1, 4, 3, 8, 12, 17, 19, 24, 22, 18, 14,
               13, 5, 21, 15, 27, 25, 4, 16, 20, 26, 2, 10, 2, 11, 23)


d5 <- bases
d5$tematica <- "Saúde"
d5$indicador <- "Proporção de Mortes por Sexo"
d5$categoria1 <- "Feminino"
d5$categoria2 <- "Ranking"

# Criando o vetor de ranking para Feminino em 2023
d5$`2023` <- c(NA, 5, 25, 19, 22, 23, 21, 27, 24, 1, 20, 16, 11, 9, 4, 6,
               10, 14, 15, 4, 7, 13, 1, 3, 2, 12, 8, 2, 3, 18, 26, 17, 5)

d6 <- bases
d6$tematica <- "Saúde"
d6$indicador <- "Proporção de Mortes por Sexo"
d6$categoria1 <- "Ignorado"
d6$categoria2 <- "Ranking"

# Criando o vetor de ranking para Ignorado em 2023
d6$`2023` <- c(NA, 2, 16, 1, 17, 8, 2, 9, 27,
                             1, 5, 3, 14, 6, 22, 12, 15, 7, 11,
                             5, 20, 21, 4, 23, 4, 19, 24, 18,
                             3, 26, 13, 10, 25)

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
  "saude/06 - Proporção de Mortes por Sexo.xlsx",
  overwrite = TRUE
)

##07 - Hospitais por 10 Mil Habitantes----
d1 <- bases
d1$tematica <- "Saúde"
d1$indicador <- "Hospitais por 10 Mil Habitantes"
d1$categoria1 <- "Quantidade"
d1$categoria2 <- "-"

# Criando os vetores para Hospitais por 10 Mil Habitantes por ano
d1$`2019` <- c(0.32, 0.30, 0.46, 0.26, 0.25, 0.21, 0.29, 0.15, 0.45,
                      0.35, 0.38, 0.35, 0.31, 0.33, 0.38, 0.29, 0.23, 0.26, 0.44,
                      0.27, 0.32, 0.27, 0.27, 0.24, 0.36, 0.42, 0.37, 0.28,
                      0.48, 0.40, 0.50, 0.61, 0.24)

d1$`2020` <- c(0.33, 0.31, 0.52, 0.23, 0.26, 0.21, 0.30, 0.16, 0.45,
                      0.36, 0.39, 0.35, 0.32, 0.33, 0.38, 0.32, 0.25, 0.25, 0.46,
                      0.28, 0.34, 0.27, 0.29, 0.25, 0.36, 0.42, 0.37, 0.29,
                      0.50, 0.41, 0.48, 0.63, 0.29)

d1$`2021` <- c(0.34, 0.33, 0.57, 0.29, 0.28, 0.21, 0.30, 0.18, 0.45,
                      0.38, 0.40, 0.36, 0.33, 0.37, 0.41, 0.33, 0.25, 0.26, 0.46,
                      0.28, 0.34, 0.28, 0.29, 0.24, 0.35, 0.41, 0.37, 0.29,
                      0.51, 0.41, 0.48, 0.63, 0.34)

d1$`2022` <- c(0.36, 0.36, 0.67, 0.31, 0.30, 0.20, 0.33, 0.31, 0.48,
                      0.40, 0.43, 0.36, 0.35, 0.42, 0.43, 0.36, 0.27, 0.26, 0.50,
                      0.29, 0.35, 0.31, 0.32, 0.25, 0.36, 0.40, 0.36, 0.30,
                      0.52, 0.43, 0.48, 0.64, 0.37)

d1$`2023` <- c(0.34, 0.34, 0.63, 0.30, 0.28, 0.20, 0.31, 0.29, 0.50,
                      0.39, 0.42, 0.35, 0.33, 0.42, 0.41, 0.34, 0.26, 0.28, 0.49,
                      0.28, 0.34, 0.29, 0.31, 0.25, 0.35, 0.39, 0.35, 0.30,
                      0.50, 0.41, 0.47, 0.62, 0.36)


d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Hospitais por 10 Mil Habitantes"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"

# Criando o vetor de ranking para 2023
d2$`2023` <- c(NA, 4, 1, 19, 23, 27, 18, 22, 3,
               2, 7, 12, 16, 6, 9, 15, 25, 24, 4,
               5, 14, 21, 17, 26, 3, 10, 13, 20,
               1, 8, 5, 2, 11)


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

# Criando os vetores para Postos e Centros de Saúde por 10 Mil Habitantes por ano
d1$`2019` <- c(2.21, 2.23, 1.88, 2.64, 1.48, 2.25, 2.50, 2.19, 2.86,
                   3.16, 3.28, 4.63, 2.66, 3.54, 4.05, 2.73, 3.10, 3.01, 3.07,
                   1.58, 2.75, 2.29, 1.21, 1.11, 2.37, 2.31, 2.61, 2.28,
                   2.02, 2.31, 2.84, 2.11, 0.58)

d1$`2020` <- c(2.25, 2.28, 1.89, 2.67, 1.50, 2.47, 2.56, 2.27, 2.94,
                   3.21, 3.32, 4.68, 2.68, 3.54, 4.08, 2.76, 3.11, 3.08, 3.18,
                   1.60, 2.80, 2.25, 1.22, 1.14, 2.41, 2.34, 2.65, 2.32,
                   2.03, 2.34, 2.93, 2.07, 0.59)

d1$`2021` <- c(2.27, 2.34, 1.90, 2.69, 1.60, 2.36, 2.64, 2.31, 2.94,
                   3.26, 3.37, 4.75, 2.73, 3.59, 4.12, 2.84, 3.13, 3.13, 3.20,
                   1.62, 2.82, 2.29, 1.23, 1.15, 2.40, 2.34, 2.62, 2.32,
                   2.02, 2.34, 2.96, 2.06, 0.58)

d1$`2022` <- c(2.40, 2.57, 2.23, 2.94, 1.77, 2.45, 2.86, 2.83, 3.17,
                   3.48, 3.62, 4.82, 2.94, 3.92, 4.23, 3.07, 3.40, 3.36, 3.40,
                   1.72, 2.97, 2.44, 1.35, 1.21, 2.44, 2.37, 2.54, 2.44,
                   2.09, 2.41, 2.88, 2.12, 0.64)

d1$`2023` <- c(2.32, 2.43, 1.99, 2.75, 1.66, 2.29, 2.75, 2.65, 3.00,
                   3.38, 3.54, 4.73, 2.87, 3.80, 4.09, 2.98, 3.31, 3.36, 3.28,
                   1.67, 2.92, 2.32, 1.25, 1.19, 2.37, 2.31, 2.45, 2.36,
                   2.01, 2.30, 2.75, 2.09, 0.60)




d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Postos e Centros de Saúde por 10 Mil Habitantes"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"

# Criando o vetor de ranking para 2023
d2$`2023` <- c(NA, 2, 23, 13, 24, 21, 14, 15, 8,
               1, 4, 1, 11, 3, 2, 9, 6, 5, 7,
               5, 10, 18, 25, 26, 3, 19, 16, 17,
               4, 20, 12, 22, 27)

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

# Criando os vetores para Médicos por 10 Mil Habitantes por ano
d1$`2019` <- c(19.68, 10.50, 14.24, 10.83, 11.13, 13.62, 8.53, 9.53, 14.58,
                    13.33, 8.09, 12.57, 12.59, 15.15, 15.48, 15.72, 13.15, 16.23, 13.47,
                    24.67, 22.10, 22.35, 24.84, 26.00, 22.50, 20.91, 22.09, 24.35,
                    20.04, 19.54, 14.83, 16.91, 33.81)

d1$`2020` <- c(20.46, 11.04, 15.42, 11.33, 11.26, 15.23, 8.93, 10.27, 15.65,
                    13.98, 8.56, 13.05, 13.55, 15.35, 16.44, 16.21, 14.20, 16.97, 14.10,
                    25.25, 23.13, 23.11, 24.58, 26.67, 23.92, 22.34, 23.71, 25.64,
                    21.50, 20.80, 16.08, 18.12, 36.28)

d1$`2021` <- c(21.52, 11.55, 16.43, 11.60, 11.54, 13.39, 9.51, 11.35, 16.49,
                    14.91, 9.25, 13.95, 14.43, 17.28, 17.57, 17.02, 15.15, 18.24, 14.90,
                    26.43, 24.16, 24.49, 25.95, 27.82, 25.03, 23.46, 25.02, 26.62,
                    22.89, 21.56, 17.93, 19.26, 38.30)

d1$`2022` <- c(23.86, 13.35, 20.17, 13.47, 12.91, 14.31, 11.05, 14.51, 18.66,
                    16.66, 10.33, 14.82, 16.26, 19.76, 18.83, 19.11, 17.52, 20.67, 16.64,
                    29.35, 26.37, 27.72, 29.57, 30.78, 26.88, 25.44, 25.86, 29.10,
                    25.17, 23.92, 18.58, 21.53, 44.05)

d1$`2023` <- c(24.48, 13.78, 19.69, 14.36, 13.37, 16.21, 11.31, 14.93, 19.96,
                    17.03, 10.61, 15.39, 16.67, 20.29, 19.32, 19.24, 17.93, 21.84, 16.91,
                    30.05, 27.39, 28.37, 29.06, 31.80, 27.97, 26.69, 27.16, 29.89,
                    25.85, 24.39, 19.82, 22.22, 43.86)


d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Médicos por 10 Mil Habitantes"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"

# Criando o vetor de ranking para 2023
d2$`2023` <- c(NA, 5, 15, 24, 25, 21, 26, 23, 13,
               4, 27, 22, 20, 12, 16, 17, 18, 11, 19,
               1, 6, 5, 4, 2, 2, 8, 7, 3,
               3, 9, 14, 10, 1)

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

# Criando os vetores para Leitos Hospitalares por Mil Habitantes por ano
d1$`2019` <- c(2.33, 1.89, 2.64, 1.91, 1.59, 2.04, 1.85, 1.49, 2.17,
                   2.22, 2.12, 2.44, 2.24, 2.28, 2.26, 2.51, 1.98, 1.61, 2.15,
                   2.33, 2.18, 2.38, 2.49, 2.32, 2.71, 2.71, 2.36, 2.93,
                   2.58, 2.25, 2.41, 2.73, 2.72)

d1$`2020` <- c(2.53, 2.08, 2.99, 2.14, 1.72, 2.43, 2.04, 1.68, 2.31,
                   2.45, 2.36, 2.76, 2.45, 2.58, 2.49, 2.78, 2.25, 1.74, 2.32,
                   2.52, 2.40, 2.45, 2.75, 2.49, 2.81, 2.82, 2.48, 3.02,
                   2.87, 2.44, 2.51, 2.97, 3.47)

d1$`2021` <- c(2.56, 2.12, 3.13, 2.03, 1.77, 2.38, 2.05, 1.59, 2.53,
                   2.51, 2.38, 2.78, 2.52, 2.67, 2.70, 2.84, 2.30, 1.73, 2.37,
                   2.53, 2.44, 2.60, 2.68, 2.51, 2.80, 2.75, 2.53, 3.02,
                   3.00, 2.40, 2.69, 3.19, 3.44)

d1$`2022` <- c(2.59, 2.21, 3.40, 2.18, 1.86, 2.38, 2.10, 2.24, 2.44,
                   2.56, 2.54, 2.57, 2.55, 2.70, 2.68, 2.98, 2.39, 1.72, 2.40,
                   2.57, 2.39, 2.67, 2.86, 2.54, 2.74, 2.66, 2.33, 3.09,
                   2.91, 2.54, 2.47, 3.02, 3.56)

d1$`2023` <- c(2.50, 2.09, 3.21, 2.07, 1.78, 2.27, 1.94, 2.13, 2.44,
                   2.47, 2.44, 2.48, 2.41, 2.54, 2.69, 2.79, 2.35, 1.71, 2.36,
                   2.49, 2.32, 2.61, 2.73, 2.47, 2.65, 2.59, 2.31, 2.95,
                   2.80, 2.36, 2.44, 2.90, 3.46)



d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Leitos Hospitalares por Mil Habitantes"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"

# Criando o vetor de ranking para 2023
d2$`2023` <- c(NA, 5, 2, 24, 26, 22, 25, 23, 13,
                           4, 14, 11, 16, 10, 7, 5, 19, 27, 17,
                           3, 20, 8, 6, 12, 2, 9, 21, 3,
                           1, 18, 15, 4, 1)



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

# Criando os vetores para Percentual de Nascidos Vivos com 7 ou mais Consultas Pré-Natal por ano
d1$`2019` <- c(72.43, 53.45, 68.91, 52.21, 50.14, 43.73, 51.65, 42.94, 70.34,
                     67.50, 52.07, 65.87, 75.56, 69.86, 72.94, 72.29, 70.28, 63.06, 66.17,
                     78.19, 79.13, 72.16, 72.95, 80.21, 81.94, 85.58, 80.02, 79.19,
                     72.01, 70.71, 72.53, 70.67, 75.68)

d1$`2020` <- c(71.02, 49.41, 67.40, 46.92, 47.42, 40.88, 46.38, 35.63, 68.15,
                     65.20, 48.07, 59.51, 74.51, 68.77, 68.88, 71.76, 63.22, 60.57, 65.10,
                     78.08, 79.45, 70.81, 71.28, 80.62, 80.88, 84.86, 78.00, 78.59,
                     70.47, 69.55, 70.22, 69.67, 73.65)

d1$`2021` <- c(73.14, 54.28, 69.59, 41.56, 53.37, 45.04, 52.85, 41.77, 70.82,
                     69.01, 54.62, 65.47, 77.33, 72.57, 73.20, 74.29, 68.34, 65.99, 67.91,
                     79.06, 80.33, 72.76, 71.93, 81.68, 82.11, 85.34, 80.06, 80.01,
                     72.10, 71.22, 72.53, 70.99, 75.06)

d1$`2022` <- c(74.79, 58.61, 72.54, 47.25, 58.27, 45.65, 56.97, 49.61, 73.88,
                     71.87, 61.85, 69.85, 79.25, 74.31, 73.74, 75.36, 71.74, 70.27, 70.07,
                     79.10, 80.19, 73.98, 73.13, 81.23, 82.50, 85.59, 80.46, 80.55,
                     73.65, 73.34, 74.16, 72.92, 74.98)

d1$`2023` <- c(77.17, 62.02, 75.20, 56.87, 61.80, 50.62, 60.15, 51.17, 75.04,
                     74.46, 65.38, 73.77, 81.47, 75.20, 73.80, 76.95, 76.17, 74.59, 73.08,
                     81.03, 81.72, 75.68, 76.87, 82.72, 84.53, 87.39, 82.82, 82.60,
                     76.47, 77.22, 76.79, 75.58, 77.40)

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Percentual de Nascidos Vivos com 7 ou mais Consultas Pré-Natal"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"

# Criando o vetor de ranking para 2023
d2$`2023` <- c(NA, 5, 15, 25, 23, 27, 24, 26, 17,
               4, 22, 20, 6, 16, 19, 9, 12, 18, 21,
               2, 5, 13, 10, 3, 1, 1, 2, 4,
               3, 8, 11, 14, 7)

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
d1$categoria1 <- "Vaginal"
d1$categoria2 <- "Percentual"

d2 <- bases
d2$tematica <- "Saúde"
d2$indicador <- "Percentual de Nascidos Vivos por Tipo de Parto"
d2$categoria1 <- "Cesário"
d2$categoria2 <- "Percentual"

d3 <- bases
d3$tematica <- "Saúde"
d3$indicador <- "Percentual de Nascidos Vivos por Tipo de Parto"
d3$categoria1 <- "Ignorado"
d3$categoria2 <- "Percentual"

d4 <- bases
d4$tematica <- "Saúde"
d4$indicador <- "Percentual de Nascidos Vivos por Tipo de Parto"
d4$categoria1 <- "Vaginal"
d4$categoria2 <- "Ranking"

d5 <- bases
d5$tematica <- "Saúde"
d5$indicador <- "Percentual de Nascidos Vivos por Tipo de Parto"
d5$categoria1 <- "Cesário"
d5$categoria2 <- "Ranking"

d6 <- bases
d6$tematica <- "Saúde"
d6$indicador <- "Percentual de Nascidos Vivos por Tipo de Parto"
d6$categoria1 <- "Ignorado"
d6$categoria2 <- "Ranking"


# Criando os vetores para Percentual de Nascidos Vivos por Tipo de Parto em 2019

# Parto Vaginal
d1$`2019` <- c(43.63, 52.22, 33.09, 55.50, 60.94, 65.35, 49.66, 63.41, 43.08,
                          47.50, 49.85, 42.22, 41.40, 37.52, 39.30, 48.86, 47.54, 55.37, 53.85,
                          41.45, 41.82, 40.00, 42.24, 41.14, 38.62, 37.66, 42.50, 36.89,
                          37.53, 37.81, 39.04, 32.98, 45.45)
# Parto Vaginal
d1$`2020` <- c(42.70, 51.25, 31.92, 53.61, 60.77, 65.22, 47.92, 62.44, 44.12,
                          46.58, 48.93, 41.38, 39.76, 35.94, 38.32, 48.06, 47.87, 55.80, 52.94,
                          40.80, 41.36, 39.80, 41.51, 40.38, 37.15, 35.05, 41.78, 36.03,
                          36.19, 37.66, 36.76, 31.54, 44.79)
# Parto Vaginal
d1$`2021` <- c(42.93, 50.77, 31.64, 52.40, 60.44, 67.35, 46.82, 62.26, 44.07,
                          46.36, 47.92, 41.46, 38.46, 35.54, 38.39, 48.81, 47.22, 56.26, 53.16,
                          41.49, 41.95, 39.75, 41.74, 41.36, 37.30, 35.33, 42.11, 35.82,
                          36.15, 36.90, 35.55, 32.48, 45.01)
# Parto Vaginal
d1$`2022` <- c(41.85, 48.90, 31.45, 51.23, 58.56, 64.91, 44.69, 61.14, 42.84,
                          44.58, 44.91, 39.07, 35.97, 34.91, 36.05, 47.90, 44.17, 55.57, 52.08,
                          40.71, 41.56, 39.38, 40.55, 40.50, 37.56, 35.21, 43.17, 35.74,
                          36.07, 36.45, 35.67, 33.06, 43.80)
# Parto Vaginal
d1$`2023` <- c(40.36, 47.32, 29.89, 49.99, 57.04, 62.01, 43.32, 57.47, 41.96,
                          42.57, 42.51, 37.34, 33.14, 32.52, 34.60, 46.26, 44.04, 53.63, 49.98,
                          39.53, 40.82, 38.83, 39.54, 39.00, 36.38, 33.63, 42.86, 34.37,
                          34.56, 33.56, 34.91, 31.84, 42.13)
# Parto Cesário
d2$`2019` <- c(56.30, 47.73, 66.73, 44.42, 39.06, 34.65, 50.26, 36.57, 56.91,
                          52.37, 49.97, 57.73, 58.43, 62.34, 60.64, 51.00, 52.38, 44.62, 46.04,
                          58.51, 58.11, 60.00, 57.69, 58.84, 61.33, 62.28, 57.39, 63.10,
                          62.46, 62.19, 60.93, 67.02, 54.53)
# Parto Cesário
d2$`2020` <- c(57.22, 48.70, 67.97, 46.30, 39.23, 34.78, 52.01, 37.50, 55.88,
                          53.23, 50.38, 58.59, 60.13, 63.98, 61.36, 51.85, 52.06, 44.18, 46.97,
                          59.16, 58.55, 60.18, 58.43, 59.59, 62.81, 64.90, 58.15, 63.96,
                          63.79, 62.30, 63.24, 68.44, 55.18)
# Parto Cesário
d2$`2021` <- c(57.01, 49.16, 68.02, 47.50, 39.56, 32.65, 53.11, 37.62, 55.93,
                          53.54, 51.91, 58.33, 61.43, 64.44, 61.52, 51.14, 52.78, 43.70, 46.70,
                          58.48, 57.98, 60.22, 58.24, 58.62, 62.66, 64.61, 57.86, 64.16,
                          63.83, 63.09, 64.40, 67.51, 54.98)
# Parto Cesário
d2$`2022` <- c(58.10, 51.05, 68.48, 48.71, 41.43, 35.09, 55.23, 38.84, 57.16,
                          55.34, 54.91, 60.85, 63.99, 65.06, 63.84, 52.01, 55.83, 44.42, 47.83,
                          59.26, 58.37, 60.61, 59.41, 59.49, 62.40, 64.75, 56.81, 64.20,
                          63.92, 63.54, 64.31, 66.92, 56.19)
# Parto Cesário
d2$`2023` <- c(59.59, 52.62, 69.96, 49.95, 42.96, 37.99, 56.61, 42.47, 58.04,
                          57.36, 57.38, 62.62, 66.79, 67.43, 65.37, 53.64, 55.94, 46.36, 49.92,
                          60.44, 59.12, 61.13, 60.42, 60.99, 63.55, 66.29, 57.11, 65.54,
                          65.42, 66.43, 65.03, 68.15, 57.86)


# Parto Ignorado
d3$`2019` <- c(0.07, 0.06, 0.18, 0.07, 0.00, 0.00, 0.08, 0.03, 0.01,
                           0.13, 0.18, 0.06, 0.17, 0.14, 0.06, 0.15, 0.08, 0.01, 0.11,
                           0.04, 0.07, 0.01, 0.07, 0.02, 0.05, 0.05, 0.11, 0.01,
                           0.02, 0.00, 0.03, 0.01, 0.02)
# Parto Ignorado
d3$`2020` <- c(0.08, 0.05, 0.10, 0.09, 0.00, 0.00, 0.07, 0.05, 0.00,
                           0.18, 0.69, 0.03, 0.10, 0.08, 0.32, 0.09, 0.07, 0.02, 0.10,
                           0.04, 0.08, 0.02, 0.06, 0.02, 0.04, 0.05, 0.06, 0.01,
                           0.02, 0.04, 0.00, 0.01, 0.04)
# Parto Ignorado
d3$`2021` <- c(0.06, 0.07, 0.34, 0.10, 0.00, 0.00, 0.07, 0.11, 0.00,
                           0.11, 0.17, 0.22, 0.11, 0.02, 0.08, 0.06, 0.01, 0.04, 0.14,
                           0.03, 0.06, 0.02, 0.02, 0.01, 0.03, 0.06, 0.02, 0.02,
                           0.02, 0.01, 0.05, 0.01, 0.01)
# Parto Ignorado
d3$`2022` <- c(0.05, 0.05, 0.07, 0.06, 0.00, 0.00, 0.08, 0.02, 0.01,
                           0.08, 0.18, 0.08, 0.05, 0.04, 0.11, 0.09, 0.00, 0.02, 0.08,
                           0.03, 0.08, 0.01, 0.04, 0.01, 0.04, 0.03, 0.02, 0.06,
                           0.02, 0.01, 0.03, 0.02, 0.01)
# Parto Ignorado
d3$`2023` <- c(0.05, 0.01, 0.05, 0.04, 0.08, 0.09, 0.03, 0.02, 0.10,
                           0.01, 0.02, 0.11, 0.00, 0.06, 0.00, 0.04, 0.06, 0.07, 0.02,
                           0.06, 0.03, 0.04, 0.05, 0.06, 0.01, 0.08, 0.10, 0.06,
                           0.02, 0.01, 0.08, 0.15, 0.02)
# Ranking Parto Vaginal 2023
d4$`2023` <- c(NA, 1, 27, 5, 3, 1, 9, 2, 13,
                            2, 11, 18, 24, 25, 20, 7, 8, 4, 6,
                            3, 14, 17, 15, 16, 4, 22, 10, 21,
                            5, 23, 19, 26, 12)
# Ranking Parto Cesário 2023
d5$`2023` <- c(NA, 5, 1, 22, 25, 27, 19, 26, 15,
                            4, 17, 10, 4, 3, 8, 21, 20, 24, 23,
                            3, 14, 11, 13, 12, 2, 6, 18, 7,
                            1, 5, 9, 2, 16)
# Ranking Parto Ignorado 2023
d6$`2023` <- c(NA, 4, 14, 17, 6, 5, 20, 21, 3,
                             5, 23, 2, 26, 11, 27, 18, 13, 9, 24,
                             1, 19, 16, 15, 10, 3, 8, 4, 12,
                             2, 25, 7, 1, 22)


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


# Vetores com os óbitos por ano
d1$`2019` <- c(235301, 12098, 1446, 561, 2729, 419, 5279, 503, 1161, 52717, 4501, 2938, 9748, 3691, 4296, 9698, 2550, 1889, 13406,
               110909, 24524, 4774, 22656, 58955, 44064, 14860, 9286, 19918, 15513, 2895, 2889, 6874, 2855)

d1$`2020` <- c(229300, 11777, 1343, 539, 2637, 406, 5238, 507, 1107, 51705, 4426, 2687, 9509, 3620, 4121, 9230, 2522, 1892, 13698,
               106972, 24653, 4532, 22081, 55706, 43480, 14939, 9344, 19197, 15366, 2860, 2953, 6768, 2785)

d1$`2021` <- c(235805, 12156, 1415, 531, 2696, 425, 5338, 516, 1235, 52810, 4561, 3004, 9545, 3613, 4249, 9524, 2682, 1817, 13815,
               110148, 25232, 4475, 22055, 58386, 44989, 15392, 9738, 19859, 15702, 3040, 2977, 6875, 2810)

d1$`2022` <- c(244009, 12887, 1528, 583, 2864, 455, 5689, 510, 1258, 54548, 4695, 3084, 9954, 3680, 4489, 9616, 2657, 1884, 14489,
               113545, 26214, 4663, 22806, 59862, 46762, 15876, 10329, 20557, 16267, 3090, 3151, 7241, 2785)

d1$`2023` <- c(253983, 13764, 1596, 652, 3069, 485, 6043, 551, 1368, 57299, 4819, 3056, 10453, 3894, 4547, 10139, 2833, 2130, 15428,
               117982, 27125, 5029, 23805, 62023, 47738, 16461, 10804, 20473, 17200, 3271, 3389, 7642, 2898)

# Vetor com os rankings de 2023 (nomes entre aspas, conforme estão na tabela)
d2$`2023` <- c(NA, 5, 23, 25, 18, 27, 11, 26, 24, 2, 13, 19, 8, 15, 14, 9, 21, 22, 6, 1, 2, 12, 3, 1, 3, 5, 7, 4, 4, 17, 16, 10, 20)



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

# Vetores de óbitos por ano
d1$`2019` <- c(235301, 12098, 1446, 561, 2729, 419, 5279, 503, 1161, 52717, 4501, 2938, 9748, 3691, 4296, 9698, 2550, 1889, 13406,
               110909, 24524, 4774, 22656, 58955, 44064, 14860, 9286, 19918, 15513, 2895, 2889, 6874, 2855)

d1$`2020` <- c(229300, 11777, 1343, 539, 2637, 406, 5238, 507, 1107, 51705, 4426, 2687, 9509, 3620, 4121, 9230, 2522, 1892, 13698,
               106972, 24653, 4532, 22081, 55706, 43480, 14939, 9344, 19197, 15366, 2860, 2953, 6768, 2785)

d1$`2021` <- c(235805, 12156, 1415, 531, 2696, 425, 5338, 516, 1235, 52810, 4561, 3004, 9545, 3613, 4249, 9524, 2682, 1817, 13815,
               110148, 25232, 4475, 22055, 58386, 44989, 15392, 9738, 19859, 15702, 3040, 2977, 6875, 2810)

d1$`2022` <- c(244009, 12887, 1528, 583, 2864, 455, 5689, 510, 1258, 54548, 4695, 3084, 9954, 3680, 4489, 9616, 2657, 1884, 14489,
               113545, 26214, 4663, 22806, 59862, 46762, 15876, 10329, 20557, 16267, 3090, 3151, 7241, 2785)

d1$`2023` <- c(253983, 13764, 1596, 652, 3069, 485, 6043, 551, 1368, 57299, 4819, 3056, 10453, 3894, 4547, 10139, 2833, 2130, 15428,
               117982, 27125, 5029, 23805, 62023, 47738, 16461, 10804, 20473, 17200, 3271, 3389, 7642, 2898)

# Vetor do ranking 2023 já em formato numérico
d2$`2023` <- c(NA, 5, 23, 25, 18, 27, 11, 26, 24, 2, 13, 19, 8, 15, 14, 9, 21, 22, 6,
               1, 2, 12, 3, 1, 3, 5, 7, 4, 4, 17, 16, 10, 20)



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


# Vetores de óbitos por ano
`d1$2019` <- c(56666,4221,340,192,983,130,2164,184,228,15222,1734,754,2375,884,1092,2905,
               916,574,3988,25701,6290,708,8147,10556,7571,2330,1602,3639,3951,635,828,1833,655)

`d1$2020` <- c(267287,24113,2212,1006,6972,959,10184,1292,1488,69227,6733,3566,14149,3966,4660,
               14119,4543,2870,14621,122551,19417,5615,39263,58256,30276,11013,6732,12531,
               21120,2823,5207,9320,3770)

`d1$2021` <- c(486667,35680,5301,1517,10337,1412,12777,1236,3100,92273,9142,5389,18254,5902
               ,7293,15786,5262,4051,21194,231051,50866,8128,49188,122869,83068,35553,16491,31024,44595,7694,10128,20301,6472)

`d1$2022` <- c(131800,8984,1022,365,1927,292,4272,350,756,30808,2736,1727,5296,1904,2687,
                              5711,1996,989,7762,61924,14093,2157,13774,31900,20132,7192,4124,8816,9952,1714,2105,4749,1384)

`d1$2023` <- c(73827,5482,488,262,1344,246,2639,177,326,18831,1901,941,3039,1098
               ,1395,3705,1320,682,4750,34383,8324,1004,9190,15865,10053,3520,2011,4522,5078,848,1047,2391,792)

# Vetor do ranking 2023 já em formato numérico
`d2$2023` <- c(NA,4,23,25,14,26,9,27,24,2,12,19,8,16,13,6,15,22,4,1,3,
               18,2,1,3,7,11,5,5,20,17,10,21)

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


# Vetores de óbitos por doenças cardiovasculares (2019-2023)
`d1$2019` <- c(364132,19603,1999,926,3386,610,9644,711,2327,96244,10726,6415,15490,
               6132,7743,17876,6287,3135,22440,170916,35493,7046,38939,89438,54161,20294,11468,22399,23208,4894,4430,10752,3132)

`d1$2020` <- c(357741,20353,2046,771,3565,601,10206,765,2399,96649,11318,6728,14955,
               6037,7606,16705,6041,3156,24103,164015,35595,6880,37073,84467,52739,20224,11447,21068,23985,5157,4570,11010,3248)

`d1$2021` <- c(382507,21279,2237,861,4009,638,10267,725,2542,99733,11151,6887,15941,
               6139,8057,17214,6177,3345,24822,177183,37799,7146,39123,93115,58508,22279,12882,23347,25804,5526,4797,12192,3289)

`d1$2022` <- c(400154,22020,2286,910,4016,675,10704,826,2603,104694,11750,7287,16426,
               6673,8623,17215,6731,3463,26526,184600,39147,7552,38509,99392,62490,23502,13523,25465,26350,5656,5048,12325,3321)

`d1$2023` <- c(386086,22413,2383,935,4319,698,10796,834,2448,101458,11394,6740,15160,
               6113,7811,18292,6625,3423,25900,177992,37588,7261,37857,95286,58166,21928,13130,23108,26057,5660,5252,11904,3241)

# Ranking 2023 
`d2$2023` <- c(NA,5,24,25,20,27,12,26,23,2,11,15,8,17,13,7,16,21,4,1,3,14,2,1,3,6,9,5,4,18,19,10,22)



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


# Vetores de óbitos por doenças do aparelho digestivo (2019-2023)
`d1$2019` <- c(68770,3858,353,174,805,114,1847,160,405,18158,1846,1002,3019,1148,
               1305,3596,1177,676,4389,31583,7124,1212,6081,17166,10276,4200,2035,4041,4895,912,868,2342,773)

`d1$2020` <- c(66667,3682,362,163,749,134,1717,160,397,17897,1884,1055,2795,1105,
               1349,3371,1168,638,4532,29883,7083,1111,5584,16105,10286,4280,2097,3909,4919,918,822,2392,787)

`d1$2021` <- c(72317,4150,363,156,829,164,1952,159,527,18893,2038,1065,3053,1127,
               1396,3601,1131,671,4811,32820,7682,1212,5884,18042,10965,4583,2227,4155,5489,973,946,2722,848)

`d1$2022` <- c(76485,4296,382,152,906,136,2110,142,468,20474,2226,1245,3451,1196,
               1519,3846,1121,739,5131,34176,7969,1253,6039,18915,11692,4921,2397,4374,5847,1120,1068,2765,894)

`d1$2023` <- c(76550,4448,416,181,974,153,2132,127,465,20536,2105,1172,3285,1188,
               1686,4048,1206,703,5143,34226,8001,1498,6026,18701,11462,4858,2366,4238,5878,1110,1065,2802,901)

# Ranking 2023 (convertido: "°" removido, "-" como NA)
`d2$2023` <- c(NA,5,24,25,20,26,11,27,23,2,12,17,8,16,13,7,15,22,4,1,2,14,3,1,3,5,10,6,4,18,19,9,21)


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

# Vetores de óbitos por doenças respiratórias (2019-2023)
`d1$2019` <- c(162005,9064,909,639,1574,236,4573,409,724,38807,3217,2389,7669,
               2445,3266,8017,2074,1275,8455,78472,18192,2510,17253,40517,25044,9155,5219,10670,10618,2248,1927,5189,1254)

`d1$2020` <- c(148773,10427,711,488,1747,225,6260,329,667,35431,4204,2071,6553,
               2179,2991,7598,1670,1129,7036,74233,15592,2193,16039,40409,19396,7293,4123,7980,9286,2008,1734,4374,1170)

`d1$2021` <- c(142468,9085,794,616,1940,266,4399,356,714,34453,3308,2048,6345,
               2151,3008,7764,1712,903,7214,69906,15096,1878,16548,36384,20089,7415,4329,8345,8935,2062,1544,4313,1016)

`d1$2022` <- c(176073,10485,975,629,1877,320,5319,460,905,44971,4013,2637,8618,
               2771,4025,8764,2351,1515,10277,81611,19660,2441,16856,42654,27473,10069,5985,11419,11533,2648,1940,5645,1300)

`d1$2023` <- c(170843,10584,1076,626,2027,351,5127,525,852,42137,3895,2494,8493,
               2541,3426,8151,2217,1320,9600,81850,20041,2334,16993,42482,24707,9311,5274,10122,11565,2497,1984,5624,1460)

# Ranking 2023 (convertido: "°" removido, "-" como NA)
`d2$2023` <- c(NA,5,23,25,19,27,11,26,24,2,12,16,7,14,13,8,18,22,5,1,2,17,3,1,3,6,10,4,4,15,20,9,21)


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


# Vetores de óbitos por afecções perinatais (2019-2023)
`d1$2019` <- c(20354,2536,166,123,644,88,1240,150,125,6545,983,406,911,338,446,933,
               383,342,1803,7423,1780,309,1473,3861,2281,881,536,864,1569,252,396,706,215)

`d1$2020` <- c(18770,2508,196,125,550,129,1211,157,140,6159,913,342,872,315,425,902,
               357,307,1726,6620,1623,303,1433,3261,2026,780,526,720,1457,257,346,625,229)

`d1$2021` <- c(18602,2588,171,154,620,108,1191,164,180,6099,889,362,760,329,458,916,
               389,255,1741,6427,1525,327,1376,3199,1982,744,517,721,1506,223,393,651,239)

`d1$2022` <- c(18029,2261,195,113,540,87,1053,131,142,5773,800,370,774,288,441,869,
               355,291,1585,6482,1588,301,1279,3314,2063,808,533,722,1450,232,404,622,192)

`d1$2023` <- c(17666,2296,155,125,544,127,1067,124,154,5452,763,343,771,252,375,764,
               369,299,1516,6424,1517,312,1300,3295,1972,821,477,674,1522,258,434,606,224)

# Ranking 2023 (convertido: "°" removido, "-" como NA)
`d2$2023` <- c(NA,3,23,26,12,25,5,27,24,2,9,17,7,21,15,8,16,19,3,1,2,18,4,1,4,6,13,10,5,20,14,11,22)



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


# Vetores de proporção de cobertura ACS (2016-2020)
`d1$2016` <- c(65.15,79.19,80.71,96.94,62.76,67.26,82.69,71.33,99.73,85.12,88.90,98.88
               ,81.62,78.50,98.23,85.97,77.18,91.38,80.91,51.22,73.31,65.35,54.76,38.26,
               60.32,60.14,73.67,52.40,62.56,85.64,78.27,64.93,18.58)

`d1$2017` <- c(64.46,77.14,77.91,94.01,61.34,68.73,81.84,50.90,99.27,85.00,89.11,98.23
               ,81.33,77.06,97.46,86.02,77.01,91.65,81.19,50.81,74.19,62.21,54.34,37.51,
               58.25,56.58,71.64,51.72,62.40,86.87,77.88,64.70,17.98)

`d1$2018` <- c(64.03,77.93,76.44,93.33,63.21,69.57,81.21,71.50,98.44,84.47,88.92,97.78
               ,80.70,76.37,96.89,85.63,77.02,90.82,80.44,50.37,75.52,61.68,52.74,36.71,
               57.13,56.04,69.42,50.62,62.33,90.10,75.10,63.97,19.84)

`d1$2019` <- c(63.27,76.99,75.62,87.24,63.64,63.13,80.76,68.27,96.94,84.61,89.75,97.43,
               80.21,77.09,97.50,84.58,76.83,90.90,81.12,49.68,75.57,62.40,47.55,37.41,
               55.60,54.70,67.75,48.92,60.78,91.49,73.29,62.55,13.80)

`d1$2020` <- c(61.03,74.53,72.10,87.52,62.98,64.14,76.73,73.59,92.88,82.66,88.80,96.19,
               78.13,77.63,96.09,81.65,76.08,85.22,78.82,47.39,74.51,60.60,41.84,35.82,
               53.12,52.94,65.81,45.30,58.51,89.35,70.73,58.50,15.98)

# Ranking 2020 (convertido: "°" removido, "-" como NA)
`d2$2020` <- c(NA,2,16,6,20,19,12,15,3,1,5,1,10,11,2,8,13,7,9,5,14,21,25,26,4,23,18,24,3,4,17,22,27)

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

# Vetores de proporção de cobertura ESF (2016-2020)
`d1$2016` <- c(62.62,62.49,71.27,78.03,53.93,74.53,55.80,68.83,95.30,78.33,82.84,
               99.06,77.34,79.29,94.50,75.56,75.11,83.59,69.69,52.10,77.28,58.97,
               54.44,38.79,65.14,65.18,78.64,56.91,58.87,65.70,67.41,64.98,29.23)

`d1$2017` <- c(63.91,62.91,67.92,78.41,55.39,68.51,58.92,49.39,95.05,80.29,84.18,
               99.40,82.40,78.45,94.24,77.21,75.80,83.99,72.38,53.20,78.66,58.10,
               57.11,39.36,65.94,65.16,78.74,58.88,60.78,67.94,69.94,65.48,33.60)

`d1$2018` <- c(64.19,64.48,68.59,78.00,57.83,72.54,59.13,68.64,93.89,79.65,84.36,
               99.73,79.84,77.07,93.96,76.71,75.28,81.71,72.47,52.98,80.04,57.61,
               56.37,38.65,66.32,64.17,79.87,60.08,65.57,70.28,69.42,66.62,54.82)

`d1$2019` <- c(64.47,64.76,70.69,71.38,61.17,63.92,59.83,59.01,94.15,81.74,85.40,
               99.95,82.90,78.28,95.85,77.16,75.99,86.08,75.84,52.58,80.75,63.22,
               50.39,39.47,66.48,64.60,81.52,58.98,64.29,70.65,69.86,68.08,43.14)

`d1$2020` <- c(63.62,64.69,69.92,75.18,64.12,66.52,57.64,63.73,92.76,82.33,85.44,
               99.03,83.88,80.56,94.99,76.98,75.54,86.63,77.54,50.99,77.53,65.11,
               47.55,38.82,63.66,63.31,78.19,54.87,65.29,74.57,70.12,64.07,54.00)

# Ranking 2020 (convertido: "°" removido, "-" como NA)
`d2$2020` <- c(NA,3,16,13,19,17,23,21,3,1,5,1,6,7,2,11,12,4,9,5,10,18,26,27,4,22,8,24,2,14,15,20,25)


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

# Vetores de proporção de cobertura da atenção básica (2021-2024)
`d1$2021` <- c(72.41,66.27,72.32,68.50,65.29,74.43,55.33,39.93,88.07,81.34,78.09,
               95.11,88.92,76.10,86.44,72.55,73.23,87.82,73.76,66.63,84.59,71.84,
               57.03,53.05,79.53,79.81,91.46,67.32,68.28,77.72,77.53,63.22,54.65)

`d1$2022` <- c(78.91,75.08,86.65,81.90,78.43,83.00,67.14,58.10,98.05,90.25,87.31,
               98.23,98.36,85.92,93.95,81.37,81.15,97.89,83.54,71.97,91.90,82.97,
               69.22,62.79,84.93,86.60,91.88,78.30,76.76,85.93,81.91,72.74,71.16)

`d1$2023` <- c(80.35,74.64,82.24,82.55,76.50,81.39,67.90,62.76,96.77,87.34,86.71,
               96.47,95.49,84.65,92.67,81.59,80.03,97.54,83.34,75.25,91.04,84.98,
               70.51,68.86,86.69,89.13,90.26,81.61,78.11,89.02,79.40,75.54,72.20)

`d1$2024` <- c(80.19,75.45,81.12,82.29,76.35,80.90,70.18,67.27,93.56,85.83,85.71
               ,95.91,92.21,83.98,89.68,79.96,77.58,93.75,83.30,75.24,89.23,83.37
               ,71.16,69.55,87.06,89.00,92.56,81.08,79.73,87.92,83.95,77.10,72.83)

# Ranking 2024 (convertido: "°" removido, "-" como NA)
`d2$2024` <- c(NA,4,16,15,22,18,25,27,3,2,10,1,5,11,6,19,20,2,14,5,7,13,24,26,1,8,4,17,3,9,12,21,23)



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

