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

d4$`2023` <- c(NA, 1, 3, 9, 6, 5, 7, 1, 4, 3, 8, 12, 17, 19, 24, 22, 18, 14,
               13, 5, 21, 15, 27, 25, 4, 16, 20, 26, 2, 10, 2, 11, 23)


d5 <- bases
d5$tematica <- "Saúde"
d5$indicador <- "Proporção de Mortes por Sexo"
d5$categoria1 <- "Feminino"
d5$categoria2 <- "Ranking"

d5$`2023` <- c(NA, 5, 25, 19, 22, 23, 21, 27, 24, 1, 20, 16, 11, 9, 4, 6,
               10, 14, 15, 4, 7, 13, 1, 3, 2, 12, 8, 2, 3, 18, 26, 17, 5)

d6 <- bases
d6$tematica <- "Saúde"
d6$indicador <- "Proporção de Mortes por Sexo"
d6$categoria1 <- "Ignorado"
d6$categoria2 <- "Ranking"

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

