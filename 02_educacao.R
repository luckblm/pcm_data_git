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
d6$categoria1 <- "Ranking 2023"
d6$categoria2 <- "-"

# Matrícula Total por ano
d1$`2019` <- c(3755092, 3651989, 1279242, 112092, 32912, 6961, 80708, 4686, 26578,
               876270, 134344, 54016, 182672, 59377, 60062, 109589, 49720, 27613,
               198877, 1818540, 330436, 73787, 257506, 1156811, 626649)

d1$`2020` <- c(3651989, 1279242, 112092, 32912, 6961, 80708, 4686, 26578, 876270,
               134344, 54016, 182672, 59377, 60062, 109589, 49720, 27613, 198877,
               1818540, 330436, 73787, 257506, 1156811, 626649)

d1$`2021` <- c(1279242, 130956, 4371, 4041, 4371, 6710, 74297, 3748, 53105, 10190,
               84124, 36629, 53469, 45317, 22397, 14434, 24658, 36585, 125625,
               60209, 280423, 48189, 107000, 134580, 121470, 251693, 308506)

d1$`2022` <- c(1269263, 129676, 4754, 5029, 4355, 6935, 77774, 3900, 57263, 10374,
               90285, 38163, 55367, 47235, 23046, 14469, 25210, 37618, 128095,
               61571, 290495, 48932, 111927, 137290, 120204, 265701, 315527)

d1$`2023` <- c(1282369, 133559, 4810, 5152, 4576, 7042, 78679, 3989, 58552, 10524,
               91654, 39142, 56142, 48557, 23567, 14723, 25876, 38069, 130454,
               62782, 296135, 49873, 114261, 141275, 122763, 270113, 323126)

d2$`2019` <- c(782, "-", "-", "-", "-", "-", "-", "-", "-", "-", 106, "-", "-",
               12, 48, 45, 31, 21, 36, 35, 124, "-", 3, 40, 28, "-", 10)

d2$`2020` <- c(798, "-", "-", "-", "-", "-", "-", "-", "-", "-", 107, "-", "-",
               13, 50, 46, 33, 22, 37, 36, 124, "-", 4, 42, 29, "-", 11)

d2$`2021` <- c(798, "-", "-", "-", "-", "-", "-", "-", "-", "-", 106, "-", "-",
               13, 50, 46, 33, 22, 37, 36, 124, "-", 4, 42, 29, "-", 11)

d2$`2022` <- c(801, "-", "-", "-", "-", "-", "-", "-", "-", "-", 110, "-", "-",
               13, 50, 46, 33, 22, 38, 37, 125, "-", 4, 43, 30, "-", 11)

d2$`2023` <- c(813, "-", "-", "-", "-", "-", "-", "-", "-", "-", 115, "-", "-",
               12, 52, 49, 34, 23, 39, 38, 128, "-", 3, 42, 31, "-", 10)

d3$`2019` <- c(2543, 105, "-", "-", "-", "-", "-", "-", 54, "-", 1013, "-", "-",
               376, 7, "-", "-", 34, 52, 91, 8, 1, 32, "-", 33, 19, 6, 34)

d3$`2020` <- c(2576, 109, "-", "-", "-", "-", "-", "-", 55, "-", 1062, "-", "-",
               374, 8, "-", "-", 36, 54, 93, 9, 1, 33, "-", 33, 20, 6, 35)

d3$`2021` <- c(2578, 110, "-", "-", "-", "-", "-", "-", 56, "-", 1072, "-", "-",
               377, 8, "-", "-", 37, 55, 94, 9, 1, 33, "-", 34, 20, 6, 35)

d3$`2022` <- c(2584, 111, "-", "-", "-", "-", "-", "-", 56, "-", 1076, "-", "-",
               380, 7, "-", "-", 37, 55, 94, 8, 1, 34, "-", 34, 20, 6, 35)

d3$`2023` <- c(2610, 114, "-", "-", "-", "-", "-", "-", 58, "-", 1100, "-", "-",
               385, 8, "-", "-", 38, 56, 96, 9, 1, 35, "-", 36, 21, 7, 36)

d4$`2019` <- c(438520, 123536, 1850, 2715, 8257, 5852, 74699, 3242, 24962, 11944,
               50642, 11391, 48062, 31987, 14851, 37242, 84482, 67845, 174828,
               76277, 137867, 82839, 70350, 48590, 151680, 19710, 177888)

d4$`2020` <- c(445035, 126889, 1909, 2812, 8586, 6050, 75656, 3371, 25742, 12388,
               52251, 11851, 50674, 32688, 15239, 38343, 86535, 68286, 177313,
               77185, 139855, 84944, 71511, 50798, 153116, 20252, 180458)

d4$`2021` <- c(446362, 126879, 1907, 2825, 8647, 6067, 75702, 3383, 25780, 12358,
               52355, 11870, 50759, 32696, 15333, 38373, 86747, 68238, 177602,
               77206, 140020, 84984, 71531, 50830, 153264, 20291, 180671)

d4$`2022` <- c(447230, 126051, 1931, 2827, 8595, 6105, 75772, 3364, 25725, 12438,
               52276, 11863, 50241, 32825, 15172, 38323, 86175, 68489, 177312,
               77089, 139506, 84872, 71488, 50776, 153544, 20237, 180942)

d4$`2023` <- c(454590, 129438, 1983, 2945, 8763, 6230, 76467, 3402, 26182, 12729,
               53010, 12248, 51557, 33419, 15798, 38857, 88791, 70129, 179389,
               77577, 142321, 86089, 73749, 51983, 155429, 20714, 184607)

d5$`2019` <- c(774801, 7463, 2752, 305, 1284, 1389, 75189, 3955, 25708, 10669,
               84867, 23837, 45021, 22799, 10222, 19230, 19127, 8837, 10648,
               31761, 55851, 31243, 35691, 46372, 7446, 73948, 44746)

d5$`2020` <- c(773307, 7465, 2767, 314, 1312, 1423, 75685, 3952, 25668, 10799,
               85879, 24229, 45674, 23014, 10471, 19972, 19860, 8896, 10792,
               32410, 55393, 31817, 36079, 47197, 7592, 74546, 45490)

d5$`2021` <- c(773526, 7463, 2768, 314, 1313, 1413, 75665, 3957, 25668, 10796,
               85854, 24218, 45674, 23014, 10470, 19972, 19861, 8899, 10793,
               32410, 55369, 31823, 36081, 47189, 7585, 74537, 45489)

d5$`2022` <- c(773817, 7481, 2765, 312, 1307, 1430, 75702, 3955, 25662, 10790, 
               85685, 24300, 45384, 23015, 10496, 19886, 19782, 8952, 10745,
               32051, 55273, 31260, 35874, 46744, 7584, 74739, 45317)

d5$`2023` <- c(802876, 7598, 2778, 307, 1313, 1412, 77194, 3987, 26757, 10926,
               87329, 24134, 46085, 23357, 10675, 20083, 19973, 8882, 10814,
               32531, 56177, 31845, 36201, 47578, 7633, 75523, 46147)

d6$`2023` <- c("-", 5, 24, 25, 20, 26, 11, 27, 23, 2, 9, 17, 8, 19, 15, 10,
               16, 22, 5, 1, 2, 14, 3, 1, 3, 4, 7, 6, 4, 18, 13, 12, 21)


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
