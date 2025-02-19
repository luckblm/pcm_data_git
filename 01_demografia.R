#DEMOGRAFIA----

##INDICADORES----

#TERRITÓRIO----
##01 - População Total - Estimativas Populacionais e Censo Demográfico----
d1 <- bases
d1$tematica <- "Demografia"
d1$indicador <- "População Total - Estimativas Populacionais e Censo Demográfico, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "População Total"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Demografia"
d2$indicador <- "População Total - Estimativas Populacionais e Censo Demográfico, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"

# Criando vetores com os valores de cada coluna

d1$`2020` <- c(211755692, 18672591, 1796460, 894470, 4207714, 631181, 8690745, 
              861773, 1590248, 57374243, 7114598, 3281480, 9187103, 3534165, 
              4039277, 9616621, 3351543, 2318822, 14930634, 89012240, 21292666, 
              4064052, 17366189, 46289333, 30192315, 11516840, 7252502, 
              11422973, 16504303, 2809394, 3526220, 7113540, 3055149)

d1$`2021` <- c(213317639, 18906962, 1815278, 906876, 4269995, 652713, 8777124, 
              877613, 1607363, 57667842, 7153262, 3289290, 9240580, 3560903, 
              4059905, 9674793, 3365351, 2338474, 14985284, 89632912, 21411923, 
              4108508, 17463349, 46649132, 30402587, 11597484, 7338473, 
              11466630, 16707336, 2839188, 3567234, 7206589, 3094325)

d1$`2022` <- c(203080756, 17355778, 1581196, 830018, 3941613, 636707, 8121025, 
              733759, 1511460, 54657621, 6775805, 3271199, 8794957, 3302729, 
              3974687, 9058931, 3127683, 2210004, 14141626, 84840113, 20539989, 
              3833712, 16055174, 44411238, 29937706, 11444380, 7610361, 
              10882965, 16289538, 2757013, 3658649, 7056495, 2817381)

d1$`2023` <- c(211695158, 18535113, 1740255, 876582, 4240571, 695270, 8616120, 
              799124, 1567191, 56970423, 7003234, 3365881, 9196672, 3436278, 
              4124468, 9514483, 3218607, 2281994, 14828806, 88387852, 21247401, 
              4076068, 17213813, 45850570, 30903366, 11753862, 7927212, 
              11222292, 16898404, 2878009, 3778389, 7274463, 2967543)

d1$`2024` <- c(212583750, 18669345, 1746227, 880631, 4281209, 716793, 8664306, 
              802837, 1577342, 57112096, 7010960, 3375646, 9233656, 3446071, 
              4145040, 9539029, 3220104, 2291077, 14850513, 88617693, 21322691, 
              4102129, 17219679, 45973194, 31113021, 11824665, 8058441, 
              11229915, 17071595, 2901895, 3836399, 7350483, 2982818)

d2$`2024` <- c(NA, 4, 23, 25, 13, 27, 9, 26, 24, 2, 12, 18, 8, 17, 14, 7, 19, 
                  22, 4, 1, 2, 15, 3, 1, 3, 5, 10, 6, 5, 21, 16, 11, 20)

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(wb, sheet = nome, cols = 1:ncol(dados_lista[[nome]]), widths = "auto")
}
# Salvar o arquivo Excel
saveWorkbook(wb, "Demografia/01 - População Total - Estimativas Populacionais e Censo Demográfico.xlsx", overwrite = TRUE)


##02 - Densidade Demográfica----
d1 <- bases
d1$tematica <- "Demografia"
d1$indicador <- "Densidade Demográfica (População/km²), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Densidade Demográfica"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Demografia"
d2$indicador <- "Densidade Demográfica (População/km²), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"

d1$`2020` = c(24.88, 4.85, 7.56, 5.45, 2.70, 2.82, 6.98, 6.05, 5.73, 36.96, 21.58, 13.03, 61.70, 66.92, 71.53, 98.06, 120.37, 105.76, 26.44, 96.27, 36.30, 88.21, 396.94, 186.49, 52.35, 57.79, 75.76, 40.55, 10.27, 7.87, 3.90, 20.91, 530.34)
d1$`2021` = c(25.07, 4.91, 7.63, 5.52, 2.74, 2.92, 7.04, 6.16, 5.79, 37.15, 21.70, 13.07, 62.06, 67.43, 71.90, 98.65, 120.92, 106.59, 26.53, 96.95, 36.51, 89.17, 399.16, 187.94, 52.71, 58.19, 76.66, 40.70, 10.40, 7.95, 3.95, 21.18, 537.14)
d1$`2022` = c(23.86, 4.51, 6.65, 5.06, 2.53, 2.85, 6.52, 5.15, 5.45, 35.21, 20.55, 12.99, 59.07, 62.54, 70.39, 92.37, 112.38, 100.74, 25.04, 91.76, 35.02, 83.21, 366.97, 178.92, 51.91, 57.42, 79.50, 38.63, 10.14, 7.72, 4.05, 20.74, 489.06)
d1$`2023` = c(24.87, 4.81, 7.32, 5.34, 2.72, 3.11, 6.92, 5.61, 5.65, 36.70, 21.24, 13.37, 61.77, 65.07, 73.04, 97.02, 115.65, 104.02, 26.26, 95.60, 36.23, 88.47, 393.45, 184.72, 53.58, 58.98, 82.81, 39.84, 10.52, 8.06, 4.18, 21.38, 515.13)
d1$`2024` = c(24.98, 4.85, 7.34, 5.36, 2.75, 3.21, 6.95, 5.64, 5.69, 36.79, 21.27, 13.41, 62.01, 65.25, 73.41, 97.27, 115.70, 104.43, 26.30, 95.85, 36.35, 89.03, 393.59, 185.21, 53.95, 59.33, 84.18, 39.86, 10.63, 8.13, 4.25, 21.60, 517.78)
d2$`2024` = c(NA, 5, 20, 24, 27, 26, 21, 23, 22, 3, 17, 18, 11, 10, 9, 6, 4, 5, 15, 1, 14, 7, 2, 3, 2, 12, 8, 13, 4, 19, 25, 16, 1)




# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(wb, sheet = nome, cols = 1:ncol(dados_lista[[nome]]), widths = "auto")
}
# Salvar o arquivo Excel
saveWorkbook(wb, "Demografia/02 - Densidade Demográfica.xlsx", overwrite = TRUE)

##03 - População por Faixa Etária----

d1 <- bases
d1$tematica <- "Demografia"
d1$indicador <- "População por Faixa Etária, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"
d1$`2023` <- c(211695158, 18535113, 1740255, 876582, 4240571, 695270, 8616120, 799124, 1567191, 
           56970423, 7003234, 3365881, 9196672, 3436278, 4124468, 9514483, 3218607, 2281994, 
           14828806, 88387852, 21247401, 4076068, 17213813, 45850570, 30903366, 11753862, 
           7927212, 11222292, 16898404, 2878009, 3778389, 7274463, 2967543) 


d2 <- bases
d2$tematica <- "Demografia"
d2$indicador <- "População por Faixa Etária, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking "
d2$categoria2 <- "-"
d2$`2023` <- c(NA, 4, 23, 25, 13, 27, 9, 26, 24, 2, 12, 18, 8, 17, 14, 7, 19, 22, 4, 1, 
                  2, 15, 3, 1, 3, 5, 10, 6, 5, 21, 16, 11, 20)

d3 <- bases
d3$tematica <- "Demografia"
d3$indicador <- "População por Faixa Etária, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d3$categoria1 <- "0 a 4 anos"
d3$categoria2 <- "-"
d3$`2023` <- c(13421242, 1489614, 125954, 74874, 372015, 71276, 655615, 72115, 117765, 
               3762280, 520916, 222451, 602199, 214008, 273021, 625091, 235513, 151763, 
               917318, 5141414, 1220755, 266281, 958493, 2695885, 1865211, 729604, 504102, 
               631505, 1162723, 210411, 294993, 468790, 188529)


d4 <- bases
d4$tematica <- "Demografia"
d4$indicador <- "População por Faixa Etária, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d4$categoria1 <- "5 a 9 anos"
d4$categoria2 <- "-"
d4$`2023` <- c(14589745, 1560194, 132328, 77678, 392192, 67205, 692423, 75213, 123155, 
               4043311, 554149, 233603, 643840, 233897, 288266, 679123, 242865, 165043, 
               1002525, 5727296, 1314435, 280280, 1104521, 3028060, 2027087, 800697, 
               524343, 702047, 1231857, 223089, 295377, 508749, 204642) 

d5 <- bases
d5$tematica <- "Demografia"
d5$indicador <- "População por Faixa Etária, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d5$categoria1 <- "10 a 14 anos"
d5$categoria2 <- "-"
d5$`2023` <- c(14515508, 1572129, 133776, 78189, 389084, 62094, 715008, 71835, 122143, 
                 4165609, 591098, 242561, 645856, 239849, 295180, 697478, 250831, 168693, 
                 1034063, 5649525, 1314332, 272178, 1062671, 3000344, 1947317, 778498, 
                 497960, 670859, 1180928, 215472, 276694, 491520, 197242)


d6 <- bases
d6$tematica <- "Demografia"
d6$indicador <- "População por Faixa Etária, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d6$categoria1 <- "15 a 19 anos"
d6$categoria2 <- "-"
d6$`2023` <- c(15169410, 1644083, 136218, 85160, 386961, 60932, 777258, 72329, 125225, 
                 4416907, 635544, 267357, 684837, 251611, 313251, 732710, 257900, 176640, 
                 1097057, 5855307, 1428101, 281132, 1086141, 3059933, 2019965, 810771, 
                 499345, 709849, 1233148, 215078, 282242, 522555, 213273)

d7 <- bases
d7$tematica <- "Demografia"
d7$indicador <- "População por Faixa Etária, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d7$categoria1 <- "20 a 29 anos"
d7$categoria2 <- "-"
d7$`2023` <- c(32576988, 3227220, 282650, 160340, 755588, 125300, 1503683, 145136, 
                 254523, 9017667, 1168515, 526101, 1493427, 533505, 634680, 1493592, 
                 523567, 366327, 2277953, 13031364, 3191433, 600934, 2550335, 6688662, 
                 4615648, 1782100, 1221871, 1611677, 2685089, 446208, 608940, 1157999, 
                 471942)

d8 <- bases
d8$tematica <- "Demografia"
d8$indicador <- "População por Faixa Etária, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d8$categoria1 <- "30 a 39 anos"
d8$categoria2 <- "-"
d8$`2023` <- c(32893858, 2912922, 277425, 131873, 651675, 110389, 1370033, 125320, 
                 246207, 8954501, 1103639, 524028, 1494141, 554819, 635296, 1464871, 
                 487736, 363123, 2326848, 13568591, 3261147, 632400, 2565412, 7109632, 
                 4750495, 1769744, 1326454, 1654297, 2707349, 442318, 614204, 1169132, 
                 481695) 

d9 <- bases
d9$tematica <- "Demografia"
d9$indicador <- "População por Faixa Etária, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d9$categoria1 <- "40 a 49 anos"
d9$categoria2 <- "-"
d9$`2023` <- c(30910229, 2498990, 251492, 113013, 552976, 86590, 1170477, 103121, 
                 221321, 8176061, 927058, 476290, 1277326, 494690, 586038, 1373950, 
                 451312, 331314, 2258083, 13287317, 3136416, 609155, 2563052, 6978694, 
                 4415914, 1678460, 1161828, 1575626, 2531947, 401546, 547664, 1100043, 
                 482694)

d10 <- bases
d10$tematica <- "Demografia"
d10$indicador <- "População por Faixa Etária, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d10$categoria1 <- "50 a 59 anos"
d10$categoria2 <- "-"
d10$`2023` <- c(24636687, 1721934, 191579, 74516, 361554, 56050, 808484, 68464, 
                 161287, 6260526, 657100, 367466, 1018096, 398741, 466048, 1068789, 
                 346606, 255493, 1682187, 10885393, 2614508, 479609, 2164129, 5627147, 
                 3840960, 1468252, 952546, 1420162, 1927874, 321217, 410299, 851055, 
                 345303)

d11 <- bases
d11$tematica <- "Demografia"
d11$indicador <- "População por Faixa Etária, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d11$categoria1 <- "60 a 69 anos"
d11$categoria2 <- "-"
d11$`2023` <- c(18291312, 1115032, 125860, 46948, 227800, 34163, 532167, 39884, 
                 108210, 4391822, 459886, 266824, 703339, 276164, 327086, 748934, 
                 236034, 168893, 1204662, 8455026, 2061803, 372988, 1744819, 4275416, 
                 3033679, 1090058, 716334, 1227287, 1295753, 230540, 268022, 579282, 
                 217909) 

d12 <- bases
d12$tematica <- "Demografia"
d12$indicador <- "População por Faixa Etária, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d12$categoria1 <- "70 a 79 anos"
d12$categoria2 <- "-"
d12$`2023` <- c(10117249, 557837, 59503, 23973, 107114, 15700, 273703, 18202, 
                 59642, 2540180, 264280, 159648, 417153, 157728, 200088, 428079, 
                 131344, 92703, 689157, 4673899, 1154592, 193148, 976878, 2349281, 
                 1681594, 597279, 374282, 710033, 663739, 120099, 129520, 299260, 
                 114860)

d13 <- bases
d13$tematica <- "Demografia"
d13$indicador <- "População por Faixa Etária, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d13$categoria1 <- "80 anos e mais"
d13$categoria2 <- "-"
d13$`2023` <- c(4572930, 235158, 23470, 10018, 43612, 5571, 117269, 7505, 
                   27713, 1241559, 121049, 79552, 216458, 81266, 105514, 201866, 
                   54899, 42002, 338953, 2112720, 549879, 87963, 437362, 1037516, 
                   705496, 248399, 148147, 308950, 277997, 52031, 50434, 126078, 
                   49454)


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2, "d3" = d3, "d4" = d4 , "d5" = d5 , "d6" = d6, "d7" = d7,
                    "d8" = d8, "d9" = d9, "d10" = d10 , "d11" = d11 , "d12" = d12,  "d13" = d13)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  
  # Ajustar automaticamente a largura das colunas
  setColWidths(wb, sheet = nome, cols = 1:ncol(dados_lista[[nome]]), widths = "auto")
}
# Salvar o arquivo Excel
saveWorkbook(wb, "Demografia/03 - População por Faixa Etária.xlsx", overwrite = TRUE)

##04 - Número de Eleitores----
d1 <- bases
linha_extra <- tibble::tibble(
  tematica = "Demografia",
  subtematica = "-",
  indicador = "Número de Eleitores, Segundo Brasil, Grandes Regiões e Unidades da Federação",
  regiao = "-",
  localidade = "Exterior",
  categoria1 = "Número de Eleitores",
  categoria2 = "-"
)
d1$tematica <- "Demografia"
d1$indicador <- "Número de Eleitores, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Número de Eleitores"
d1$categoria2 <- "-"
d1 <- rbind(d1,linha_extra)

# Criando vetores para cada coluna
# Vetor com o número de eleitores para os anos de 2020 a 2024
d1$`2020` <- c(145958233, 11488004, 1156175, 548013, 2442675, 342297, 5450710, 516027, 1032107, 
            39238502, 4552310, 2451974, 6206078, 2439327, 2960294, 6711341, 2212605, 1604212, 
            10100361, 62505545, 15768482, 2767545, 12394670, 31574848, 21502285, 7988648, 
            5146976, 8366661, 10713941, 1833683, 2207355, 4591079, 2081824, 509956)

d1$`2021` <- c(146765823, 11655466, 1168877, 554117, 2452719, 345641, 5572505, 522363, 1039244, 
            39390390, 4549506, 2466415, 6263279, 2456411, 2976620, 6621632, 2233890, 1616583, 
            10206054, 62658032, 15402847, 2800578, 12488650, 31965957, 21609792, 8075138, 
            5138326, 8396328, 10885926, 1851498, 2241835, 4681995, 2110598, 566217)

d1$`2022` <- c(156210885, 12554727, 1229800, 588613, 2650238, 366033, 6077760, 549983, 1092300, 
            42328372, 5040354, 2570024, 6808321, 2550568, 3086948, 7009662, 2322857, 1668853, 
            11270785, 66585876, 16257859, 2920666, 12800588, 34606763, 22511302, 8458085, 
            5484573, 8568644, 11533682, 1993695, 2468810, 4869125, 2202052, 696926)

d1$`2023` <- c(155387262, 12518893, 1231425, 590161, 2653878, 372582, 5996658, 554827, 1119362, 
            41952664, 4964093, 2600066, 6724701, 2575451, 3121141, 7006216, 2363622, 1684641, 
            10912733, 65979462, 16128128, 2935608, 12885916, 34029810, 22557843, 8466190, 
            5504424, 8587229, 11632724, 1979549, 2484830, 4962143, 2206202, 745676)

d1$`2024` <- c(155912680, 12987166, 1266546, 612448, 2749346, 389863, 6226373, 571248, 1171342, 
            43302692, 5180738, 2698764, 6935539, 2649282, 3225312, 7152871, 2442894, 1733785, 
            11283507, 66906335, 16469155, 2999642, 13033929, 34403609, 22969108, 8645891, 
            5640659, 8682558, 9747379, 2032487, 2588457, 5126435, NA, NA)


d2 <- bases
linha_extra <- tibble::tibble(
  tematica = "Demografia",
  subtematica = "-",
  indicador = "Número de Eleitores, Segundo Brasil, Grandes Regiões e Unidades da Federação",
  regiao = "-",
  localidade = "Exterior",
  categoria1 = "Ranking",
  categoria2 = "-"
)
d2$tematica <- "Demografia"
d2$indicador <- "Número de Eleitores, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"
d2 <- rbind(d2,linha_extra)

# Vetor com o ranking de 2024
d2$`2024` <- c(NA, 4, 22, 24, 15, 26, 9, 25, 23, 2, 
                    11, 16, 8, 17, 13, 7, 19, 21, 4, 1, 
                    2, 14, 3, 1, 3, 6, 10, 5, 5, 20, 
                    18, 12, NA, NA)


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  
  # Ajustar automaticamente a largura das colunas
  setColWidths(wb, sheet = nome, cols = 1:ncol(dados_lista[[nome]]), widths = "auto")
}
# Salvar o arquivo Excel
saveWorkbook(wb, "Demografia/04 - Número de Eleitores.xlsx", overwrite = TRUE)

##05 - Expectativa de Vida ao Nascer (em anos)----
d1 <- bases
d1$tematica <- "Demografia"
d1$indicador <- "Expectativa de Vida ao Nascer (em anos), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Expectativa de Vida ao Nascer (em anos)"
d1$categoria2 <- "-"

# Criando vetores para cada coluna
d1$`2019` <- c(76.22, 75.63, 76.57, 74.83, 75.34, 74.12, 75.69, 73.95, 76.70, 75.84, 
            75.97, 76.71, 76.78, 76.24, 76.20, 75.17, 74.39, 75.49, 75.63, 76.42, 
            77.16, 77.09, 74.49, 76.81, 76.82, 76.63, 77.83, 76.43, 76.51, 76.05, 
            75.91, 76.02, 79.06)

d1$`2020` <- c(74.81, 73.04, 74.15, 72.84, 71.81, 71.42, 73.30, 70.69, 75.24, 74.07, 
            73.86, 75.29, 74.19, 75.10, 74.88, 73.13, 72.50, 73.77, 74.37, 75.14, 
            76.67, 75.09, 72.56, 75.55, 76.32, 75.82, 77.23, 76.27, 74.67, 74.78, 
            73.43, 74.46, 76.69)

d1$`2021` <- c(72.78, 71.70, 70.54, 71.23, 69.72, 69.07, 73.16, 70.45, 72.52, 73.23, 
            73.18, 74.07, 73.54, 74.10, 73.39, 72.46, 71.98, 73.30, 73.40, 72.90, 
            73.98, 73.75, 71.49, 72.90, 73.05, 71.91, 74.52, 73.29, 72.01, 71.34, 
            70.96, 71.68, 74.95)

d1$`2022` <- c(75.45, 74.97, 74.94, 75.20, 74.82, 73.32, 75.19, 73.64, 75.40, 74.92, 
            74.82, 75.48, 75.88, 75.96, 74.91, 74.29, 73.39, 75.27, 74.74, 75.83, 
            76.51, 76.31, 74.74, 75.93, 75.76, 75.38, 76.92, 75.44, 75.70, 74.53, 
            74.91, 75.45, 78.78)

d1$`2023` <- c(76.39, 75.67, 75.77, 75.63, 75.24, 74.09, 76.04, 73.76, 76.47, 75.92, 
            75.38, 76.76, 77.07, 77.65, 76.62, 75.29, 74.13, 76.14, 75.58, 76.65, 
            77.33, 76.99, 75.39, 76.82, 77.12, 76.64, 78.10, 76.97, 76.55, 75.14, 
            75.47, 76.57, 79.60)


d2 <- bases
d2$tematica <- "Demografia"
d2$indicador <- "Expectativa de Vida ao Nascer (em anos), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking "
d2$categoria2 <- "-"

# Vetor com o ranking de 2023
d2$`2023` <- c(NA, 5, 16, 17, 23, 26, 15, 27, 13, 4, 
                    21, 9, 5, 3, 11, 22, 25, 14, 18, 2, 
                    4, 6, 20, 8, 1, 10, 2, 7, 3, 24, 
                    19, 12, 1)

# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  
  # Ajustar automaticamente a largura das colunas
  setColWidths(wb, sheet = nome, cols = 1:ncol(dados_lista[[nome]]), widths = "auto")
}
# Salvar o arquivo Excel
saveWorkbook(wb, "Demografia/05 - Expectativa de Vida ao Nascer (em anos).xlsx", overwrite = TRUE)

##06 - Taxa de Fecundidade Total----
d1 <- bases
d1$tematica <- "Demografia"
d1$indicador <- "Taxa de Fecundidade Total, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Taxa"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Demografia"
d2$indicador <- "Taxa de Fecundidade Total, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking "
d2$categoria2 <- "-"

# Criando vetores para cada coluna
d1$`2019` <- c(1.73, 2.04, 1.89, 2.20, 2.22, 2.80, 1.90, 2.28, 1.96, 1.75, 
            1.92, 1.77, 1.72, 1.67, 1.78, 1.75, 1.86, 1.72, 1.65, 1.63, 
            1.58, 1.77, 1.58, 1.66, 1.67, 1.73, 1.65, 1.62, 1.83, 2.00, 
            2.04, 1.70, 1.70)

d1$`2020` <- c(1.66, 1.96, 1.82, 2.05, 2.15, 2.51, 1.83, 2.18, 1.90, 1.68, 
            1.81, 1.69, 1.63, 1.59, 1.74, 1.70, 1.81, 1.67, 1.59, 1.57, 
            1.53, 1.74, 1.52, 1.59, 1.62, 1.65, 1.63, 1.58, 1.75, 1.92, 
            1.97, 1.64, 1.59)

d1$`2021` <- c(1.64, 2.00, 1.81, 2.09, 2.21, 2.49, 1.88, 2.23, 1.90, 1.68, 
            1.86, 1.72, 1.61, 1.60, 1.73, 1.68, 1.84, 1.66, 1.58, 1.51, 
            1.51, 1.70, 1.46, 1.52, 1.57, 1.61, 1.59, 1.52, 1.73, 1.92, 
            1.98, 1.60, 1.55)

d1$`2022` <- c(1.63, 1.96, 1.96, 2.01, 2.12, 2.35, 1.84, 2.10, 1.85, 1.63, 
            1.73, 1.65, 1.57, 1.53, 1.66, 1.64, 1.78, 1.57, 1.57, 1.54, 
            1.52, 1.79, 1.52, 1.53, 1.60, 1.63, 1.64, 1.54, 1.72, 1.90, 
            1.98, 1.60, 1.54)

d1$`2023` <- c(1.57, 1.83, 1.73, 1.91, 1.96, 2.26, 1.73, 1.94, 1.87, 1.56, 
            1.68, 1.61, 1.51, 1.48, 1.62, 1.56, 1.79, 1.57, 1.47, 1.48, 
            1.48, 1.72, 1.39, 1.48, 1.56, 1.59, 1.58, 1.51, 1.71, 1.85, 
            1.98, 1.62, 1.47)

d2$`2023` <- c(NA, 1, 10, 5, 3, 1, 9, 4, 6, 3, 
                    12, 15, 20, 24, 14, 19, 8, 18, 26, 5, 
                    23, 11, 27, 22, 4, 16, 17, 21, 2, 7, 
                    2, 13, 25)


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  
  # Ajustar automaticamente a largura das colunas
  setColWidths(wb, sheet = nome, cols = 1:ncol(dados_lista[[nome]]), widths = "auto")
}
# Salvar o arquivo Excel
saveWorkbook(wb, "Demografia/06 - Taxa de Fecundidade Total.xlsx", overwrite = TRUE)

##07 - Razão de Dependência----

d1 <- bases
d1$tematica <- "Demografia"
d1$indicador <- "Razão de Dependência, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Razão de Dependência"
d1$categoria2 <- "-"
 
 #Criando vetores para cada coluna
d1$`2019` <- c(44.44, 48.98, 43.72, 52.13, 52.08, 52.48, 48.16, 50.95, 47.45, 46.19, 
            51.12, 47.04, 45.45, 44.45, 47.34, 46.18, 47.17, 44.63, 44.36, 43.16, 
            42.71, 43.42, 43.49, 43.21, 43.28, 43.76, 40.47, 44.70, 42.71, 46.65, 
            43.68, 41.53, 40.75)

d1$`2020` <- c(44.46, 48.34, 43.54, 51.04, 51.41, 52.18, 47.37, 50.23, 47.26, 45.86, 
            50.27, 46.65, 45.24, 44.22, 47.00, 45.88, 46.89, 44.41, 44.13, 43.39, 
            42.90, 43.73, 43.83, 43.43, 43.66, 44.00, 41.00, 45.16, 42.74, 46.72, 
            43.87, 41.53, 40.60)

d1$`2021` <- c(44.44, 47.72, 43.42, 50.00, 50.73, 52.10, 46.60, 49.73, 46.89, 45.50, 
            49.23, 46.26, 44.91, 43.93, 46.69, 45.57, 46.63, 44.11, 43.96, 43.60, 
            43.11, 44.07, 44.10, 43.60, 43.98, 44.16, 41.41, 45.61, 42.68, 46.75, 
            43.86, 41.50, 40.33)

d1$`2022` <- c(44.50, 47.18, 43.45, 49.09, 50.07, 51.91, 45.93, 49.30, 46.64, 45.21, 
            48.32, 45.92, 44.64, 43.75, 46.45, 45.34, 46.47, 43.84, 43.85, 43.86, 
            43.43, 44.53, 44.43, 43.80, 44.38, 44.42, 41.88, 46.14, 42.77, 46.94, 
            44.06, 41.60, 40.14)

d1$`2023` <- c(44.62, 46.66, 43.59, 48.18, 49.36, 51.48, 45.31, 48.71, 46.53, 44.97, 
            47.49, 45.70, 44.43, 43.61, 46.27, 45.15, 46.30, 43.61, 43.75, 44.21, 
            43.86, 45.08, 44.83, 44.06, 44.91, 44.80, 42.42, 46.83, 42.95, 47.12, 
            44.37, 41.84, 40.02)


d2 <- bases
d2$tematica <- "Demografia"
d2$indicador <- "Razão de Dependência, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"

#Vetor com o ranking de 2023
d2$`2023` <- c(NA, 1, 24, 4, 2, 1, 12, 3, 8, 2, 
                    5, 11, 17, 22, 10, 13, 9, 23, 21, 4, 
                    20, 14, 15, 19, 3, 16, 25, 7, 5, 6, 
                    18, 26, 27)


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(wb, sheet = nome, cols = 1:ncol(dados_lista[[nome]]), widths = "auto")
}
# Salvar o arquivo Excel
saveWorkbook(wb, "Demografia/07 - Razão de Dependência.xlsx", overwrite = TRUE)

##08 - Índice de Envelhecimento----
d1 <- bases
d1$tematica <- "Demografia"
d1$indicador <- "Índice de Envelhecimento, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Índice de Envelhecimento"
d1$categoria2 <- "-"

# Criando vetores para cada coluna
d1$`2019` <- c(44.97, 22.48, 28.08, 18.90, 17.80, 15.25, 24.20, 15.66, 31.18, 40.10, 
            29.15, 42.63, 42.88, 43.78, 45.86, 40.56, 33.13, 35.07, 43.84, 53.62, 
            56.63, 45.39, 59.90, 50.79, 54.65, 48.50, 47.66, 66.48, 35.25, 36.22, 
            29.25, 38.69, 34.12)

d1$`2020` <- c(47.01, 23.64, 29.76, 19.94, 18.61, 15.83, 25.55, 16.49, 32.59, 41.81, 
            30.53, 44.65, 44.35, 45.49, 47.43, 42.21, 34.69, 36.74, 45.88, 56.09, 
            59.54, 47.66, 62.25, 53.13, 57.04, 50.77, 49.41, 69.62, 36.86, 37.70, 
            30.36, 40.46, 36.25)

d1$`2021` <- c(48.97, 24.73, 31.23, 20.94, 19.30, 16.37, 26.87, 17.26, 33.95, 43.50, 
            31.95, 46.61, 45.80, 47.33, 48.90, 43.78, 36.18, 38.45, 47.91, 58.45, 
            62.29, 49.75, 64.46, 55.40, 59.28, 52.89, 51.10, 72.56, 38.39, 39.03, 
            31.42, 42.14, 38.38)

d1$`2022` <- c(51.07, 26.00, 32.79, 22.08, 20.22, 16.98, 28.40, 18.21, 35.37, 45.36, 
            33.51, 48.68, 47.45, 49.28, 50.43, 45.59, 37.77, 40.40, 50.12, 60.99, 
            65.07, 52.03, 67.14, 57.83, 61.43, 54.97, 52.73, 75.34, 40.02, 40.34, 
            32.55, 43.93, 40.79)

d1$`2023` <- c(53.60, 27.59, 34.75, 23.52, 21.51, 17.80, 30.23, 19.43, 37.06, 47.62, 
            35.34, 51.11, 49.53, 51.72, 52.34, 47.84, 39.68, 42.73, 52.77, 64.04, 
            68.27, 54.70, 70.48, 60.74, 64.00, 57.51, 54.70, 78.57, 41.99, 42.04, 
            33.93, 46.06, 43.66)

d2 <- bases
d2$tematica <- "Demografia"
d2$indicador <- "Índice de Envelhecimento, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d2$categoria1 <- "Ranking"
d2$categoria2 <- "-"

#Vetor com o ranking 2023
d2$`2023` <- c(NA, 5, 21, 24, 25, 27, 23, 26, 19, 3, 
                    20, 11, 12, 10, 9, 13, 18, 16, 8, 1, 
                    3, 6, 2, 4, 2, 5, 7, 1, 4, 17, 
                    22, 14, 15)


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list("d1" = d1, "d2" = d2)
# Loop para adicionar cada data frame ao Excel
for (nome in names(dados_lista)) {
  addWorksheet(wb, nome)  # Criar aba
  writeData(wb, nome, dados_lista[[nome]])  # Escrever os dados
  # Ajustar automaticamente a largura das colunas
  setColWidths(wb, sheet = nome, cols = 1:ncol(dados_lista[[nome]]), widths = "auto")
}
# Salvar o arquivo Excel
saveWorkbook(wb, "Demografia/08 - Índice de Envelhecimento.xlsx", overwrite = TRUE)