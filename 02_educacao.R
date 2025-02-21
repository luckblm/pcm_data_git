#EDUCAÇÃO----
##EDUCAÇÃO BÁSICA----

bases$subtematica <- "Educação Básica"

###01 - Número de Matrículas na Creche - Ensino Regular por Dependência Administrativa----

######### MATRCULAS TOTAL ###########
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Matrículas na Creche - Ensino Regular por Dependência Administrativa"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

# Matrícula Total por ano
d1$`2019` <- c(3755092, 3651989, 1279242, 112092, 32912, 6961, 80708, 4686, 26578, 876270, 134344, 54016, 182672, 59377, 60062, 109589, 49720, 27613, 198877, 1818540, 330436, 73787, 257506, 1156811, 626649)
d1$`2020` <- c(3651989, 1279242, 112092, 32912, 6961, 80708, 4686, 26578, 876270, 134344, 54016, 182672, 59377, 60062, 109589, 49720, 27613, 198877, 1818540, 330436, 73787, 257506, 1156811, 626649)
d1$`2021` <- c(1279242, 130956, 4371, 4041, 4371, 6710, 74297, 3748, 53105, 10190, 84124, 36629, 53469, 45317, 22397, 14434, 24658, 36585, 125625, 60209, 280423, 48189, 107000, 134580, 121470, 251693, 308506)
d1$`2022` <- c(1269263, 129676, 4754, 5029, 4355, 6935, 77774, 3900, 57263, 10374, 90285, 38163, 55367, 47235, 23046, 14469, 25210, 37618, 128095, 61571, 290495, 48932, 111927, 137290, 120204, 265701, 315527)
d1$`2023` <- c(1282369, 133559, 4810, 5152, 4576, 7042, 78679, 3989, 58552, 10524, 91654, 39142, 56142, 48557, 23567, 14723, 25876, 38069, 130454, 62782, 296135, 49873, 114261, 141275, 122763, 270113, 323126)

######### FEDERAL ##########
d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Matrículas na Creche - Ensino Regular por Dependência Administrativa"
d2$categoria1 <- "Federal"
d2$categoria2 <- "-"

# Número de Matrículas por ano - Federal
d2$`2019` <- c(782, "-", "-", "-", "-", "-", "-", "-", "-", "-", 106, "-", "-", 12, 48, 45, 31, 21, 36, 35, 124, "-", 3, 40, 28, "-", 10)
d2$`2020` <- c(798, "-", "-", "-", "-", "-", "-", "-", "-", "-", 107, "-", "-", 13, 50, 46, 33, 22, 37, 36, 124, "-", 4, 42, 29, "-", 11)
d2$`2021` <- c(798, "-", "-", "-", "-", "-", "-", "-", "-", "-", 106, "-", "-", 13, 50, 46, 33, 22, 37, 36, 124, "-", 4, 42, 29, "-", 11)
d2$`2022` <- c(801, "-", "-", "-", "-", "-", "-", "-", "-", "-", 110, "-", "-", 13, 50, 46, 33, 22, 38, 37, 125, "-", 4, 43, 30, "-", 11)
d2$`2023` <- c(813, "-", "-", "-", "-", "-", "-", "-", "-", "-", 115, "-", "-", 12, 52, 49, 34, 23, 39, 38, 128, "-", 3, 42, 31, "-", 10)

######### ESTADUAL  ############
d3$tematica <- "Educação"
d3$indicador <- "Número de Matrículas na Creche - Ensino Regular por Dependência Administrativa"
d3$categoria1 <- "Estadual"
d3$categoria2 <- "-"

# Número de Matrículas por ano - Estadual
d3$`2019` <- c(2543, 105, "-", "-", "-", "-", "-", "-", 54, "-", 1013, "-", "-", 376, 7, "-", "-", 34, 52, 91, 8, 1, 32, "-", 33, 19, 6, 34)
d3$`2020` <- c(2576, 109, "-", "-", "-", "-", "-", "-", 55, "-", 1062, "-", "-", 374, 8, "-", "-", 36, 54, 93, 9, 1, 33, "-", 33, 20, 6, 35)
d3$`2021` <- c(2578, 110, "-", "-", "-", "-", "-", "-", 56, "-", 1072, "-", "-", 377, 8, "-", "-", 37, 55, 94, 9, 1, 33, "-", 34, 20, 6, 35)
d3$`2022` <- c(2584, 111, "-", "-", "-", "-", "-", "-", 56, "-", 1076, "-", "-", 380, 7, "-", "-", 37, 55, 94, 8, 1, 34, "-", 34, 20, 6, 35)
d3$`2023` <- c(2610, 114, "-", "-", "-", "-", "-", "-", 58, "-", 1100, "-", "-", 385, 8, "-", "-", 38, 56, 96, 9, 1, 35, "-", 36, 21, 7, 36)


######### MUNICIPAL  ############
d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Matrículas na Creche - Ensino Regular por Dependência Administrativa"
d4$categoria1 <- "Municipal"
d4$categoria2 <- "-"

# Número de Matrículas por ano - Municipal
d4$`2019` <- c(438520, 123536, 1850, 2715, 8257, 5852, 74699, 3242, 24962, 11944, 50642, 11391, 48062, 31987, 14851, 37242, 84482, 67845, 174828, 76277, 137867, 82839, 70350, 48590, 151680, 19710, 177888)
d4$`2020` <- c(445035, 126889, 1909, 2812, 8586, 6050, 75656, 3371, 25742, 12388, 52251, 11851, 50674, 32688, 15239, 38343, 86535, 68286, 177313, 77185, 139855, 84944, 71511, 50798, 153116, 20252, 180458)
d4$`2021` <- c(446362, 126879, 1907, 2825, 8647, 6067, 75702, 3383, 25780, 12358, 52355, 11870, 50759, 32696, 15333, 38373, 86747, 68238, 177602, 77206, 140020, 84984, 71531, 50830, 153264, 20291, 180671)
d4$`2022` <- c(447230, 126051, 1931, 2827, 8595, 6105, 75772, 3364, 25725, 12438, 52276, 11863, 50241, 32825, 15172, 38323, 86175, 68489, 177312, 77089, 139506, 84872, 71488, 50776, 153544, 20237, 180942)
d4$`2023` <- c(454590, 129438, 1983, 2945, 8763, 6230, 76467, 3402, 26182, 12729, 53010, 12248, 51557, 33419, 15798, 38857, 88791, 70129, 179389, 77577, 142321, 86089, 73749, 51983, 155429, 20714, 184607)

######### PRIVADA  ############
d5 <- bases
d5$tematica <- "Educação"
d5$indicador <- "Número de Matrículas na Creche - Ensino Regular por Dependência Administrativa"
d5$categoria1 <- "Privada"
d5$categoria2 <- "-"

# Número de Matrículas por ano - Privada
d5$`2019` <- c(774801, 7463, 2752, 305, 1284, 1389, 75189, 3955, 25708, 10669, 84867, 23837, 45021, 22799, 10222, 19230, 19127, 8837, 10648, 31761, 55851, 31243, 35691, 46372, 7446, 73948, 44746)
d5$`2020` <- c(773307, 7465, 2767, 314, 1312, 1423, 75685, 3952, 25668, 10799, 85879, 24229, 45674, 23014, 10471, 19972, 19860, 8896, 10792, 32410, 55393, 31817, 36079, 47197, 7592, 74546, 45490)
d5$`2021` <- c(773526, 7463, 2768, 314, 1313, 1413, 75665, 3957, 25668, 10796, 85854, 24218, 45674, 23014, 10470, 19972, 19861, 8899, 10793, 32410, 55369, 31823, 36081, 47189, 7585, 74537, 45489)
d5$`2022` <- c(773817, 7481, 2765, 312, 1307, 1430, 75702, 3955, 25662, 10790, 85685, 24300, 45384, 23015, 10496, 19886, 19782, 8952, 10745, 32051, 55273, 31260, 35874, 46744, 7584, 74739, 45317)
d5$`2023` <- c(802876, 7598, 2778, 307, 1313, 1412, 77194, 3987, 26757, 10926, 87329, 24134, 46085, 23357, 10675, 20083, 19973, 8882, 10814, 32531, 56177, 31845, 36201, 47578, 7633, 75523, 46147)

######### RANKING 2023  ############
d6 <- bases
d6$tematica <- "Educação"
d6$indicador <- "Número de Matrículas na Creche - Ensino Regular por Dependência Administrativa"
d6$categoria1 <- "Ranking 2023"
d6$categoria2 <- "-"

# Vetor com o Ranking 2023
d6$`2023` <- c("-", 5, 24, 25, 20, 26, 11, 27, 23, 2, 9, 17, 8, 19, 15, 10, 16, 22, 5, 1, 2, 14, 3, 1, 3, 4, 7, 6, 4, 18, 13, 12, 21)

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

######### MATRICULAS TOTAL ###########
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

# Total por ano
d1$`2019` <- c(5217686, 512633, 35782, 27326, 125758, 21103, 239819, 22349, 40496, 1473035, 225507, 93448, 239464, 87867, 99803, 233659, 81899, 57928, 353460, 2100695, 472408, 106225, 380198, 1141864, 720860, 274477, 191697, 254686, 410463, 73334, 104312, 161072, 71745)
d1$`2020` <- c(5177806, 510412, 38184, 26827, 124818, 20931, 236752, 21993, 40907, 1445175, 218353, 92771, 242655, 87381, 96280, 226569, 77705, 55601, 347860, 2079563, 464390, 106512, 374269, 1134392, 723963, 276344, 194149, 253470, 418693, 75425, 105712, 165556, 72000)
d1$`2021` <- c(4902189, 487040, 38097, 24850, 117950, 20327, 225996, 19875, 39945, 1359506, 209723, 89822, 231842, 80260, 92159, 203874, 74470, 50649, 326707, 1958899, 445727, 99793, 343758, 1069621, 708102, 276973, 192689, 238440, 388642, 70874, 97957, 152942, 66869)
d1$`2022` <- c(5093075, 504262, 41719, 24881, 120914, 20929, 232472, 20597, 42750, 1417424, 211618, 89732, 237032, 82093, 96475, 215238, 78646, 55464, 351126, 2022309, 470802, 102682, 359987, 1088838, 737382, 289331, 202296, 245755, 411698, 75679, 103605, 163415, 68999)
d1$`2023` <- c(5338282, 539499, 45748, 26989, 127415, 23358, 246717, 22433, 46839, 1493682, 222531, 94139, 248087, 86503, 103443, 233012, 84669, 58803, 362495, 2101212, 498721, 108351, 370857, 1123283, 763785, 300866, 210123, 252796, 440104, 80845, 113081, 174916, 71262)


######### FEDERAL ###########
d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa"
d2$categoria1 <- "Federal"
d2$categoria2 <- "-"

# Federal por ano
d2$`2019` <- c(1519, 94, "-", 25, "-", "-", 69, "-", "-", 249, "-", "-", 27, 118, 104, "-", "-", "-", "-", 872, 195, 78, 336, 263, 263, 69, 142, 52, 41, "-", "-", 41, "-")
d2$`2020` <- c(1399, 73, "-", 25, "-", "-", 48, "-", "-", 268, "-", "-", 39, 111, 118, "-", "-", "-", "-", 849, 198, 78, 370, 203, 166, "-", 127, 39, 43, "-", "-", 43, "-")
d2$`2021` <- c(1285, 48, "-", 21, "-", "-", 27, "-", "-", 254, "-", "-", 39, 104, 111, "-", "-", "-", "-", 805, 259, 48, 269, 229, 138, "-", 100, 38, 40, "-", "-", 40, "-")
d2$`2022` <- c(1415, 85, "-", 21, "-", "-", 64, "-", "-", 214, "-", "-", 28, 110, 76, "-", "-", "-", "-", 869, 261, 55, 342, 211, 205, "-", 140, 65, 42, "-", "-", 42, "-")
d2$`2023` <- c(1525, 93, "-", 25, "-", "-", 68, "-", "-", 254, "-", "-", 42, 105, 107, "-", "-", "-", "-", 930, 257, 72, 360, 241, 200, "-", 148, 52, 48, "-", "-", 48, "-")



######### ESTADUAL ###########
d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa"
d3$categoria1 <- "Estadual"
d3$categoria2 <- "-"

# Estadual por ano
d3$`2019` <- c(55206, 810, 60, 279, 6, "-", 59, 337, 69, 3929, "-", "-", 756, "-", 213, 1783, 374, "-", 803, 1619, 531, "-", 223, 865, 2817, 870, 48, 1899, 46031, 148, 309, "-", 45574)
d3$`2020` <- c(55467, 800, 27, 278, 13, "-", 60, 370, 52, 3770, "-", "-", 768, "-", 188, 1740, 305, "-", 769, 973, 428, "-", 190, 355, 2357, 909, 55, 1393, 47567, 138, 339, 19, 47071)
d3$`2021` <- c(52986, 857, "-", 276, 8, "-", 58, 453, 62, 3845, "-", "-", 777, "-", 173, 1741, 301, "-", 853, 1013, 529, "-", 186, 298, 2165, 927, 56, 1182, 45106, 92, 248, "-", 44766)
d3$`2022` <- c(53642, 968, "-", 285, 5, "-", 183, 424, 71, 3711, 51, "-", 786, "-", 184, 1813, 331, "-", 546, 974, 465, "-", 217, 292, 2087, 824, 60, 1203, 45902, 109, 24, "-", 45769)
d3$`2023` <- c(56192, 1057, "-", 373, 10, "-", 246, 341, 87, 4030, 30, "-", 822, "-", 172, 2020, 344, "-", 642, 1014, 483, "-", 178, 353, 1953, 802, 60, 1091, 48138, 126, 21, "-", 47991)

######### MUNICIPAL ###########
d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa"
d4$categoria1 <- "Municipal"
d4$categoria2 <- "-"

# Municipal por ano
d4$`2019` <- c(3953633, 451441, 29919, 25730, 112687, 19038, 210376, 19405, 34286, 1074691, 181788, 77153, 176175, 64286, 67055, 146445, 61746, 39399, 260644, 1598030, 366676, 95493, 238837, 897024, 557674, 218323, 159238, 180113, 271797, 61187, 89812, 120798, "-")
d4$`2020` <- c(4000709, 453622, 32903, 25421, 112184, 19095, 209697, 19082, 35240, 1085614, 176730, 78772, 183391, 66069, 68565, 148614, 61011, 38927, 263535, 1607178, 365838, 96216, 243035, 902089, 569402, 225060, 162123, 182219, 284893, 64023, 92875, 127995, "-")
d4$`2021` <- c(3949829, 441481, 33954, 23610, 108214, 18898, 203117, 17822, 35866, 1088562, 170481, 78416, 186911, 62265, 70239, 144603, 60932, 38596, 276119, 1580411, 374754, 91068, 242890, 871699, 568615, 232703, 160923, 174989, 270760, 60513, 87077, 123170, "-")
d4$`2022` <- c(3960069, 447791, 36060, 23253, 108838, 19162, 204955, 17941, 37582, 1074600, 169224, 75976, 183174, 61702, 68055, 145346, 62207, 39065, 269851, 1578080, 379749, 92855, 241436, 864040, 580244, 237545, 166262, 176437, 279354, 63428, 91104, 124822, "-")
d4$`2023` <- c(4112950, 479291, 39914, 25212, 114618, 21131, 217806, 19662, 40948, 1120477, 176980, 79012, 190358, 64514, 72781, 153916, 67301, 41077, 274538, 1615983, 398344, 97534, 247273, 872832, 598993, 246116, 171854, 181023, 298206, 66948, 98863, 132395, "-")


######### PRIVADA ###########
d5 <- bases
d5$tematica <- "Educação"
d5$indicador <- "Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa"
d5$categoria1 <- "Privada"
d5$categoria2 <- "-"

# Privada por ano
d5$`2019` <- c(1207328, 60288, 5803, 1292, 13065, 2065, 29315, 2607, 6141, 394166, 43719, 16295, 62506, 23463, 32431, 85431, 19779, 18529, 92013, 500174, 105006, 10654, 140802, 243712, 160106, 55215, 32269, 72622, 92594, 11999, 14191, 40233, 26171)
d5$`2020` <- c(1120231, 55917, 5254, 1103, 12621, 1836, 26947, 2541, 5615, 355523, 41623, 13999, 58457, 21201, 27409, 76215, 16389, 16674, 83556, 470563, 97926, 10218, 130674, 231745, 152038, 50375, 31844, 69819, 86190, 11264, 12498, 37499, 24929)
d5$`2021` <- c(898089, 44654, 4143, 943, 9728, 1429, 22794, 1600, 4017, 266845, 39242, 11406, 44115, 17891, 21636, 57530, 13237, 12053, 49735, 376670, 70185, 8677, 100413, 197395, 137184, 43343, 31610, 62231, 72736, 10269, 10632, 29732, 22103)
d5$`2022` <- c(1077949, 55418, 5659, 1322, 12071, 1767, 27270, 2232, 5097, 338899, 42343, 13756, 53044, 20281, 28160, 68079, 16108, 16399, 80729, 442386, 90327, 9772, 117992, 224295, 154846, 50962, 35834, 68050, 86400, 12142, 12477, 38551, 23230)
d5$`2023` <- c(1167615, 59058, 5834, 1379, 12787, 2227, 28597, 2430, 5804, 368921, 45521, 15127, 56865, 21884, 30383, 77076, 17024, 17726, 87315, 483285, 99637, 10745, 123046, 249857, 162639, 53948, 38061, 70630, 93712, 13771, 14197, 42473, 23271)


######### RANKING 2023 ###########
d6 <- bases
d6$tematica <- "Educação"
d6$indicador <- "Número de Matrículas na Pré-Escola - Ensino Regular por Dependência Administrativa"
d6$categoria1 <- "Ranking 2023"
d6$categoria2 <- "-"

# Ranking 2023
d6$`2023` <- c("-", 4, 24, 25, 13, 26, 8, 27, 23, 2, 10, 17, 7, 18, 16, 9, 19, 22, 4, 1, 2, 15, 3, 1, 3, 5, 11, 6, 5, 20, 14, 12, 21)


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

############ TOTAL ###########
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Matrículas no Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"

# Total por ano
d1$`2019` <- c(26923730, 3015573, 261087, 156388, 700172, 102264, 1418496, 135999, 241167, 7889261, 1153446, 470183, 1177983, 459400, 548659, 1284245, 475704, 328558, 1991083, 10349288, 2461094, 502696, 1976311, 5409187, 3550498, 1404493, 865262, 1280743, 2119110, 397032, 478220, 868931, 374927)
d1$`2020` <- c(26718830, 2978206, 252638, 156679, 700104, 104202, 1394011, 134820, 235752, 7736717, 1123973, 458077, 1165368, 448764, 538748, 1267292, 464704, 322614, 1947177, 10346173, 2462047, 501920, 1967998, 5414208, 3553679, 1407978, 876392, 1269309, 2104055, 392015, 477717, 861291, 373032)
d1$`2021` <- c(26515601, 2955281, 244815, 153015, 702763, 103123, 1389983, 133839, 227743, 7698779, 1112636, 459871, 1161434, 447692, 540919, 1249850, 458782, 320638, 1946957, 10252321, 2407107, 503003, 1945408, 5396803, 3506528, 1348296, 900240, 1257992, 2102692, 391975, 486568, 855021, 369128)
d1$`2022` <- c(26452228, 2915590, 237067, 149645, 696897, 107058, 1368026, 132504, 224393, 7597217, 1084585, 452465, 1153226, 447245, 534889, 1246567, 453292, 314459, 1910489, 10250077, 2407621, 501919, 1960826, 5379711, 3566859, 1380369, 922946, 1263544, 2122485, 393401, 501193, 862812, 365079)
d1$`2023` <- c(26108208, 2864390, 227246, 147350, 692623, 111096, 1331723, 131948, 222404, 7420513, 1045700, 439111, 1134636, 437133, 523996, 1229419, 439014, 305930, 1865574, 10130005, 2343279, 498357, 1945098, 5343271, 3562935, 1365869, 941272, 1255794, 2130365, 393616, 507089, 868222, 361438)


######### FEDERAL ###########
d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Matrículas no Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d2$categoria1 <- "Federal"
d2$categoria2 <- "-"

# Federal por ano
d2$`2019` <- c(23102, 3116, "-", 340, 528, 327, 1921, "-", "-", 2530, 442, "-", 425, 182, 154, 732, "-", 236, 359, 12814, 2666, "-", 9916, 232, 2394, 473, 690, 1231, 2248, 484, "-", 502, 1262)
d2$`2020` <- c(22772, 3271, "-", 342, 738, 327, 1864, "-", "-", 2523, 434, "-", 442, 224, 144, 708, "-", 199, 372, 12360, 2622, "-", 9738, "-", 2376, 490, 687, 1199, 2242, 475, "-", 493, 1274)
d2$`2021` <- c(22279, 3191, "-", 306, 807, 325, 1753, "-", "-", 2569, 442, "-", 459, 224, 137, 702, "-", 237, 368, 11899, 2642, "-", 9051, 206, 2366, 472, 698, 1196, 2254, 477, "-", 432, 1345)
d2$`2022` <- c(22798, 3282, "-", 303, 793, 317, 1869, "-", "-", 2595, 463, "-", 461, 229, 116, 709, "-", 245, 372, 12156, 2474, "-", 9487, 195, 2392, 481, 671, 1240, 2373, 497, "-", 475, 1401)
d2$`2023` <- c(23728, 3362, "-", 335, 745, 317, 1965, "-", "-", 2625, 464, "-", 460, 225, 137, 736, "-", 240, 363, 12872, 2546, "-", 9766, 560, 2438, 488, 662, 1288, 2431, 483, "-", 505, 1443)

######### ESTADUAL ###########
d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Matrículas no Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d3$categoria1 <- "Estadual"
d3$categoria2 <- "-"

# Estadual por ano
d3$`2019` <- c(6921857, 808311, 112544, 91111, 221444, 47075, 179995, 72822, 83320, 674355, 29818, 35465, 18155, 90881, 89227, 151188, 50028, 70231, 139362, 3296155, 1034628, 105998, 161516, 1994013, 1299370, 539221, 291450, 468699, 843666, 120700, 203140, 245662, 274164)
d3$`2020` <- c(6836438, 795230, 109307, 91488, 219339, 47589, 172310, 72555, 82642, 636060, 24969, 34026, 15511, 88173, 79459, 146385, 49614, 69372, 128551, 3296947, 1038442, 101607, 165586, 1991312, 1283259, 546804, 291931, 444524, 824942, 108504, 198758, 242507, 275173)
d3$`2021` <- c(6606309, 771892, 107812, 89374, 210888, 45864, 168438, 71884, 77632, 590762, 24234, 32102, 13956, 84987, 71993, 139775, 46128, 67597, 109990, 3218759, 983134, 99220, 154749, 1981656, 1213001, 503014, 293285, 416702, 811895, 101380, 201668, 235798, 273049)
d3$`2022` <- c(6462246, 748496, 103509, 87555, 206418, 46898, 158367, 70196, 75553, 558033, 23474, 28018, 12556, 83640, 63033, 136610, 44341, 65623, 100738, 3121863, 937770, 88755, 154869, 1940469, 1239206, 529931, 290863, 418412, 794648, 95100, 196009, 234125, 269414)
d3$`2023` <- c(6221872, 719037, 96867, 85689, 201653, 48231, 143708, 68883, 74006, 524563, 22330, 27454, 11419, 79388, 53603, 130320, 40516, 62615, 96918, 3004310, 875086, 84377, 152551, 1892296, 1200871, 510033, 284406, 406432, 773091, 90110, 185972, 231651, 265358)

######### MUNICIPAL ###########
d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Matrículas no Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d4$categoria1 <- "Municipal"
d4$categoria2 <- "-"

# Municipal por ano
d4$`2019` <- c(15261665, 1947983, 123174, 57710, 428192, 46982, 1106054, 50407, 135464, 5726919, 997259, 365528, 912245, 268430, 339883, 817200, 339447, 181702, 1505225, 4928380, 1103196, 332742, 1194200, 2298242, 1752633, 655054, 459110, 638469, 905750, 227580, 216665, 461505, "-")
d4$`2020` <- c(15210213, 1929356, 118245, 57844, 430758, 48956, 1092445, 49878, 131230, 5658188, 974229, 360394, 910445, 264458, 345135, 811175, 333390, 178008, 1480954, 4930442, 1098939, 336112, 1193510, 2301881, 1771950, 652999, 467896, 651055, 920277, 234682, 222153, 463442, "-")
d4$`2021` <- c(15472487, 1942581, 114057, 56608, 444555, 49822, 1096669, 50613, 130257, 5774362, 966718, 368285, 937316, 267654, 363407, 825533, 334647, 182375, 1528427, 5004147, 1113976, 342105, 1227463, 2320603, 1806320, 649428, 487236, 669656, 945077, 242447, 229770, 472860, "-")
d4$`2022` <- c(15373541, 1911975, 109296, 54649, 438607, 52629, 1078779, 50340, 127675, 5648541, 935302, 361059, 923355, 264920, 360118, 811654, 329446, 176153, 1486534, 5023283, 1138949, 348267, 1222907, 2313160, 1821361, 648428, 505216, 667717, 968381, 247550, 247373, 473458, "-")
d4$`2023` <- c(15172004, 1885612, 106079, 53896, 436894, 55347, 1055890, 50979, 126527, 5476716, 898750, 345924, 900912, 257369, 355484, 793114, 320725, 170280, 1434158, 4985741, 1130832, 347482, 1209324, 2298103, 1836015, 649028, 520907, 666080, 987920, 251315, 260602, 476003, "-")

######### PRIVADA ###########
d5 <- bases
d5$tematica <- "Educação"
d5$indicador <- "Número de Matrículas no Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d5$categoria1 <- "Privada"
d5$categoria2 <- "-"

# Privada por ano
d5$`2019` <- c(4717106, 256163, 25369, 7227, 50008, 7880, 130526, 12770, 22383, 1485457, 125927, 69190, 247158, 99907, 119395, 315125, 86229, 76389, 346137, 2111939, 320604, 63956, 610679, 1116700, 496101, 209745, 114012, 172344, 367446, 48268, 58415, 161262, 99501)
d5$`2020` <- c(4649407, 250349, 25086, 7005, 49269, 7330, 127392, 12387, 21880, 1439946, 124341, 63657, 238970, 95909, 114010, 309024, 81700, 75035, 337300, 2106424, 322044, 64201, 599164, 1121015, 496094, 207685, 115878, 172531, 356594, 48354, 56806, 154849, 96585)
d5$`2021` <- c(4414526, 237617, 22946, 6727, 46513, 7112, 123123, 11342, 19854, 1331086, 121242, 59484, 209703, 94827, 105382, 283840, 78007, 70429, 308172, 2017516, 307355, 61678, 554145, 1094338, 484841, 195382, 119021, 170438, 343466, 47671, 55130, 145931, 94734)
d5$`2022` <- c(4593643, 251837, 24262, 7138, 51079, 7214, 129011, 11968, 21165, 1388048, 125346, 63388, 216854, 98456, 111622, 297594, 79505, 72438, 322845, 2092775, 328428, 64897, 573563, 1125887, 503900, 201529, 126196, 176175, 357083, 50254, 57811, 154754, 94264)
d5$`2023` <- c(4690604, 256379, 24300, 7430, 53331, 7201, 130160, 12086, 21871, 1416609, 124156, 65733, 221845, 100151, 114772, 305249, 77773, 72795, 334135, 2127082, 334815, 66498, 573457, 1152312, 523611, 206320, 135297, 181994, 366923, 51708, 60515, 160063, 94637)

######### RANKING 2023 ###########
d6 <- bases
d6$tematica <- "Educação"
d6$indicador <- "Número de Matrículas no Ensino Fundamental - Ensino Regular por Dependência Administrativa"
d6$categoria1 <- "Ranking 2023"
d6$categoria2 <- "-"

# Ranking 2023
d6$`2023` <- c("-", 4, 23, 25, 13, 27, 6, 26, 24, 2, 10, 17, 9, 19, 14, 8, 18, 22, 4, 1, 2, 16, 3, 1, 3, 5, 11, 7, 5, 20, 15, 12, 21)

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

############ TOTAL ###########
d1 <- bases
d1$tematica <- "Educação"
d1$indicador <- "Número de Matrículas no Ensino Médio - Ensino Regular por Dependência Administrativa"
d1$categoria1 <- "Total"
d1$categoria2 <- "-"


############ FEDERAL ###########
d2 <- bases
d2$tematica <- "Educação"
d2$indicador <- "Número de Matrículas no Ensino Médio - Ensino Regular por Dependência Administrativa"
d2$categoria1 <- "Federal"
d2$categoria2 <- "-"


############ ESTADUAL ###########
d3 <- bases
d3$tematica <- "Educação"
d3$indicador <- "Número de Matrículas no Ensino Médio - Ensino Regular por Dependência Administrativa"
d3$categoria1 <- "Estadual"
d3$categoria2 <- "-"


############ MUNICIPAL ###########
d4 <- bases
d4$tematica <- "Educação"
d4$indicador <- "Número de Matrículas no Ensino Médio - Ensino Regular por Dependência Administrativa"
d4$categoria1 <- "Municipal"
d4$categoria2 <- "-"


############ PRIVADA ###########
d5 <- bases
d5$tematica <- "Educação"
d5$indicador <- "Número de Matrículas no Ensino Médio - Ensino Regular por Dependência Administrativa"
d5$categoria1 <- "Privada"
d5$categoria2 <- "-"


############ RAKING 2023 ###########
d6 <- bases
d6$tematica <- "Educação"
d6$indicador <- "Número de Matrículas no Ensino Médio - Ensino Regular por Dependência Administrativa"
d6$categoria1 <- "Ranking 2023"
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
