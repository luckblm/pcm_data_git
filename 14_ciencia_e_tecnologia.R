#CIÊNCIA E TECNOLOGIA----
bases$subtematica <- "-"

##01 - Total dos investimentos realizados em bolsas (R$ Mil) e no fomento à pesquisa----

d1 <- bases
d1$tematica <- "Ciência e Tecnologia"
d1$indicador <- "Total dos Investimentos Realizados em Bolsas (R$ Mil) e no Fomento à Pesquisa, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Total dos Investimentos Realizados em Bolsas (R$ Mil) e no Fomento à Pesquisa"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Ciência e Tecnologia"
d2$indicador <- "Total dos Investimentos Realizados em Bolsas (R$ Mil) e no Fomento à Pesquisa, segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "ciencia_tecnologia/01 - Total dos investimentos realizados em bolsas e no fomento à pesquisa.xlsx",
  overwrite = TRUE
)
##02 - Número de Bolsas de Estudo no País e no Exterior----
d1 <- bases
d1$tematica <- "Ciência e Tecnologia"
d1$indicador <- "Número de Bolsas de Estudo no País e no Exterior, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Número de Bolsas de Estudo"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Ciência e Tecnologia"
d2$indicador <- "Número de Bolsas de Estudo no País e no Exterior, segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "ciencia_tecnologia/02 - Número de Bolsas de Estudo no País e no Exterior.xlsx",
  overwrite = TRUE
)
##03 - Investimentos realizados em bolsas e no fomento à pesquisa per capita (R$)----
d1 <- bases
d1$tematica <- "Ciência e Tecnologia"
d1$indicador <- "Investimentos realizados em bolsas e no fomento à pesquisa per capita (R$), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Investimento per capita (R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Ciência e Tecnologia"
d2$indicador <- "Investimentos realizados em bolsas e no fomento à pesquisa per capita (R$), segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "ciencia_tecnologia/03 - Investimentos realizados em bolsas e no fomento à pesquisa per capita.xlsx",
  overwrite = TRUE
)
##04 - Número de Doutores----
d1 <- bases
d1$tematica <- "Ciência e Tecnologia"
d1$indicador <- "Número de Doutores, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Número de Doutores"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Ciência e Tecnologia"
d2$indicador <- "Número de Doutores, segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "ciencia_tecnologia/04 - Número de Doutores.xlsx",
  overwrite = TRUE
)
##05 - Número de doutores por 100.000 habitantes----
d1 <- bases
d1$tematica <- "Ciência e Tecnologia"
d1$indicador <- "Número de doutores por 100.000 habitantes, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Número de doutores por 100.000 habitantes"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Ciência e Tecnologia"
d2$indicador <- "Número de doutores por 100.000 habitantes, segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "ciencia_tecnologia/05 - Número de doutores por 100_000 habitantes.xlsx",
  overwrite = TRUE
)
##06 - Dispêndios dos Governos Estaduais em Ciência e Tecnologia----
d1 <- bases
d1$tematica <- "Ciência e Tecnologia"
d1$indicador <- "Dispêndios dos Governos Estaduais em Ciência e Tecnologia (C&T), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- ""
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Ciência e Tecnologia"
d2$indicador <- "Dispêndios dos Governos Estaduais em Ciência e Tecnologia (C&T), segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "ciencia_tecnologia/06 - Dispêndios dos Governos Estaduais em Ciência e Tecnologia.xlsx",
  overwrite = TRUE
)
##07 - Percentual dos Dispêndios em Ciência e Tecnologia -----
d1 <- bases
d1$tematica <- "Ciência e Tecnologia"
d1$indicador <- "Percentual dos Dispêndios em Ciência e Tecnologia dos Governos Estaduais em Relação às Suas Receitas Totais, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Percentual dos dispêndios (%)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Ciência e Tecnologia"
d2$indicador <- "Percentual dos Dispêndios em Ciência e Tecnologia dos Governos Estaduais em Relação às Suas Receitas Totais, segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "ciencia_tecnologia/07 - Percentual dos Dispêndios em Ciência e Tecnologia.xlsx",
  overwrite = TRUE
)
##08 - Dispêndios dos Governos Estaduais em Pesquisa e Desenvolvimento -----

# Vetor das Unidades da Federação e siglas das IES

macroregião <- c(
  "Brasil", "Norte", "Norte", "Norte", "Norte", "Norte", "Norte", "Norte", "Nordeste",
  "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste",
  "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste",
  "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Sudeste", "Sudeste", "Sudeste",
  "Sudeste", "Sudeste", "Sudeste", "Sudeste", "Sudeste", "Sudeste", "Sudeste",
  "Sudeste", "Sudeste", "Sudeste", "Sudeste", "Sudeste", "Sudeste", "Sul", "Sul",
  "Sul", "Sul", "Sul", "Sul", "Sul", "Sul", "Sul", "Sul",
  "Sul", "Sul", "Sul", "Centro-Oeste", "Centro-Oeste",
  "Centro-Oeste", "Centro-Oeste", "Centro-Oeste", "Centro-Oeste", "Centro-Oeste", "Distrito Federal", "Distrito Federal"
)

unidades_federacao <- c(
  "Brasil", "Norte", "Amazonas", "UEA", "Pará", "UEPa", "Roraima", "UERR", "Nordeste",
  "Alagoas", "UNEAL", "Piauí", "UESPI", "Maranhão", "UEMA", "Ceará", "UECE", "URCA",
  "UVA", "Rio Grande do Norte", "UERN", "Paraíba", "UEPB", "Pernambuco", "UPE",
  "Bahia", "UEFS", "UESB", "UESC", "UNEB", "Sudeste", "Minas Gerais", "UEMG",
  "UNIMONTES", "Rio de Janeiro", "UENF", "UERJ", "UEZO", "São Paulo", "CEETEPS",
  "FAENQUIL", "FAMEMA", "FAMERP", "UNESP", "UNICAMP", "USP", "Sul", "Paraná",
  "UEL", "UEM", "UENP", "UEPG", "UNESPAR", "UNICENTRO", "UNIOESTE", "Santa Catarina",
  "UDESC", "Rio Grande do Sul", "UERGS", "Centro-Oeste", "Mato Grosso do Sul",
  "UEMS", "Mato Grosso", "UNEMAT", "Goiás", "UEG", "Distrito Federal", "FEPECS"
)

# Vetor das Instituições Estaduais de Ensino Superior
instituicoes_superiores <- c(
  NA, NA, NA, "Universidade do Estado do Amazonas", NA, "Universidade do Estado do Pará", NA,
  "Universidade Estadual de Roraima", NA, NA, "Universidade Estadual de Alagoas", NA, 
  "Universidade Estadual do Piauí", NA, "Universidade Estadual do Maranhão", NA, 
  "Universidade Estadual do Ceará", "Universidade Regional do Cariri", 
  "Universidade Estadual Vale do Acaraú", NA, "UERN - Universidade do Estado do Rio Grande do Norte", 
  NA, "Universidade Estadual da Paraíba", NA, "Universidade de Pernambuco", NA, 
  "Universidade Estadual de Feira de Santana", "Universidade Estadual do Sudoeste da Bahia", 
  "Universidade Estadual de Santa Cruz", "Universidade do Estado da Bahia", NA, NA, 
  "Universidade do Estado de Minas Gerais", "Universidade Estadual de Montes Claros", NA, 
  "Universidade Estadual do Norte Fluminense Darcy Ribeiro", "Universidade do Estado do Rio de Janeiro",
  "Centro Universitário Estadual da Zona Oeste", NA, "Centro Estadual de Educação Tecnológica Paula Souza",
  "Faculdade de Engenharia Química de Lorena", "Faculdade de Medicina de Marília",
  "Faculdade de Medicina de São José do Rio Preto", "Universidade Estadual Paulista 'Júlio de Mesquita Filho'",
  "Universidade Estadual de Campinas", "Universidade de São Paulo", NA, NA, 
  "Universidade Estadual de Londrina", "Universidade Estadual de Maringá", 
  "Universidade Estadual do Norte do Paraná", "Universidade Estadual de Ponta Grossa",
  "Universidade Estadual do Paraná", "Universidade Estadual do Centro-Oeste",
  "Universidade Estadual do Oeste do Paraná", NA, "Universidade do Estado de Santa Catarina",
  NA, "Universidade Estadual do Rio Grande do Sul", NA, NA, 
  "Universidade do Estado de Mato Grosso do Sul", NA, "Universidade do Estado de Mato Grosso",
  NA, "Universidade Estadual de Goiás", NA, "Fundação de Ensino e Pesquisa em Ciências da Saúde"
)

d1 <- bases
d1$tematica <- "Ciência e Tecnologia"
d1$regiao <- macroregião
d1$localidade <- unidades_federacao
d1$indicador <- "Estimativa dos dispêndios em pesquisa e desenvolvimento (P&D) das instituições estaduais de ensino superior, por região, unidade da federação e instituição"
d1$categoria1 <- instituicoes_superiores
d1$categoria2 <- "(em milhões de R$ correntes)"

d2 <- bases
d2$tematica <- "Ciência e Tecnologia"
d1$regiao <- macroregião
d1$localidade <- unidades_federacao
d2$indicador <- "Estimativa dos dispêndios em pesquisa e desenvolvimento (P&D) das instituições estaduais de ensino superior, por região, unidade da federação e instituição"
d2$categoria1 <- instituicoes_superiores
d2$categoria2 <- "Ranking "


# Criar um novo arquivo Excel
wb <- createWorkbook()
# Criar lista de bases de dados
dados_lista <- list(
  "d1" = d1,
  "d" = d
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
  "ciencia_tecnologia/.xlsx",
  overwrite = TRUE
)
##09 - Percentual dos Dispêndios em Pesquisa e Desenvolvimento -----
d1 <- bases
d1$tematica <- "Ciência e Tecnologia"
d1$indicador <- "Percentual dos Dispêndios em Pesquisa e Desenvolvimento (P&D)(1) dos Governos Estaduais em Relação às suas Receitas Totais, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Percentual dos Dispêndios em Pesquisa e Desenvolvimento (P&D) dos Governos Estaduais em Relação às suas Receitas Totais"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Ciência e Tecnologia"
d2$indicador <- "Percentual dos Dispêndios em Pesquisa e Desenvolvimento (P&D)(1) dos Governos Estaduais em Relação às suas Receitas Totais, segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "ciencia_tecnologia/09 - Percentual dos Dispêndios em Pesquisa e Desenvolvimento.xlsx",
  overwrite = TRUE
)
##10 - Dispêndios em Pesquisa e Desenvolvimento das Instituições Estaduais de Ensino Superior----

macroregião <- c(
  "Brasil", "Norte", "Norte", "Norte", "Norte", "Norte", "Norte", "Norte", "Nordeste",
  "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste",
  "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste",
  "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Nordeste", "Sudeste", "Sudeste", "Sudeste",
  "Sudeste", "Sudeste", "Sudeste", "Sudeste", "Sudeste", "Sudeste", "Sudeste",
  "Sudeste", "Sudeste", "Sudeste", "Sudeste", "Sudeste", "Sudeste", "Sul", "Sul",
  "Sul", "Sul", "Sul", "Sul", "Sul", "Sul", "Sul", "Sul",
  "Sul", "Sul", "Sul", "Centro-Oeste", "Centro-Oeste",
  "Centro-Oeste", "Centro-Oeste", "Centro-Oeste", "Centro-Oeste", "Centro-Oeste", "Distrito Federal", "Distrito Federal"
)

unidades_federacao <- c(
  "Brasil", "Norte", "Amazonas", "UEA", "Pará", "UEPa", "Roraima", "UERR", "Nordeste",
  "Alagoas", "UNEAL", "Piauí", "UESPI", "Maranhão", "UEMA", "Ceará", "UECE", "URCA",
  "UVA", "Rio Grande do Norte", "UERN", "Paraíba", "UEPB", "Pernambuco", "UPE",
  "Bahia", "UEFS", "UESB", "UESC", "UNEB", "Sudeste", "Minas Gerais", "UEMG",
  "UNIMONTES", "Rio de Janeiro", "UENF", "UERJ", "UEZO", "São Paulo", "CEETEPS",
  "FAENQUIL", "FAMEMA", "FAMERP", "UNESP", "UNICAMP", "USP", "Sul", "Paraná",
  "UEL", "UEM", "UENP", "UEPG", "UNESPAR", "UNICENTRO", "UNIOESTE", "Santa Catarina",
  "UDESC", "Rio Grande do Sul", "UERGS", "Centro-Oeste", "Mato Grosso do Sul",
  "UEMS", "Mato Grosso", "UNEMAT", "Goiás", "UEG", "Distrito Federal", "FEPECS"
)

# Vetor das Instituições Estaduais de Ensino Superior
instituicoes_superiores <- c(
  NA, NA, NA, "Universidade do Estado do Amazonas", NA, "Universidade do Estado do Pará", NA,
  "Universidade Estadual de Roraima", NA, NA, "Universidade Estadual de Alagoas", NA, 
  "Universidade Estadual do Piauí", NA, "Universidade Estadual do Maranhão", NA, 
  "Universidade Estadual do Ceará", "Universidade Regional do Cariri", 
  "Universidade Estadual Vale do Acaraú", NA, "UERN - Universidade do Estado do Rio Grande do Norte", 
  NA, "Universidade Estadual da Paraíba", NA, "Universidade de Pernambuco", NA, 
  "Universidade Estadual de Feira de Santana", "Universidade Estadual do Sudoeste da Bahia", 
  "Universidade Estadual de Santa Cruz", "Universidade do Estado da Bahia", NA, NA, 
  "Universidade do Estado de Minas Gerais", "Universidade Estadual de Montes Claros", NA, 
  "Universidade Estadual do Norte Fluminense Darcy Ribeiro", "Universidade do Estado do Rio de Janeiro",
  "Centro Universitário Estadual da Zona Oeste", NA, "Centro Estadual de Educação Tecnológica Paula Souza",
  "Faculdade de Engenharia Química de Lorena", "Faculdade de Medicina de Marília",
  "Faculdade de Medicina de São José do Rio Preto", "Universidade Estadual Paulista 'Júlio de Mesquita Filho'",
  "Universidade Estadual de Campinas", "Universidade de São Paulo", NA, NA, 
  "Universidade Estadual de Londrina", "Universidade Estadual de Maringá", 
  "Universidade Estadual do Norte do Paraná", "Universidade Estadual de Ponta Grossa",
  "Universidade Estadual do Paraná", "Universidade Estadual do Centro-Oeste",
  "Universidade Estadual do Oeste do Paraná", NA, "Universidade do Estado de Santa Catarina",
  NA, "Universidade Estadual do Rio Grande do Sul", NA, NA, 
  "Universidade do Estado de Mato Grosso do Sul", NA, "Universidade do Estado de Mato Grosso",
  NA, "Universidade Estadual de Goiás", NA, "Fundação de Ensino e Pesquisa em Ciências da Saúde"
)


d1 <- bases
d1$tematica <- "Ciência e Tecnologia"
d1$regiao <- macroregião
d1$localidade <- unidades_federacao
d1$indicador <- "Estimativa dos dispêndios em pesquisa e desenvolvimento (P&D) das instituições estaduais de ensino superior, por região, unidade da federação e instituição"
d1$categoria1 <- instituicoes_superiores
d1$categoria2 <- "(em milhões de R$ correntes)"

d2 <- bases
d2$tematica <- "Ciência e Tecnologia"
d2$indicador <- "Estimativa dos dispêndios em pesquisa e desenvolvimento (P&D) das instituições estaduais de ensino superior, por região, unidade da federação e instituição"
d2$categoria2 <- instituicoes_superiores
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
  "ciencia_tecnologia/10 - Dispêndios em Pesquisa e Desenvolvimento das Instituições Estaduais de Ensino Superior.xlsx",
  overwrite = TRUE
)
##11 - Investimento da CAPES em bolsa e fomento----
d1 <- bases
d1$tematica <- "Ciência e Tecnologia"
d1$indicador <- "Investimento da CAPES em bolsa e fomento (Mil R$), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Investimento da CAPES em bolsa e fomento (Mil R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Ciência e Tecnologia"
d2$indicador <- "Investimento da CAPES em bolsa e fomento (Mil R$), segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "ciencia_tecnologia/11 - Investimento da CAPES em bolsa e fomento.xlsx",
  overwrite = TRUE
)
##12 - Concessão de Bolsas de Pós-graduação da CAPES----
d1 <- bases
d1$tematica <- "Ciência e Tecnologia"
d1$indicador <- "Concessão de Bolsas de Pós-graduação da CAPES, segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Concessão de Bolsas de Pós-graduação da CAPES (Quantidade)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Ciência e Tecnologia"
d2$indicador <- "Concessão de Bolsas de Pós-graduação da CAPES, segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "ciencia_tecnologia/12 - Concessão de Bolsas de Pós-graduação da CAPES.xlsx",
  overwrite = TRUE
)
##13 - Concessão de Bolsas de Doutorado da CAPES----
d <- bases
d$tematica <- "Ciência e Tecnologia"
d$indicador <- "Concessão de Bolsas de Doutorado da CAPES, segundo Brasil, Grandes Regiões e Unidades da Federação"
d$categoria1 <- "Concessão de Bolsas de Doutorado da CAPES (Quantidade)"
d$categoria2 <- "-"

d <- bases
d$tematica <- "Ciência e Tecnologia"
d$indicador <- "Concessão de Bolsas de Doutorado da CAPES, segundo Brasil, Grandes Regiões e Unidades da Federação"
d$categoria1 <- "Ranking"
d$categoria2 <- "-"


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
  "ciencia_tecnologia/13 - Concessão de Bolsas de Doutorado da CAPES.xlsx",
  overwrite = TRUE
)
##14 - Concessão de Bolsas de Mestrado da CAPES----
d <- bases
d$tematica <- "Ciência e Tecnologia"
d$indicador <- "Concessão de Bolsas de Mestrado da CAPES, segundo Brasil, Grandes Regiões e Unidades da Federação"
d$categoria1 <- "Concessão de Bolsas de Mestrado da CAPES (Quantidade)"
d$categoria2 <- "-"

d <- bases
d$tematica <- "Ciência e Tecnologia"
d$indicador <- "Concessão de Bolsas de Mestrado da CAPES, segundo Brasil, Grandes Regiões e Unidades da Federação"
d$categoria1 <- "Ranking"
d$categoria2 <- "-"


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
  "ciencia_tecnologia/14 - Concessão de Bolsas de Mestrado da CAPES.xlsx",
  overwrite = TRUE
)