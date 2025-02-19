#FINANÇAS PÚBLICAS----

bases$subtematica <- "-"

##01 - Receita Orçamentária Total (Mil R$)----

d1 <- bases
d1$tematica <- "Finanças Públicas"
d1$indicador <- "Receita Orçamentaria Total (Mil R$), Preços Correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Receita Orçamentaria Total (Mil R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Finanças Públicas"
d2$indicador <- "Receita Orçamentaria Total (Mil R$), Preços Correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "financas_publicas/01 - Receita Orçamentária Total.xlsx",
  overwrite = TRUE
)
##02 - Receita Correntes (Mil R$)----
d1 <- bases
d1$tematica <- "Finanças Públicas"
d1$indicador <- "Receitas Correntes (Mil R$), Preços Correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Receitas Correntes (Mil R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Finanças Públicas"
d2$indicador <- "Receitas Correntes (Mil R$), Preços Correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "financas_publicas/02 - Receita Correntes.xlsx",
  overwrite = TRUE
)
##03 - Receitas de Transferências Correntes (Mil R$)----
d1 <- bases
d1$tematica <- "Finanças Públicas"
d1$indicador <- "Receitas de Transferências Correntes (Mil R$), Preços Correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Receitas de Transferências Correntes (Mil R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Finanças Públicas"
d2$indicador <- "Receitas de Transferências Correntes (Mil R$), Preços Correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "financas_publicas/03 - Receitas de Transferências Correntes.xlsx",
  overwrite = TRUE
)
##04 - Despesa Total - Empenhada (R$) - Preços Correntes----
d1 <- bases
d1$tematica <- "Finanças Públicas"
d1$indicador <- "Despesa Total - Empenhada (Mil R$) - Preços Correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Despesa Total - Empenhada (Mil R$) - Preços Correntes"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Finanças Públicas"
d2$indicador <- "Despesa Total - Empenhada (Mil R$) - Preços Correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "financas_publicas/04 - Despesa Total_Empenhada_Preços Correntes.xlsx",
  overwrite = TRUE
)
##05 - Despesas com Pessoal e Encargos Sociais - Empenhada (R$) - Preços Correntes----
d1 <- bases
d1$tematica <- "Finanças Públicas"
d1$indicador <- "Despesas com Pessoal e Encargos Sociais - Empenhada (Mil R$) - Preços Correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Despesas com Pessoal e Encargos Sociais - Empenhada (R$) - Preços Correntes"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Finanças Públicas"
d2$indicador <- "Despesas com Pessoal e Encargos Sociais - Empenhada (Mil R$) - Preços Correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "financas_publicas/05 - Despesas com Pessoal e Encargos Sociais_Empenhada_Preços Correntes.xlsx",
  overwrite = TRUE
)
##06 - Despesas com Investimentos - Empenhada (R$) - Preços Correntes----
d1 <- bases
d1$tematica <- "Finanças Públicas"
d1$indicador <- "Despesas com Investimentos - Empenhada (Mil R$) - preços correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Despesas com Investimentos  - Empenhada (Mil R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Finanças Públicas"
d2$indicador <- "Despesas com Investimentos - Empenhada (Mil R$) - preços correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "financas_publicas/06 - Despesas com Investimentos_Empenhada_ Preços Correntes.xlsx",
  overwrite = TRUE
)
##07 - Transferências Constitucionais - FPE----
d1 <- bases
d1$tematica <- "Finanças Públicas"
d1$indicador <- "Transferências Constitucionais - Fundo de Participação dos Estados - FPE (Mil R$) - preços correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "FPE (Mil R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Finanças Públicas"
d2$indicador <- "Transferências Constitucionais - Fundo de Participação dos Estados - FPE (Mil R$) - preços correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "financas_publicas/07 - Transferências Constitucionais_FPE.xlsx",
  overwrite = TRUE
)
##08 - Transferências Constitucionais - FUNDEB----
d1 <- bases
d1$tematica <- "Finanças Públicas"
d1$indicador <- "Transferências Constitucionais - Fundo de Manutenção e Desenvolvimento da Educação Básica e de Valorização dos Profissionais da Educação - FUNDEB (Mil R$) - preços correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "FUNDEB (Mil R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Finanças Públicas"
d2$indicador <- "Transferências Constitucionais - Fundo de Manutenção e Desenvolvimento da Educação Básica e de Valorização dos Profissionais da Educação - FUNDEB (Mil R$) - preços correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "financas_publicas/08 - Transferências Constitucionais_FUNDEB.xlsx",
  overwrite = TRUE
)
##09 - Transferências Constitucionais - IPI-EXP----
d1 <- bases
d1$tematica <- "Finanças Públicas"
d1$indicador <- "Transferências Constitucionais - Imposto sobre Produtos Industrializados destinados a Exportação - IPI-Exp (Mil R$) - preços correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "IPI-Exp (Mil R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Finanças Públicas"
d2$indicador <- "Transferências Constitucionais - Imposto sobre Produtos Industrializados destinados a Exportação - IPI-Exp (Mil R$) - preços correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "financas_publicas/09 - Transferências Constitucionais_IPI_EXP.xlsx",
  overwrite = TRUE
)
##10 - Transferências Constitucionais - LC 87-96 Lei Kandir----
d1 <- bases
d1$tematica <- "Finanças Públicas"
d1$indicador <- "Transferências Constitucionais -  LC 87/96 Lei Kandir (Mil R$) - preços correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "LC 87/96 Lei Kandir (Mil R$)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Finanças Públicas"
d2$indicador <- "Transferências Constitucionais -  LC 87/96 Lei Kandir (Mil R$) - preços correntes, Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "financas_publicas/10 - Transferências Constitucionais_LC_87-96_Lei_Kandir.xlsx",
  overwrite = TRUE
)
##11 - Arrecadação do ICMS----
d1 <- bases
d1$tematica <- "Finanças Públicas"
d1$indicador <- "Arrecadação do ICMS - Valores correntes (R$ Mil), segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Valores correntes (R$ Mil)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Finanças Públicas"
d2$indicador <- "Arrecadação do ICMS - Valores correntes (R$ Mil), segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "financas_publicas/11 - Arrecadação do ICMS.xlsx",
  overwrite = TRUE
)
##12 - Arrecadação do IPVA----
d1 <- bases
d1$tematica <- "Finanças Públicas"
d1$indicador <- "Arrecadação do IPVA - Valores correntes (R$ Mil), Segundo Brasil, Grandes Regiões e Unidades da Federação"
d1$categoria1 <- "Valores correntes (R$ Mil)"
d1$categoria2 <- "-"

d2 <- bases
d2$tematica <- "Finanças Públicas"
d2$indicador <- "Arrecadação do IPVA - Valores correntes (R$ Mil), Segundo Brasil, Grandes Regiões e Unidades da Federação"
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
  "financas_publicas/12 - Arrecadação do IPVA.xlsx",
  overwrite = TRUE
)