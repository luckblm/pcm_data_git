#Start

#Carregar pacotes

library(tidyverse)
library(openxlsx)

#Base geral estrutura
bases <- tibble(
  tematica = "-",
  subtematica = "-",
  indicador = "-",
  regiao = c(
    "País",
    "Macrorregião",  "Estado",  "Estado",  "Estado",  "Estado",  "Estado",  "Estado",
    "Estado",  "Macrorregião",  "Estado",  "Estado",  "Estado",  "Estado",  "Estado",
    "Estado",  "Estado",  "Estado",  "Estado",  "Macrorregião",  "Estado",  "Estado",
    "Estado",  "Estado",  "Macrorregião",  "Estado",  "Estado",  "Estado",  "Macrorregião",
    "Estado",  "Estado",  "Estado",  "Estado"),
  localidade = c(
    "BRASIL",
    "Região Norte", "Rondônia", "Acre", "Amazonas", "Roraima", "Pará", "Amapá", "Tocantins",
    "Região Nordeste", "Maranhão", "Piauí", "Ceará", "Rio Grande do Norte", "Paraíba", 
    "Pernambuco", "Alagoas", "Sergipe", "Bahia",
    "Região Sudeste", "Minas Gerais", "Espírito Santo", "Rio de Janeiro", "São Paulo",
    "Região Sul", "Paraná", "Santa Catarina", "Rio Grande do Sul",
    "Região Centro-Oeste", "Mato Grosso do Sul", "Mato Grosso", "Goiás", "Distrito Federal"
  ),
  categoria1 = "-",
  categoria2 = "-",
  categoria3 = "-"
  )