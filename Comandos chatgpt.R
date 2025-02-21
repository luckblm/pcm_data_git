Sem Categorias
Se eu fosse você, eu escreveria o pedido assim:  
  
**Pedido:**  
"Extraia os valores da tabela e crie vetores em R com as 
taxas de mortalidade infantil para os anos `2021`, `2022` e `2023`
armazenados dentro de `d1$`. Além disso, crie um vetor para o ranking de `2023`,
armazenado dentro de `d2$`, mas nomeado apenas como `d2$`2023`` (sem `Ranking_` no nome).
Os vetores devem conter exatamente os mesmos valores da tabela, na mesma ordem, 
e os anos devem ser escritos entre crases (``)."  

Dessa forma, o pedido fica claro e direto, sem margem para interpretação errada.


### Prompt contextualizado

Agora Crie a mesma estrutura para número de matrículas na pré escola - ensino regular por dependência administrativa.

Crie os vetores em R para as colunas de dados, Total, Federal, Estadual, Municipal e Privada, para cada um dos anos, 2019, 2020, 2021, 2022 e 2023, estruturados de forma organizada e com os valores dos anos entre `. 
Em cada vetor, os valores devem corresponder aos dados de cada ano, sem modificar os dados originais, incluindo os valores que contêm "-", mantenha como está e não substitua-os por zero. 

Acrescente ainda um vetor com o ranking de 2023, de acordo com a coluna correspondente, mantendo apenas os números, sem caracteres especiais. 

Os dados devem estar estruturados como elucidado no modelo, a baixo:
`
# 2023
total_`2023` <- c(1282369, 133559, 4810, 5152, 4576, 7042, 78679, 3989, 58552, 10524, 91654, 39142, 56142, 48557, 23567, 14723, 25876, 38069, 130454, 62782, 296135, 49873, 114261, 141275, 122763, 270113, 323126)
federal_`2023` <- c(813, "-", "-", "-", "-", "-", "-", "-", "-", "-", 115, "-", "-", 12, 52, 49, 34, 23, 39, 38, 128, "-", 3, 42, 31, "-", 10)
estadual_`2023` <- c(2610, 114, "-", "-", "-", "-", "-", "-", 58, "-", 1100, "-", "-", 385, 8, "-", "-", 38, 56, 96, 9, 1, 35, "-", 36, 21, 7, 36)
municipal_`2023` <- c(454590, 129438, 1983, 2945, 8763, 6230, 76467, 3402, 26182, 12729, 53010, 12248, 51557, 33419, 15798, 38857, 88791, 70129, 179389, 77577, 142321, 86089, 73749, 51983, 155429, 20714, 184607)
privada_`2023` <- c(802876, 7598, 2778, 307, 1313, 1412, 77194, 3987, 26757, 10926, 87329, 24134, 46085, 23357, 10675, 20083, 19973, 8882, 10814, 32531, 56177, 31845, 36201, 47578, 7633, 75523, 46147)
ranking_`2023` <- c("-", 5, 24, 25, 20, 26, 11, 27, 23, 2, 9, 17, 8, 19, 15, 10, 16, 22, 5, 1, 2, 14, 3, 1, 3, 4, 7, 6, 4, 18, 13, 12, 21)


e devem ser agrupados corretamente em cada um dos anos indicados na tabela 2019, 2020, 2021, 2022 e 2023, nesta ordem, substituindo com os dados do conjunto de indicado  entre {}:

Conjunto de dados:

{

}


