# Projeto de Geração de Base de Dados e Exportação para Excel

## Descrição

Este projeto tem como objetivo criar uma estrutura de base de dados para a atualização do produto **Pará no Contexto Nacional**, utilizando a linguagem R. O projeto organiza os dados em diferentes categorias e os exporta para um arquivo Excel, permitindo a geração de planilhas organizadas por temática, subtemática e região. Isso facilita a análise e o compartilhamento das informações atualizadas sobre o estado do Pará no cenário nacional. 

As planilhas geradas têm como principal finalidade alimentar a base de dados unificada, facilitando a integração e atualização contínua das informações.

## Requisitos

Antes de executar o script, certifique-se de ter os seguintes pacotes instalados no R:

```r
install.packages("tidyverse")
install.packages("openxlsx")
```

## Como executar

1. Clone este repositório ou baixe o arquivo do script R.
2. Abra o script em um ambiente R (RStudio recomendado).
3. Execute o script para gerar as bases de dados e exportá-las para um arquivo Excel.

## Estrutura do Projeto

O script segue os seguintes passos:

1. **Carregar pacotes:** Importa os pacotes `tidyverse` e `openxlsx`.
2. **Criar estrutura base:** Gera um `tibble` contendo regiões, estados e categorias iniciais.
3. **Criar variantes da base:** Divide a base principal em diversas categorias (`d1`, `d2`, `d3`, etc.), diferenciando colunas por temática e indicador.
4. **Criar e salvar arquivos Excel:** Cria um workbook, adiciona abas para cada conjunto de dados e ajusta automaticamente a largura das colunas antes de salvar os arquivos.

## Organização dos Scripts

Os scripts R estão organizados conforme a temática abordada. A seguir, uma breve descrição de cada um:

- **01\_demografia.R**: Processamento e análise de dados demográficos gerais.
- **02\_educacao.R**: Indicadores educacionais, como taxa de escolarização e níveis de ensino.
- **03\_saude.R**: Dados sobre saúde pública, expectativa de vida e cobertura vacinal.
- **04\_inclusaosocial.R**: Estatísticas sobre inclusão social e equidade.
- **05\_previdenciasocial.R**: Informações sobre previdência social e aposentadorias.
- **06\_seguranca.R**: Dados sobre segurança pública e criminalidade.
- **07\_habitacao\_e\_saneamento.R**: Indicadores de habitação e saneamento básico.
- **08\_Economia.R**: Estatísticas econômicas gerais.
- **09\_mercado\_de\_trabalho.R**: Informações sobre emprego, desemprego e mercado de trabalho.
- **10\_produto\_interno\_bruto.R**: Cálculo e análise do PIB por região.
- **11\_financas\_publicas.R**: Indicadores de finanças públicas e orçamento governamental.
- **12\_balanca\_comercial.R**: Estatísticas sobre exportações e importações.
- **13\_meio\_ambiente.R**: Indicadores ambientais e de sustentabilidade.
- **14\_ciencia\_e\_tecnologia.R**: Estatísticas sobre inovação e pesquisa científica.
- **15\_infraestrutura.R**: Dados sobre infraestrutura e desenvolvimento urbano.
- **16\_territorio.R**: Indicadores geográficos e uso do território.

## Localização dos Arquivos Gerados

Os arquivos Excel serão salvos no diretório especificado no comando `saveWorkbook()`. Certifique-se de que o caminho está correto antes de executar o script.

## Contribuição

Sinta-se à vontade para abrir issues e pull requests para sugerir melhorias ou correções no projeto.

## Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

