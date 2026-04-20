# MS_SISTEMAS

Projeto em Python para mapear e extrair dados publicamente acessiveis de conveniados a partir de endpoints abertos do portal `mssistemas`, com foco em empresas do segmento de assistencia e convenio funerario.

O script faz uma varredura de codigos, identifica estruturas validas, organiza classes e subclasses de conveniados e gera um arquivo Excel consolidado com os resultados.

## Objetivo

Este projeto foi criado para:

- identificar codigos ativos no portal analisado;
- relacionar convenios encontrados na faixa de codigos pesquisada;
- extrair nome do estabelecimento, classe, subclasse e cidade;
- consolidar tudo em uma planilha Excel para analise posterior.

## Como funciona

O arquivo principal e o `Gabetta_explorer.py`. ( Foi colocado esse nome porque o objetivo principal era apenas extrair o convenio Gabetta )

Ele executa, em alto nivel, as seguintes etapas:

1. consulta a faixa de codigos `850` a `900`;
2. verifica quais codigos possuem estrutura valida de conveniados;
3. tenta mapear nomes conhecidos de convenios com base em evidencias observadas no portal;
4. monta uma aba inicial com os codigos oficiais e codigos adjacentes;
5. extrai os conveniados por classe e subclasse;
6. salva tudo no arquivo `gabetta_explorer.xlsx`.

## Tecnologias usadas

- Python 3
- pandas
- urllib.request
- re
- pathlib

## Como executar

1. Instale o Python 3.
2. Instale a dependencia:

```bash
pip install pandas openpyxl
```

3. Execute o script:

```bash
python Gabetta_explorer.py
```

## Arquivos do projeto

- `Gabetta_explorer.py`: script principal de coleta, tratamento e exportacao dos dados.
- `gabetta_explorer.xlsx`: arquivo Excel gerado com os resultados da execucao.

## Estrutura do Excel gerado

O arquivo Excel pode conter:

- `Inicial`: resumo dos codigos oficiais identificados e codigos relacionados;
- `scan_850_900`: resultado bruto da varredura dos codigos analisados;
- `convenios_nomeados`: relacao de convenios identificados e evidencias;
- abas por convenio: listagem de conveniados com nome, classe, subclasse e cidade.

## Exemplo de dados extraidos

Os registros coletados podem incluir campos como:

- `Nome do lugar`
- `classe`
- `subclasse`
- `cidade`

## Observacoes importantes

- Este projeto utiliza informacoes publicamente acessiveis nos endpoints consultados pelo script.
- A identificacao de alguns convenios e baseada em evidencias observadas no portal, podendo haver casos de inferencia e nao de confirmacao oficial.
- O uso dos dados deve respeitar a legislacao aplicavel, a privacidade, os termos da plataforma consultada e as boas praticas de uso responsavel.

## Autor

Gabriel Mayr  
GitHub: `gabrielmayr003`