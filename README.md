# Ferramenta de Automação ArcGIS
![Logo animada da Ferramenta Arcpy Automap em animação de respiração em frente a um globo terestre, imagem em estilo cartoon.](https://github.com/DesignerDjalma/ArcpyAutoMap/blob/main/doc/arcpyautomap.gif)
## Descrição

Esta ferramenta foi desenvolvida para aprimorar e automatizar tarefas de produção e análise de mapas de regularização fundiária dentro do ArcGIS Desktop. Com ela, é possível executar uma série de operações geoespaciais de forma automatizada e prática, utilizando uma interface simples e intuitiva. A ferramenta foi projetada para garantir que todos os processos sejam realizados de maneira rápida e precisa, economizando tempo e esforço dos usuários.  

## Showcase

### Interface da Ferramente e Mapa Exportado

![Imagem ilustrando a Interface da Ferramente e Mapa Exportado](https://github.com/DesignerDjalma/ArcpyAutoMap/blob/main/doc/show_case.png) 

## Funcionalidades
- **Integração com ArcGIS Desktop**: A ferramenta é totalmente compatível com o ArcGIS Desktop e ArcGIS Pro, facilitando a automação de processos geoespaciais complexos.
- **Exportar Mapas em PDF.** A partir de um layout modelo é possível exportar mapas com diversas informações.
- **Exportar Relatórios em DOCX.** A partir de um modelo padrão é possivel exportar um documento de relatório.
- **Exportar Planilhas em XLSX.** A partir de uma planilha padrão é possivel fazer o controle de quais e quantos processos foram feitos.
- **Atualizar Banco de dados PostgreSQL.**

### Principais Funcionalidades Especificas
- **Importa shapefiles** de diretorios na maquina e adiciona ao projeto ao dataframe **Principal**, ou utilizada camadas do projeto. Adicionando-as como **Area de Interesse**, porém com o nome original que foi importado ou que está no projeto, no dataframe **Principal**
- **Ignora importações desnecessárias** caso já estajam no projeto
- **Aplica simbologia** na **Area de Interesse** adequada, podendo varias entre: _Point_, _Line_, _Polyline_ ou _Polygon_
- **Aplica Zoom**, e tranforma a escala para um valor adequado (fechada em zeros)
- **Afastas textos** de outras feições que possam sobrepor a geometria dentro do Layout
- **Obtem os valores da Carta Casdastral** onde se encontra geograficamente a **Area de Interesse**
- **Obtem os valores das Zonas Econômicas Exclusivas (ZEE)** onde se encontra geograficamente a **Area de Interesse**
- **Obtem os valoer das Macrozoneamento Ecológico-Econômico (MZEE)** onde se encontra geograficamente a **Area de Interesse**
- **Aplica Zoom** para o Municipio informado, no dataframe **Mapa de Situação**
- **Atualiza** automaticamente as **Definition Querys** das feições do **Mapa de Situação**
    - Definition Query: Municipios de Interesse
    - Definition Query: Limites Municipais
    - Definition Query: Localidades
- **Atualiza o Layout** no Projeto com base no valores informados na ferramenta.
- **Exporta um Mapa em .PDF** do layout.
- **Exporta um Relatorios em .DOCX** com base no valores informados na ferramenta.
- **Exporta um Planilhas em .XLSX** com base no valores informados na ferramenta.
- **Atualiza o Banco de dados/Shapefile (Local)**, referente as Areas de Interesse, conectado a ferramenta.
- **Atualiza o Banco de dados/GeoDatabase (Local)**, referente as Areas de Interesse,  conectado a ferramenta.
- **Atualiza a Banco de dados/Tabela no Banco de Dados PostgreSQL**, referente as Areas de Interesse, conectado a ferramenta.

### Tabelas para os Bancos de Dados

| ID | MUNICIPIO | SITUAÇÃO | ANO | NUMERO  | INTERESSADO     | IMOVEL     | PARCELA | GEORREF   | DATA       | COMPLEMENTO | CARTA         | ZEE                 | PATH/GEOM    |
|----|-----------|----------|-----|---------|-----------------|------------|---------|-----------|------------|-------------|---------------|---------------------|--------------|
| Integer  | String | String | Integer | Integer | String | String | Integer       | String | Date | String | String | String | String |

#### Descrição dos Campos

- **ID**: Identificação única do registro.
- **MUNICIPIO**: Nome do município.
- **SITUAÇÃO**: Situação do processo de regularização.
- **ANO**: Ano de referência.
- **NUMERO**: Número do processo.
- **INTERESSADO**: Nome do interessado.
- **IMOVEL**: Nome do imóvel.
- **PARCELA**: Número da parcela.
- **GEORREF**: Tipo de georreferenciamento.
- **DATA**: Data do registro.
- **COMPLEMENTO**: Complemento das informações.
- **CARTA**: Cartas Cadastrais associadas.
- **ZEE**: Zonas de zoneamento ecológico-econômico.
- **PATH/GEOM**: Caminho para o arquivo shapefile associado, ou Geometria.

## Requisitos

Antes de instalar e utilizar a ferramenta, certifique-se de que os seguintes requisitos estejam atendidos:

- **ArcGIS Desktop**: É necessário ter o programa instalado. A ferramenta foi desenvolvida para ser utilizada com o ArcGIS Desktop.
- **Python 2.7**: O Python 2.7 vem como complementa da instalação do ArcGIS Desktop. A instalação do Python deve estar preferencialmente no seguinte caminho: `C:\Python27\ArcGIS10.x\python.exe`. Caso não esteja a instalação manual das bibliotecas deverá conter o caminho especificado do Python 2.7 do seu ArcGIS Desktop 10.x.
- **Atualização do Pip**(Opcional): Garante que você tenha a versão mais recente do gerenciador de pacotes Python.
- **Bibliotecas Essenciais**: As bibliotecas necessárias para o funcionamento da ferramenta:
  - PyYAML==5.4.1
  - psycopg2==2.8.6
  - openpyxl==2.6.4
  - lxml==4.6.1
  - python-docx==0.8.11 ()


## Download da Ferramenta

 - [Windows](doc/windows.md)

## Dados Complementares

- [Base Cartográfica (Shapefiles)](https://drive.google.com/file/d/1o3J3j2Df0bAiNAglx-w_cARaeKUOU5l6/view?usp=drive_link)
- [Mapa de Situação (Shapefiles)](https://drive.google.com/file/d/1qFUI4bz6wsqGchw2QcpXubJYvnBwL49Z/view?usp=drive_link)

## Instalação

1. **Download do Instalador**

   Faça o download do arquivo zip conténdo o instalador da ferramenta a partir da seção de **Download da Ferramenta** logo acima, no GitHub. 

2. **Executar o Instalador**

   Extraia o arquivo zipado e execute o instalador baixado. Durante a instalação, um script será executado para garantir que todas as bibliotecas Python necessárias sejam instaladas.

3. **Mensagem de Instalação**

   - Durante a instalação, será exibida uma mensagem informando sobre a atualização do pip e a instalação das bibliotecas necessárias.
   - Após a conclusão, uma mensagem confirmará que a instalação foi bem-sucedida.
       - Caso ocorra algum erro, a instalação das bibliotecas adicionais não será feita, porém **a extração dos arquivos ocorrerá normalmente**. Caso a ferramenta não esteja utilizável por problemas de bibliotecas não encotradas, a instalação das bibliotecas pode ser feita manualmente no passo seguinte.

4. **Instalação Manual**

   Caso algum erro tenha ocorrido ou se você preferir instalar as dependências manualmente, siga estas etapas:

   - **Atualizar o Pip**:
     Abra um prompt de comando e execute o seguinte comando para atualizar o pip, onde 'x' corresponde a versão (menor) do seu ArcGIS:
     ```shell
     "C:\Python27\ArcGIS10.x\python.exe" -m pip install --upgrade pip
     ```

   - **Instalar Bibliotecas Python**:
     Em seguida, instale as bibliotecas necessárias com o seguinte comando:
     ```shell
     "C:\Python27\ArcGIS10.x\python.exe" -m pip install PyYAML==5.4.1 psycopg2==2.8.6 openpyxl==2.6.4 lxml==4.6.1 python-docx==0.8.11
     ```

## Uso

Após a instalação, a ferramenta estará pronta para uso dentro do ArcGIS Desktop. Você pode acessá-la navegando na pasta de instalação a partir do Catalog dentro do ArcMap. Basta abrir a Python Toolbox **ArcPyAutoMap v1.0.0.pyt** e usa-la da melhor maneira.

## Contribuições

Contribuições são bem-vindas! Se você deseja colaborar com o desenvolvimento da ferramenta.

## Suporte

Se você encontrar problemas ou tiver dúvidas sobre a ferramenta, abra uma issue neste repositório no GitHub.
