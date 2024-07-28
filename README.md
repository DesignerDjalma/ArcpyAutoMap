# Ferramenta de Automação ArcGIS

## Descrição

Esta ferramenta foi desenvolvida para aprimorar e automatizar tarefas dentro do ArcGIS Desktop. Com ela, você pode executar uma série de operações geoespaciais de forma eficiente, utilizando bibliotecas específicas do Python. A ferramenta foi projetada para garantir que todos os processos sejam realizados de maneira rápida e precisa, economizando tempo e esforço dos usuários.

## Funcionalidades

- **Atualização Automática do Pip**: Garante que você tenha a versão mais recente do gerenciador de pacotes Python.
- **Instalação de Bibliotecas Essenciais**: Instala automaticamente as bibliotecas necessárias para o funcionamento da ferramenta:
  - PyYAML==5.4.1
  - psycopg2==2.8.6
  - openpyxl==2.6.4
  - lxml==4.6.1
  - python-docx==0.8.11
- **Integração com ArcGIS Desktop**: A ferramenta é totalmente compatível com o ArcGIS Desktop, facilitando a automação de processos geoespaciais complexos.

## Requisitos

Antes de instalar e utilizar a ferramenta, certifique-se de que os seguintes requisitos estejam atendidos:

- **ArcGIS Desktop**: É necessário ter o programa instalado. A ferramenta foi desenvolvida para ser utilizada com o ArcGIS Desktop.
- **Python 2.7**: O Python 2.7 vem como complementa da instalação do ArcGIS Desktop. A instalação do Python deve estar no seguinte caminho: `C:\Python27\ArcGIS10.8\python.exe`.

## Download da Ferramenta

 - [Windows](doc/windows.md)

## Instalação

1. **Download do Instalador**

   Faça o download do instalador da ferramenta a partir da seção de [Releases](https://github.com/DesignerDjalma/ArcpyAutoMap/releases) do repositório no GitHub. 

2. **Executar o Instalador**

   Execute o instalador baixado. Durante a instalação, um script será executado para garantir que todas as bibliotecas Python necessárias sejam instaladas.

3. **Mensagem de Instalação**

   - Durante a instalação, será exibida uma mensagem informando sobre a atualização do pip e a instalação das bibliotecas necessárias.
   - Após a conclusão, uma mensagem confirmará que a instalação foi bem-sucedida.

4. **Instalação Manual (Opcional)**

   Se você preferir instalar as dependências manualmente, siga estas etapas:

   - **Atualizar o Pip**:
     Abra um prompt de comando e execute o seguinte comando para atualizar o pip:
     ```shell
     "C:\Python27\ArcGIS10.8\python.exe" -m pip install --upgrade pip
     ```

   - **Instalar Bibliotecas Python**:
     Em seguida, instale as bibliotecas necessárias com o seguinte comando:
     ```shell
     "C:\Python27\ArcGIS10.8\python.exe" -m pip install PyYAML==5.4.1 psycopg2==2.8.6 openpyxl==2.6.4 lxml==4.6.1 python-docx==0.8.11
     ```

## Uso

Após a instalação, a ferramenta estará pronta para uso dentro do ArcGIS Desktop. Você pode acessá-la navegando na pasta de instalação a partir do Catalog dentro do ArcMap.

## Contribuições

Contribuições são bem-vindas! Se você deseja colaborar com o desenvolvimento da ferramenta, siga os passos abaixo:

1. Faça um fork do repositório.
2. Crie uma branch para a sua funcionalidade (`git checkout -b funcionalidade/incrivel`).
3. Faça commit das suas alterações (`git commit -am 'Adiciona funcionalidade incrível'`).
4. Faça push para a branch (`git push origin funcionalidade/incrivel`).
5. Abra um Pull Request.

## Suporte

Se você encontrar problemas ou tiver dúvidas sobre a ferramenta, abra uma issue no nosso repositório no GitHub ou entre em contato diretamente.

## Licença

Este projeto está licenciado sob os termos da licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.
