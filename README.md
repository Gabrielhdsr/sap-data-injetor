#🚀 SAP Data Injetor (XML Spreadsheet 2003)

Este projeto automatiza a extração de dados do SQL Server e a injeção inteligente em templates XML do SAP (S/4HANA). O script foi desenvolvido para a Facchini, garantindo que grandes volumes de dados sejam processados sem corromper a estrutura exigida pelo SAP.

📋 Pré-requisitos
Certifique-se de ter o Python instalado e execute o comando abaixo para instalar as bibliotecas necessárias:

pip install pandas lxml sqlalchemy pyodbc

Nota: É obrigatório ter o ODBC Driver 18 for SQL Server instalado no sistema para a comunicação com o banco de dados.

⚙️ Configuração (db_config.json)
Antes de rodar, edite o arquivo db_config.json na raiz do projeto:

{ "server": "NOME_DO_SERVIDOR", "database": "NOME_DO_BANCO", "username": "USUARIO", "password": "SENHA" }

🛠️ Como Executar
Salve o template XML original do SAP na pasta /layouts.

No terminal, execute o script passando o nome do arquivo: python main.py "CAR.SUP.002 - Fornecedor Criação.xml"

Resultado: O script criará uma subpasta em /saida com o nome do prefixo (ex: CAR_SUP_002) contendo os arquivos fatiados.

🧠 Lógica de Aprovação de Abas
O script utiliza uma Busca Híbrida para garantir integridade total:

Identificação por Prefixo: Extrai o prefixo do nome do arquivo (ex: CAR.SUP.002 vira CAR_SUP_002).

Match de Nome (Fuzzy > 60%): Compara o nome da aba do XML com o sufixo das tabelas no banco (ignora acentos e espaços).

Validação de DNA (Match 100%): O script lê as colunas técnicas na Linha 5 do XML e verifica se TODAS elas existem na tabela do SQL. Se faltar uma única coluna, a aba é ignorada.

📦 Fatiamento de Arquivos (Chunking)
Para respeitar o limite de 90MB por arquivo no SAP:

Tamanho do Lote: 1.500 registros por arquivo.

Comportamento: Se uma aba tiver 5.000 registros, serão gerados 4 arquivos. Os últimos arquivos de uma sequência podem ser menores, pois contêm apenas o saldo remanescente dos dados.

📝 Auditoria e Logs
Toda execução gera um relatório na pasta /logs:

Sucesso: Lista abas preenchidas e total de registros.

Reprovação: Se uma aba for pulada, o log detalha o motivo (ex: DNA incompleto. Faltam: ['LIFNR']).

Estrutura do Projeto
├── main.py # Script principal ├── db_config.json # Configurações de banco ├── layouts/ # Templates (Input) ├── saida/ # XMLs gerados (Output) └── logs/ # Histórico de auditoria