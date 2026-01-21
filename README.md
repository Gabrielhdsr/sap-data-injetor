#🚀 SAP Data Injetor (XML Spreadsheet 2003)
Este projeto automatiza a extração de dados do SQL Server e a injeção em templates XML do SAP (S/4HANA). Desenvolvido para a Facchini, o script garante integridade referencial, baixo consumo de memória e segue rigorosamente a formatação exigida pelo SAP Migration Cockpit.

📋 Pré-requisitos
Certifique-se de ter o Python instalado e execute o comando abaixo no terminal:

pip install pandas lxml sqlalchemy pyodbc

⚠️ Importante: É obrigatório ter o ODBC Driver 18 for SQL Server instalado no Windows para a comunicação com o banco de dados.

⚙️ Configuração (db_config.json)
Edite o arquivo db_config.json na raiz do projeto com as credenciais do banco:

{ "server": "SEU_SERVIDOR", "database": "SEU_BANCO", "username": "USUARIO", "password": "SENHA" }

🛠️ Ferramentas Disponíveis
1. Inspecionar Tabelas (Check de Segurança)
Antes de processar milhares de registros, use este script para validar se os nomes das tabelas no banco seguem o novo padrão e se a chave primária será detectada corretamente.

python tabelas.py "NOME_DO_ARQUIVO.xml"

2. Gerar XMLs (Execução Principal)
Processa os dados em lotes (chunks) e gera os arquivos finais na pasta de saída.

python main.py "NOME_DO_ARQUIVO.xml"

🧠 Lógica do "Sniper"
O script foi reescrito para ser totalmente autônomo, eliminando configurações manuais a cada novo layout:

Vinculação Direta (Aba -> Tabela): O script normaliza o nome da aba do Excel e busca a tabela exata no banco. Regra: Remove acentos, transforma "Nº" em "N" e troca espaços/caracteres especiais por "_". Exemplo: Aba "Nºs identificação fiscal" vira a tabela "PREFIXO_NS_IDENTIFICACAO_FISCAL".

Auto-Detecção de Chave Primária: O script não precisa mais de uma lista prévia (LIFNR, KUNNR, etc). Ele identifica a aba Mestra (ex: "Dados gerais"), lê a 1ª Coluna dessa tabela no SQL e a define automaticamente como a chave âncora para todo o projeto.

Carga Sob Demanda: Diferente de versões anteriores, o script não carrega o banco inteiro na memória. Ele baixa apenas a lista de IDs e faz consultas fracionadas (SELECT WHERE ID IN ...), permitindo processar volumes massivos de dados sem lentidão ou crash.

📦 Fatiamento de Arquivos
Para respeitar os limites de tamanho do SAP e garantir a integridade:

Lote Padrão: 500 Chaves (Fornecedores/Clientes) por arquivo.

Integridade Total: Todos os dados de um mesmo ID (Endereços, Bancos, Contatos) são mantidos no mesmo arquivo XML, evitando quebras de referência durante a importação no SAP.

📝 Auditoria e Logs
O projeto preza por um terminal limpo e um log detalhado:

Terminal: Mostra apenas o status de sucesso e o progresso da geração.

Pasta /logs: Gera um .txt completo com cada tentativa de vinculação, erros de tabelas inexistentes, chave detectada e tempo total de execução.

📂 Estrutura do Projeto
├── main.py # Script principal de processamento ├── tabelas.py # Script de inspeção e validação ├── db_config.json # Configurações de acesso ao banco ├── layouts/ # Local dos templates XML originais ├── saida/ # Onde os arquivos fatiados serão criados └── logs/ # Histórico detalhado de execuções