import json
import os
from sqlalchemy import create_engine, text
from urllib.parse import quote_plus

def testar_conexao_sap():
    try:
        # Carrega as configurações do seu JSON
        caminho_config = os.path.join(os.path.dirname(__file__), "db_config.json")
        with open(caminho_config, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        
        print("--- Iniciando Teste de Conexão URI ---")
        
        # O PULO DO GATO: Codificar a senha para que o '+' seja interpretado corretamente
        user = quote_plus(cfg['username'])
        pw = quote_plus(cfg['password'])
        srv = cfg['server']
        db = cfg['database']
        
        # Montagem da URI no formato que você solicitou
        # Usando Driver 18 e TrustServerCertificate=yes
        uri = f"mssql+pyodbc://{user}:{pw}@{srv}/{db}?driver=ODBC+Driver+18+for+SQL+Server&TrustServerCertificate=yes"
        
        print(f"Tentando conectar em: {srv}...")
        
        engine = create_engine(uri)
        
        with engine.connect() as conn:
            # Consulta simples para confirmar a identidade no banco
            res = conn.execute(text("SELECT SUSER_NAME(), DB_NAME()")).fetchone()
            print("\n>> SUCESSO ABSOLUTO!")
            print(f">> Logado como: {res[0]}")
            print(f">> Banco atual: {res[1]}")
            return True

    except Exception as e:
        print("\n>> FALHA NA CONEXÃO")
        erro_str = str(e)
        if "18456" in erro_str:
            print("Erro [18456]: Login falhou. Provavelmente o SQL Server não aceitou a senha.")
            print("Dica: Verifique se o '+' na senha foi codificado corretamente ou se o login 'gabriel' permite SQL Auth.")
        elif "ODBC Driver 18" in erro_str:
            print("Erro de Driver: Verifique se o 'ODBC Driver 18' está instalado na sua máquina.")
        else:
            print(f"Erro técnico: {erro_str}")
        return False

if __name__ == "__main__":
    testar_conexao_sap()