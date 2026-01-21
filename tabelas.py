import os, sys, json, unicodedata, re
from lxml import etree
import pandas as pd
from sqlalchemy import create_engine
from urllib.parse import quote_plus

# ==========================================
# CONFIGURAÇÕES
# ==========================================
ABAS_IGNORADAS = {"Lista de campos", "Field List", "Introdução", "Introduction"}
NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
def _q(tag): return f"{{{NS['ss']}}}{tag}"

# ==========================================
# FUNÇÕES
# ==========================================
def normalizar_nome_tabela(texto):
    if not texto: return ""
    texto = str(texto).replace("Nº", "N").replace("nº", "n")
    norm = "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    norm = re.sub(r'[^a-zA-Z0-9]', '_', norm.upper())
    return re.sub(r'_+', '_', norm).strip('_')

def conectar_db():
    try:
        with open("db_config.json", "r", encoding="utf-8") as f: cfg = json.load(f)
        user, pw = quote_plus(cfg['username']), quote_plus(cfg['password'])
        uri = f"mssql+pyodbc://{user}:{pw}@{cfg['server']}/{cfg['database']}?driver=ODBC+Driver+18+for+SQL+Server&TrustServerCertificate=yes"
        return create_engine(uri)
    except Exception as e:
        print(f"[ERRO] Não conectou no banco: {e}")
        sys.exit(1)

# ==========================================
# INSPEÇÃO
# ==========================================
def inspecionar_tabelas(nome_input):
    print(f"\n{'='*60}")
    print(f"🔍 INSPEÇÃO DE TABELAS: {nome_input}")
    print(f"{'='*60}")

    prefixo = os.path.basename(nome_input).split(" - ")[0].replace(".", "_")
    print(f"📌 Prefixo Extraído: {prefixo}\n")

    engine = conectar_db()
    
    layout_path = os.path.join("layouts", nome_input if nome_input.endswith(".xml") else f"{nome_input}.xml")
    if not os.path.exists(layout_path):
        print(f"❌ Arquivo não encontrado: {layout_path}"); return

    # --- AQUI ESTÁ A CORREÇÃO: MODO RECOVER=TRUE ---
    parser = etree.XMLParser(recover=True) 
    tree = etree.parse(layout_path, parser)
    
    encontradas = 0
    total_abas = 0
    tabela_mestra_info = None

    print(f"{'ABA (Excel)':<40} | {'TABELA ESPERADA (SQL)':<45} | {'STATUS'}")
    print("-" * 100)

    for ws in tree.xpath("//ss:Worksheet", namespaces=NS):
        aba_nome = ws.get(_q("Name"))
        if any(ign in aba_nome for ign in ABAS_IGNORADAS): continue
        
        total_abas += 1
        
        # 1. Aplica a transformação
        sufixo = normalizar_nome_tabela(aba_nome)
        nome_tabela = f"{prefixo}_{sufixo}"
        
        status = "❓"
        
        try:
            # 2. Testa existência (rápido, sem baixar dados)
            df_cols = pd.read_sql(f"SELECT TOP 0 * FROM dbo.[{nome_tabela}]", engine)
            status = "✅ OK"
            encontradas += 1
            
            # Se for a aba mestra (Dados Gerais), guarda as colunas para mostrar depois
            if "DADOS_GERAIS" in sufixo:
                tabela_mestra_info = (nome_tabela, df_cols.columns.tolist())
                
        except:
            status = "❌ NÃO EXISTE"

        print(f"{aba_nome:<40} | {nome_tabela:<45} | {status}")

    print("-" * 100)
    print(f"📊 RESUMO: {encontradas} encontradas de {total_abas} abas processáveis.")

    # MOSTRA A CHAVE DA MESTRA
    if tabela_mestra_info:
        nome_mestra, colunas = tabela_mestra_info
        print(f"\n🔑 ANÁLISE DA TABELA MESTRA ({nome_mestra}):")
        if colunas:
            print(f"   1ª Coluna (Será a Chave): [ {colunas[0]} ]")
            print(f"   Primeiras 5 Colunas: {colunas[:5]}")
        else:
            print("   ⚠️ Tabela existe mas não tem colunas!")
    else:
        print("\n⚠️ ALERTA: Tabela 'Dados gerais' não foi encontrada! O script principal vai falhar.")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python tabelas.py \"NOME_DO_ARQUIVO.xml\"")
    else:
        inspecionar_tabelas(sys.argv[1])