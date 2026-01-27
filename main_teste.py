import os, sys, json, math, unicodedata, re, datetime, time
from lxml import etree
import pandas as pd
from sqlalchemy import create_engine
from urllib.parse import quote_plus
import xlwings as xw # Adicione no topo do arquivo
    
# ==========================================
# CONFIGURAÇÕES
# ==========================================
TAMANHO_LOTE_CHAVES = 2500 

# Namespaces
NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
# Namespace adicional do Excel (necessário para achar as opções de proteção)
NS_EXCEL = "urn:schemas-microsoft-com:office:excel" 

def _q(tag): return f"{{{NS['ss']}}}{tag}"
ABAS_IGNORADAS = {"Lista de campos", "Field List", "Introdução", "Introduction"}

# ==========================================
# UTILITÁRIOS
# ==========================================
def log_msg(lista_logs, msg, console=True):
    if console:
        print(msg)
    lista_logs.append(msg)

def salvar_log_arquivo(prefixo, lista_logs):
    os.makedirs("logs", exist_ok=True)
    data_hora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nome_arquivo = f"logs/{prefixo}_{data_hora}.txt"
    with open(nome_arquivo, "w", encoding="utf-8") as f:
        f.write("\n".join(lista_logs))
    print(f"\n[LOG SALVO] {nome_arquivo}")

def normalizar_nome_tabela(texto):
    if not texto: return ""
    texto = str(texto).replace("Nº", "N").replace("nº", "n")
    norm = "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    norm = re.sub(r'[^a-zA-Z0-9]', '_', norm.upper())
    return re.sub(r'_+', '_', norm).strip('_')

def xml_safe(value):
    if value is None: return ""
    s = str(value)
    s = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\u007F-\u0084\u0086-\u009F]", "", s)
    return unicodedata.normalize("NFC", s)

def conectar_db():
    try:
        with open("db_config.json", "r", encoding="utf-8") as f: cfg = json.load(f)
        user, pw = quote_plus(cfg['username']), quote_plus(cfg['password'])
        uri = f"mssql+pyodbc://{user}:{pw}@{cfg['server']}/{cfg['database']}?driver=ODBC+Driver+18+for+SQL+Server&TrustServerCertificate=yes"
        return create_engine(uri)
    except Exception as e:
        print(f"[ERRO FATAL] Falha na conexão: {e}")
        sys.exit(1)

def dar_refresh_conexao(engine, log):
    log_msg(log, "--- REFRESH NO BANCO ---")
    try:
        engine.dispose()
        pd.read_sql("SELECT 1", engine)
        log_msg(log, "[OK] Conexão renovada!")
    except Exception as e:
        log_msg(log, f"[AVISO] Refresh falhou: {e}")

# ==========================================
# MANIPULAÇÃO XML
# ==========================================
def obter_header_tecnico(ws):
    table = ws.find(".//ss:Table", namespaces=NS)
    if table is None: return None
    rows = table.findall("ss:Row", namespaces=NS)
    if len(rows) < 5: return None
    
    header = []
    for c in rows[4].findall("ss:Cell", namespaces=NS):
        data = c.find("ss:Data", namespaces=NS)
        val = data.text.strip().upper() if data is not None and data.text else ""
        header.append(val)
    return [h for h in header if h] if header else None

def desproteger_aba(ws):
    """
    Remove as tags que protegem a planilha dentro de WorksheetOptions.
    """
    # Procura a tag <x:WorksheetOptions>
    options = ws.find(f".//{{{NS_EXCEL}}}WorksheetOptions")
    
    if options is not None:
        # Tags que causam o bloqueio da aba
        tags_bloqueio = [
            f"{{{NS_EXCEL}}}ProtectObjects",
            f"{{{NS_EXCEL}}}ProtectScenarios",
            f"{{{NS_EXCEL}}}Protected",
            f"{{{NS_EXCEL}}}ProtectWindows"
        ]
        
        # Itera sobre os filhos e remove se for tag de proteção
        # Usamos list(options) para criar uma cópia e poder remover enquanto iteramos
        for child in list(options):
            if child.tag in tags_bloqueio:
                options.remove(child)

def obter_tipos_coluna(ws):
    """
    Localiza a tabela na aba e varre a 6ª linha (índice 5)
    para capturar as definições técnicas de tipo do SAP.
    """
    table = ws.find(".//ss:Table", namespaces=NS)
    rows = table.findall("ss:Row", namespaces=NS)
    
    if len(rows) < 6:
        return []

    tipos = []
    # Varremos as células da linha 6
    for c in rows[5].findall("ss:Cell", namespaces=NS):
        data = c.find("ss:Data", namespaces=NS)
        # Se a célula existir, pegamos o texto (ex: ENU;13;3), senão, string vazia
        val = data.text.strip().upper() if data is not None and data.text else ""
        tipos.append(val)
    return tipos

def to_number_for_xml(value):
    if value is None or str(value).strip() == "" or str(value).lower() == "nan":
        return None
    
    # Remove pontos de milhar e troca a vírgula decimal por ponto
    s = str(value).strip().replace(".", "").replace(",", ".")
    
    try:
        n = float(s)
        # Se for inteiro (ex: 10.0), vira "10". Se for decimal, mantém as casas.
        if n.is_integer():
            return str(int(n))
        else:
            return ("%.10f" % n).rstrip("0").rstrip(".")
    except:
        return None
    
def preencher_aba_xml(ws, df, header_xml):
    table = ws.find(".//ss:Table", namespaces=NS)
    rows = table.findall("ss:Row", namespaces=NS)
    for r in range(len(rows) - 1, 7, -1): table.remove(rows[r])

    tipos = obter_tipos_coluna(ws) # Pegamos a lista de tipos da linha 6

    for _, row in df.iterrows():
        row_node = etree.SubElement(table, _q("Row"))
        for idx, col_name in enumerate(header_xml):
            val = row[col_name] if col_name in df.columns else ""
            
            # Identificamos o tipo para esta coluna específica
            tipo_sap = tipos[idx] if idx < len(tipos) else ""
            is_num = "NUMERO" in tipo_sap or tipo_sap.startswith("ENU;")

            cell = etree.SubElement(row_node, _q("Cell"), Index=str(idx + 1))
            
            if is_num:
                num_limpo = to_number_for_xml(val)
                if num_limpo:
                    # AQUI É O PULO DO GATO: Type="Number"
                    etree.SubElement(cell, _q("Data"), {f"{{{NS['ss']}}}Type": "Number"}).text = num_limpo
                else:
                    etree.SubElement(cell, _q("Data"), {f"{{{NS['ss']}}}Type": "String"}).text = ""
            else:
                # Texto normal
                etree.SubElement(cell, _q("Data"), {f"{{{NS['ss']}}}Type": "String"}).text = xml_safe(val)
    
    table.set(_q("ExpandedRowCount"), str(len(table.findall("ss:Row", namespaces=NS))))

def sanitizar_arquivos_com_excel(pasta_saida):
    print(f"\n--- TENTANDO SANITIZAÇÃO COM XLWINGS na pasta: {pasta_saida} ---")
    
    arquivos = [f for f in os.listdir(pasta_saida) if f.endswith(".xml")]
    if not arquivos:
        return

    # 1. Inicia o Excel de forma invisível
    # 'app' aqui é o processo do Excel
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False

    try:
        for arquivo in arquivos:
            caminho_completo = os.path.join(os.getcwd(), pasta_saida, arquivo)
            print(f"Sanitizando: {arquivo}...", end="")
            
            try:
                # 2. Abre o arquivo
                wb = app.books.open(caminho_completo)
                
                # 3. Salva (isso força o Excel a reescrever o cabeçalho)
                wb.save()
                
                # 4. Fecha a planilha
                wb.close()
                print(" [OK]")
            except Exception as e:
                print(f" [ERRO no arquivo] {e}")
        
        print("--- SANITIZAÇÃO CONCLUÍDA ---")
    except Exception as e:
        print(f"Erro no processo do Excel: {e}")
    finally:
        # 5. Garante que o Excel feche mesmo se der erro
        app.quit()
        
# ==========================================
# EXECUÇÃO PRINCIPAL
# ==========================================
def processar_layout(nome_input):
    t_inicio = datetime.datetime.now()
    log = []
    
    log_msg(log, f"=== INICIO: {t_inicio} ===")
    prefixo = os.path.basename(nome_input).split(" - ")[0].replace(".", "_")
    log_msg(log, f"Prefixo: {prefixo}\n")
    
    engine = conectar_db()
    dar_refresh_conexao(engine, log)

    layout_path = os.path.join("layouts", nome_input if nome_input.endswith(".xml") else f"{nome_input}.xml")
    
    parser = etree.XMLParser(recover=True, remove_comments=True)
    tree_template = etree.parse(layout_path, parser)
    mapa_execucao = {}
    
    # --- FASE 1: VINCULAÇÃO DIRETA ---
    log_msg(log, "\n--- 1. VINCULACAO ---")
    for ws in tree_template.xpath("//ss:Worksheet", namespaces=NS):
        aba_original = ws.get(_q("Name"))
        if any(ign in aba_original for ign in ABAS_IGNORADAS): continue
        
        sufixo = normalizar_nome_tabela(aba_original)
        nome_tabela = f"{prefixo}_{sufixo}"
        
        header_xml = obter_header_tecnico(ws)
        if not header_xml: continue

        try:
            # Verifica se tabela existe
            pd.read_sql(f"SELECT TOP 0 * FROM dbo.[{nome_tabela}]", engine)
            
            # SE ACHOU: Mostra na tela (console=True é padrão)
            log_msg(log, f"[OK] {aba_original} -> {nome_tabela}")
            mapa_execucao[aba_original] = {"tabela": nome_tabela, "header": header_xml}
        except:
            # SE NÃO ACHOU: Esconde da tela (console=False), mas grava no log
            log_msg(log, f"[FALHA] Tabela não encontrada: {nome_tabela}", console=False)

    if not mapa_execucao:
        log_msg(log, "\n[ERRO] Nenhuma tabela vinculada."); salvar_log_arquivo(prefixo, log); return

    # --- FASE 2: DETECÇÃO CEGA DA CHAVE ---
    log_msg(log, "\n--- 2. DETECÇÃO DE CHAVE ---")
    
    aba_mestre = next((aba for aba in mapa_execucao if "DADOS_GERAIS" in normalizar_nome_tabela(aba)), list(mapa_execucao.keys())[0])
    tabela_mestre = mapa_execucao[aba_mestre]["tabela"]
    
    try:
        df_cols = pd.read_sql(f"SELECT TOP 1 * FROM dbo.[{tabela_mestre}]", engine)
        CHAVE_DETECTADA = df_cols.columns[0].upper()
        
        log_msg(log, f"TABELA MESTRE: {tabela_mestre}")
        log_msg(log, f"CHAVE DEFINIDA: {CHAVE_DETECTADA}")

        log_msg(log, "Contando registros...")
        df_ids = pd.read_sql(f"SELECT DISTINCT TOP(1) {CHAVE_DETECTADA} FROM dbo.[{tabela_mestre}] WHERE {CHAVE_DETECTADA} IN ('10001884')", engine)
        ids_unicos = df_ids[CHAVE_DETECTADA].dropna().unique().tolist()
        try: ids_unicos.sort(key=int)
        except: ids_unicos.sort()

        total = len(ids_unicos)
        lotes = math.ceil(total / TAMANHO_LOTE_CHAVES)
        log_msg(log, f"REGISTROS: {total} | ARQUIVOS: {lotes}")

    except Exception as e:
        log_msg(log, f"[ERRO CRITICO] Falha na Mestra: {e}"); salvar_log_arquivo(prefixo, log); return

    # --- FASE 3: GERAÇÃO ---
    log_msg(log, "\n--- 3. GERACAO ---")
    pasta_saida = os.path.join("saida", prefixo)
    os.makedirs(pasta_saida, exist_ok=True)

    for i in range(lotes):
        lote_ids = ids_unicos[i*TAMANHO_LOTE_CHAVES : (i+1)*TAMANHO_LOTE_CHAVES]
        ids_sql = "', '".join(map(str, lote_ids))
        
        tree_lote = etree.parse(layout_path, etree.XMLParser(recover=True))
        
        for aba, meta in mapa_execucao.items():
            tabela = meta["tabela"]
            try:
                query = f"SELECT * FROM dbo.[{tabela}] WHERE {CHAVE_DETECTADA} IN ('{ids_sql}')"
                df = pd.read_sql(query, engine)
                df = df.where(pd.notnull(df), None)
            except:
                query = f"SELECT * FROM dbo.[{tabela}]"
                df = pd.read_sql(query, engine)
                df = df.where(pd.notnull(df), None)

            df.columns = [c.upper() for c in df.columns]
            if CHAVE_DETECTADA in df.columns: df = df.sort_values(by=CHAVE_DETECTADA)
            
            ws_node = None
            for ws in tree_lote.xpath("//ss:Worksheet", namespaces=NS):
                if ws.get(_q("Name")) == aba:
                    ws_node = ws
                    break
            
            if ws_node is not None:
                # --- AQUI ESTÁ A MÁGICA ---
                # Remove a proteção antes de preencher
                desproteger_aba(ws_node) 
                preencher_aba_xml(ws_node, df, meta["header"])

        out_name = f"{prefixo}_Parte_{i+1:02d}.xml"
        
        xml_str = etree.tostring(tree_lote, encoding='unicode', xml_declaration=False).replace('<?mso-application progid="Excel.Sheet"?>', '')
        with open(os.path.join(pasta_saida, out_name), 'w', encoding='utf-8') as f:
            f.write('<?xml version="1.0">\n<?mso-application progid="Excel.Sheet"?>\n' + xml_str.strip())
            
        log_msg(log, f"[OK] {out_name}")
        
    sanitizar_arquivos_com_excel(pasta_saida)
    t_fim = datetime.datetime.now()
    log_msg(log, f"\nTEMPO TOTAL: {t_fim - t_inicio}")
    salvar_log_arquivo(prefixo, log)

if __name__ == "__main__":
    if len(sys.argv) < 2: print("Uso: python main.py \"ARQUIVO.xml\"")
    else: processar_layout(sys.argv[1])