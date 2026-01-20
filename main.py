import os, sys, json, math, unicodedata, re, datetime
from lxml import etree
import pandas as pd
from sqlalchemy import create_engine
from urllib.parse import quote_plus
from difflib import SequenceMatcher

# Namespaces SAP/Excel
NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
def _q(tag): return f"{{{NS['ss']}}}{tag}"
ABAS_IGNORADAS = {"Lista de campos", "Field List", "Introdução", "Introduction"}
CHUNK_SIZE = 1500 

def remover_acentos(texto):
    if not texto: return ""
    return "".join(c for c in unicodedata.normalize('NFD', texto)
                   if unicodedata.category(c) != 'Mn').upper().strip()

def registrar_log(prefixo, mensagens):
    os.makedirs("logs", exist_ok=True)
    data_hora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nome_arquivo = f"logs/{prefixo}_{data_hora}.txt"
    with open(nome_arquivo, "w", encoding="utf-8") as f:
        f.write("\n".join(mensagens))
    print(f"\n[LOG] Auditoria salva em: {nome_arquivo}")

def conectar_db():
    with open("db_config.json", "r", encoding="utf-8") as f:
        cfg = json.load(f)
    user, pw = quote_plus(cfg['username']), quote_plus(cfg['password'])
    uri = f"mssql+pyodbc://{user}:{pw}@{cfg['server']}/{cfg['database']}?driver=ODBC+Driver+18+for+SQL+Server&TrustServerCertificate=yes"
    return create_engine(uri)

def validar_dna_estrito(engine, tabela, campos_xml):
    campos_xml_limpos = [remover_acentos(c) for c in campos_xml if c and len(c) > 1]
    if not campos_xml_limpos:
        return False, "Aba sem colunas tecnicas na linha 5"
    try:
        df_cols = pd.read_sql(f"SELECT TOP 0 * FROM dbo.[{tabela}]", engine)
        cols_banco = [remover_acentos(c) for c in df_cols.columns.tolist()]
        matches = [c for c in campos_xml_limpos if c in cols_banco]
        if len(matches) == len(campos_xml_limpos):
            return True, f"DNA 100% OK ({len(matches)} colunas)"
        else:
            faltantes = [c for c in campos_xml_limpos if c not in cols_banco]
            return False, f"DNA incompleto. Faltam: {faltantes}"
    except Exception as e:
        return False, f"Erro SQL: {str(e)}"

def buscar_tabela_inteligente(engine, prefixo, nome_aba, campos_xml, log_exec):
    aba_norm = remover_acentos(nome_aba).replace(" ", "_")
    query = f"SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME LIKE '{prefixo}%'"
    try:
        tabelas_banco = pd.read_sql(query, engine)['TABLE_NAME'].tolist()
        for tabela in tabelas_banco:
            sufixo = tabela.replace(f"{prefixo}_", "")
            if SequenceMatcher(None, aba_norm, sufixo).ratio() >= 0.60:
                passou, motivo = validar_dna_estrito(engine, tabela, campos_xml)
                if passou:
                    log_exec.append(f"  [SUCESSO] Aba '{nome_aba}' casou com '{tabela}'")
                    return tabela
                log_exec.append(f"  [REPROVADA] Aba '{nome_aba}' -> Tabela '{tabela}': {motivo}")
    except: pass
    return None

def processar_layout(nome_input):
    log_exec = [f"=== EXECUCAO: {datetime.datetime.now()} ===", f"Arquivo: {nome_input}"]
    engine = conectar_db()
    base_name = os.path.basename(nome_input).split(" - ")[0]
    prefixo = base_name.replace(".", "_")
    log_exec.append(f"Prefixo Detectado: {prefixo}\n")
    
    layout_path = os.path.join("layouts", nome_input if nome_input.endswith(".xml") else f"{nome_input}.xml")
    parser = etree.XMLParser(recover=True, remove_comments=True)
    tree_template = etree.parse(layout_path, parser)
    
    dados_por_aba, headers_por_aba, max_rows = {}, {}, 0

    print(f"\n--- PROCESSANDO: {prefixo} ---")
    for ws in tree_template.xpath("//ss:Worksheet", namespaces=NS):
        aba_nome = ws.get(_q("Name"))
        if any(x in aba_nome for x in ABAS_IGNORADAS): continue
        
        table_node = ws.find(".//ss:Table", namespaces=NS)
        rows_xml = table_node.findall("ss:Row", namespaces=NS) if table_node is not None else []
        if len(rows_xml) < 5: continue
        
        header_xml = [c.find("ss:Data", namespaces=NS).text.strip() if c.find("ss:Data", namespaces=NS) is not None else None 
                      for c in rows_xml[4].findall("ss:Cell", namespaces=NS)]
        
        tabela = buscar_tabela_inteligente(engine, prefixo, aba_nome, header_xml, log_exec)
        
        if tabela:
            print(f"    [MATCH] Aba '{aba_nome}' -> Tabela '{tabela}'")
            df = pd.read_sql(f"SELECT * FROM dbo.[{tabela}]", engine).astype(str).replace({'nan': '', 'None': ''})
            dados_por_aba[aba_nome], headers_por_aba[aba_nome] = df, header_xml
            max_rows = max(max_rows, len(df))
            log_exec.append(f"    -> Registros: {len(df)}")

    if max_rows == 0:
        registrar_log(prefixo, log_exec)
        return

    # CRIAÇÃO DA PASTA COM O PREFIXO DENTRO DE SAIDA
    pasta_saida_prefixo = os.path.join("saida", prefixo)
    os.makedirs(pasta_saida_prefixo, exist_ok=True)

    num_chunks = math.ceil(max_rows / CHUNK_SIZE)
    for i in range(num_chunks):
        start, end = i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE
        current_tree = etree.parse(layout_path, etree.XMLParser(recover=True))
        
        for aba_nome, df_full in dados_por_aba.items():
            df_chunk = df_full.iloc[start:end]
            ws_node = current_tree.xpath(f"//ss:Worksheet[@ss:Name='{aba_nome}']", namespaces=NS)[0]
            table_xml = ws_node.find(".//ss:Table", namespaces=NS)
            rows = table_xml.findall("ss:Row", namespaces=NS)
            for r in range(len(rows) - 1, 7, -1): table_xml.remove(rows[r])

            for _, sql_row in df_chunk.iterrows():
                row_node = etree.SubElement(table_xml, _q("Row"))
                for idx, col_name in enumerate(headers_por_aba[aba_nome]):
                    cell_node = etree.SubElement(row_node, _q("Cell"), Index=str(idx + 1))
                    val = sql_row[col_name] if col_name in df_chunk.columns else ""
                    data_node = etree.SubElement(cell_node, _q("Data"), {f"{{{NS['ss']}}}Type": "String"})
                    data_node.text = val
            table_xml.set(_q("ExpandedRowCount"), str(len(table_xml.findall("ss:Row", namespaces=NS))))

        out_name = f"{prefixo}_Parte_{i+1:02d}.xml"
        caminho_arquivo = os.path.join(pasta_saida_prefixo, out_name)
        
        xml_str = etree.tostring(current_tree, encoding='unicode', xml_declaration=False)
        xml_str = xml_str.replace('<?mso-application progid="Excel.Sheet"?>', '')
        header = '<?xml version="1.0" encoding="UTF-8"?>\n<?mso-application progid="Excel.Sheet"?>\n'
        
        with open(caminho_arquivo, 'w', encoding='utf-8') as f:
            f.write(header + xml_str.strip())
        
        log_exec.append(f"[OK] Arquivo gerado: {caminho_arquivo}")
        print(f"[OK] Gerado: {caminho_arquivo}")

    registrar_log(prefixo, log_exec)

if __name__ == "__main__":
    if len(sys.argv) < 2: print("Uso: python main.py \"ARQUIVO.xml\"")
    else: processar_layout(sys.argv[1])