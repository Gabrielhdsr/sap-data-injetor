import xml.etree.ElementTree as ET
import difflib
import webbrowser
import os

def get_xml_structure(element, level=0):
    lines = []
    indent = "  " * level
    # Pegando a tag e limpando o namespace se ele for muito poluído (opcional)
    tag = element.tag
    
    attrs = " ".join([f'{k}="{v}"' for k, v in element.attrib.items()])
    attr_str = f" {attrs}" if attrs else ""
    
    lines.append(f"{indent}<{tag}{attr_str}>")
    for child in element:
        lines.extend(get_xml_structure(child, level + 1))
    lines.append(f"{indent}</{tag}>")
    return lines

def compare_xmls_html(file1, file2):
    try:
        # Tenta ler os arquivos
        tree1 = ET.parse(file1)
        tree2 = ET.parse(file2)
        
        struct1 = get_xml_structure(tree1.getroot())
        struct2 = get_xml_structure(tree2.getroot())

        # O segredo está aqui: use fromdesc e todesc para os labels das colunas
        diff_generator = difflib.HtmlDiff()
        diff_html = diff_generator.make_file(
            struct1, 
            struct2, 
            fromdesc='Modelo Original SAP', 
            todesc='Seu Arquivo Gerado'
        )

        output_path = "diff_sap.html"
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(diff_html)

        print(f"Sucesso! Comparação gerada em: {output_path}")
        webbrowser.open(f"file://{os.path.realpath(output_path)}")

    except Exception as e:
        print(f"Erro: {e}")

# Lembre-se do 'r' por causa do seu erro de unicode anterior
file_sap = r"C:\Users\gsribeiro\Desktop\sap-data-injetor\CAR.SUP.001 - Mestre de Material_ROH_1 1.xml"
file_gerado = r"C:\Users\gsribeiro\Desktop\sap-data-injetor\saida\CAR_SUP_001\CAR_SUP_001_Parte_01.xml"

compare_xmls_html(file_sap, file_gerado)