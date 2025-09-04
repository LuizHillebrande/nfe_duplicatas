import os
import xml.etree.ElementTree as ET
import re

pasta_xml = r"F:/2025/07.2025/NOTAS/MABE ESTRUTURAS/ENTRADAS"

def limpar_caracteres_invalidos(texto):
    """
    Remove caracteres inválidos para XML (controle, não UTF-8 válidos, etc).
    Mantém apenas caracteres válidos para XML 1.0.
    """
    # faixa de caracteres válidos para XML: https://www.w3.org/TR/xml/#charsets
    return re.sub(
        r'[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]',
        '', texto
    )

for arquivo in os.listdir(pasta_xml):
    if arquivo.lower().endswith(".xml"):
        caminho = os.path.join(pasta_xml, arquivo)
        try:
            ET.parse(caminho)
        except ET.ParseError as e:
            # Extrair linha e coluna do erro
            m = re.search(r'line (\d+), column (\d+)', str(e))
            if m:
                linha_err = int(m.group(1))
                col_err = int(m.group(2))
            else:
                linha_err = col_err = None

            print(f"❌ Erro no XML: {arquivo} | {e}")

            # Tentar corrigir
            with open(caminho, 'r', encoding='utf-8', errors='ignore') as f:
                linhas = f.readlines()

            if linha_err and linha_err <= len(linhas):
                # remover caractere problemático na coluna
                linha_original = linhas[linha_err - 1]
                if col_err and col_err <= len(linha_original):
                    linha_corrigida = linha_original[:col_err - 1] + linha_original[col_err:]
                    linhas[linha_err - 1] = linha_corrigida
                    print(f"🛠 Corrigido caractere na linha {linha_err}, coluna {col_err}")

            # Limpar caracteres inválidos de todas as linhas
            linhas = [limpar_caracteres_invalidos(l) for l in linhas]

            # Salvar arquivo corrigido (sobrescreve original)
            with open(caminho, 'w', encoding='utf-8') as f:
                f.writelines(linhas)

            # Testar novamente
            try:
                ET.parse(caminho)
            except ET.ParseError as e2:
                print(f"❌ Ainda com erro: {arquivo} | {e2}")
            else:
                print(f"✅ Corrigido: {arquivo}")

        else:
            print(f"✅ OK: {arquivo}")
