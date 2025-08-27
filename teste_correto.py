import tkinter as tk
from tkinter import filedialog, messagebox
import xmltodict
import os
import pandas as pd
from datetime import datetime, timedelta

# --- FUNÇÕES AUXILIARES ---

def formatar_data(data):
    """Converte data para formato DD/MM/YYYY"""
    try:
        return datetime.strptime(data, '%Y-%m-%d').strftime('%d/%m/%Y')
    except ValueError:
        return data

def adicionar_30_dias(data):
    """Retorna data 30 dias após a data de emissão"""
    try:
        dt = datetime.strptime(data, '%Y-%m-%d')
        return (dt + timedelta(days=30)).strftime('%Y-%m-%d')
    except:
        return data

def processar_xml(xml_path, mapa_cfop):
    """Processa um XML: corrige CFOPs e datas de duplicata"""
    with open(xml_path, "rb") as arquivo_xml:
        try:
            nfe_dict = xmltodict.parse(arquivo_xml)
            info_nfe = nfe_dict["NFeLog"]["procNFe"]["NFe"]["infNFe"]

            # --- Corrigir CFOPs ---
            det = info_nfe.get("det")
            if det:
                if isinstance(det, list):
                    for item in det:
                        cfop_atual = item["prod"]["CFOP"]
                        item["prod"]["CFOP"] = mapa_cfop.get(cfop_atual, cfop_atual)
                elif isinstance(det, dict):
                    cfop_atual = det["prod"]["CFOP"]
                    det["prod"]["CFOP"] = mapa_cfop.get(cfop_atual, cfop_atual)

            # --- Corrigir datas de duplicata ---
            cobr = info_nfe.get("cobr")
            dup = cobr.get("dup") if cobr else None

            data_emissao = info_nfe["ide"].get("dhEmi")  # pega dhEmi se existir
            if data_emissao:
                # separa só a parte da data (ano-mês-dia)
                data_emissao = data_emissao.split("T")[0]
            else:
                # caso não exista, define uma data padrão ou pula a nota
                data_emissao = None


            if dup:
                if isinstance(dup, list):
                    for parcela in dup:
                        if "dVenc" not in parcela or not parcela["dVenc"]:
                            parcela["dVenc"] = adicionar_30_dias(data_emissao)
                elif isinstance(dup, dict):
                    if "dVenc" not in dup or not dup["dVenc"]:
                        dup["dVenc"] = adicionar_30_dias(data_emissao)
            else:
                # Se não existir 'cobr', criar estrutura com 1 duplicata
                info_nfe["cobr"] = {"dup": {"nDup": "001", "dVenc": adicionar_30_dias(data_emissao), "vDup": "0.00"}}

            # --- Salvar XML corrigido ---
            with open(xml_path, "w", encoding="utf-8") as f:
                f.write(xmltodict.unparse(nfe_dict, pretty=True, encoding="utf-8"))


            return True

        except Exception as e:
            print(f"Erro ao processar {xml_path}: {e}")
            return False
from urllib.parse import unquote, urlparse
# --- INTERFACE TKINTER ---

def selecionar_pasta():
    pasta = filedialog.askdirectory()
    if pasta:
        entry_pasta.delete(0, tk.END)
        entry_pasta.insert(0, pasta)

def selecionar_excel():
    arquivo = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if arquivo:
        entry_excel.delete(0, tk.END)
        entry_excel.insert(0, arquivo)

def processar_todos():
    pasta = entry_pasta.get()
    excel = entry_excel.get()

    if not os.path.isdir(pasta):
        messagebox.showerror("Erro", "Pasta de XML inválida")
        return
    if not os.path.isfile(excel):
        messagebox.showerror("Erro", "Arquivo Excel inválido")
        return

    # --- Carregar mapeamento CFOP ---
    tabela = pd.read_excel(excel)
    mapa_cfop = dict(zip(tabela['CFOP Origem'].astype(str), tabela['CFOP Destino'].astype(str)))

    from urllib.parse import unquote, urlparse

    # --- Listar todos os XMLs, mesmo se vierem com "file:///" ---
    xmls = []
    print(f"Procurando arquivos XML em: {pasta}")

    for f in os.listdir(pasta):
        if f.lower().endswith(".xml"):
            caminho_completo = os.path.join(pasta, f)

            # Normaliza o caminho (resolve barras invertidas / normais)
            caminho_completo = os.path.normpath(caminho_completo)

            # Se vier como file:/// (Edge), converte
            if caminho_completo.startswith("file:///"):
                caminho_completo = unquote(urlparse(caminho_completo).path)
                if caminho_completo.startswith("/"):
                    caminho_completo = caminho_completo[1:]  # remove a barra inicial extra

            # Verifica se o arquivo existe e não está vazio
            if os.path.isfile(caminho_completo) and os.path.getsize(caminho_completo) > 0:
                xmls.append(caminho_completo)
                print(f"Arquivo encontrado: {caminho_completo}")
            else:
                print(f"Atenção: arquivo não encontrado ou vazio: {caminho_completo}")

    total = len(xmls)
    print(f"Total de XMLs encontrados: {total}")
    sucesso = 0

    for i, xml_path in enumerate(xmls, 1):
        print(f"Processando {i}/{total}: {xml_path}")
        if processar_xml(xml_path, mapa_cfop):
            sucesso += 1
        print(f"[{i}/{total}] Processado: {os.path.basename(xml_path)}")

    print(f"Processamento finalizado. {sucesso}/{total} XMLs processados com sucesso.")



    messagebox.showinfo("Finalizado", f"{sucesso}/{total} XMLs processados com sucesso!")

# --- Construção da interface ---
root = tk.Tk()
root.title("Correção de CFOP e Duplicatas NFe")

tk.Label(root, text="Pasta dos XMLs:").grid(row=0, column=0, padx=5, pady=5)
entry_pasta = tk.Entry(root, width=50)
entry_pasta.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Selecionar", command=selecionar_pasta).grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="Arquivo Excel de mapeamento:").grid(row=1, column=0, padx=5, pady=5)
entry_excel = tk.Entry(root, width=50)
entry_excel.grid(row=1, column=1, padx=5, pady=5)
tk.Button(root, text="Selecionar", command=selecionar_excel).grid(row=1, column=2, padx=5, pady=5)

tk.Button(root, text="Processar XMLs", command=processar_todos, bg="green", fg="white").grid(row=2, column=0, columnspan=3, pady=20)

root.mainloop()
