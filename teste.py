import os
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox

# ---------------- FUNÇÕES ----------------

def limpar_icms_c100_e_c190(linha):
    campos = linha.strip().split("|")

    try:
        if campos[1] == "C100":
            # Garante que tem pelo menos 35 campos
            while len(campos) < 35:
                campos.append("")
            
            campos[21] = "0,00"  # VL_BC_ICMS
            campos[22] = "0,00"  # VL_ICMS
            campos[23] = "0,00"  # VL_BC_ICMS_ST
            campos[24] = "0,00"  # VL_ICMS_ST
            campos[25] = "0,00"  # VL_IPI
            campos[26] = "0,00"  # VL_PIS
            campos[27] = "0,00"  # VL_COFINS
            campos[28] = "0,00"  # VL_PIS_ST
            campos[29] = "0,00"  # VL_COFINS_ST

            print(f"[INFO] C100 zerado com sucesso: {linha.strip()}")

        elif campos[1] == "C190":
            # Garante que tem pelo menos 20 campos
            while len(campos) < 20:
                campos.append("")

            # Checando valor original da ALIQ_ICMS
            aliq_icms_original = campos[4]
            campos[4] = "0,00"  # ALIQ ICMS
            #campos[5] = "0,00"  # VL CONTABIL NUNCA ZERAR
            campos[6] = "0,00"  # VL_ICMS
            campos[7] = "0,00"  # VL_BC_ICMS_ST
            campos[8] = "0,00"  # VL_ICMS_ST
            campos[10] = "0,00" # VL_IPI
            #print(f"[INFO] C190 zerado com sucesso (ALIQ_ICMS antes: {aliq_icms_original}): {linha.strip()}")
            linha_zerada = "|".join(campos)
            print(f"[INFO] C190 zerado com sucesso (ALIQ_ICMS antes: {aliq_icms_original}): {linha_zerada}")

    except IndexError as e:
        print(f"[ERRO] Linha malformada: {linha.strip()} | Erro: {e}")

    return "|".join(campos) + "|\n"


def ler_xml_notas(pasta_xml):
    """
    Lê todos os XMLs da pasta e extrai:
    - chave da nota
    - número da nota
    - data de emissão
    - lista de duplicatas (nDup, dVenc, vDup)
    """
    notas = {}
    print("\n🔎 Lendo XMLs da pasta:", pasta_xml)
    for arquivo in os.listdir(pasta_xml):
        if not arquivo.endswith(".xml"):
            continue
        caminho = os.path.join(pasta_xml, arquivo)
        print(f"\n📄 Processando XML: {arquivo}")
        tree = ET.parse(caminho)
        root = tree.getroot()

        # Pega a chave da NF
        chave = root.find(".//{*}chNFe")
        chave = chave.text.strip() if chave is not None else None
        print("➡️ Chave:", chave)
        

        # Número da NF
        n_nf = root.find(".//{*}nNF").text.strip()
        print("➡️ Número NF:", n_nf)

        # Data de emissão
        dhEmi = root.find(".//{*}dhEmi")
        if dhEmi is None:
            dhEmi = root.find(".//{*}dEmi")
        if dhEmi is not None:
            data_emissao = datetime.fromisoformat(dhEmi.text).date()
            print("➡️ Data emissão (dhEmi/dEmi):", data_emissao)
        else:
            print("⚠️ Nenhuma data de emissão encontrada!")
            data_emissao = None

        # Duplicatas
        duplicatas = []
        for dup in root.findall(".//{*}dup"):
            duplicata = {
                "nDup": dup.find("{*}nDup").text,
                "dVenc": dup.find("{*}dVenc").text,
                "vDup": dup.find("{*}vDup").text
            }
            duplicatas.append(duplicata)
        print("➡️ Duplicatas encontradas:", duplicatas if duplicatas else "Nenhuma")

        notas[chave] = {
            "numero": n_nf,
            "emissao": data_emissao,
            "duplicatas": duplicatas
        }
    return notas

import re
from datetime import datetime, timedelta
def _num2(valor):
    """
    Normaliza número de parcela para 2 dígitos (01..99).
    Remove não-dígitos, converte para int e volta em 2 dígitos.
    Se não der pra converter, retorna '01'.
    """
    s = re.sub(r"\D", "", str(valor or ""))
    try:
        return f"{int(s):02d}"
    except:
        return "01"

def processar_sped(arquivo_sped, notas_xml, saida_sped):
    """
    Lê o SPED, insere registros C140/C141, ajusta C990, 9900 e |9999|.
    """
    print("\n📑 Lendo SPED:", arquivo_sped)
    with open(arquivo_sped, "r", encoding="latin1") as f:
        linhas = f.readlines()

    novas_linhas = []
    bloco9 = []
    dentro_bloco9 = False
    count_c140, count_c141 = 0, 0

    for linha in linhas:
        if linha.startswith("|9001|"):
            dentro_bloco9 = True

        if dentro_bloco9:
            bloco9.append(linha)
            continue

        if linha.startswith("|C990|"):
            continue  # vamos recalcular depois

        campos = linha.strip().split("|")
        if campos[1] in ("C100", "C190"):
            linha = limpar_icms_c100_e_c190(linha)


        novas_linhas.append(linha)
        campos = linha.strip().split("|")

        # Encontrando C100
        if len(campos) > 9 and campos[1] == "C100":
            mod = campos[5].strip()
            chave_atual = campos[9].strip()
            print(f"\n🔍 Encontrado C100 com chave {chave_atual}, modelo {mod}")

            # Apenas NFe (55) e NFCe (65) aceitam duplicatas
            if mod not in ["55", "65"]:
                print(f"⚠️ Modelo {mod} não aceita duplicatas. Ignorando C140/C141.")
                continue

            if chave_atual not in notas_xml:
                print(f"⚠️ Chave {chave_atual} não encontrada nos XMLs. Ignorando C140/C141.")
                continue

            nota = notas_xml[chave_atual]
            duplicatas = nota["duplicatas"]

            # Se não houver duplicatas, cria fictícia
            if not duplicatas:
                emissao = nota["emissao"]
                vl_total = campos[12]
                duplicatas = [{
                    "nDup": "01",
                    "dVenc": (emissao + timedelta(days=30)).strftime("%Y-%m-%d"),
                    "vDup": vl_total
                }]

            qtd_parc = len(duplicatas)
            vl_total = sum([float(d["vDup"].replace(",", ".")) for d in duplicatas])
            vl_total_str = f"{vl_total:.2f}".replace(".", ",")

            # ---------------- C140 ----------------
            ind_emit = "1"  # emissão própria
            ind_tit = "00"  # duplicata
            desc_tit = ""   # só usado se ind_tit=99
            num_tit  = _num2(duplicatas[0].get("nDup", "1"))

            c140 = f"|C140|{ind_emit}|{ind_tit}|{desc_tit}|{num_tit}|{qtd_parc}|{vl_total_str}|\n"
            novas_linhas.append(c140)
            count_c140 += 1
            print("➕ Adicionado C140:", c140.strip())

            # ---------------- C141 ----------------
            for dup in duplicatas:
                dVenc = datetime.fromisoformat(dup["dVenc"]).strftime("%d%m%Y")  # ddmmaaaa
                valor = float(dup['vDup'].replace(",", "."))
                valor_sped = f"{valor:.2f}".replace(".", ",")
                nDup_sped = _num2(dup.get('nDup', '1'))  # garante 2 dígitos SEM 3 dígitos
                c141 = f"|C141|{nDup_sped}|{dVenc}|{valor_sped}|\n"
                novas_linhas.append(c141)
                count_c141 += 1
                print("➕ Adicionado C141:", c141.strip())

    # ---------------- Recalcular C990 ----------------
    qtd_lin_c = sum(1 for l in novas_linhas if l.startswith("|C")) + 1
    c990 = f"|C990|{qtd_lin_c}|\n"
    print(f"\n♻️ Recalculado C990: {c990.strip()}")

    ult_c_index = max((i for i, l in enumerate(novas_linhas) if l.startswith("|C")), default=-1)
    if ult_c_index >= 0:
        novas_linhas_corrigidas = (
            novas_linhas[:ult_c_index + 1] +
            [c990] +
            novas_linhas[ult_c_index + 1:]
        )
    else:
        novas_linhas_corrigidas = [c990] + novas_linhas

    # ---------------- Atualizar bloco 9 ----------------
    bloco9_atualizado, count_9900, count_lin9 = atualizar_bloco9(bloco9, count_c140, count_c141)

    # Junta tudo: registros normais + bloco 9 atualizado
    novas_linhas_corrigidas += bloco9_atualizado

    # ---------------- Atualizar |9999| ----------------
    
        # Atualizar |9900|9900|<total>
    for i, linha in enumerate(novas_linhas_corrigidas):
        if linha.startswith("|9900|9900|"):
            novas_linhas_corrigidas[i] = f"|9900|9900|{count_9900}|\n"
            print(f"♻️ Atualizado total de 9900 para {count_9900}")
            break

    # Atualizar |9990|<linhas do bloco 9>
    for i, linha in enumerate(novas_linhas_corrigidas):
        if linha.startswith("|9990|"):
            novas_linhas_corrigidas[i] = f"|9990|{count_lin9}|\n"
            print(f"♻️ Atualizado total de linhas do Bloco 9 para {count_lin9}")
            break
    
    # Atualizar |9999| para total de linhas do SPED
    total_linhas = len(novas_linhas_corrigidas)
    for i, linha in enumerate(novas_linhas_corrigidas):
        if linha.startswith("|9999|"):
            novas_linhas_corrigidas[i] = f"|9999|{total_linhas}|\n"
            print(f"♻️ Atualizado |9999| para {total_linhas}")
            break

    # ---------------- Salvar arquivo ----------------
    with open(saida_sped, "w", encoding="latin1") as f:
        f.writelines(novas_linhas_corrigidas)

    print("\n✅ SPED corrigido gerado em:", saida_sped)





def atualizar_bloco9(bloco9, count_c140, count_c141):
    novas_linhas = []
    atualizado_c140 = False
    atualizado_c141 = False

    for linha in bloco9:
        if linha.startswith("|9900|C140|"):
            qtd = int(linha.split("|")[3])
            linha = f"|9900|C140|{qtd + count_c140}|\n"
            atualizado_c140 = True
        elif linha.startswith("|9900|C141|"):
            qtd = int(linha.split("|")[3])
            linha = f"|9900|C141|{qtd + count_c141}|\n"
            atualizado_c141 = True

        # Adiciona todas as 9900 exceto |9999|
        if not linha.startswith("|9999|"):
            novas_linhas.append(linha)

    # Inserir C140/C141 se não existirem, após C990
    idx_c990 = next((i for i, l in enumerate(novas_linhas) if l.startswith("|9900|C990|")), None)
    if idx_c990 is not None:
        if not atualizado_c140 and count_c140 > 0:
            novas_linhas.insert(idx_c990 + 1, f"|9900|C140|{count_c140}|\n")
            idx_c990 += 1
        if not atualizado_c141 and count_c141 > 0:
            novas_linhas.insert(idx_c990 + 1, f"|9900|C141|{count_c141}|\n")

    # Adiciona |9999| temporário
    novas_linhas.append("|9999|0|\n")

    # Contagem real do Bloco 9
    count_9900 = sum(1 for l in novas_linhas if l.startswith("|9900|") and not l.startswith("|9999|"))
    count_lin9 = len(novas_linhas)  # inclui todas as linhas do bloco 9

    return novas_linhas, count_9900, count_lin9


# ---------------- INTERFACE GRÁFICA ----------------

def escolher_pasta_xml():
    pasta = filedialog.askdirectory()
    if pasta:
        entry_xml.delete(0, tk.END)
        entry_xml.insert(0, pasta)

def escolher_sped():
    arquivo = filedialog.askopenfilename(filetypes=[("SPED TXT", "*.txt")])
    if arquivo:
        entry_sped.delete(0, tk.END)
        entry_sped.insert(0, arquivo)

def executar():
    pasta_xml = entry_xml.get()
    arquivo_sped = entry_sped.get()

    if not pasta_xml or not arquivo_sped:
        messagebox.showerror("Erro", "Selecione a pasta de XML e o arquivo SPED!")
        return

    saida_sped = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("SPED TXT", "*.txt")],
        title="Salvar SPED Corrigido",
        initialfile="sped_corrigido.txt",
        initialdir=os.path.dirname(arquivo_sped)
    )

    if not saida_sped:
        return

    try:
        notas = ler_xml_notas(pasta_xml)
        processar_sped(arquivo_sped, notas, saida_sped)
        messagebox.showinfo("Sucesso", f"Novo SPED gerado: {saida_sped}")
    except Exception as e:
        messagebox.showerror("Erro", str(e))


# ---------------- JANELA ----------------

root = tk.Tk()
root.title("Gerar SPED com Duplicatas")

frame1 = tk.Frame(root)
frame1.pack(padx=10, pady=5)
tk.Label(frame1, text="Pasta de XML:").pack(side=tk.LEFT)
entry_xml = tk.Entry(frame1, width=50)
entry_xml.pack(side=tk.LEFT, padx=5)
tk.Button(frame1, text="Procurar", command=escolher_pasta_xml).pack(side=tk.LEFT)

frame2 = tk.Frame(root)
frame2.pack(padx=10, pady=5)
tk.Label(frame2, text="Arquivo SPED:").pack(side=tk.LEFT)
entry_sped = tk.Entry(frame2, width=50)
entry_sped.pack(side=tk.LEFT, padx=5)
tk.Button(frame2, text="Procurar", command=escolher_sped).pack(side=tk.LEFT)

frame3 = tk.Frame(root)
frame3.pack(pady=15)
tk.Button(frame3, text="Gerar SPED Corrigido", command=executar, bg="green", fg="white").pack()

root.mainloop()
