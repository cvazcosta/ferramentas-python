import zipfile
import os
import shutil
from tkinter import Tk, filedialog
import stat

def handle_remove_readonly(func, path, exc_info):
    os.chmod(path, stat.S_IWRITE)
    func(path)

def remover_protecao_excel(caminho_arquivo):
    if not (caminho_arquivo.endswith('.xlsx') or caminho_arquivo.endswith('.xlsm')):
        raise ValueError("O arquivo deve ser .xlsx ou .xlsm")

    nome_base = os.path.splitext(os.path.basename(caminho_arquivo))[0]
    extensao = os.path.splitext(caminho_arquivo)[1]
    arquivo_sem_protecao = f"{nome_base}_DESPROTEGIDO{extensao}"

    pasta_temp = "temp_excel"

    with zipfile.ZipFile(caminho_arquivo, 'r') as zip_ref:
        zip_ref.extractall(pasta_temp)

    pasta_planilhas = os.path.join(pasta_temp, "xl", "worksheets")
    for nome_arquivo in os.listdir(pasta_planilhas):
        caminho_xml = os.path.join(pasta_planilhas, nome_arquivo)

        # Ignora se não for um arquivo ou se não for XML
        if not os.path.isfile(caminho_xml) or not nome_arquivo.endswith('.xml'):
            continue

        with open(caminho_xml, 'r', encoding='utf-8') as f:
            conteudo = f.read()

        # Remove a tag de proteção
        conteudo = conteudo.replace('<sheetProtection', '<!-- <sheetProtection')
        conteudo = conteudo.replace('/>', '/> -->')

        with open(caminho_xml, 'w', encoding='utf-8') as f:
            f.write(conteudo)

    shutil.make_archive("arquivo_temp", 'zip', pasta_temp)
    os.rename("arquivo_temp.zip", arquivo_sem_protecao)

    shutil.rmtree(pasta_temp, onerror=handle_remove_readonly)

    print(f"Planilha desprotegida salva como: {arquivo_sem_protecao}")

if __name__ == "__main__":
    Tk().withdraw()
    caminho = filedialog.askopenfilename(title="Selecione o arquivo Excel protegido", filetypes=[("Planilhas Excel", "*.xlsm *.xlsx")])
    if caminho:
        remover_protecao_excel(caminho)
    else:
        print("Nenhum arquivo selecionado.")
