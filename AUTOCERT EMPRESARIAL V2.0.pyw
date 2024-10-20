import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
import os
import zipfile
import win32com.client as win32
import time

# Função para carregar o arquivo .xlsx
def carregar_arquivo():
    global df
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if caminho_arquivo:
        try:
            # Carregar a planilha usando os índices das colunas (D = índice 3, E = índice 4) e ajustando para o cabeçalho na linha 2
            df = pd.read_excel(caminho_arquivo, sheet_name='Sheet1', usecols=[3, 4], skiprows=2)
            df.columns = ["COD. PRODUTO", "QUANTIDADE"]  # Renomeia as colunas para garantir consistência

            lista_produtos.delete(0, tk.END)
            for index, row in df.iterrows():
                lista_produtos.insert(tk.END, f"Produto: {row['COD. PRODUTO']}, Quantidade: {row['QUANTIDADE']}")
            messagebox.showinfo("Sucesso", "DOC. SAÍDA IMPORTADO COM SUCESSO!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar arquivo: {e}")

# Função para preencher o template com as informações e salvar o novo arquivo
def salvar_relatorio():
    data_faturamento = entry_data_faturamento.get()
    numero_nota = entry_numero_nota.get()

    if not data_faturamento or not numero_nota:
        messagebox.showwarning("Atenção", "Por favor, preencha todos os campos.")
        return

    # Definir o caminho do template e da pasta de certificados
    script_directory = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_directory, "TEMPLATE_CERTIFICADO_MOBENSANI.xltm")  # Usar o template .xltm
    output_directory = os.path.join(script_directory, "Certificados_Gerados")

    # Verificar se o template existe
    if not os.path.exists(template_path):
        messagebox.showerror("Erro", f"Template não encontrado no diretório: {template_path}")
        return

    # Criar a pasta de certificados, se não existir
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Configurar o Excel para executar o VBA
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False  # Desativar alertas para evitar prompt de salvamento

    # Lista para armazenar os caminhos dos arquivos Excel gerados
    arquivos_excel = []
    arquivos_pdf = []  # Lista para armazenar os caminhos dos PDFs gerados

    # Processar cada linha da base de dados
    for index, row in df.iterrows():
        try:
            # Carregar o template
            workbook = excel.Workbooks.Open(template_path)
            worksheet = workbook.Worksheets("MOBENSANI")

            # Inserir a data de faturamento e número da nota nas células especificadas
            worksheet.Range("G3").Value = numero_nota
            worksheet.Range("G5").Value = data_faturamento

            # Preencher as informações da linha atual
            codigo_produto = row["COD. PRODUTO"]
            quantidade = row["QUANTIDADE"]
            worksheet.Range("C9").Value = codigo_produto
            worksheet.Range("D11").Value = quantidade
            worksheet.Range("D11:E11").Merge()

            # Extrair o número do código do produto, removendo "NX " do início
            codigo_produto_numero = codigo_produto.replace("NX ", "")

            # Definir o caminho de salvamento do novo arquivo
            output_file_name = f"{numero_nota}{codigo_produto_numero}.xlsm"  # Salvar como .xlsm para suportar macros
            output_file_path = os.path.join(output_directory, output_file_name)

            # Salvar a cópia como .xlsm
            workbook.SaveAs(output_file_path, FileFormat=52)  # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
            
            # Armazenar o caminho do arquivo Excel gerado para excluir depois
            arquivos_excel.append(output_file_path)

            # Executar a macro no novo arquivo salvo
            excel.Application.Run(f"{os.path.basename(output_file_path)}!InserirImagemComBaseNoNumero")

            # Salvar o arquivo em PDF
            pdf_file_name = f"{numero_nota}{codigo_produto_numero}.pdf"
            pdf_file_path = os.path.join(output_directory, pdf_file_name)
            workbook.ExportAsFixedFormat(0, pdf_file_path)  # 0 = xlTypePDF

            # Armazenar o caminho do arquivo PDF gerado para inclusão no .zip
            arquivos_pdf.append(pdf_file_path)

            workbook.Close(SaveChanges=False)  # Fechar sem salvar alterações
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao preencher o template para o produto {codigo_produto}: {e}")
            workbook.Close(SaveChanges=False)

    # Fechar o Excel
    excel.Quit()

    # Deletar todos os arquivos Excel (.xlsm) gerados
    for file_path in arquivos_excel:
        if os.path.exists(file_path):
            os.remove(file_path)

    # Criar um arquivo .zip com os PDFs
    zip_file_name = f"CERTIFICADO MOBENSANI NF {numero_nota}.zip"
    zip_file_path = os.path.join(output_directory, zip_file_name)
    with zipfile.ZipFile(zip_file_path, 'w') as zipf:
        for pdf_file in arquivos_pdf:
            zipf.write(pdf_file, os.path.basename(pdf_file))

    messagebox.showinfo("Sucesso", f"CERTIFICADOS DIGITAIS MOBENSANI GERADOS COM SUCESSO!!!")

# Configurações da Interface Gráfica
root = tk.Tk()
root.title("AutoCert Mobensani V2.0")

# Botão para carregar o arquivo .xlsx
btn_carregar = tk.Button(root, text="IMPORTE O DOC. SAÍDA", command=carregar_arquivo)
btn_carregar.pack(pady=10)

# Listbox para mostrar os produtos e quantidades
lista_produtos = tk.Listbox(root, width=50, height=10)
lista_produtos.pack(pady=10)

# Entradas para data de faturamento e número da nota fiscal
frame_dados = tk.Frame(root)
frame_dados.pack(pady=10)

tk.Label(frame_dados, text="DATA DO FATURAMENTO:").grid(row=0, column=0, padx=5, pady=5)
entry_data_faturamento = tk.Entry(frame_dados)
entry_data_faturamento.grid(row=0, column=1, padx=5, pady=5)

tk.Label(frame_dados, text="NÚMERO DA NOTA FISCAL:").grid(row=1, column=0, padx=5, pady=5)
entry_numero_nota = tk.Entry(frame_dados)
entry_numero_nota.grid(row=1, column=1, padx=5, pady=5)

# Botão para salvar o relatório
btn_salvar = tk.Button(root, text="GERAR CERTIFICADOS DIGITAIS", command=salvar_relatorio)
btn_salvar.pack(pady=10)

# Label de rodapé com mensagem de aviso em vermelho
footer_label = tk.Label(root, text="Para o devido funcionamento, FECHE O EXCEL!", fg="red")
footer_label.pack(pady=10)

# Executa a Interface Gráfica
root.mainloop()
