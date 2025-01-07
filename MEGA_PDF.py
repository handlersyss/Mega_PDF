import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess
import platform
import zipfile

# Função auxiliar para verificar e obter caminhos absolutos dos arquivos
def verificar_e_obter_caminhos(files):
    ''''
    caminhos = []
    for file in files:
        if not os.path.isfile(file):
            messagebox.showwarning("Aviso", f"Arquivo não encontrado: {file}")
            continue
        caminhos.append(os.path.abspath(file))
    return caminhos
    '''
    caminhos_validos = []
    arquivos_invalidos = []

    for file in files:
        if os.path.isfile(file):
            caminhos_validos.append(os.path.abspath(file))
        else:
            arquivos_invalidos.append(file)

    if arquivos_invalidos:
        mensagem = "\n".join(arquivos_invalidos)
        messagebox.showwarning("Aviso", f"Arquivos não encontrados: \n{mensagem}")

    return caminhos_validos

def mesclar_pdf(files, output_path):
    try:
        import PyPDF2
    except ImportError as e:
        messagebox.showerror("Erro", f"Erro ao importar PypPDF2:", {str(e)})
        return

    merger = PyPDF2.PdfMerger()
    for pdf in files:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()

def selecionar_arquivos_pdf_e_mesclar():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if files:
        try:
            output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if output_file:
                mesclar_pdf(files, output_file)
                messagebox.showinfo("Sucesso", f"Arquivos PDF mesclados com sucesso em: {output_file}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao mesclar arquivos PDF: {str(e)}")

def word_para_pdf(files):
    if platform.system() == "Windows":
        try:
            import win32com.client
        except ImportError as e:
            messagebox.showerror("Erro", f"Dependência faltando: {str(e)}")
            return

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            arquivos = verificar_e_obter_caminhos(files)
            if not arquivos:
                messagebox.showwarning("Aviso", "Nenhum arquivo foi selecionado.")
                return
            
            for abs_path in arquivos:
                try:
                    doc = word.Documents.Open(abs_path)
                    pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
                    doc.SaveAs(pdf_path, FileFormat=17)
                    doc.Close()
                except Exception as err:
                    messagebox.showerror("Erro", f"Erro ao processar '{abs_path}': {str(err)}")
                    if doc:
                        doc.close()
            word.Quit()
            messagebox.showinfo("Sucesso", "Arquivos Word convertidos para PDF com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos Word para PDF: {str(e)}")
            word.Quit()
    else:
        try:
            arquivos = verificar_e_obter_caminhos(files)
            for abs_path in arquivos:
                try:
                    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', abs_path], check=True)           
                    messagebox.showinfo("Sucesso", f"Arquivo Word '{abs_path}' convertido para PDF com sucesso.")
                except subprocess.CalledProcessError as err:
                    messagebox.showerror("Erro", f"Erro ao converter '{abs_path}' para PDF: {str(err)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos Word em PDF: {str(e)}")

def selecionar_arquivos_de_palavras_e_converter():
    files = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx;*.doc;*.odt")])
    if files:
        try: 
            word_para_pdf(files)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos Word em PDF: {str(e)}")

def pdf_para_word(files):
    try:
        import pdf2docx
    except ImportError as e: 
        messagebox.showerror("Erro", f"Dependência faltando: {str(e)}")
        return

    for file in files:
        pdf_path = file
        pdf_name = os.path.basename(pdf_path)
        docx_path = os.path.splitext(pdf_path)[0] + ".docx"
        try:
            #Use pdf2docx para converter PDF para Docx
            pdf2docx.parse(pdf_path, docx_path)
            messagebox.showinfo("Sucesso", f"Arquivo PDF '{pdf_name}' convertido para Word com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivo PDF '{pdf_name}' para Docx: {str(e)}")
            continue

        if platform.system() != "Windows":
            od_path = os.path.splitext(pdf_path)[0] + ".odt"
            try:
               # Use Libreoffice para converter Docx para ODT
                result = subprocess.run(['libreoffice', '--headless', '--convert-to', 'odt', docx_path], check=True, capture_output=True, text=True)
                if result.returncode == 0:
                   messagebox.showinfo("Sucesso", f"Arquivo Docx '{pdf_name}' convertido para ODT com sucesso.")
                else:
                   messagebox.showerror("Erro", f"Erro ao converter '{pdf_name}' para ODT: {result.stderr}")
            except subprocess.CalledProcessError as err:
                messagebox.showerror("Erro", f"Erro ao converter arquivo Docx '{pdf_name}' para ODT: {str(err)}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao converter arquivo Docx '{pdf_name}' para ODT: {str(e)}")

def selecionar_arquivos_pdf_e_converter():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if files:
        try:
            pdf_para_word(files)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos PDF para Word: {str(e)}")

def pdf_para_excel(files):
    try:
        import pandas as pd
        import pdfplumber
        import os
    except ImportError as e:
        messagebox.showerror("Erro", f"Dependência faltando: {str(e)}")
        return

    try:
        arquivos = verificar_e_obter_caminhos(files)
        for abs_path in arquivos:
            try:    
                print(f"Processando arquivo: {abs_path}")
                with pdfplumber.open(abs_path) as pdf:
                    all_tables = []
                    for page in pdf.pages:
                        tables = page.extract_tables()
                        for table in tables:
                            df = pd.DataFrame(table[1:], columns=table[0])
                            all_tables.append(df)

                excel_path = os.path.splitext(abs_path)[0] + ".xlsx"
                with pd.ExcelWriter(excel_path) as writer:
                    if all_tables:
                        for idx, df in enumerate(all_tables):
                            df.to_excel(writer, sheet_name=f'Tabla {idx + 1}', index=False)
                    else:
                        #Cria um DataFrame vazio para evitar erro ao salvar o arquivo Excel
                        df = pd.DataFrame(["Nenhuma tabela encontrada no arquivo PDF."])
                        df.to_excel(writer, sheet_name='Sem Tables', index=False)
                        
                    messagebox.showinfo("Sucesso", f"Arquivo PDF '{abs_path}' convertido para Excel com sucesso.")
            except UnicodeDecodeError as ude:
                messagebox.showerror("Erro", f"Erro ao processar arquivo PDF '{abs_path}': {str(ude)}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao converter arquivo PDF '{abs_path}' para Excel: [str(e)]")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar arquivos: {str(e)}")

def excel_para_pdf(files):
    if platform.system() == "Windows":
        try:
            import win32com.client
        except ImportError as e:
            messagebox.showerror("Erro", f"Dependência faltando: {str(e)}")
            return
            
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        try:
            for file in files:
                if not os.path.isfile(file):
                    messagebox.showwarning("Atenção", f"Arquivo não encontrado: {file}")
                    continue
                abs_path = os.path.abspath(file)
                pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
                try:
                    workbook = excel.Workbooks.Open(abs_path)
                    workbook.ExportAsFixedFormat(0, pdf_path)  # 0 corresponde ao formato PDF
                    workbook.Close(False)
                    messagebox.showinfo("Sucesso", f"Arquivo Excel '{file}' convertido para PDF com sucesso.")
                except Exception as convert_err:
                    messagebox.showerror("Erro", f"Erro ao converter arquivo '{file}' para PDF: {str(convert_err)}")
                    if workbook:
                        workbook.Close(False)
            excel.Quit()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos Excel para PDF: {str(e)}")
            excel.Quit()
    else:
        messagebox.showerror("Erro", "A conversão de Excel para PDF só é compatível com Windows.")


def selecionar_arquivos_pdf_e_converter_para_excel():
    files = filedialog.askopenfilenames(filetype=[("PDF files", "*.pdf")])
    if files:
        try:
            pdf_para_excel(files)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos PDF para Excel: {str(e)}")

def selecionar_arquivos_excel_e_converter_para_pdf():
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if files:
        try:
            excel_para_pdf(files)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos Excel para PDF: {str(e)}")

def imprimir_pdfs(files):
    for file in files:
        if platform.system() == "Windows":
            try:
                os.startfile(file, "print")
            except Exception as e: 
                messagebox.showerror("Erro", f"Erro ao imprimir arquivo '{file}': {str(e)}")
        else:
            try:
                subprocess.run(['lp', file], check=True)
                messagebox.showeinfo("Sucesso", f"Arquivo PDF '{file}' enviado para a impressora com sucesso.")
            except subprocess.CalledProcessError as arr:
                messagebox.showerror("Erro", f"Erro ao imprimir arquivo '{file}': {str(e)}")

def selecionar_arquivos_pdf_e_imprimir():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if files:
        try:
            imprimir_pdfs(files)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao imprimir arquivos PDF: {str(e)}")

def comprimir_arquivos(files):
    output_file = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP files", "*.zip")])
    if output_file:
        try:
            with zipfile.ZipFile(output_file, 'w') as zipf:
                for file in files:
                    zipf.write(file, os.path.basename(file))
            messagebox.showinfo("Sucesso", f"Arquivos compactados com sucesso em: {output_file}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao compactar arquivos: {str(e)}")

def selecionar_arquivos_e_compactar():
    files = filedialog.askopenfilenames()
    if files:
        try:
            comprimir_arquivos(files)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao compactar arquivos: {str(e)}")

def criar_gui():
    root = tk.Tk()
    root.title("MEGA PDF")
    root.configure(bg='black')

    ascii_art_lines = [
        ("███╗   ███╗███████╗ ██████╗  █████╗     ██████╗ ██████╗ ███████╗", "#703C98"),
        ("████╗ ████║██╔════╝██╔════╝ ██╔══██╗    ██╔══██╗██╔══██╗██╔════╝", "#A569BD"),
        ("██╔████╔██║█████╗  ██║  ███╗███████║    ██████╔╝██║  ██║█████╗  ", "#7D3C98"),
        ("██║╚██╔╝██║██╔══╝  ██║   ██║██╔══██║    ██╔═══╝ ██║  ██║██╔══╝  ", "#A569BD"),
        ("██║ ╚═╝ ██║███████╗╚██████╔╝██║  ██║    ██║     ██████╔╝██║     ", "#7D3C98"),    
        ("╚═╝     ╚═╝╚══════╝ ╚═════╝ ╚═╝  ╚═╝    ╚═╝     ╚═════╝ ╚═╝     ", "#A569BD"),
#        ("Criador: Edson França Neto", "purple"),
#        ("Contato: (33)998341977", "purple"),
#        ("E-mail: edsontaylor@outlook.com.br", "purple")
    ]
   
    for line, color in ascii_art_lines:
        label = tk.Label(root, text=line, font=("Courier", 10), fg=color, bg='black')
        label.pack()

    frame = tk.Frame(root, padx=20, pady=20, bg='black')
    frame.pack(padx=20, pady=20)

    select_word_button = tk.Button(frame, text="Converter Word para PDF", command=selecionar_arquivos_de_palavras_e_converter, bg='#7D3C98', fg='white', highlightbackground='black', highlightcolor='black', activebackground='#A569BD', activeforeground='white')
    select_word_button.pack(pady=10)

    select_pdf_button = tk.Button(frame, text="Converter PDF para Word", command=selecionar_arquivos_pdf_e_converter, bg='#7D3C98', fg='white', highlightbackground='black', activeforeground='white')
    select_pdf_button.pack(pady=10)

    select_pdf_para_excel_button = tk.Button(frame, text="Converter PDF para Excel", command=selecionar_arquivos_pdf_e_converter_para_excel, bg='#7D3C98', fg='white', highlightbackground='black')
    select_pdf_para_excel_button.pack(pady=10)

    select_excel_para_pdf_button = tk.Button(frame, text="Converter Excel para PDF", command=selecionar_arquivos_excel_e_converter_para_pdf, bg='#7D3C98', fg='white', highlightbackground='black')
    select_excel_para_pdf_button.pack(pady=10)

    comprimir_arquivos_button = tk.Button(frame, text="Compactar Arquivos", command=selecionar_arquivos_e_compactar, bg='#7D3C98', fg='white', highlightbackground='black')
    comprimir_arquivos_button.pack(pady=10)

    select_button = tk.Button(frame, text="Juntar PDF", command=selecionar_arquivos_pdf_e_mesclar, bg="#7D3C98", fg='white', highlightbackground='black', highlightcolor='black', activebackground='#A569BD', activeforeground='white')
    select_button.pack(pady=10)

    print_pdf_button = tk.Button(frame, text="Imprimir PDF", command=selecionar_arquivos_pdf_e_imprimir, bg='#7D3C98', fg='white', highlightbackground='black', activeforeground='white')
    print_pdf_button.pack(pady=10)



    root.mainloop()

if __name__=="__main__":
    criar_gui()