import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess
import platform
import zipfile

def merge_pdfs(files, output_path):
    import PyPDF2
    merger = PyPDF2.PdfMerger()
    for pdf in files:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()

def select_pdf_files_and_merge():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if files:
        try:
            output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if output_file:
                merge_pdfs(files, output_file)
                messagebox.showinfo("Sucesso", f"Arquivos PDF mesclados com sucesso em: {output_file}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao mesclar arquivos PDF: {str(e)}")

def word_to_pdf(files):
    if platform.system() == "Windows":
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            for file in files:
                if not os.path.isfile(file):
                    messagebox.showwarning("Aviso", f"Arquivo não encontrado: {file}")
                    continue
                abs_path = os.path.abspath(file)
                try:
                    doc = word.Documents.Open(abs_path)
                except Exception as open_err:
                    messagebox.showerror("Erro", f"Erro ao abrir arquivo '{file}': {str(open_err)}")
                    continue
                pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
                try: 
                    doc.SaveAs(pdf_path, FileFormat=17) # 17 corresponde ao formato PDF
                    doc.Close()
                except Exception as save_err:
                    messagebox.showerror("Erro", f"Erro ao salvar arquivo '{file}' como pdf: {str(save_err)}")
                    doc.Close()
                    continue
            word.Quit()
            messagebox.showinfo("Sucesso", "Arquivos Word convertidos para PDF com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos Word para PDF: {str(e)}")
            word.Quit()
    else: # Para Linux e outros sistemas operacionais
        try:
            for file in files:
                if not os.path.isfile(file):
                    messagebox.showwarning("Aviso", f"Arquivo não encontrado: {file}")
                    continue
                abs_path = os.path.abspath(file)
                pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
                try:
                    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', abs_path], check=True)
                    messagebox.showinfo("Sucesso", f"Arquivo Word '{file}' convertido para PDF com sucesso.")
                except subprocess.CalledProcessError as err:
                    messagebox.showerror("Erro", f"Erro ao converter arquivo '{file}' para PDF: {str(err)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos Word em PDF: {str(e)}")

def select_word_files_and_convert():
    files = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx;*.doc")])
    if files:
        try: 
            word_to_pdf(files)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos Word em PDF: {str(e)}")

def pdf_to_word(files):
    import pdf2docx
    for file in files:
        pdf_path = file
        pdf_name = os.path.basename(pdf_path)
        docx_path = os.path.splitext(pdf_path)[0] + ".docx"
        try:
            # Use pdf2docx para converter PDF para Docx
            pdf2docx.parse(pdf_path, docx_path)
            messagebox.showinfo("Sucesso", f"Arquivo PDF '{pdf_name}' convertido para Word com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivo PDF '{pdf_name}' para Word: {str(e)}")

def select_pdf_files_and_convert():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if files:
        try:
            pdf_to_word(files)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos PDF para Word: {str(e)}")

def pdf_to_excel(files):
    import pandas as pd
    from tabula import read_pdf
    import os

    try:
        for file in files:
            if not os.path.isfile(file):
                messagebox.showwarning("Atenção", f"Arquivo não encontrado: {file}")
                continue
            try:
                # Lê tabelas do PDF
                dfs = read_pdf(file, pages="all", multiple_tables=True)
                excel_path = os.path.splitext(file)[0] + ".xlsx"
                
                # Filtra DataFrames vazios
                dfs = [df for df in dfs if not df.empty]

                if not dfs:
                    messagebox.showinfo("Info", f"Nenhuma tabela encontrada no arquivo PDF '{file}' para converter para Excel.")
                    continue

                # Escreve todas as tabelas em um arquivo Excel
                with pd.ExcelWriter(excel_path) as writer:
                    for idx, df in enumerate(dfs):
                        df.to_excel(writer, sheet_name=f'Table {idx + 1}', index=False)
                messagebox.showinfo("Sucesso", f"Arquivo PDF '{file}' convertido para Excel com sucesso.")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao converter arquivo PDF '{file}' para Excel: {str(e)}")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao processar os arquivos: {str(e)}")

def excel_to_pdf(files):
    if platform.system() == "Windows":
        import win32com.client
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


def select_pdf_files_and_convert_to_excel():
    files = filedialog.askopenfilenames(filetype=[("PDF files", "*.pdf")])
    if files:
        try:
            pdf_to_excel(files)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos PDF para Excel: {str(e)}")

def select_excel_files_and_convert_to_pdf():
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if files:
        try:
            excel_to_pdf(files)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos Excel para PDF: {str(e)}")

def print_pdfs(files):
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

def select_pdf_files_and_print():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if files:
        try:
            print_pdfs(files)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao imprimir arquivos PDF: {str(e)}")

def compress_files(files):
    output_file = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP files", "*.zip")])
    if output_file:
        try:
            with zipfile.ZipFile(output_file, 'w') as zipf:
                for file in files:
                    zipf.write(file, os.path.basename(file))
            messagebox.showinfo("Sucesso", f"Arquivos compactados com sucesso em: {output_file}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao compactar arquivos: {str(e)}")

def select_files_and_compress():
    files = filedialog.askopenfilenames()
    if files:
        try:
            compress_files(files)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao compactar arquivos: {str(e)}")

def creat_gui():
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

    select_word_button = tk.Button(frame, text="Converter Word para PDF", command=select_word_files_and_convert, bg='#7D3C98', fg='white', highlightbackground='black', highlightcolor='black', activebackground='#A569BD', activeforeground='white')
    select_word_button.pack(pady=10)

    select_pdf_button = tk.Button(frame, text="Converter PDF para Word", command=select_pdf_files_and_convert, bg='#7D3C98', fg='white', highlightbackground='black', activeforeground='white')
    select_pdf_button.pack(pady=10)

    select_pdf_to_excel_button = tk.Button(frame, text="Converter PDF para Excel", command=select_pdf_files_and_convert_to_excel, bg='#7D3C98', fg='white', highlightbackground='black')
    select_pdf_to_excel_button.pack(pady=10)

    select_excel_to_pdf_button = tk.Button(frame, text="Converter Excel para PDF", command=select_excel_files_and_convert_to_pdf, bg='#7D3C98', fg='white', highlightbackground='black')
    select_excel_to_pdf_button.pack(pady=10)

    compress_files_button = tk.Button(frame, text="Compactar Arquivos", command=select_files_and_compress, bg='#7D3C98', fg='white', highlightbackground='black')
    compress_files_button.pack(pady=10)

    select_button = tk.Button(frame, text="Juntar PDF", command=select_pdf_files_and_merge, bg="#7D3C98", fg='white', highlightbackground='black', highlightcolor='black', activebackground='#A569BD', activeforeground='white')
    select_button.pack(pady=10)

    print_pdf_button = tk.Button(frame, text="Imprimir PDF", command=select_pdf_files_and_print, bg='#7D3C98', fg='white', highlightbackground='black', activeforeground='white')
    print_pdf_button.pack(pady=10)



    root.mainloop()

if __name__=="__main__":
    creat_gui()