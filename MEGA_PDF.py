import tkinter as tk
from tkinter import filedialog, messagebox
import PyPDF2
import os
import win32com.client
from pdf2docx import parse

# Função de mesclagem de PDFs e seleção de arquivos
def merge_pdfs(files, output_path):
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
                messagebox.showinfo("Success", f"PDF files merged successfully into: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error merging PDF files: {str(e)}")

# Função para converter arquivos Word para PDF
def word_to_pdf(files):
    word = win32com.client.Dispatch("Word.Application")
    try:
        for file in files:
            if not os.path.isfile(file):
                messagebox.showwarning("Warning", f"File not found: {file}")
                continue
            doc = word.Documents.Open(file)
            pdf_path = os.path.splitext(file)[0] + ".pdf"
            doc.SaveAs(pdf_path, FileFormat=17) # 17 corresponde ao formato PDF
            doc.close()
        word.Quit()
        messagebox.showinfo("Success", "Word files converted to PDF successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"Error converting Word files to PDF: {str(e)}")
        word.Quit()

# Função para selecionar arquivos Word e converter para PDF
def select_word_files_and_convert():
    files = filedialog.askopenfilenames(filetype=[("Word files", "*.docx;*.doc")])
    if files:
        try: 
            word_to_pdf(files)
        except Exception as e:
            messagebox.showerror("Error", f"Error converting Word files to PDF: {str(e)}")

# Função para converter arquivos PDF para Word            
def pdf_to_word(files):
    for file in files:
        pdf_path = file
        pdf_name = os.path.basename(pdf_path)
        docx_path = os.path.splitext(pdf_path)[0] + ".docx"
        try:
            # Use pdf2docx to convert PDF to DOCX
            parse(pdf_path, docx_path)
            messagebox.showinfo("Success", f"PDF file '{pdf_name}' converted to Word successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Error converting PDF file '{pdf_name}' to Word: {str(e)}")

# Função para selecionar arquivos PDF e converter para Word
def select_pdf_files_and_convert():
    files = filedialog.askopenfilenames(filetype=[("PDF files", "*.pdf")])
    if files:
        try:
            pdf_to_word(files)
        except Exception as e:
            messagebox.showerror("Error", f"Error converting PDF files to Word: {str(e)}")

# Função para criar a interface grafica
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
        ("Criador: Edson França Neto", "purple"),
        ("Contato: (33)998341977", "purple"),
        ("E-mail: edsontaylor@outlook.com.br", "purple")
    ]
   
    #Adicionando label com a arte ASCII colorida
    for line, color in ascii_art_lines:
        label = tk.Label(root, text=line, font=("Courier", 10), fg=color, bg='black')
        label.pack()

    frame = tk.Frame(root, padx=40, pady=40, bg='black')
    frame.pack(padx=20, pady=20)

    select_button = tk.Button(frame, text="Select PDF Files to Merge", command=select_pdf_files_and_merge, bg="#7D3C98", fg='white', highlightbackground='black', highlightcolor='black', activebackground='#A569BD', activeforeground='white')
    select_button.pack(pady=10)

    select_word_button = tk.Button(frame, text="Select Word Files", command=select_word_files_and_convert, bg='#7D3C98', fg='white', highlightbackground='black', highlightcolor='black', activebackground='#A569BD', activeforeground='white')
    select_word_button.pack(pady=10)

    select_pdf_button = tk.Button(frame, text="Select PDF Files", command=select_pdf_files_and_convert, bg='#7D3C98', fg='white', highlightbackground='black', activeforeground='white')
    select_pdf_button.pack(pady=10)

    root.mainloop()

if __name__=="__main__":
    creat_gui()

#desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
#arquivos_path = os.path.join(desktop_path, "arquivos")

#if not os.path.exists(arquivos_path):
#    print("O diretório 'arquivos' não foi encontrado no desktop.")
#    exit()

#merger = PyPDF2.PdfMerger()

#lista_arquivos = os.listdir(arquivos_path)
#lista_arquivos.sort()
#print(lista_arquivos)

#for arquivo in lista_arquivos:
#    if arquivo.endswith(".pdf"):
#        arquivo_path = os.path.join(arquivos_path, arquivo)
#        merger.append(arquivo_path)

# Solicitar o nome do arquivo final ao usuario
#nome_arquivo_final = input("Digite o nome do arquivo PDF final: ")

# certifique-se de adicionar a extensão .pdf ao nome do arquivo, se ainda não estiver presente
#if not nome_arquivo_final.endswith(".pdf"):
#    nome_arquivo_final += ".pdf"

#arquivo_final_path = os.path.join(desktop_path, nome_arquivo_final)


#merger.write(arquivo_final_path)
#merger.close()