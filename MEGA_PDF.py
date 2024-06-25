import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess
import platform

# Função de mesclagem de PDFs e seleção de arquivos
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
                messagebox.showinfo("Success", f"PDF files merged successfully into: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error merging PDF files: {str(e)}")

# Função para converter arquivos Word para PDF
def word_to_pdf(files):
    if platform.system() == "Windows":
        import win32com.client # Apenas para Windows
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            for file in files:
                if not os.path.isfile(file):
                    messagebox.showwarning("Warning", f"File not found: {file}")
                    continue
                abs_path = os.path.abspath(file)
                try:
                    doc = word.Documents.Open(abs_path)
                except Exception as open_err:
                    messagebox.showerror("Error", f"Error opening file '{file}': {str(open_err)}")
                    continue
                pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
                try: 
                    doc.SaveAs(pdf_path, FileFormat=17) # 17 corresponde ao formato PDF
                    doc.Close()
                except Exception as save_err:
                    messagebox.showerror("Error", f"Error saving file '{file}' as PDF: {str(save_err)}")
                    doc.Close()
                    continue
            word.Quit()
            messagebox.showinfo("Success", "Word files converted to PDF successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Error converting Word files to PDF: {str(e)}")
            word.Quit()
    else: # Para Linux e outros sistemas operacionais
        try:
            for file in files:
                if not os.path.isfile(file):
                    messagebox.showwarning("warnig", f"File not foud: {file}")
                    continue
                abs_path = os.path.abspath(file)
                pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
                try:
                    subprocess.run(['libreoffice', '--headless', '--convert-to','pdf', abs_path], check=True)
                    messagebox.showinfo("Sucess", f"Word file '{file}' converted to PDF successfully.")
                except subprocess.CalledProcessError as err:
                    messagebox.showerror("Error", f"Error converting file '{file}' to PDF: {str(err)}")
        except Exception as e:
            messagebox.showerror("Error", f"Error converting Word files to PDF: {str(e)}")


# Função para selecionar arquivos Word e converter para PDF
def select_word_files_and_convert():
    files = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx;*.doc")])
    if files:
        try: 
            word_to_pdf(files)
        except Exception as e:
            messagebox.showerror("Error", f"Error converting Word files to PDF: {str(e)}")

# Função para converter arquivos PDF para Word            
def pdf_to_word(files):
    import pdf2docx
    for file in files:
        pdf_path = file
        pdf_name = os.path.basename(pdf_path)
        docx_path = os.path.splitext(pdf_path)[0] + ".docx"
        try:
            # Use pdf2docx to convert PDF to DOCX
            pdf2docx.parse(pdf_path, docx_path)
            messagebox.showinfo("Success", f"PDF file '{pdf_name}' converted to Word successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Error converting PDF file '{pdf_name}' to Word: {str(e)}")

# Função para selecionar arquivos PDF e converter para Word
def select_pdf_files_and_convert():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
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

    select_button = tk.Button(frame, text="Juntar PDF", command=select_pdf_files_and_merge, bg="#7D3C98", fg='white', highlightbackground='black', highlightcolor='black', activebackground='#A569BD', activeforeground='white')
    select_button.pack(pady=10)

    select_word_button = tk.Button(frame, text="Converter Word para PDF", command=select_word_files_and_convert, bg='#7D3C98', fg='white', highlightbackground='black', highlightcolor='black', activebackground='#A569BD', activeforeground='white')
    select_word_button.pack(pady=10)

    select_pdf_button = tk.Button(frame, text="Converter PDF para Word", command=select_pdf_files_and_convert, bg='#7D3C98', fg='white', highlightbackground='black', activeforeground='white')
    select_pdf_button.pack(pady=10)

    root.mainloop()

if __name__=="__main__":
    creat_gui()
