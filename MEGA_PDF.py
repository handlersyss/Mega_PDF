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
                messagebox.showinfo("Success", f"PDF files merged successfully into: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error merging PDF files: {str(e)}")

def word_to_pdf(files):
    if platform.system() == "Windows":
        import win32com.client
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
                    messagebox.showwarning("Warning", f"File not found: {file}")
                    continue
                abs_path = os.path.abspath(file)
                pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
                try:
                    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', abs_path], check=True)
                    messagebox.showinfo("Success", f"Word file '{file}' converted to PDF successfully.")
                except subprocess.CalledProcessError as err:
                    messagebox.showerror("Error", f"Error converting file '{file}' to PDF: {str(err)}")
        except Exception as e:
            messagebox.showerror("Error", f"Error converting Word files to PDF: {str(e)}")

def select_word_files_and_convert():
    files = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx;*.doc")])
    if files:
        try: 
            word_to_pdf(files)
        except Exception as e:
            messagebox.showerror("Error", f"Error converting Word files to PDF: {str(e)}")

def pdf_to_word(files):
    import pdf2docx
    for file in files:
        pdf_path = file
        pdf_name = os.path.basename(pdf_path)
        docx_path = os.path.splitext(pdf_path)[0] + ".docx"
        try:
            # Use pdf2docx para converter PDF para Docx
            pdf2docx.parse(pdf_path, docx_path)
            messagebox.showinfo("Success", f"PDF file '{pdf_name}' converted to Word successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Error converting PDF file '{pdf_name}' to Word: {str(e)}")

def select_pdf_files_and_convert():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if files:
        try:
            pdf_to_word(files)
        except Exception as e:
            messagebox.showerror("Error", f"Error converting PDF files to Word: {str(e)}")
#So salvando
#def pdf_to_excel(files):
    #import pandas as pd
    #from tabula import read_pdf
    #for file in files:
        #try:
            #Le tabelas do PDF
            #dfs = read_pdf(file, pages="all", multiple_tables=True)
            #excel_path = os.path.splitext(file)[0] + ".xlsx"

            #Escreve todas as tabelas em um arquivo excel
            #with pd.ExcelWriter(excel_path) as writer:
                #for idx, df in enumerate(dfs):
                    #df.to_excel(writer, sheet_name=f'Table {idx + 1}', index=False)
            #messagebox.showinfo("Sucess", f"PDF file '{file}' converted to Excel successfully.")
        #except Exception as e:
            #messagebox.showerror("Error", f"Error converting PDF file '{file}' to Excel: {str(e)}")
            
def pdf_to_excel(files):
    import pandas as pd
    from tabula import read_pdf
    import os

    try:
        for file in files:
            if not os.path.isfile(file):
                messagebox.showwarning("Warning", f"File not found: {file}")
                continue
            try:
                # Lê tabelas do PDF
                dfs = read_pdf(file, pages="all", multiple_tables=True)
                excel_path = os.path.splitext(file)[0] + ".xlsx"
                
                # Filtra DataFrames vazios
                dfs = [df for df in dfs if not df.empty]

                if not dfs:
                    messagebox.showinfo("Info", f"No tables found in PDF file '{file}' to convert to Excel.")
                    continue

                # Escreve todas as tabelas em um arquivo Excel
                with pd.ExcelWriter(excel_path) as writer:
                    for idx, df in enumerate(dfs):
                        df.to_excel(writer, sheet_name=f'Table {idx + 1}', index=False)
                messagebox.showinfo("Success", f"PDF file '{file}' converted to Excel successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Error converting PDF file '{file}' to Excel: {str(e)}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while processing the files: {str(e)}")

#So salvando
#def excel_to_pdf(files):
#    import win32com.client
#    if platform.system() == "Windows":
#        #import win32com.client
#        excel = win32com.client.Dispatch("Excel.Application")
#        excel.Visible = False
#        try:
#            for file in files:
#                if not os.path.isfile(file):
#                    messagebox.showwarning("Warning", f"File not found: {file}")
#                    continue
#                abs_path = os.path.abspath(file)
#                pdf_path = os.path.splitext(abs_path) [0] + ".pdf"
#                workbook = excel.Workbook.Open(abs_path)
#                workbook.ExportAsFixedFormat(0, pdf_path) # 0 corresponde ao formato PDF
#                workbook.Close()
#            excel.Quit()
#            messagebox.showinfo("Success", "Excel files converted to PDF successfully.")
#        except Exception as e:
#            messagebox.showerror("Error", f"Error converting Excel files to PDF: {str(e)}")
#            excel.Quit()               

def excel_to_pdf(files):
    if platform.system() == "Windows":
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        try:
            for file in files:
                if not os.path.isfile(file):
                    messagebox.showwarning("Warning", f"File not found: {file}")
                    continue
                abs_path = os.path.abspath(file)
                pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
                try:
                    workbook = excel.Workbooks.Open(abs_path)
                    workbook.ExportAsFixedFormat(0, pdf_path)  # 0 corresponde ao formato PDF
                    workbook.Close(False)
                    messagebox.showinfo("Success", f"Excel file '{file}' converted to PDF successfully.")
                except Exception as convert_err:
                    messagebox.showerror("Error", f"Error converting file '{file}' to PDF: {str(convert_err)}")
                    if workbook:
                        workbook.Close(False)
            excel.Quit()
        except Exception as e:
            messagebox.showerror("Error", f"Error converting Excel files to PDF: {str(e)}")
            excel.Quit()
    else:
        messagebox.showerror("Error", "Excel to PDF conversion is only supported on Windows.")


def select_pdf_files_and_convert_to_excel():
    files = filedialog.askopenfilenames(filetype=[("PDF files", "*.pdf")])
    if files:
        try:
            pdf_to_excel(files)
        except Exception as e:
            messagebox.showerror("Error", f"Error converting PDF files to Excel: {str(e)}")

def select_excel_files_and_convert_to_pdf():
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if files:
        try:
            excel_to_pdf(files)
        except Exception as e:
            messagebox.showerror("Error", f"Error converting Excel files to PDF: {str(e)}")

def print_pdfs(files):
    for file in files:
        if platform.system() == "Windows":
            try:
                os.startfile(file, "print")
            except Exception as e: 
                messagebox.showerror("Error", f"Error printing file '{file}': {str(e)}")
        else:
            try:
                subprocess.run(['lp', file], check=True)
                messagebox.showeinfo("Sucess", f"PDF file '{file}' sent to printer sucessfully.")
            except subprocess.CalledProcessError as arr:
                messagebox.showerror("Error", f"Error printing file '{file}': {str(e)}")

def select_pdf_files_and_print():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if files:
        try:
            print_pdfs(files)
        except Exception as e:
            messagebox.showerror("Error", f"Error printing PDF files: {str(e)}")

def compress_files(files):
    output_file = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP files", "*.zip")])
    if output_file:
        try:
            with zipfile.ZipFile(output_file, 'w') as zipf:
                for file in files:
                    zipf.write(file, os.path.basename(file))
            messagebox.showinfo("Success", f"Files compressed successfully into: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error compressing files: {str(e)}")

def select_files_and_compress():
    files = filedialog.askopenfilenames()
    if files:
        try:
            compress_files(files)
        except Exception as e:
            messagebox.showerror("Error", f"Error compressing files: {str(e)}")

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
   
    for line, color in ascii_art_lines:
        label = tk.Label(root, text=line, font=("Courier", 10), fg=color, bg='black')
        label.pack()

    frame = tk.Frame(root, padx=40, pady=40, bg='black')
    frame.pack(padx=40, pady=40)

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