import tkinter as tk
from tkinter import messagebox, filedialog
from docx import Document
import os
import win32com.client

def limpar_campos():
    entry_preco.delete(0, tk.END)
    entry_cidade.delete(0, tk.END)
    entry_entrega1.delete(0, tk.END)
    entry_entrega2.delete(0, tk.END)
    entry_vendedor.delete(0, tk.END)
    entry_nome_arquivo.delete(0, tk.END)

def salvar_dados():
    nome_arquivo = entry_nome_arquivo.get()
    if not nome_arquivo.endswith(".docx"):
        nome_arquivo += ".docx"
    
    caminho_arquivo = filedialog.asksaveasfilename(defaultextension=".docx", 
        filetypes=[("Word Files", "*.docx")],
            initialfile=nome_arquivo)
    
    if not caminho_arquivo:
        return

    try:
        documento = Document("Orçamento.docx")

        referencias = {
            "AAAA": opcao_selecionada.get(),
            "BBBB": entry_preco.get(),
            "CCCC": entry_cidade.get(),
            "DDDD": opcao_selecionada_area.get(),
            "EEEE": entry_entrega1.get(),
            "FFFF": entry_entrega2.get(),
            "GGGG": entry_vendedor.get(),
            "HHHH": entry_m2.get()
        }

        for paragrafo in documento.paragraphs:
            for codigo in referencias:
                if codigo in paragrafo.text:
                    paragrafo.text = paragrafo.text.replace(codigo, referencias[codigo])

        # Salva o documento Word
        documento.save(caminho_arquivo)
        
        # Converte o documento para PDF usando o Microsoft Word via COM
        caminho_arquivo_absoluto = os.path.abspath(caminho_arquivo)
        caminho_pdf = caminho_arquivo_absoluto.replace(".docx", ".pdf")
        try:
            word = win32com.client.Dispatch("Word.Application")
            doc = word.Documents.Open(caminho_arquivo_absoluto)
            doc.SaveAs(caminho_pdf, FileFormat=17)  # 17 é o código para PDF
            doc.Close()
            word.Quit()
            messagebox.showinfo("Sucesso", f"Dados salvos e convertidos para PDF com sucesso: '{caminho_pdf}'")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter o arquivo para PDF: {e}")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar os dados: {e}")

# Configuração da janela principal
root = tk.Tk()
root.title("Orçamento Alves Gramas")

# Criando os widgets
label_m2 = tk.Label(root, text="M²")
label_m2.grid(row=0, column=0, padx=10, pady=5)
entry_m2 = tk.Entry(root)
entry_m2.grid(row=0, column=1, padx=10, pady=5)

label_grama = tk.Label(root, text="Grama:")
label_grama.grid(row=0, column=2, padx=10, pady=5)
opcoes_grama = ["Selecione a grama", "Esmeralda", "Bermuda", "São Carlos", "Santo Agostinho", "Batatais", "Coreana"]
opcao_selecionada = tk.StringVar()
opcao_selecionada.set(opcoes_grama[0])
menu_opcoes_grama = tk.OptionMenu(root, opcao_selecionada, *opcoes_grama)
menu_opcoes_grama.grid(row=0, column=3, padx=10, pady=5)


label_preco = tk.Label(root, text="Preço R$:")
label_preco.grid(row=1, column=0, padx=10, pady=5)
entry_preco = tk.Entry(root)
entry_preco.grid(row=1, column=1, padx=10, pady=5)


label_area = tk.Label(root, text="Área:")
label_area.grid(row=1, column=2, padx=10, pady=5)
opcoes_area = ["Selecione a Área", "Urbana", "Rural"]
opcao_selecionada_area = tk.StringVar()
opcao_selecionada_area.set(opcoes_area[0])
menu_opcoes_area = tk.OptionMenu(root, opcao_selecionada_area, *opcoes_area)
menu_opcoes_area.grid(row=1, column=3, padx=10, pady=5)

label_cidade = tk.Label(root, text="Cidade-Estado:")
label_cidade.grid(row=2, column=0, padx=10, pady=5)
entry_cidade = tk.Entry(root)
entry_cidade.grid(row=2, column=1, padx=10, pady=5)

label_entrega1 = tk.Label(root, text="Dia de entrega(min):")
label_entrega1.grid(row=2, column=2, padx=10, pady=5)
entry_entrega1 = tk.Entry(root)
entry_entrega1.grid(row=2, column=3, padx=10, pady=5)

label_entrega2 = tk.Label(root, text="Dia de entrega(max):")
label_entrega2.grid(row=3, column=0, padx=10, pady=5)
entry_entrega2 = tk.Entry(root)
entry_entrega2.grid(row=3, column=1, padx=10, pady=5)

label_vendedor = tk.Label(root, text="Vendedor:")
label_vendedor.grid(row=4, column=0, padx=10, pady=5)
entry_vendedor = tk.Entry(root)
entry_vendedor.grid(row=4, column=1, padx=10, pady=5)

# Adiciona a entrada para o nome do arquivo
label_nome_arquivo = tk.Label(root, text="Nome do Arquivo:")
label_nome_arquivo.grid(row=4, column=2, padx=10, pady=5)
entry_nome_arquivo = tk.Entry(root)
entry_nome_arquivo.grid(row=4, column=3, padx=10, pady=5)

button_salvar = tk.Button(root, text="Salvar", command=salvar_dados)
button_salvar.grid(row=5, column=1, pady=10)

button_limpar = tk.Button(root, text="Limpar", command=limpar_campos)
button_limpar.grid(row=5, column=2, pady=10)

# Inicia a aplicação
root.mainloop()