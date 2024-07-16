import sqlite3
from docx import Document
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Conectar ao banco de dados (ou criar se não existir)
conexao = sqlite3.connect('frases.db')
cursor = conexao.cursor()

# Criar a tabela de frases (adicionar a coluna 'cluster' se não existir)
cursor.execute('''
CREATE TABLE IF NOT EXISTS frases (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cluster TEXT,
    frase TEXT NOT NULL
)
''')

# Adicionar coluna 'cluster' se não existir
try:
    cursor.execute('ALTER TABLE frases ADD COLUMN cluster TEXT')
except sqlite3.OperationalError:
    # Se a coluna já existir, ignorar o erro
    pass


# Função para inserir uma frase
def inserir_frase():
    cluster = entrada_cluster.get()
    frase = entrada_frase.get()
    if cluster.strip() == "" or frase.strip() == "":
        messagebox.showerror("Erro", "Digite um cluster e uma frase válidos!")
        return

    cursor.execute('INSERT INTO frases (cluster, frase) VALUES (?, ?)', (cluster, frase))
    conexao.commit()
    entrada_cluster.delete(0, END)
    entrada_frase.delete(0, END)  # Limpar os campos após a inserção


# Função para visualizar todas as frases
def visualizar_frases():
    cursor.execute('SELECT * FROM frases')
    frases = cursor.fetchall()
    if not frases:
        messagebox.showinfo("Aviso", "Não há frases armazenadas!")
    else:
        lista_frases.delete(0, END)
        for frase in frases:
            lista_frases.insert(END, f"{frase[0]}: {frase[1]} - {frase[2]}")


# Função para deletar uma frase pelo ID
def deletar_frase():
    id_frase = entrada_id.get()
    if id_frase.strip() == "":
        messagebox.showerror("Erro", "Digite o ID da frase a ser deletada!")
        return

    try:
        cursor.execute('DELETE FROM frases WHERE id = ?', (id_frase,))
        conexao.commit() 
        entrada_id.delete(0, END)  # Limpar o campo após a deleção
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao deletar a frase: {e}")


# Função para exportar frases para um documento do Word
def exportar_para_word():
    nome_arquivo = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Arquivo do Word", "*.docx")])
    if not nome_arquivo:
        return

    cursor.execute('SELECT * FROM frases')
    frases = cursor.fetchall()

    documento = Document()
    documento.add_heading('Segmentos de Textos', 0)

    # Criar tabela com 2 colunas
    tabela = documento.add_table(rows=1, cols=2)
    tabela.style = 'Table Grid'  # Estilo da tabela com bordas

    # Configurar largura das colunas
    largura_coluna_cluster = Inches(1.0)
    largura_coluna_frase = Inches(4.0)
    tabela.columns[0].width = largura_coluna_cluster
    tabela.columns[1].width = largura_coluna_frase

    # Adicionar cabeçalho à tabela
    hdr_cells = tabela.rows[0].cells
    hdr_cells[0].text = 'Cluster'
    hdr_cells[1].text = 'Segmento de Texto'
    hdr_cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    hdr_cells[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Adicionar frases à tabela
    for frase in frases:
        row_cells = tabela.add_row().cells
        row_cells[0].text = frase[1]  # Cluster
        row_cells[1].text = frase[2]  # Frase
        row_cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        row_cells[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    # Salvar documento
    try:
        documento.save(nome_arquivo)
        messagebox.showinfo("Sucesso", f"Tabela exportada para {nome_arquivo}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar a tabela: {e}")


# Criar a interface gráfica
root = Tk()
root.title("Gerenciador de Frases")

# Frame para inserir e visualizar frases
frame_inserir_visualizar = Frame(root)
frame_inserir_visualizar.pack(padx=10, pady=10)

Label(frame_inserir_visualizar, text="Digite o cluster:").grid(row=0, column=0, padx=5, pady=5)
entrada_cluster = Entry(frame_inserir_visualizar, width=10)
entrada_cluster.grid(row=0, column=1, padx=5, pady=5)

Label(frame_inserir_visualizar, text="Digite uma nova frase:").grid(row=1, column=0, padx=5, pady=5)
entrada_frase = Entry(frame_inserir_visualizar, width=50)
entrada_frase.grid(row=1, column=1, padx=5, pady=5)

botao_inserir = Button(frame_inserir_visualizar, text="Inserir Frase", command=inserir_frase)
botao_inserir.grid(row=1, column=2, padx=5, pady=5)

Label(frame_inserir_visualizar, text="ID da frase a ser deletada:").grid(row=2, column=0, padx=5, pady=5)
entrada_id = Entry(frame_inserir_visualizar, width=10)
entrada_id.grid(row=2, column=1, padx=5, pady=5)

botao_deletar = Button(frame_inserir_visualizar, text="Deletar Frase", command=deletar_frase)
botao_deletar.grid(row=2, column=2, padx=5, pady=5)

botao_visualizar = Button(frame_inserir_visualizar, text="Visualizar Frases", command=visualizar_frases)
botao_visualizar.grid(row=3, column=1, padx=5, pady=5)

# Botão para exportar frases para o Word
botao_exportar_word = Button(frame_inserir_visualizar, text="Exportar Frases para Word", command=exportar_para_word)
botao_exportar_word.grid(row=4, column=1, padx=5, pady=5)

# Lista de frases
frame_lista = Frame(root)
frame_lista.pack(padx=10, pady=10)

Label(frame_lista, text="Frases Armazenadas:").pack()

scrollbar = Scrollbar(frame_lista)
scrollbar.pack(side=RIGHT, fill=Y)

lista_frases = Listbox(frame_lista, width=70, yscrollcommand=scrollbar.set)
lista_frases.pack()

scrollbar.config(command=lista_frases.yview)

# Executar a interface
root.mainloop()

# Fechar conexão com o banco de dados
conexao.close()
