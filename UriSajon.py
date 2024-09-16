import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
import xlwings as xw
from pandastable import Table

# Definindo o layout principal
class DictionaryApp(tk.Tk):
    def __init__(self):
        super().__init__()

        # Definindo o título da janela
        self.title("우리 사전")
        self.geometry("900x900")
        self.configure(bg="#d8c0d9")
        self.toggle_button = None
        # Inicializando a base de dados
        wb = xw.Book('translations_db.xlsx')  # Nome do arquivo Excel
        sheet = wb.sheets['Sheet1']  # Acessa a planilha correta, altere conforme necessário
        # Carregar os dados da planilha para um DataFrame
        data = sheet.range('A1').expand().value  # Expande os dados até o último preenchido
        self.data = pd.DataFrame(data[1:], columns=data[0])
        self.search_language = tk.StringVar(value="Português") 


        # Variáveis para os campos de entrada
        self.vars = {}

        # Frame lateral (nav bar)
        self.sidebar = tk.Frame(self, bg="#d188b5", width=100)
        self.sidebar.pack(side="left", fill="y")

        # Botões da navbar
        self.add_nav_button("Buscar", self.search_screen)
        self.add_nav_button("Adicionar", self.add_word_screen)
        self.add_nav_button("Ver Tudo", self.view_all_words)
        self.add_nav_button("Correção", self.correction_screen)

        # Frame principal
        self.main_frame = tk.Frame(self, bg="#d8c0d9")
        self.main_frame.pack(side="right", expand=True, fill="both")

        # Carregar a tela inicial
        self.search_screen()

    def add_nav_button(self, text, command):
        button = tk.Button(self.sidebar, text=text, font=("Helvetica", 12), bg="white", relief="raised", padx=10, pady=5, command=command)
        button.pack(pady=10, padx=10, fill="x")

    def search_screen(self):
        self.clear_main_frame()

        # Título
        title = tk.Label(
                self.main_frame, 
                text="   ~     URI    ~    우리 사전    ~", 
                font=("Helvetica", 28, "bold"),  # Aumente o tamanho e adicione estilo 'italic'
                bg="#8B008B",  # Fundo roxo escuro
                fg="white",  # Cor da fonte branca
                relief="raised",  # Crie uma borda em relevo
                bd=5,  # Espessura da borda
                padx=20,  # Adiciona padding horizontal
                pady=10  # Adiciona padding vertical
                )
        title.pack(pady=30)

        # Adicionar uma sombra ao título
        title.bind("<Configure>", lambda e: title.config(
            highlightthickness=4, 
            highlightbackground="#4B0082",  # Cor da sombra
        ))

        # Área de busca
        self.search_frame = tk.Frame(self.main_frame, bg="#d8c0d9")
        self.search_frame.pack(pady=10)

        # Dropdown para selecionar o idioma de busca
        self.language_options = ["Português", "한국어"]
        self.language_dropdown = tk.OptionMenu(self.search_frame, self.search_language, *self.language_options, command=self.update_search_layout)
        self.language_dropdown.config(font=("Helvetica", 12), bg="#b664a0", width=10)
        self.language_dropdown.grid(row=0, column=3, padx=10, pady=5)

        self.search_entry = tk.Entry(self.search_frame, font=("Helvetica", 14), width=15)

        # Inicializa o layout
        self.update_search_layout()

        # Botão de busca
        search_button = tk.Button(self.main_frame, text="Buscar", font=("Helvetica", 12), bg="#d188b5", command=self.search_word)
        search_button.pack(pady=10)

        # Frame para mostrar os resultados da busca
        self.result_frame = tk.Frame(self.main_frame, bg="#d8c0d9")
        self.result_frame.pack(pady=20)

    def update_search_layout(self, *args):
        # Limpa o layout atual
        self.search_entry.grid_forget()
        self.search_entry.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

    def search_word(self):
        search_term = self.search_entry.get()
        
        # Verifica qual idioma foi selecionado para realizar a busca na coluna correta
        if self.search_language.get() == "Português":
        # Busca case-insensitive no campo "Português"
            result = self.data[self.data["Português"].str.contains(search_term, case=False, na=False)]
        else:
        # Busca normal no campo "Coreano"
            result = self.data[self.data["Coreano"] == search_term]

        # Limpar resultados anteriores
        for widget in self.result_frame.winfo_children():
            widget.destroy()

        if not result.empty:
            # Exibir resultados em Labels
            for idx, row in result.iterrows():
                tk.Label(self.result_frame, text=f"Coreano: {row['Coreano']}", font=("Helvetica", 14), bg="#d8c0d9").pack(anchor="w", pady=2)
                tk.Label(self.result_frame, text=f"Português: {row['Português']}", font=("Helvetica", 14), bg="#d8c0d9").pack(anchor="w", pady=2)
                tk.Label(self.result_frame, text=f"Inglês: {row['Inglês']}", font=("Helvetica", 14), bg="#d8c0d9").pack(anchor="w", pady=2)
                tk.Label(self.result_frame, text=f"Significado: {row['Significado']}", font=("Helvetica", 14), bg="#d8c0d9").pack(anchor="w", pady=2)
                tk.Label(self.result_frame, text=f"Variações/Sinônimos: {row['Variações/Sinônimos']}", font=("Helvetica", 14), bg="#d8c0d9").pack(anchor="w", pady=2)
        else:
            # Exibir mensagem de não encontrado
            tk.Label(self.result_frame, text="Palavra não encontrada.", font=("Helvetica", 14), bg="#d8c0d9").pack(anchor="w", pady=2)


    def add_word_screen(self):
        self.clear_main_frame()
        title = tk.Label(self.main_frame, text="Adicionar Palavra", font=("Helvetica", 20, "bold"), bg="#d8c0d9")
        title.pack(pady=20)

        form_frame = tk.Frame(self.main_frame, bg="#d8c0d9")
        form_frame.pack(pady=10)

        # Adicionando campos para inserção de dados
        self.create_input(form_frame, "Coreano:", "coreano")
        self.create_input(form_frame, "Português:", "portugues")
        self.create_input(form_frame, "Inglês:", "ingles")
        self.create_input(form_frame, "Significado:", "significado")
        self.create_input(form_frame, "Variações/Sinônimos:", "variacoes")

        add_button = tk.Button(self.main_frame, text="Adicionar", font=("Helvetica", 12), bg="#d188b5", command=self.add_word)
        add_button.pack(pady=10)

    def view_all_words(self):
        self.clear_main_frame()

    # Título
        title = tk.Label(self.main_frame, text="Ver Todas as Palavras", font=("Helvetica", 20, "bold"), bg="#d8c0d9")
        title.pack(pady=20)

    # Frame para a tabela
        table_frame = tk.Frame(self.main_frame)
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        if not self.data.empty:
        # Display da tabela usando pandastable
            pt = Table(table_frame, dataframe=self.data, showtoolbar=True, showstatusbar=True)
            pt.show()
        else:
            tk.Label(self.main_frame, text="Nenhuma palavra cadastrada.", font=("Helvetica", 14), bg="#d8c0d9").pack(pady=20)

    # Botão de download
        download_button = tk.Button(self.main_frame, text="Baixar Excel", font=("Helvetica", 12), bg="#d188b5", command=self.download_excel)
        download_button.pack(pady=10)


    def add_word(self):
    # Verificar se a palavra já existe
        existing_entry = self.data[
            (self.data["Coreano"] == self.vars["coreano"].get()) | (self.data["Português"] == self.vars["portugues"].get())
            ]

        if not existing_entry.empty:
            messagebox.showwarning("Erro", "Essa palavra já está registrada.")
        else:
        # Adiciona a palavra à base de dados
            new_data = {
            "Coreano": self.vars["coreano"].get(),
            "Português": self.vars["portugues"].get(),
            "Inglês": self.vars["ingles"].get(),
            "Significado": self.vars["significado"].get(),
            "Variações/Sinônimos": self.vars["variacoes"].get()
            }
            new_data_df = pd.DataFrame([new_data])
            self.data = pd.concat([self.data, new_data_df], ignore_index=True)

        # Abre o arquivo Excel usando xlwings
            try:
                wb = xw.Book('translations_db.xlsx')  # Abre o arquivo Excel
                sheet = wb.sheets['Sheet1']  # Acessa a planilha correta, altere se necessário

            # Encontra a primeira linha vazia na coluna 1 (A)
                last_row = len(sheet.range('A:A').value) - sheet.range('A:A').value.count(None)

            # Escreve os novos dados na próxima linha vazia
                sheet.range(f'A{last_row + 1}').value = new_data["Coreano"]
                sheet.range(f'B{last_row + 1}').value = new_data["Português"]
                sheet.range(f'C{last_row + 1}').value = new_data["Inglês"]
                sheet.range(f'D{last_row + 1}').value = new_data["Significado"]
                sheet.range(f'E{last_row + 1}').value = new_data["Variações/Sinônimos"]

            # Salva o arquivo Excel
                wb.save()
                wb.close()

                messagebox.showinfo("Sucesso", "Palavra adicionada e salva com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar no Excel: {e}")

    def download_excel(self):
        file_path = "translation_db.xlsx"
        self.data.to_excel(file_path, index=False)
        messagebox.showinfo("Download", f"O arquivo foi salvo como {file_path}.")

    def correction_screen(self):
        self.clear_main_frame()
        title = tk.Label(self.main_frame, text="Correção de Palavra", font=("Helvetica", 20, "bold"), bg="#d8c0d9")
        title.pack(pady=20)

        search_frame = tk.Frame(self.main_frame, bg="#d8c0d9")
        search_frame.pack(pady=10)

        search_label = tk.Label(search_frame, text="Buscar palavra (Coreano ou Português):", font=("Helvetica", 14), bg="#d8c0d9")
        search_label.pack(side="left", padx=5)

        self.search_entry1 = tk.Entry(search_frame, font=("Helvetica", 14), width=15)
        self.search_entry1.pack(side="left", padx=5)

        search_button = tk.Button(self.main_frame, text="Buscar", font=("Helvetica", 12), bg="#d188b5", command=self.find_for_correction)
        search_button.pack(pady=10)

    def find_for_correction(self):
        search_term = self.search_entry1.get()
        result = self.data[(self.data["Coreano"] == search_term) | (self.data["Português"] == search_term)]

        if not result.empty:
            self.display_correction_fields(result)
        else:
            messagebox.showwarning("Erro", "Palavra não encontrada.")

    def display_correction_fields(self, result):
        for idx, entry in result.iterrows():
            # Exibe campos para correção com os dados atuais
            self.clear_main_frame()
            form_frame = tk.Frame(self.main_frame, bg="#d8c0d9")
            form_frame.pack(pady=10)

            self.create_input(form_frame, "Coreano:", "coreano", entry["Coreano"])
            self.create_input(form_frame, "Português:", "portugues", entry["Português"])
            self.create_input(form_frame, "Inglês:", "ingles", entry["Inglês"])
            self.create_input(form_frame, "Significado:", "significado", entry["Significado"])
            self.create_input(form_frame, "Variações/Sinônimos:", "variacoes", entry["Variações/Sinônimos"])

            correct_button = tk.Button(self.main_frame, text="Corrigir", font=("Helvetica", 12), bg="#d188b5", command=lambda: self.update_word(idx))
            correct_button.pack(pady=10)

    def update_word(self, index):
        self.data.loc[index, "Coreano"] = self.vars["coreano"].get()
        self.data.loc[index, "Português"] = self.vars["portugues"].get()
        self.data.loc[index, "Inglês"] = self.vars["ingles"].get()
        self.data.loc[index, "Significado"] = self.vars["significado"].get()
        self.data.loc[index, "Variações/Sinônimos"] = self.vars["variacoes"].get()

        messagebox.showinfo("Sucesso", "Palavra corrigida com sucesso!")

    def create_input(self, frame, label_text, var_name, default_text=""):
        label = tk.Label(frame, text=label_text, font=("Helvetica", 14), bg="#d8c0d9")
        label.pack(pady=5)

        entry = tk.Entry(frame, font=("Helvetica", 14))
        entry.insert(0, default_text)
        entry.pack(pady=5)

        self.vars[var_name] = entry

    def clear_main_frame(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()


# Iniciando o aplicativo
if __name__ == "__main__":
    app = DictionaryApp()
    app.mainloop()
