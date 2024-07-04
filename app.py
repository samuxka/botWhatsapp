import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import tkinter.font as tkFont
from datetime import datetime
from tkinter import filedialog
import os
from openpyxl import load_workbook, Workbook

                
class WhatsAppBotApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatize - WhatsApp Bot")
        self.root.iconphoto(False, tk.PhotoImage(file="./icon.png"))
        self.data_handler = DataHandler('clientes.xlsx')
        self.df = self.data_handler.load_data()

        self.tree = ttk.Treeview(root, columns=self.df.columns, show='headings')
        self.tree.grid(row=1, column=0, columnspan=3, sticky='nsew')
        self.update_treeview_columns()
        self.load_treeview()
        self.tree.bind("<ButtonRelease-1>", self.on_tree_select)
        self.tree.bind("<Double-1>", self.on_double_click)

        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.grid(row=1, column=3, sticky='ns')

        self.scrollbar_x = ttk.Scrollbar(root, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=self.scrollbar_x.set)
        self.scrollbar_x.grid(row=2, column=0, columnspan=3, sticky='ew')

        self.top_frame = tk.Frame(root)
        self.top_frame.grid(row=0, column=0, columnspan=4, sticky='ew')

        self.add_column_button = tk.Button(self.top_frame, text="Adicionar Coluna", background='white', font='Montserrat', foreground='black', command=self.show_add_column_dialog)
        self.add_column_button.pack(side=tk.LEFT)

        self.add_row_button = tk.Button(self.top_frame, text="Adicionar Linha", background='white', font='Montserrat', foreground='black', command=self.add_row)
        self.add_row_button.pack(side=tk.LEFT)

        self.remove_column_button = tk.Button(self.top_frame, text="Remover Coluna", background='white', font='Montserrat', foreground='black', command=self.remove_column)
        self.remove_column_button.pack(side=tk.LEFT)

        self.remove_row_button = tk.Button(self.top_frame, text="Remover Linha", background='#CCCCCC', font='Montserrat', foreground='black', command=self.remove_row, state=tk.DISABLED)
        self.remove_row_button.pack(side=tk.LEFT)

        self.bottom_frame = tk.Frame(root)
        self.bottom_frame.grid(row=3, column=0, columnspan=4, sticky='ew')

        self.message_entry = tk.Entry(self.bottom_frame, width=80)
        self.message_entry.pack(side=tk.LEFT, padx=5, pady=5)

        self.send_button = tk.Button(self.bottom_frame, text="Enviar Mensagens", font='Montserrat', background='#33CC33', command=self.send_messages)
        self.send_button.pack(side=tk.RIGHT, padx=5, pady=5)

        self.clear_sheet_button = tk.Button(self.top_frame, text="Limpar Planilha", background='red', font='Montserrat', foreground='white', command=self.clear_sheet)
        self.clear_sheet_button.pack(side=tk.RIGHT)

        root.grid_columnconfigure(0, weight=1)
        root.grid_rowconfigure(1, weight=1)

    def clear_sheet(self):
        confirm = messagebox.askyesno("Confirmar", "Tem certeza que deseja limpar toda a planilha? Isso excluirá todas as colunas e linhas.")
        if confirm:
            self.df = pd.DataFrame()  # Cria um DataFrame vazio
            self.update_treeview_columns()
            self.load_treeview()
            self.data_handler.save_data(self.df)
            self.remove_column_button.config(state=tk.DISABLED)
            self.remove_row_button.config(state=tk.DISABLED)

    def load_treeview(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for index, row in self.df.iterrows():
            row_values = []
            for col in self.df.columns:
                value = row[col]
                if pd.notna(value) and isinstance(value, pd.Timestamp):
                    value = value.strftime('%d/%m/%Y')
                row_values.append(value)
            self.tree.insert("", "end", values=row_values)

    def update_treeview_columns(self):
        self.tree["columns"] = list(self.df.columns)
        for col in self.df.columns:
            try:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=tkFont.Font().measure(col))
            except Exception as e:
                print(f"Erro ao atualizar a coluna {col}: {e}")
        if not self.df.columns.any():
            self.tree["columns"] = ()

    def show_add_column_dialog(self):
        def on_confirm():
            column_name = entry_name.get()
            column_type = column_type_var.get()
            if column_name and column_type:
                self.add_column(column_name, column_type)
                dialog.destroy()
            else:
                messagebox.showwarning("Atenção", "Por favor, preencha todos os campos.")
        
        def on_cancel():
            dialog.destroy()

        dialog = tk.Toplevel(self.root)
        dialog.title("Adicionar Coluna")
        
        tk.Label(dialog, text="Nome da Coluna").pack(pady=(10, 0))
        entry_name = tk.Entry(dialog, width=50)
        entry_name.pack(pady=(0, 10))
        
        column_type_var = tk.StringVar(value="texto")
        
        types_frame = tk.Frame(dialog)
        types_frame.pack()
        
        tk.Radiobutton(types_frame, text="Link", variable=column_type_var, value="link").grid(row=0, column=0, padx=5, pady=5)
        tk.Radiobutton(types_frame, text="Valor (R$)", variable=column_type_var, value="valor").grid(row=0, column=1, padx=5, pady=5)
        tk.Radiobutton(types_frame, text="Texto", variable=column_type_var, value="texto").grid(row=0, column=2, padx=5, pady=5)
        tk.Radiobutton(types_frame, text="Data", variable=column_type_var, value="data").grid(row=1, column=0, padx=5, pady=5)
        tk.Radiobutton(types_frame, text="Número", variable=column_type_var, value="numero").grid(row=1, column=1, padx=5, pady=5)
        
        buttons_frame = tk.Frame(dialog)
        buttons_frame.pack(pady=10)
        
        tk.Button(buttons_frame, text="Confirmar", command=on_confirm).grid(row=0, column=0, padx=5)
        tk.Button(buttons_frame, text="Cancelar", command=on_cancel).grid(row=0, column=1, padx=5)

    def add_column(self, col_name, col_type):
        if col_name not in self.df.columns:
            if col_type == "link" or col_type == "texto":
                self.df[col_name] = ""
            elif col_type == "valor":
                self.df[col_name] = "R$0,00"
            elif col_type == "numero":
                self.df[col_name] = 0
            elif col_type == "data":
                self.df[col_name] = pd.NaT
            self.update_treeview_columns()
            self.load_treeview()
            self.data_handler.save_data(self.df)
        else:
            messagebox.showwarning("Atenção", "A coluna já existe.")

    def on_tree_select(self, event):
        selected_items = self.tree.selection()
        if selected_items:
            self.remove_row_button.config(state=tk.NORMAL)
        else:
            self.remove_row_button.config(state=tk.DISABLED)

    def remove_column(self):
        col_name = simpledialog.askstring("Remover Coluna", "Nome da Coluna para Remover:")
        if col_name and col_name in self.df.columns:
            self.df.drop(columns=[col_name], inplace=True)
            self.update_treeview_columns()
            self.load_treeview()
            self.data_handler.save_data(self.df)
        else:
            messagebox.showwarning("Atenção", "Coluna não encontrada.")

    def remove_row(self):
        selected_item = self.tree.selection()
        if selected_item:
            tree_index = self.tree.item(selected_item[0])['values'][0]
            df_index = self.df.index[self.df['telefone'] == tree_index].tolist()
            if df_index:
                self.df.drop(df_index[0], inplace=True)
                self.df.reset_index(drop=True, inplace=True)
                self.update_treeview_columns()
                self.load_treeview()
                self.data_handler.save_data(self.df)
                self.remove_row_button.config(state=tk.DISABLED)
            else:
                messagebox.showwarning("Atenção", "Linha não encontrada no DataFrame.")
        else:
            messagebox.showwarning("Atenção", "Linha não encontrada.")

    def on_double_click(self, event):
        if not self.tree.selection():
            return
        item = self.tree.selection()[0]
        column = self.tree.identify_column(event.x)[1:]
        column_index = int(column) - 1
        column_name = self.df.columns[column_index]
        old_value = self.tree.item(item, "values")[column_index]

        new_value = simpledialog.askstring("Editar Valor", f"Editar valor de {column_name}:", initialvalue=old_value)
        
        if new_value is not None:
            if self.df[column_name].dtype == 'datetime64[ns]':
                try:
                    new_value = datetime.strptime(new_value, '%d/%m/%Y')
                    self.df.at[int(self.tree.index(item)), column_name] = pd.Timestamp(new_value)
                except ValueError:
                    messagebox.showwarning("Atenção", "Data no formato inválido.")
                    return
            else:
                if column_name == "valor":
                    try:
                        new_value = float(new_value.replace("R$", "").replace(",", ".").strip())
                        new_value = f"R${new_value:.2f}".replace(".", ",")
                    except ValueError:
                        messagebox.showwarning("Atenção", "Valor no formato inválido.")
                        return
                self.df.at[int(self.tree.index(item)), column_name] = new_value
            self.load_treeview()
            self.data_handler.save_data(self.df)

    def add_row(self):
        new_row_data = {}
        for col in self.df.columns:
            if self.df[col].dtype == 'datetime64[ns]':
                new_value = simpledialog.askstring("Adicionar Linha", f"Valor para {col} (dd/mm/yyyy):")
                if new_value:
                    try:
                        new_value = datetime.strptime(new_value, '%d/%m/%Y')
                        new_value = pd.Timestamp(new_value)
                    except ValueError:
                        messagebox.showwarning("Atenção", "Data no formato inválido.")
                        return
                else:
                    new_value = pd.NaT
            elif col == "valor":
                new_value = simpledialog.askstring("Adicionar Linha", f"Valor para {col} (R$):")
                if new_value:
                    try:
                        new_value = float(new_value.replace("R$", "").replace(",", ".").strip())
                        new_value = f"R${new_value:.2f}".replace(".", ",")
                    except ValueError:
                        messagebox.showwarning("Atenção", "Valor no formato inválido.")
                        return
            else:
                new_value = simpledialog.askstring("Adicionar Linha", f"Valor para {col}:")

            new_row_data[col] = new_value if new_value is not None else ""

        self.df = self.df._append(new_row_data, ignore_index=True)
        self.load_treeview()
        self.data_handler.save_data(self.df)

    def send_messages(self):
        mensagem_template = self.message_entry.get().strip()
        if not mensagem_template:
            messagebox.showwarning("Atenção", "A mensagem não pode estar vazia.")
            return

        for index, row in self.df.iterrows():
            mensagem = mensagem_template
            for col in self.df.columns:
                value = row[col]
                if pd.notna(value) and isinstance(value, pd.Timestamp):
                    value = value.strftime('%d/%m/%Y')
                mensagem = mensagem.replace(f"${{{col}}}", str(value))
            telefone = row['telefone']
            link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
            try:
                webbrowser.open(link_mensagem_whatsapp)
                sleep(10)
                pyautogui.press('enter')
                sleep(2)
                pyautogui.hotkey('ctrl', 'w')
            except Exception as e:
                print(f'Não foi possível enviar mensagem para {row["nome"]}, {telefone}: {e}')
                with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                    arquivo.write(f'{row["nome"]}, {telefone}\n')

class DataHandler:
    def __init__(self, filepath):
        self.filepath = filepath

    def load_data(self):
        try:
            df = pd.read_excel(self.filepath)
        except FileNotFoundError:
            df = pd.DataFrame()
        df.columns = [col.strip() for col in df.columns]
        for col in df.columns:
            df[col] = df[col].apply(lambda x: str(x).strip() if isinstance(x, str) else x)
        return df

    def save_data(self, df):
        with pd.ExcelWriter(self.filepath, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)


if __name__ == "__main__":
    root = tk.Tk()
    messagebox.showwarning("Atenção", "Primeiro abra o seu whatsapp web para poder executar o app")
    webbrowser.open("https://web.whatsapp.com/")
    sleep(60)
    app = WhatsAppBotApp(root)
    root.mainloop()
