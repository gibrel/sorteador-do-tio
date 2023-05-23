import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as messagebox
import win32clipboard
from random import sample


class ResultWindow(tk.Toplevel):
    def __init__(self, names):
        super().__init__()
        self.title("Resultados")
        self.geometry("800x600")

        self.result_text = tk.Text(self, height=10, width=30)
        self.result_text.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.result_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.configure(yscrollcommand=scrollbar.set)

        self.display_results(names)

    def display_results(self, names):
        result = "Nomes Sorteados:\n\n"
        for i, name in enumerate(names, start=1):
            result += f"{i}. {name}\n"
        self.result_text.insert(tk.END, result)
        self.result_text.config(state=tk.DISABLED)


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Aplicativo de Planilha de Nomes")
        self.geometry("800x600")

        self.toolbar = ttk.Frame(self)
        self.toolbar.pack(side=tk.TOP, fill=tk.X)

        self.tree_frame = ttk.Frame(self)
        self.tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(self.tree_frame)
        self.tree["columns"] = ("ID", "Name")
        self.tree.column("#0", width=0, stretch=tk.NO)
        self.tree.column("ID", anchor=tk.CENTER, width=50)
        self.tree.column("Name", anchor=tk.W, width=200)
        self.tree.heading("ID", text="ID")
        self.tree.heading("Name", text="Nome")

        self.tree_scrollbar = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.summary_label = ttk.Label(self, text="Resumo das Importações:")
        self.summary_label.pack(pady=5)

        self.summary_text = tk.Text(self, height=3)
        self.summary_text.pack(pady=5)

        self.quantity_label = ttk.Label(self, text="Quantidade:")
        self.quantity_label.pack(pady=5)

        self.quantity_entry = ttk.Entry(self)
        self.quantity_entry.pack(pady=5)

        get_names_button = ttk.Button(self.toolbar, text="Obter Nomes (Ctrl+V)", command=self.get_clipboard_text)
        get_names_button.pack(side=tk.LEFT, padx=5)

        clear_button = ttk.Button(self.toolbar, text="Limpar Planilha", command=self.clear_sheet)
        clear_button.pack(side=tk.LEFT, padx=5)

        select_button = ttk.Button(self.toolbar, text="Sortear", command=self.select_names)
        select_button.pack(side=tk.LEFT, padx=5)

        self.bind("<Control-v>", self.on_paste)

        self.current_id = 1
        self.names = []
        self.new_entries = 0
        self.duplicate_entries = 0

    def get_clipboard_text(self, event=None):
        try:
            win32clipboard.OpenClipboard()
            clipboard_text = win32clipboard.GetClipboardData()
            win32clipboard.CloseClipboard()
            name_list = clipboard_text.split("\n")
            if name_list:
                self.insert_names(name_list)
            else:
                messagebox.showinfo("Erro", "Nenhum nome encontrado na área de transferência.")
        except Exception as e:
            messagebox.showinfo("Erro", f"Erro ao obter os nomes da área de transferência: {str(e)}")

    def insert_names(self, name_list):
        self.clear_sheet()
        self.new_entries = 0
        self.duplicate_entries = 0

        for name in name_list:
            name = name.strip()
            if name:
                if name not in self.names:
                    self.names.append(name)
                    self.tree.insert("", tk.END, values=(self.current_id, name))
                    self.current_id += 1
                    self.new_entries += 1
                else:
                    self.duplicate_entries += 1

        self.update_summary()

    def clear_sheet(self):
        self.tree.delete(*self.tree.get_children())
        self.current_id = 1
        self.names.clear()
        self.clear_summary()

    def clear_summary(self):
        self.summary_text.config(state=tk.NORMAL)
        self.summary_text.delete("1.0", tk.END)
        self.summary_text.config(state=tk.DISABLED)

    def update_summary(self):
        total_entries = len(self.names)
        self.summary_text.config(state=tk.NORMAL)
        self.summary_text.delete("1.0", tk.END)
        self.summary_text.insert(tk.END, f"Total de Nomes: {total_entries}\n")
        self.summary_text.insert(tk.END, f"Novas Entradas: {self.new_entries}\n")
        self.summary_text.insert(tk.END, f"Entradas Duplicadas: {self.duplicate_entries}\n")
        self.summary_text.config(state=tk.DISABLED)

    def on_paste(self, event):
        self.get_clipboard_text()

    def select_names(self):
        quantity = self.quantity_entry.get()
        if quantity.isdigit():
            quantity = int(quantity)
            if quantity <= len(self.names):
                selected_names = sample(self.names, quantity)
                ResultWindow(selected_names)
            else:
                messagebox.showinfo("Erro", "A quantidade inserida é maior que o número de nomes na planilha.")
        else:
            messagebox.showinfo("Erro", "Digite um número válido.")


if __name__ == "__main__":
    app = Application()
    app.mainloop()

