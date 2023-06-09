from tkinter import scrolledtext, messagebox, filedialog
from PIL import Image
import random
import json
import customtkinter as ctk
import pandas as pd

# Constantes
IMPORT_BUTTON_TEXT = "Importar"
CLEAR_BUTTON_TEXT = "Limpar"
EXPORT_FORMAT_DEFAULT = "-- SELECIONE --"
EXPORT_FORMAT_OPTIONS = ["-- SELECIONE --", "JSON", "EXCEL"]
EXPORT_BUTTON_TEXT = "Exportar"
EXIT_BUTTON_TEXT = "Sair"
RESUME_LABEL_TEXT = "OS IMPERIALISTAS SÃO APENAS TIGRES DE PAPEL"
QUANTITY_LABEL_TEXT = "Quantidade:"
PICK_BUTTON_TEXT = "Sortear"
ERROR_TITLE = "Erro"
ERROR_INVALID_QUANTITY = "A quantidade inserida é inválida."
ERROR_INVALID_EXPORT_FORMAT = "Selecione um formato de exportação válido."
JSON_FILE_TYPE = [("Arquivo JSON", "*.json")]
EXCEL_FILE_TYPE = [("Planilha Excel", "*.xlsx")]

ctk.set_default_color_theme("dark-blue")
ctk.set_appearance_mode("system")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Aplicativo de Sorteio")
        self.geometry("800x600")

        # load and create background image
        self.bg_image = ctk.CTkImage(
            Image.open("assets/mao-zedong.jpg"),
            size=(1600, 900))

        self.background_label = ctk.CTkLabel(self, image=self.bg_image)
        self.background_label.place(x=0, y=0, relwidth=1, relheight=1)

        self.top_bar = ctk.CTkFrame(self)
        self.top_bar.pack(side=ctk.TOP, fill=ctk.X, padx=5, pady=5)

        self.import_button = ctk.CTkButton(
            self.top_bar, text=IMPORT_BUTTON_TEXT,
            command=self.import_names)
        self.import_button.pack(side=ctk.LEFT, padx=5, pady=5)

        self.clear_button = ctk.CTkButton(
            self.top_bar,
            text=CLEAR_BUTTON_TEXT,
            command=self.clear_sheet)
        self.clear_button.pack(side=ctk.LEFT, padx=5, pady=5)

        self.export_format_var = ctk.StringVar()
        self.export_format_var.set(EXPORT_FORMAT_DEFAULT)

        self.export_dropdown = ctk.CTkOptionMenu(
            self.top_bar,
            variable=self.export_format_var,
            values=EXPORT_FORMAT_OPTIONS)
        self.export_dropdown.pack(side=ctk.LEFT, padx=5, pady=5)

        self.export_button = ctk.CTkButton(
            self.top_bar,
            text=EXPORT_BUTTON_TEXT,
            command=self.export_results)
        self.export_button.pack(side=ctk.LEFT, padx=5, pady=5)

        self.exit_button = ctk.CTkButton(
            self.top_bar,
            text=EXIT_BUTTON_TEXT,
            command=self.quit)
        self.exit_button.pack(side=ctk.RIGHT, padx=5, pady=5)

        self.center_frame = ctk.CTkFrame(self)
        self.center_frame.pack(fill=ctk.BOTH, expand=True, padx=5, pady=5)

        self.sheet_textbox = scrolledtext.ScrolledText(self.center_frame, width=40, height=10)
        self.sheet_textbox.pack(side=ctk.LEFT, padx=5, pady=5, fill=ctk.BOTH, expand=True)

        self.results_textbox = scrolledtext.ScrolledText(self.center_frame, width=40, height=10)
        self.results_textbox.pack(side=ctk.LEFT, padx=5, pady=5, fill=ctk.BOTH, expand=True)

        self.bottom_bar = ctk.CTkFrame(self)
        self.bottom_bar.pack(side=ctk.BOTTOM, fill=ctk.X, padx=5, pady=5)

        self.summary_label = ctk.CTkLabel(
            self.bottom_bar,
            text=RESUME_LABEL_TEXT)
        self.summary_label.pack()

        self.bottom_elements_subframe = ctk.CTkFrame(self.bottom_bar)
        self.bottom_elements_subframe.pack()

        self.quantity_label = ctk.CTkLabel(
            self.bottom_elements_subframe,
            text=QUANTITY_LABEL_TEXT)
        self.quantity_label.pack(side=ctk.LEFT)

        self.quantity_entry = ctk.CTkEntry(self.bottom_elements_subframe)
        self.quantity_entry.pack(side=ctk.LEFT)

        self.pick_button = ctk.CTkButton(
            self.bottom_elements_subframe,
            text=PICK_BUTTON_TEXT,
            command=self.pick_names)
        self.pick_button.pack(side=ctk.LEFT)

        self.bottom_elements_subframe.pack_configure(expand=True)

    def import_names(self):
        clipboard_text = self.clipboard_get()
        names = clipboard_text.split("\n")
        new_entries = 0
        repeated_entries = 0

        for name in names:
            name = name.strip()
            if name and not any(name in entry for entry in self.sheet_textbox.get("1.0", ctk.END).split("\n")):
                new_entries += 1
                self.sheet_textbox.insert(ctk.END, f"ID: {new_entries}\tNome: {name}\n")
            elif name:
                repeated_entries += 1

        self.summary_label.configure(
            text=f"Novas entradas: {new_entries} | "
                 f"Entradas repetidas: {repeated_entries} | "
                 f"Total de nomes: {new_entries + repeated_entries}")

    def clear_sheet(self):
        self.sheet_textbox.delete('1.0', ctk.END)
        self.results_textbox.delete('1.0', ctk.END)
        self.summary_label.configure(text=RESUME_LABEL_TEXT)

    def pick_names(self):
        quantity = self.quantity_entry.get()
        names_list = self.sheet_textbox.get("1.0", ctk.END).split("\n")
        names_list = [name.split("\t")[1] for name in names_list if name]

        if quantity.isdigit() and int(quantity) <= len(names_list):
            results = random.sample(names_list, int(quantity))
            self.results_textbox.delete('1.0', ctk.END)
            self.results_textbox.insert(ctk.END, "\n".join(results))
        else:
            messagebox.showerror(ERROR_TITLE, ERROR_INVALID_QUANTITY)

    def export_results(self):
        export_format = self.export_format_var.get()

        if export_format == "JSON":
            self.export_to_json()
        elif export_format == "EXCEL":
            self.export_to_excel()
        else:
            messagebox.showerror(ERROR_TITLE, ERROR_INVALID_EXPORT_FORMAT)

    def export_to_json(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=JSON_FILE_TYPE)

        if file_path:
            results = self.results_textbox.get("1.0", ctk.END).splitlines()
            results_dict = {}

            for index, result in enumerate(results):
                results_dict[index + 1] = result

            with open(file_path, "w") as json_file:
                json.dump(results_dict, json_file, indent=4)

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=EXCEL_FILE_TYPE)

        if file_path:
            results = self.results_textbox.get("1.0", ctk.END).splitlines()
            results_list = []

            for result in results:
                results_list.append({"Resultado": result})

            df = pd.DataFrame(results_list)
            df.to_excel(file_path, index=False, engine="openpyxl")


app = App()
app.mainloop()
