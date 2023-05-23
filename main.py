# import tkinter
# import tkinter.messagebox
# import customtkinter
# import tktabl
#
# customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
# customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"
#
# headers = ["Id", "Nome", "Local", ]
#
#
# def change_appearance_mode_event(new_appearance_mode: str):
#     customtkinter.set_appearance_mode(new_appearance_mode)
#
#
# def open_input_dialog_event():
#     dialog = customtkinter.CTkInputDialog(text="Type in a number:", title="CTkInputDialog")
#     print("CTkInputDialog:", dialog.get_input())
#
#
# def change_scaling_event(new_scaling: str):
#     new_scaling_float = int(new_scaling.replace("%", "")) / 100
#     customtkinter.set_widget_scaling(new_scaling_float)
#
#
# def sidebar_button_event():
#     print("sidebar_button click")
#
#
# class App(customtkinter.CTk):
#     def __init__(self):
#         super().__init__()
#
#         # configure window
#         self.title("CustomTkinter complex_example.py")
#         self.geometry(f"{1100}x{580}")
#
#         # configure grid layout (4x4)
#         # self.grid_columnconfigure(1, weight=1)
#         # self.grid_columnconfigure((2, 3), weight=0)
#         # self.grid_rowconfigure((0, 1, 2), weight=1)
#         #
#         # # create sidebar frame with widgets
#         # self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
#         # self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
#         # self.sidebar_frame.grid_rowconfigure(4, weight=1)
#         # self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="CustomTkinter",
#         #                                          font=customtkinter.CTkFont(size=20, weight="bold"))
#         # self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
#         # self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, command=sidebar_button_event)
#         # self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
#         # self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, command=sidebar_button_event)
#         # self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)
#         # self.sidebar_button_3 = customtkinter.CTkButton(self.sidebar_frame, command=sidebar_button_event)
#         # self.sidebar_button_3.grid(row=3, column=0, padx=20, pady=10)
#         # self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
#         # self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
#         # self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame,
#         #                                                                values=["Light", "Dark", "System"],
#         #                                                                command=change_appearance_mode_event)
#         # self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
#         # self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
#         # self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
#         # self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame,
#         #                                                        values=["80%", "90%", "100%", "110%", "120%"],
#         #                                                        command=change_scaling_event)
#         # self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))
#         #
#         # # create main entry and button
#         # self.entry = customtkinter.CTkEntry(self, placeholder_text="CTkEntry")
#         # self.entry.grid(row=3, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")
#         #
#         # self.main_button_1 = customtkinter.CTkButton(master=self, fg_color="transparent",
#         #                                              border_width=2, text_color=("gray10", "#DCE4EE"))
#         # self.main_button_1.grid(row=3, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")
#         #
#         # # create textbox
#         # self.textbox = customtkinter.CTkTextbox(self, width=250)
#         # self.textbox.grid(row=0, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")
#         #
#         # # create tabview
#         # self.tabview = customtkinter.CTkTabview(self, width=250)
#         # self.tabview.grid(row=0, column=2, padx=(20, 0), pady=(20, 0), sticky="nsew")
#         # self.tabview.add("CTkTabview")
#         # self.tabview.add("Tab 2")
#         # self.tabview.add("Tab 3")
#         # self.tabview.tab("CTkTabview").grid_columnconfigure(0, weight=1)  # configure grid of individual tabs
#         # self.tabview.tab("Tab 2").grid_columnconfigure(0, weight=1)
#         #
#         # self.optionmenu_1 = customtkinter.CTkOptionMenu(self.tabview.tab("CTkTabview"), dynamic_resizing=False,
#         #                                                 values=["Value 1", "Value 2", "Value Long Long Long"])
#         # self.optionmenu_1.grid(row=0, column=0, padx=20, pady=(20, 10))
#         # self.combobox_1 = customtkinter.CTkComboBox(self.tabview.tab("CTkTabview"),
#         #                                             values=["Value 1", "Value 2", "Value Long....."])
#         # self.combobox_1.grid(row=1, column=0, padx=20, pady=(10, 10))
#         # self.string_input_button = customtkinter.CTkButton(self.tabview.tab("CTkTabview"), text="Open CTkInputDialog",
#         #                                                    command=open_input_dialog_event)
#         # self.string_input_button.grid(row=2, column=0, padx=20, pady=(10, 10))
#         # self.label_tab_2 = customtkinter.CTkLabel(self.tabview.tab("Tab 2"), text="CTkLabel on Tab 2")
#         # self.label_tab_2.grid(row=0, column=0, padx=20, pady=20)
#         #
#         # # create radiobutton frame
#         # self.radiobutton_frame = customtkinter.CTkFrame(self)
#         # self.radiobutton_frame.grid(row=0, column=3, padx=(20, 20), pady=(20, 0), sticky="nsew")
#         # self.radio_var = tkinter.IntVar(value=0)
#         # self.label_radio_group = customtkinter.CTkLabel(master=self.radiobutton_frame, text="CTkRadioButton Group:")
#         # self.label_radio_group.grid(row=0, column=2, columnspan=1, padx=10, pady=10, sticky="")
#         # self.radio_button_1 = customtkinter.CTkRadioButton(master=self.radiobutton_frame,
#         #                                                    variable=self.radio_var, value=0)
#         # self.radio_button_1.grid(row=1, column=2, pady=10, padx=20, sticky="n")
#         # self.radio_button_2 = customtkinter.CTkRadioButton(master=self.radiobutton_frame,
#         #                                                    variable=self.radio_var, value=1)
#         # self.radio_button_2.grid(row=2, column=2, pady=10, padx=20, sticky="n")
#         # self.radio_button_3 = customtkinter.CTkRadioButton(master=self.radiobutton_frame,
#         #                                                    variable=self.radio_var, value=2)
#         # self.radio_button_3.grid(row=3, column=2, pady=10, padx=20, sticky="n")
#         #
#         # # create slider and progressbar frame
#         # self.slider_progressbar_frame = customtkinter.CTkFrame(self, fg_color="transparent")
#         # self.slider_progressbar_frame.grid(row=1, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")
#         # self.slider_progressbar_frame.grid_columnconfigure(0, weight=1)
#         # self.slider_progressbar_frame.grid_rowconfigure(4, weight=1)
#         # self.seg_button_1 = customtkinter.CTkSegmentedButton(self.slider_progressbar_frame)
#         # self.seg_button_1.grid(row=0, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
#         # self.progressbar_1 = customtkinter.CTkProgressBar(self.slider_progressbar_frame)
#         # self.progressbar_1.grid(row=1, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
#         # self.progressbar_2 = customtkinter.CTkProgressBar(self.slider_progressbar_frame)
#         # self.progressbar_2.grid(row=2, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
#         # self.slider_1 = customtkinter.CTkSlider(self.slider_progressbar_frame, from_=0, to=1, number_of_steps=4)
#         # self.slider_1.grid(row=3, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
#         # self.slider_2 = customtkinter.CTkSlider(self.slider_progressbar_frame, orientation="vertical")
#         # self.slider_2.grid(row=0, column=1, rowspan=5, padx=(10, 10), pady=(10, 10), sticky="ns")
#         # self.progressbar_3 = customtkinter.CTkProgressBar(self.slider_progressbar_frame, orientation="vertical")
#         # self.progressbar_3.grid(row=0, column=2, rowspan=5, padx=(10, 20), pady=(10, 10), sticky="ns")
#         #
#         # # create scrollable frame
#         # self.scrollable_frame = customtkinter.CTkScrollableFrame(self, label_text="CTkScrollableFrame")
#         # self.scrollable_frame.grid(row=1, column=2, padx=(20, 0), pady=(20, 0), sticky="nsew")
#         # self.scrollable_frame.grid_columnconfigure(0, weight=1)
#         # self.scrollable_frame_switches = []
#         # for i in range(100):
#         #     switch = customtkinter.CTkSwitch(master=self.scrollable_frame, text=f"CTkSwitch {i}")
#         #     switch.grid(row=i, column=0, padx=10, pady=(0, 20))
#         #     self.scrollable_frame_switches.append(switch)
#         #
#         # # create checkbox and switch frame
#         # self.checkbox_slider_frame = customtkinter.CTkFrame(self)
#         # self.checkbox_slider_frame.grid(row=1, column=3, padx=(20, 20), pady=(20, 0), sticky="nsew")
#         # self.checkbox_1 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
#         # self.checkbox_1.grid(row=1, column=0, pady=(20, 0), padx=20, sticky="n")
#         # self.checkbox_2 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
#         # self.checkbox_2.grid(row=2, column=0, pady=(20, 0), padx=20, sticky="n")
#         # self.checkbox_3 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
#         # self.checkbox_3.grid(row=3, column=0, pady=20, padx=20, sticky="n")
#         #
#         # # set default values
#         # self.sidebar_button_3.configure(state="disabled", text="Disabled CTkButton")
#         # self.checkbox_3.configure(state="disabled")
#         # self.checkbox_1.select()
#         # self.scrollable_frame_switches[0].select()
#         # self.scrollable_frame_switches[4].select()
#         # self.radio_button_3.configure(state="disabled")
#         # self.appearance_mode_optionemenu.set("Dark")
#         # self.scaling_optionemenu.set("100%")
#         # self.optionmenu_1.set("CTkOptionmenu")
#         # self.combobox_1.set("CTkComboBox")
#         # self.slider_1.configure(command=self.progressbar_2.set)
#         # self.slider_2.configure(command=self.progressbar_3.set)
#         # self.progressbar_1.configure(mode="indeterminnate")
#         # self.progressbar_1.start()
#         # self.textbox.insert("0.0", "CTkTextbox\n\n" + 'Lorem ipsum dolor sit amet, consetetur sadipscing elitr, '
#         #                                               'sed diam nonumy eirmod tempor invidunt ut labore et dolore '
#         #                                               'magna aliquyam erat, sed diam voluptua.\n\n' * 20)
#         # self.seg_button_1.configure(values=["CTkSegmentedButton", "Value 2", "Value 3"])
#         # self.seg_button_1.set("Value 2")
#
#
# if __name__ == "__main__":
#     app = App()
#     app.mainloop()


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

