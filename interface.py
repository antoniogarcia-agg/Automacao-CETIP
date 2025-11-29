import customtkinter as ctk
from tkinter import filedialog, messagebox


class InterfaceUI():
    def __init__(self, janela):
        # Configuração da janela principal
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("dark-blue")
        self.janela = janela
        self.janela.title("GERADOR PU")
        self.janela.geometry("500x400")
        self.entrada_planilha = None
        self.entrada_pasta = None
        self.setup_ui()

    def selecionar_planilha(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione a planilha",
            filetypes=[("Arquivos Excel", "*.xlsx;*.xls;*.xlsm")]
        )
        if arquivo:
            self.entry_planilha.delete(0, ctk.END)
            self.entry_planilha.insert(0, arquivo)
            self.entrada_planilha = arquivo

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta")
        if pasta:
            self.entry_pasta.delete(0, ctk.END)
            self.entry_pasta.insert(0, pasta)
            self.entrada_pasta = pasta

    def executar(self):
        if not self.entry_planilha.get() or not self.entry_pasta.get():
            messagebox.showwarning(
                "Atenção", "Por favor, selecione a planilha e a pasta antes de executar.")
        else:
            self.janela.destroy()

    def setup_ui(self):
        # Título
        titulo = ctk.CTkLabel(self.janela, text="GERADOR PU EVENTOS",
                              font=ctk.CTkFont(size=20, weight="bold"))
        titulo.pack(pady=20)

        # Campo para seleção de planilha
        frame_planilha = ctk.CTkFrame(self.janela)
        frame_planilha.pack(pady=10, padx=20, fill="x")

        label_planilha = ctk.CTkLabel(
            frame_planilha, text="Selecione a Planilha:", font=ctk.CTkFont(size=14))
        label_planilha.pack(side="left", padx=10, pady=10)

        self.entry_planilha = ctk.CTkEntry(
            frame_planilha, placeholder_text="Nenhuma planilha selecionada")
        self.entry_planilha.pack(side="left", fill="x", expand=True, padx=10)

        btn_planilha = ctk.CTkButton(
            frame_planilha, text="Procurar", command=self.selecionar_planilha)
        btn_planilha.pack(side="right", padx=10)

        # Campo para seleção de pasta
        frame_pasta = ctk.CTkFrame(self.janela)
        frame_pasta.pack(pady=10, padx=20, fill="x")

        label_pasta = ctk.CTkLabel(
            frame_pasta, text="Selecione a Pasta:", font=ctk.CTkFont(size=14))
        label_pasta.pack(side="left", padx=10, pady=10)

        self.entry_pasta = ctk.CTkEntry(
            frame_pasta, placeholder_text="Nenhuma pasta selecionada")
        self.entry_pasta.pack(side="left", fill="x", expand=True, padx=10)

        btn_pasta = ctk.CTkButton(frame_pasta, text="Procurar",
                                  command=self.selecionar_pasta)
        btn_pasta.pack(side="right", padx=10)

        # Botão de executar
        btn_executar = ctk.CTkButton(
            self.janela, text="Executar", command=self.executar, font=ctk.CTkFont(size=16, weight="bold"))
        btn_executar.pack(pady=40)

        # Rodar a aplicação
        self.janela.mainloop()
