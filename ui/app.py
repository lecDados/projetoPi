import customtkinter as ctk
from tkinter import filedialog
from tkcalendar import DateEntry
import shutil
import os

from services.excel_service import salvar_no_excel
from config.colunas import colunas


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Cadastro Excel")
        self.geometry("700x800")

        self.arquivo_path = None
        self.entries = {}
        self.foto_path = None

        self.criar_interface()

    def criar_interface(self):

        ctk.CTkLabel(
            self,
            text="Sistema de Cadastro Excel",
            font=("Arial", 20)
        ).pack(pady=10)

        ctk.CTkButton(
            self,
            text="Selecionar Arquivo Excel",
            command=self.selecionar_arquivo
        ).pack(pady=10)

        self.label_arquivo = ctk.CTkLabel(
            self,
            text="Nenhum arquivo selecionado"
        )
        self.label_arquivo.pack()

        frame = ctk.CTkScrollableFrame(self, width=600, height=500)
        frame.pack(pady=10)

        for col in colunas:

            ctk.CTkLabel(frame, text=col).pack(anchor="w", padx=10)

            # CAMPO DATA
            if col == "Data":

                data_frame = ctk.CTkFrame(frame)
                data_frame.pack(padx=10, pady=5)

                calendario = DateEntry(
                    data_frame,
                    width=20,
                    background='darkblue',
                    foreground='white',
                    borderwidth=2,
                    date_pattern='dd/mm/yyyy'
                )
                calendario.pack(side="left", padx=5)

                hora_entry = ctk.CTkEntry(
                    data_frame,
                    width=100,
                    placeholder_text="Hora"
                )
                hora_entry.pack(side="left", padx=5)

                self.entries["Data"] = calendario
                self.entries["Hora"] = hora_entry

            # CAMPO FOTO
            elif col == "Foto da ação":

                botao_foto = ctk.CTkButton(
                    frame,
                    text="Selecionar Foto",
                    command=self.selecionar_foto
                )
                botao_foto.pack(padx=10, pady=5)

                self.label_foto = ctk.CTkLabel(
                    frame,
                    text="Nenhuma foto selecionada"
                )
                self.label_foto.pack()

            else:
                entry = ctk.CTkEntry(frame, width=500)
                entry.pack(padx=10, pady=5)

                self.entries[col] = entry

        ctk.CTkButton(
            self,
            text="Salvar",
            command=self.salvar
        ).pack(pady=20)

    def selecionar_arquivo(self):

        self.arquivo_path = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx")]
        )

        if self.arquivo_path:
            self.label_arquivo.configure(
                text=os.path.basename(self.arquivo_path)
            )

    def selecionar_foto(self):

        self.foto_path = filedialog.askopenfilename(
            filetypes=[
                ("Imagens", "*.png *.jpg *.jpeg")
            ]
        )

        if self.foto_path:
            self.label_foto.configure(
                text=os.path.basename(self.foto_path)
            )

    def salvar(self):

        if not self.arquivo_path:
            print("Selecione um arquivo Excel")
            return

        dados = {}

        for col, entry in self.entries.items():

            if col == "Data":
                dados["Data"] = entry.get()

            elif col == "Hora":
                dados["Hora"] = entry.get()

            else:
                dados[col] = entry.get()

        # SALVAR FOTO
        if self.foto_path:

            pasta_fotos = "fotos"

            if not os.path.exists(pasta_fotos):
                os.makedirs(pasta_fotos)

            nome_foto = os.path.basename(self.foto_path)

            destino = os.path.join(
                pasta_fotos,
                nome_foto
            )

            shutil.copy(self.foto_path, destino)

            dados["Foto da ação"] = destino

        salvar_no_excel(
            self.arquivo_path,
            dados,
            colunas
        )

        self.mostrar_sucesso()

    def mostrar_sucesso(self):

        popup = ctk.CTkToplevel(self)
        popup.geometry("300x150")

        ctk.CTkLabel(
            popup,
            text="Salvo com sucesso!"
        ).pack(pady=30)

        ctk.CTkButton(
            popup,
            text="OK",
            command=popup.destroy
        ).pack()