import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
from openpyxl import load_workbook
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

        ctk.CTkButton(
            self,
            text="Gerar Relatório TXT",
            command=self.gerar_relatorio
        ).pack(pady=5)

        self.label_arquivo = ctk.CTkLabel(
            self,
            text="Nenhum arquivo selecionado"
        )

        self.label_arquivo.pack()

        frame = ctk.CTkScrollableFrame(
            self,
            width=600,
            height=500
        )

        frame.pack(pady=10)

        for col in colunas:

            # TEXTO BONITO DAS COLUNAS
            texto_label = col

            if "conte" in col.lower():

                texto_label = "Os conteúdos e temas abordados foram interessantes e contribuíram para o conhecimento."

            elif "visita" in col.lower():

                texto_label = "O que achou da visita do RioPretoPrev Itinerante no seu local de trabalho?"

            elif "tempo dedicado" in col.lower():

                texto_label = "Tempo dedicado ao diálogo/exposição foi adequado."

            # LABEL
            ctk.CTkLabel(
                frame,
                text=texto_label
            ).pack(anchor="w", padx=10)

            # TURMAS
            if col == "Turmas":

                turma_var = ctk.StringVar(value="")

                turma_frame = ctk.CTkFrame(frame)

                turma_frame.pack(
                    padx=10,
                    pady=5,
                    anchor="w"
                )

                for opcao in [
                    "Manhã",
                    "Tarde"
                ]:

                    radio = ctk.CTkRadioButton(
                        turma_frame,
                        text=opcao,
                        value=opcao,
                        variable=turma_var
                    )

                    radio.pack(
                        side="left",
                        padx=5
                    )

                self.entries[col] = turma_var

            # STATUS/SITUAÇÃO
            elif col == "Status/situação atendimento":

                status_var = ctk.StringVar(value="")

                status_frame = ctk.CTkFrame(frame)

                status_frame.pack(
                    padx=10,
                    pady=5,
                    anchor="w"
                )

                for opcao in [
                    "Agendado",
                    "Realizado"
                ]:

                    radio = ctk.CTkRadioButton(
                        status_frame,
                        text=opcao,
                        value=opcao,
                        variable=status_var
                    )

                    radio.pack(
                        side="left",
                        padx=5
                    )

                self.entries[col] = status_var

            # ANO
            elif col == "Ano":

                anos = [
                    str(ano)
                    for ano in range(2020, 2036)
                ]

                combo_ano = ctk.CTkComboBox(
                    frame,
                    values=anos,
                    width=200
                )

                combo_ano.pack(
                    padx=10,
                    pady=5
                )

                self.entries[col] = combo_ano

            # NÚMERO PARTICIPANTES
            elif col == "Número participantes":

                numeros = [
                    str(i)
                    for i in range(1, 501)
                ]

                combo_numero = ctk.CTkComboBox(
                    frame,
                    values=numeros,
                    width=200
                )

                combo_numero.pack(
                    padx=10,
                    pady=5
                )

                self.entries[col] = combo_numero

            # CAMPO DATA
            elif col == "Data":

                data_frame = ctk.CTkFrame(frame)

                data_frame.pack(
                    padx=10,
                    pady=5
                )

                calendario = DateEntry(
                    data_frame,
                    width=20,
                    background='darkblue',
                    foreground='white',
                    borderwidth=2,
                    date_pattern='dd/mm/yyyy'
                )

                calendario.pack(
                    side="left",
                    padx=5
                )

                hora_entry = ctk.CTkEntry(
                    data_frame,
                    width=100,
                    placeholder_text="Hora"
                )

                hora_entry.pack(
                    side="left",
                    padx=5
                )

                self.entries["Data"] = calendario
                self.entries["Hora"] = hora_entry

            # FOTO
            elif col == "Foto da ação":

                botao_foto = ctk.CTkButton(
                    frame,
                    text="Selecionar Foto",
                    command=self.selecionar_foto
                )

                botao_foto.pack(
                    padx=10,
                    pady=5
                )

                self.label_foto = ctk.CTkLabel(
                    frame,
                    text="Nenhuma foto selecionada"
                )

                self.label_foto.pack()

            # CAMPOS DE AVALIAÇÃO
            elif (

                "conte" in col.lower()
                or "visita" in col.lower()
                or "tempo dedicado" in col.lower()
            ):

                radio_var = ctk.StringVar(value="")

                opcoes_frame = ctk.CTkFrame(frame)

                opcoes_frame.pack(
                    padx=10,
                    pady=5,
                    anchor="w"
                )

                for opcao in [
                    "Ótimo",
                    "Bom",
                    "Regular",
                    "Ruim"
                ]:

                    radio = ctk.CTkRadioButton(
                        opcoes_frame,
                        text=opcao,
                        value=opcao,
                        variable=radio_var
                    )

                    radio.pack(
                        side="left",
                        padx=5
                    )

                self.entries[col] = radio_var

            # CAMPOS NORMAIS
            else:

                entry = ctk.CTkEntry(
                    frame,
                    width=500
                )

                entry.pack(
                    padx=10,
                    pady=5
                )

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
                text=os.path.basename(
                    self.arquivo_path
                )
            )

    def selecionar_foto(self):

        self.foto_path = filedialog.askopenfilename(
            filetypes=[
                ("Imagens", "*.png *.jpg *.jpeg")
            ]
        )

        if self.foto_path:

            self.label_foto.configure(
                text=os.path.basename(
                    self.foto_path
                )
            )

    def salvar(self):

        if not self.arquivo_path:

            print("Selecione um arquivo Excel")
            return

        dados = {}

        for col, entry in self.entries.items():

            dados[col] = entry.get()

        # SALVAR FOTO
        if self.foto_path:

            pasta_fotos = "fotos"

            if not os.path.exists(
                pasta_fotos
            ):

                os.makedirs(
                    pasta_fotos
                )

            nome_foto = os.path.basename(
                self.foto_path
            )

            destino = os.path.join(
                pasta_fotos,
                nome_foto
            )

            shutil.copy(
                self.foto_path,
                destino
            )

            dados["Foto da ação"] = destino

        salvar_no_excel(
            self.arquivo_path,
            dados,
            colunas
        )

        self.mostrar_sucesso()

    def gerar_relatorio(self):

        if not self.arquivo_path:

            messagebox.showerror(
                "Erro",
                "Selecione um arquivo Excel"
            )

            return

        wb = load_workbook(self.arquivo_path)

        ws = wb.active

        colunas_relatorio = [

            {
                "busca": "conte",
                "titulo": "Os conteúdos e temas abordados foram interessantes e contribuíram para o conhecimento."
            },

            {
                "busca": "visita",
                "titulo": "O que achou da visita do RioPretoPrev Itinerante no seu local de trabalho?"
            },

            {
                "busca": "tempo dedicado",
                "titulo": "Tempo dedicado ao diálogo/exposição foi adequado."
            }
        ]

        headers = []

        for cell in ws[1]:

            headers.append(
                str(cell.value).strip()
            )

        relatorio = ""

        for coluna_info in colunas_relatorio:

            coluna_busca = coluna_info["busca"]
            titulo_coluna = coluna_info["titulo"]

            indice = None

            for i, header in enumerate(headers):

                if coluna_busca.lower() in header.lower():

                    indice = i + 1
                    break

            if indice is None:

                relatorio += f"{titulo_coluna}\n"
                relatorio += "Coluna não encontrada.\n\n"

                continue

            contagem = {

                "Ótimo": 0,
                "Bom": 0,
                "Regular": 0,
                "Ruim": 0
            }

            for linha in range(2, ws.max_row + 1):

                valor = ws.cell(
                    row=linha,
                    column=indice
                ).value

                if valor:

                    valor = str(valor).strip()

                    valor = valor.capitalize()

                    if valor == "Otimo":

                        valor = "Ótimo"

                    if valor in contagem:

                        contagem[valor] += 1

            relatorio += f"{titulo_coluna}\n\n"

            relatorio += f"Ótimo: {contagem['Ótimo']}\n"
            relatorio += f"Bom: {contagem['Bom']}\n"
            relatorio += f"Regular: {contagem['Regular']}\n"
            relatorio += f"Ruim: {contagem['Ruim']}\n"

            relatorio += "\n--------------------------\n\n"

        with open(
            "relatorio.txt",
            "w",
            encoding="utf-8"
        ) as arquivo:

            arquivo.write(relatorio)

        messagebox.showinfo(
            "Sucesso",
            "Relatório TXT gerado com sucesso!"
        )

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