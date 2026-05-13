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
        ctk.CTkLabel(self, text="Sistema de Cadastro Excel", font=("Arial", 20)).pack(pady=10)
        ctk.CTkButton(self, text="Selecionar Arquivo Excel", command=self.selecionar_arquivo).pack(pady=10)
        ctk.CTkButton(self, text="Gerar Relatório TXT", command=self.gerar_relatorio).pack(pady=5)
        self.label_arquivo = ctk.CTkLabel(self, text="Nenhum arquivo selecionado")
        self.label_arquivo.pack()

        frame = ctk.CTkScrollableFrame(self, width=600, height=500)
        frame.pack(pady=10)

        for col in colunas:
            texto_label = col
            if "conte" in col.lower():
                texto_label = "Os conteúdos e temas abordados foram interessantes e contribuíram para o conhecimento."
            elif "visita" in col.lower():
                texto_label = "O que achou da visita do RioPretoPrev Itinerante no seu local de trabalho?"
            elif "tempo dedicado" in col.lower():
                texto_label = "Tempo dedicado ao diálogo/exposição foi adequado."

            ctk.CTkLabel(frame, text=texto_label).pack(anchor="w", padx=10)

            if col == "Turmas":
                var = ctk.StringVar(value="")
                f = ctk.CTkFrame(frame)
                f.pack(padx=10, pady=5, anchor="w")
                for opt in ["Manhã", "Tarde"]:
                    ctk.CTkRadioButton(f, text=opt, value=opt, variable=var).pack(side="left", padx=5)
                self.entries[col] = var
            elif col == "Data":
                df = ctk.CTkFrame(frame)
                df.pack(padx=10, pady=5)
                cal = DateEntry(df, width=20, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
                cal.pack(side="left", padx=5)
                hora = ctk.CTkEntry(df, width=100, placeholder_text="Hora")
                hora.pack(side="left", padx=5)
                self.entries["Data"] = cal
                self.entries["Hora"] = hora
            elif col == "Foto da ação":
                ctk.CTkButton(frame, text="Selecionar Foto", command=self.selecionar_foto).pack(padx=10, pady=5)
                self.label_foto = ctk.CTkLabel(frame, text="Nenhuma foto selecionada")
                self.label_foto.pack()
                self.entries[col] = ctk.StringVar(value="")
            elif any(x in col.lower() for x in ["conte", "visita", "tempo dedicado"]):
                var = ctk.StringVar(value="")
                f = ctk.CTkFrame(frame)
                f.pack(padx=10, pady=5, anchor="w")
                for opt in ["Ótimo", "Bom", "Regular", "Ruim"]:
                    ctk.CTkRadioButton(f, text=opt, value=opt, variable=var).pack(side="left", padx=5)
                self.entries[col] = var
            else:
                entry = ctk.CTkEntry(frame, width=500)
                entry.pack(padx=10, pady=5)
                self.entries[col] = entry

        ctk.CTkButton(self, text="Salvar", command=self.salvar).pack(pady=20)

    def selecionar_arquivo(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            self.arquivo_path = path
            self.label_arquivo.configure(text=os.path.basename(path))

    def selecionar_foto(self):
        path = filedialog.askopenfilename(filetypes=[("Imagens", "*.png *.jpg *.jpeg")])
        if path:
            self.foto_path = path
            self.label_foto.configure(text=os.path.basename(path))

    def salvar(self):
        if not self.arquivo_path:
            messagebox.showerror("Erro", "Selecione um arquivo Excel")
            return
        dados = {col: (e.get() if hasattr(e, 'get') else "") for col, e in self.entries.items()}
        if self.foto_path:
            pasta = "fotos"
            if not os.path.exists(pasta): os.makedirs(pasta)
            destino = os.path.join(pasta, os.path.basename(self.foto_path))
            shutil.copy(self.foto_path, destino)
            dados["Foto da ação"] = destino
        if salvar_no_excel(self.arquivo_path, dados, colunas):
            self.mostrar_sucesso()

    def gerar_relatorio(self):
        if not self.arquivo_path:
            messagebox.showerror("Erro", "Selecione um arquivo Excel"); return
        wb = load_workbook(self.arquivo_path); ws = wb.active
        colunas_rel = [
            {"busca": "conte", "titulo": "Conteúdos abordados"},
            {"busca": "visita", "titulo": "Visita RioPretoPrev"},
            {"busca": "tempo dedicado", "titulo": "Tempo dedicado"}
        ]
        headers = [str(cell.value).strip() for cell in ws[1]]
        relatorio = ""
        for info in colunas_rel:
            idx = next((i+1 for i, h in enumerate(headers) if info["busca"].lower() in h.lower()), None)
            if idx:
                cont = {"Ótimo": 0, "Bom": 0, "Regular": 0, "Ruim": 0}
                for r in range(2, ws.max_row + 1):
                    val = str(ws.cell(row=r, column=idx).value or "").strip().capitalize()
                    if val == "Otimo": val = "Ótimo"
                    if val in cont: cont[val] += 1
                relatorio += f"{info['titulo']}\n" + "\n".join([f"{k}: {v}" for k, v in cont.items()]) + "\n\n"
        with open("relatorio.txt", "w", encoding="utf-8") as f: f.write(relatorio)
        messagebox.showinfo("Sucesso", "Relatório gerado!")

    def mostrar_sucesso(self):
        popup = ctk.CTkToplevel(self)
        popup.geometry("300x150")
        popup.attributes("-topmost", True)
        ctk.CTkLabel(popup, text="Salvo com sucesso!").pack(pady=30)
        ctk.CTkButton(popup, text="OK", command=popup.destroy).pack()