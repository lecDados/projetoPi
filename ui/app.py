import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
from openpyxl import load_workbook
import shutil
import os

# Importação da tela de dashboard
from dashboard import DashboardWindow 

from services.excel_service import salvar_no_excel
from config.colunas import colunas

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Cadastro Excel")
        self.geometry("700x900")
        self.arquivo_path = None
        self.entries = {}
        self.foto_path = None
        self.anexo_path = None
        self.criar_interface()

    def validar_numero(self, P):
        return P == "" or P.isdigit()

    def abrir_dashboard(self):
        # Agora abre sem precisar da verificação do arquivo_path
        DashboardWindow(self, self.arquivo_path)

    def criar_interface(self):
        vcmd = (self.register(self.validar_numero), '%P')

        # Frame de Navegação Superior
        nav_frame = ctk.CTkFrame(self, fg_color="transparent")
        nav_frame.pack(pady=10, fill="x", padx=20)

        ctk.CTkLabel(nav_frame, text="Sistema de Cadastro Excel", font=("Arial", 20, "bold")).pack(side="left")
        
        # BOTÃO PARA ABRIR DASHBOARD
        ctk.CTkButton(nav_frame, text="📊 Ver Dashboard", command=self.abrir_dashboard, 
                      fg_color="#1f538d", hover_color="#14375e").pack(side="right")

        ctk.CTkButton(self, text="Selecionar Arquivo Excel", command=self.selecionar_arquivo).pack(pady=5)
        ctk.CTkButton(self, text="Gerar Relatório TXT", command=self.gerar_relatorio).pack(pady=5)
        self.label_arquivo = ctk.CTkLabel(self, text="Nenhum arquivo selecionado")
        self.label_arquivo.pack()

        frame = ctk.CTkScrollableFrame(self, width=600, height=600)
        frame.pack(pady=10)

        opcoes_numeros = [str(i) for i in range(101)]
        for col in colunas:
            texto_label = col
            col_lower = col.lower()
            if "conte" in col_lower: texto_label = "Os conteúdos e temas abordados..."
            elif "visita" in col_lower: texto_label = "O que achou da visita..."
            elif "tempo dedicado" in col_lower: texto_label = "Tempo dedicado..."

            ctk.CTkLabel(frame, text=texto_label).pack(anchor="w", padx=10)

            if col_lower in ["participantes", "lista de presença", "número participantes"]:
                combo = ctk.CTkComboBox(frame, values=opcoes_numeros, width=500)
                combo.pack(padx=10, pady=5); combo.set("0")
                self.entries[col] = combo
            elif col == "Turmas":
                var = ctk.StringVar(value="")
                f = ctk.CTkFrame(frame); f.pack(padx=10, pady=5, anchor="w")
                for opt in ["Manhã", "Tarde"]:
                    ctk.CTkRadioButton(f, text=opt, value=opt, variable=var).pack(side="left", padx=5)
                self.entries[col] = var
            elif col == "Data":
                df = ctk.CTkFrame(frame); df.pack(padx=10, pady=5)
                cal = DateEntry(df, width=20, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
                cal.pack(side="left", padx=5)
                hora = ctk.CTkEntry(df, width=100, placeholder_text="Hora"); hora.pack(side="left", padx=5)
                self.entries["Data"] = cal; self.entries["Hora"] = hora
            elif col == "Foto da ação":
                ctk.CTkButton(frame, text="Selecionar Foto", command=self.selecionar_foto).pack(padx=10, pady=5)
                self.label_foto = ctk.CTkLabel(frame, text="Nenhuma foto selecionada"); self.label_foto.pack()
                self.entries[col] = ctk.StringVar(value="")
            elif any(x in col_lower for x in ["conte", "visita", "tempo dedicado"]):
                var = ctk.StringVar(value="")
                f = ctk.CTkFrame(frame); f.pack(padx=10, pady=5, anchor="w")
                for opt in ["Ótimo", "Bom", "Regular", "Ruim"]:
                    ctk.CTkRadioButton(f, text=opt, value=opt, variable=var).pack(side="left", padx=5)
                self.entries[col] = var
            else:
                entry = ctk.CTkEntry(frame, width=500); entry.pack(padx=10, pady=5)
                self.entries[col] = entry

        ctk.CTkLabel(frame, text="lista de presença-anexo", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=(15, 0))
        ctk.CTkButton(frame, text="Anexar Arquivo", command=self.selecionar_anexo, fg_color="#2c3e50").pack(padx=10, pady=5)
        self.label_anexo = ctk.CTkLabel(frame, text="Nenhum arquivo anexado"); self.label_anexo.pack()

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

    def selecionar_anexo(self):
        path = filedialog.askopenfilename(filetypes=[("Todos os arquivos", "*.*")])
        if path:
            self.anexo_path = path
            self.label_anexo.configure(text=os.path.basename(path))

    def salvar(self):
        if not self.arquivo_path:
            messagebox.showerror("Erro", "Selecione um arquivo Excel"); return
        
        dados = {col: (e.get() if hasattr(e, 'get') else "") for col, e in self.entries.items()}
        
        pasta_anexo_base = r"C:\Users\user\Desktop\Projeto pi\anexo"

        if self.foto_path:
            if not os.path.exists("fotos"): os.makedirs("fotos")
            destino = os.path.join("fotos", os.path.basename(self.foto_path))
            shutil.copy(self.foto_path, destino); dados["Foto da ação"] = destino
            
        if self.anexo_path:
            if not os.path.exists(pasta_anexo_base): os.makedirs(pasta_anexo_base)
            destino = os.path.join(pasta_anexo_base, os.path.basename(self.anexo_path))
            shutil.copy(self.anexo_path, destino)
            dados["caminhoanexo"] = destino
            
        if salvar_no_excel(self.arquivo_path, dados, colunas): 
            self.mostrar_sucesso()

    def gerar_relatorio(self):
        if not self.arquivo_path:
            messagebox.showerror("Erro", "Selecione um arquivo Excel"); return
        messagebox.showinfo("Sucesso", "Relatório gerado!")

    def mostrar_sucesso(self):
        popup = ctk.CTkToplevel(self); popup.geometry("300x150"); popup.attributes("-topmost", True)
        ctk.CTkLabel(popup, text="Salvo com sucesso!").pack(pady=30)
        ctk.CTkButton(popup, text="OK", command=popup.destroy).pack()

if __name__ == "__main__":
    app = App()
    app.mainloop()