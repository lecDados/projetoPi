import customtkinter as ctk
from tkinter import messagebox
from pymongo import MongoClient

from ui.app import App

# conexão mongo
client = MongoClient("mongodb://127.0.0.1:27017/")
db = client["sistema"]
usuarios = db["usuarios"]


class LoginWindow(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("Login")
        self.geometry("400x300")

        ctk.CTkLabel(
            self,
            text="Sistema Login",
            font=("Arial", 24, "bold")
        ).pack(pady=30)

        self.user_entry = ctk.CTkEntry(
            self,
            placeholder_text="Usuário",
            width=250
        )
        self.user_entry.pack(pady=10)

        self.pass_entry = ctk.CTkEntry(
            self,
            placeholder_text="Senha",
            show="*",
            width=250
        )
        self.pass_entry.pack(pady=10)

        ctk.CTkButton(
            self,
            text="Entrar",
            command=self.fazer_login
        ).pack(pady=20)

    def fazer_login(self):

        usuario = self.user_entry.get()
        senha = self.pass_entry.get()

        user = usuarios.find_one({
            "usuario": usuario,
            "senha": senha
        })

        if user:
            self.destroy()

            app = App()
            app.mainloop()

        else:
            messagebox.showerror(
                "Erro",
                "Usuário ou senha inválidos"
            )