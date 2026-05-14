from pymongo import MongoClient

client = MongoClient("mongodb://127.0.0.1:27017/")

db = client["sistema"]
usuarios = db["usuarios"]

usuarios.insert_one({
    "usuario": "lucimara",
    "senha": "1234"
})

print("Usuário criado!")