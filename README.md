# Sistema de Cadastro Excel

Sistema desktop desenvolvido em Python para gerenciamento de cadastros, avaliações e geração de relatórios utilizando interface gráfica com CustomTkinter, integração com Excel e autenticação via MongoDB.

---

## Sobre o Projeto

O sistema foi desenvolvido com foco em facilitar o registro de informações de ações, visitas e atividades realizadas em instituições, permitindo:

- Cadastro de informações em planilhas Excel
- Upload de fotos e anexos
- Controle de participantes
- Avaliações por formulário
- Dashboard para visualização de dados
- Sistema de login integrado ao MongoDB
- Geração de relatórios

A aplicação possui interface gráfica moderna e intuitiva para facilitar a utilização pelos usuários.

---

# Tecnologias Utilizadas

## Linguagem

- Python 3.12

## Interface gráfica

- CustomTkinter
- Tkinter
- tkcalendar

## Banco de dados

- MongoDB
- PyMongo

## Manipulação de arquivos

- OpenPyXL
- Shutil
- OS

## Empacotamento

- PyInstaller

---

# Estrutura do Projeto

```bash
Projeto PI/
│
├── anexo/                 
├── build/                 
├── config/
│   └── colunas.py
│
├── dist/
│
├── fotos/
│
├── relatorio_mes/
│
├── services/
│   └── excel_service.py
│
├── ui/
│   └── app.py
│
├── criar_user.py
├── dashboard.py
├── login.py
├── main.py
├── main.spec
└── relatorio.txt
```

---

# Sistema de Login

O sistema utiliza autenticação via MongoDB.

Exemplo de usuário:

```json
{
  "usuario": "lucimara",
  "senha": "1234"
}
```

O login é validado antes do acesso ao sistema principal.

---

# Funcionalidades

## Cadastro de dados

- Inserção de informações em planilhas `.xlsx`
- Campos dinâmicos
- Seleção de datas
- Avaliações por opção

## Upload de arquivos

- Fotos da ação
- Lista de presença
- Arquivos anexos

## Dashboard

- Visualização dos dados cadastrados
- Acompanhamento das informações registradas

## Relatórios

- Geração de relatórios `.txt`
- Organização por período

## Integração com Excel

- Leitura e escrita automática em planilhas

---

# Como Executar o Projeto

## 1. Clone o repositório

```bash
git clone https://github.com/seu-usuario/seu-repositorio.git
```

---

## 2. Entre na pasta do projeto

```bash
cd projeto-pi
```

---

## 3. Crie o ambiente virtual

```bash
python -m venv .venv
```

---

## 4. Ative o ambiente virtual

### Windows

```bash
.venv\Scripts\activate
```

---

## 5. Instale as dependências

```bash
pip install -r requirements.txt
```

Ou manualmente:

```bash
pip install customtkinter pymongo tkcalendar openpyxl pyinstaller
```

---

# Configuração do MongoDB

O sistema utiliza MongoDB local:

```python
mongodb://127.0.0.1:27017/
```

Crie o usuário executando:

```bash
python criar_user.py
```

---

# Executando o Sistema

```bash
python main.py
```

---

# Gerando Executável (.exe)

```bash
pyinstaller --onefile --windowed main.py
```

O executável será criado em:

```bash
dist/main.exe
```

---

# Dependências Principais

| Biblioteca      | Função                          |
|-----------------|---------------------------------|
| customtkinter   | Interface gráfica moderna       |
| pymongo         | Integração com MongoDB          |
| openpyxl        | Manipulação de Excel            |
| tkcalendar      | Seleção de datas                |
| pyinstaller     | Geração do executável           |

---

# Interface do Sistema

O sistema possui:

- Tela de login
- Cadastro de ações
- Upload de imagens
- Dashboard de acompanhamento
- Relatórios automáticos

---

# Melhorias Futuras

- Criptografia de senhas
- Dashboard avançado
- Exportação PDF
- Sistema multiusuário
- Backup automático
- Banco de dados online
- Controle de permissões

---

# Desenvolvedores

Projeto desenvolvido para fins acadêmicos utilizando Python, MongoDB e tecnologias desktop modernas.