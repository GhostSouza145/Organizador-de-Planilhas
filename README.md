<img width="1040" height="732" alt="image" src="https://github.com/user-attachments/assets/823bc9b1-eaec-4366-93b3-a7a8057b76bb" />
------------------------------------------------------------------------

📊 Organizador Automático de Planilhas em Python

Um aplicativo desktop moderno, bonito e fácil de usar que organiza planilhas automaticamente em poucos cliques.
Ideal para usuários leigos, empresas e qualquer pessoa que trabalhe com Excel ou CSV no dia a dia.

✨ Visão Geral

Este projeto é um organizador geral de planilhas desenvolvido em Python, com interface gráfica criada em CustomTkinter.
O usuário escolhe o tipo de planilha e o sistema faz toda a organização automaticamente, sem necessidade de configurações técnicas.

🚀 Funcionalidades

✔ Interface gráfica moderna (tema dark)
✔ Suporte a Excel (.xlsx) e CSV
✔ Organização automática de dados
✔ Remoção de linhas vazias e duplicadas
✔ Padronização de textos e datas
✔ Ordenação inteligente por tipo de planilha
✔ Preserva o arquivo original
✔ Gera uma nova planilha organizada automaticamente

🧩 Tipos de Planilhas Suportados

O usuário pode escolher o tipo de planilha diretamente na interface:

📁 Financeira

📦 Estoque

💰 Vendas

👥 Leads / Clientes

🧑‍💼 Funcionários

📄 Genérica (organização padrão)

Cada tipo aplica regras específicas para garantir a melhor organização possível.

🖥️ Interface Gráfica

A interface foi pensada para usuários não programadores:

Design moderno e profissional

Botões grandes e intuitivos

Processo simples em poucos cliques

Feedback visual durante a organização

Mensagens claras e amigáveis

🛠️ Tecnologias Utilizadas

Python 3

CustomTkinter (interface gráfica)

Pandas (manipulação de dados)

OpenPyXL (arquivos Excel)

📂 Estrutura do Projeto
organizador-planilhas/
│
├── main.py
├── modules/
│   ├── financeiro.py
│   ├── estoque.py
│   ├── vendas.py
│   ├── leads.py
│   ├── funcionarios.py
│   └── generico.py
│
├── utils/
│   ├── file_utils.py
│   └── data_utils.py
│
├── README.md
└── requirements.txt
▶️ Como Executar o Projeto
1️⃣ Clone o repositório
git clone https://github.com/seu-usuario/organizador-planilhas.git
2️⃣ Acesse a pasta
cd organizador-planilhas
3️⃣ Instale as dependências
pip install -r requirements.txt
4️⃣ Execute o aplicativo
python main.py
📦 Requisitos

Python 3.8 ou superior

Sistema operacional: Windows, Linux ou macOS

🔒 Segurança dos Dados

O arquivo original nunca é alterado

Uma nova planilha organizada é criada automaticamente

Todos os dados permanecem locais no computador do usuário

📈 Possíveis Melhorias Futuras

Exportação em PDF

Histórico de planilhas organizadas

Interface multilíngue

Versão executável (.exe)

Sistema de plugins para novos tipos de planilha

🤝 Contribuições

Contribuições são bem-vindas!
Sinta-se à vontade para abrir issues, enviar pull requests ou sugerir melhorias.

📄 Licença

Este projeto está sob a licença MIT.
Você pode usar, modificar e distribuir livremente.

💡 Projeto pensado para produtividade, simplicidade e organização profissional.
