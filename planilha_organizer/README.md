# 📊 Organizador de Planilhas Pro

Aplicativo desktop para organizar planilhas Excel (.xlsx) e CSV automaticamente.
Interface moderna, escura e intuitiva — feita para usuários não-técnicos.

---

## ✅ Pré-requisitos

- Python 3.10 ou superior
- pip (gerenciador de pacotes do Python)

---

## 🚀 Como instalar e rodar

### 1. Instale as dependências

Abra o terminal (ou Prompt de Comando) na pasta do projeto e rode:

```bash
pip install -r requirements.txt
```

### 2. Execute o aplicativo

```bash
python main.py
```

---

## 🗂️ Estrutura do projeto

```
planilha_organizer/
│
├── main.py                  ← Interface gráfica + ponto de entrada
│
├── modules/
│   ├── financeira.py        ← Organizador de planilhas financeiras
│   ├── estoque.py           ← Organizador de estoque/produtos
│   ├── vendas.py            ← Organizador de vendas/pedidos
│   ├── leads.py             ← Organizador de leads e clientes
│   ├── funcionarios.py      ← Organizador de RH/funcionários
│   └── generica.py          ← Organizador genérico (qualquer planilha)
│
├── utils/
│   └── helpers.py           ← Funções auxiliares compartilhadas
│
├── requirements.txt         ← Dependências do projeto
└── README.md                ← Este arquivo
```

---

## 🎯 Como usar

1. **Selecione seu arquivo** — clique em "Escolher Arquivo" e selecione um `.xlsx` ou `.csv`
2. **Escolha o tipo** — clique no card que melhor representa sua planilha
   - Se não souber, escolha **Genérica**
3. **Clique em "Organizar Planilha Agora"**
4. **Pronto!** O arquivo organizado será salvo na mesma pasta do original

---

## 📋 O que cada tipo faz

| Tipo           | O que organiza                                               |
|----------------|--------------------------------------------------------------|
| **Financeira** | Datas, valores monetários, tipo de transação, ordenação     |
| **Estoque**    | Produtos, SKU, quantidades, status de estoque, valor total  |
| **Vendas**     | Pedidos, clientes, valores, status, mês/ano                 |
| **Leads**      | Nomes, e-mails, telefones, CPF, cidades, deduplicação       |
| **Funcionários**| CPF, salários, tempo de empresa, faixa salarial, cargos    |
| **Genérica**   | Detecta automaticamente tipos de coluna e aplica limpeza    |

---

## 🔧 Regras aplicadas em todos os tipos

- ✅ Remove linhas completamente vazias
- ✅ Remove duplicatas exatas
- ✅ Strip de espaços extras em textos
- ✅ Cabeçalho colorido e linhas alternadas no Excel gerado
- ✅ Colunas com largura automática
- ✅ Cabeçalho fixo (congelado) na planilha gerada
- ✅ Arquivo original **nunca** é modificado

---

## 📦 Arquivo gerado

O arquivo organizado é salvo na **mesma pasta** do original com o nome:

```
NomeOriginal_organizado_TIPO_YYYYMMDD_HHMM.xlsx
```

Exemplo: `vendas_2024_organizado_vendas_20241015_1430.xlsx`

---

## ❓ Problemas comuns

| Problema | Solução |
|----------|---------|
| `ModuleNotFoundError: No module named 'customtkinter'` | Rode `pip install customtkinter` |
| Arquivo não abre depois de gerado | Feche o Excel antes de organizar |
| Planilha com encoding errado (CSV) | Salve o CSV com encoding UTF-8 |
| Erro "planilha vazia" | Verifique se a planilha tem pelo menos 1 linha de dados |
