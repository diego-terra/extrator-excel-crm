# 📋 Extrator de Contatos para CRM

Script Python que varre uma pasta com pedidos em `.xlsx`, extrai automaticamente o **nome do cliente** e o **telefone de contato** de cada arquivo, remove duplicatas e gera uma lista de CRM pronta para uso.

---

## 🚀 Como funciona

1. Percorre todos os arquivos `.xlsx` da pasta (e subpastas)
2. Lê a aba `Plan1` de cada arquivo buscando os rótulos `CLIENTE:` e `CONTATO:`
3. Consolida todos os registros em um único DataFrame
4. Remove duplicatas pelo telefone (ou pelo nome, quando não há telefone)
5. Gera o arquivo `CRM_contatos.xlsx` com os contatos únicos

---

## 📁 Estrutura esperada dos arquivos

Os arquivos de pedido devem seguir este padrão na aba `Plan1`:

| Célula | Conteúdo |
|--------|----------|
| Coluna A | `CLIENTE:` |
| Coluna B | Nome do cliente |
| Coluna D | `CONTATO:` |
| Coluna E | Telefone |

---

## ⚙️ Instalação

Certifique-se de ter o **Python** instalado. Baixe em [python.org](https://www.python.org/downloads/) e marque **"Add Python to PATH"** durante a instalação.

Instale as dependências:

```bash
pip install pandas openpyxl
```

---

## ▶️ Como usar

**Opção 1 — Interativo** (o script pergunta o caminho):
```bash
python extrair_contatos_crm.py
```

**Opção 2 — Direto pela linha de comando:**
```bash
python extrair_contatos_crm.py "C:\Caminho\Da\Sua\Pasta"
```

---

## 📊 Resultado

O arquivo `CRM_contatos.xlsx` é gerado na mesma pasta do script com as colunas:

| Coluna | Descrição |
|--------|-----------|
| `Nome` | Nome do cliente (formatado em Title Case) |
| `Telefone` | Apenas dígitos, sem formatação |
| `Qtd_Pedidos` | Quantos pedidos esse cliente possui |

---

## 🔧 Requisitos

- Python 3.8+
- pandas
- openpyxl
