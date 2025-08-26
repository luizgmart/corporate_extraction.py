# ⚙️ Automação da Extração dos Dados do Relatório de Cargas e Atualização do Banco de Dados

## 📌 Descrição

Este projeto automatiza a extração de dados do sistema **Corporate**, exporta relatórios em Excel, valida os dados contra um consolidado existente e insere **somente os registros novos**. Tudo isso com rastreabilidade via log e compatibilidade com agendamento automático.

---

## 🚀 Funcionalidades

- Login automático no sistema Corporate
- Navegação e exportação do relatório de cargas
- Leitura e validação de dados com `pandas`
- Inserção segura no Excel consolidado com `xlwings`
- Registro de atualização em aba de log
- Proteção contra duplicidade de registros
- Pronto para agendamento via Task Scheduler

---

## 🧰 Tecnologias utilizadas

- Python 3.13+
- [pywinauto](https://pywinauto.readthedocs.io/)
- [pyautogui](https://pyautogui.readthedocs.io/)
- [xlwings](https://docs.xlwings.org/)
- [pandas](https://pandas.pydata.org/)
- [dotenv](https://pypi.org/project/python-dotenv/)

---

## 🔐 Segurança

Este projeto utiliza variáveis de ambiente para proteger credenciais e caminhos sensíveis.  
Antes de executar, crie um arquivo `.env` com os seguintes campos:

```ini
CORP_USER=seu_usuario
CORP_PASS=sua_senha
ARQUIVO_ORIGEM=CAMINHO/DO/ARQUIVO/ORIGEM.Xls
ARQUIVO_DESTINO=CAMINHO/DO/ARQUIVO/DESTINO.xlsx
