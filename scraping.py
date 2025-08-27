from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import time
import pyautogui
from pywinauto import Desktop
from datetime import datetime, timedelta
import pandas as pd
import xlwings as xw
import os
from dotenv import load_dotenv
from pathlib import Path
import unicodedata

load_dotenv()
usuario = os.getenv("CORP_USER")
senha = os.getenv("CORP_PASS")

# ========================= 
# Automação Corporate
# =========================

Desktop(backend="win32").window(title_re=".*Visual Studio Code.*").minimize()

app = Application(backend="win32").start(r"X:\Corporate.exe")
main_window = app.window(title="Corporate Systems")
main_window.wait("visible", timeout=20)

time.sleep(3)
login_window = app.window(title="Login")
login_window.wait("visible", timeout=20)

login_window.child_window(class_name="TEdit", found_index=1).click_input()
send_keys(usuario)
login_window.child_window(class_name="TEdit", found_index=0).click_input()
send_keys(senha)
login_window.child_window(title="OK", class_name="TBitBtn").click_input()

time.sleep(2)
send_keys("{ENTER}")
time.sleep(5)

gestao_window = app.window(title_re=".*Gestão Empresarial.*")
gestao_window.wait("visible", timeout=20)
gestao_window.set_focus()
time.sleep(1)

pyautogui.click(x=114, y=55)
time.sleep(2)
pyautogui.click(x=199, y=98)
time.sleep(3)
pyautogui.click(x=1034, y=552)
send_keys("{ENTER}")
time.sleep(2)
pyautogui.click(x=1704, y=37)
time.sleep(1)
pyautogui.click(x=1716, y=77)
pyautogui.click(x=1611, y=77)
time.sleep(5)

relatorio_window = Desktop(backend="win32").window(title_re=".*Cargas.*")
relatorio_window.wait("visible", timeout=20)
relatorio_window.set_focus()
time.sleep(1)

inicio = (datetime.today() - timedelta(days=1)).replace(hour=0, minute=0, second=0).strftime("%d%m%Y %H:%M:%S")
final  = (datetime.today() - timedelta(days=1)).replace(hour=23, minute=59, second=0).strftime("%d%m%Y %H%M%S")

pyautogui.write(inicio, interval=0.2)
pyautogui.press('tab')
pyautogui.write(final, interval=0.2)

time.sleep(0.5)
pyautogui.click(x=1176, y=483)
time.sleep(2)
pyautogui.write("relatorio_de_cargas", interval=0.05)
pyautogui.press('enter')
time.sleep(2)
pyautogui.press('enter')
time.sleep(5)
pyautogui.press('enter')

try:
    main_window.close()
except:
    send_keys("%{F4}")

time.sleep(2)

# =========================
# Atualização segura do Excel consolidado
# =========================

arquivo_origem = Path(os.getenv("ARQUIVO_ORIGEM"))
arquivo_destino = Path(os.getenv("ARQUIVO_DESTINO"))
aba_destino = "Carga"

timeout = 30
contador = 0
while not os.path.exists(arquivo_origem) and contador < timeout:
    time.sleep(1)
    contador += 1

if not os.path.exists(arquivo_origem):
    raise FileNotFoundError(f"Arquivo não encontrado: {arquivo_origem}")

df_novo = pd.read_excel(arquivo_origem)

app_excel = xw.App(visible=False)
wb = app_excel.books.open(arquivo_destino)
ws = wb.sheets[aba_destino]

dados_existentes = ws.range("A1").expand().value
df_existente = pd.DataFrame(dados_existentes[1:], columns=dados_existentes[0])

def deduplicar_colunas(colunas):
    contador = {}
    novas_colunas = []
    for col in colunas:
        col = str(col).strip()
        if col in contador:
            contador[col] += 1
            novas_colunas.append(f"{col}_{contador[col]}")
        else:
            contador[col] = 0
            novas_colunas.append(col)
    return novas_colunas

def padronizar_df(df):
    df = df.copy()
    for col in df.columns:
        df[col] = (
            df[col]
            .astype(str)
            .str.strip()
            .str.lower()
            .apply(lambda x: unicodedata.normalize("NFKD", x).encode("ASCII", "ignore").decode("ASCII"))
        )
    return df

# Define colunas relevantes para comparação
colunas_chave = [
    "Autorização", "Data Chegada", "Cliente",
    "Material", "Placa"
]

def gerar_chave(df, colunas):
    df = df.copy()
    df = df[colunas].fillna("").astype(str)
    df = df.applymap(lambda x: unicodedata.normalize("NFKD", x).encode("ASCII", "ignore").decode("ASCII").strip().lower())
    df = df.applymap(lambda x: x.replace(".0", "") if x.endswith(".0") else x)
    return df.apply(lambda row: '|'.join(row.values), axis=1)


# Prepara os DataFrames
df_existente.columns = deduplicar_colunas(df_existente.columns)
df_novo.columns = [str(c).strip() for c in df_novo.columns]
df_existente.columns = [str(c).strip() for c in df_existente.columns]

df_novo_pad = padronizar_df(df_novo)
df_existente_pad = padronizar_df(df_existente)

# Verifica se todas as colunas-chave existem
for col in colunas_chave:
    if col not in df_novo.columns:
        raise KeyError(f"Coluna '{col}' não encontrada no relatório novo.")
    if col not in df_existente.columns:
        raise KeyError(f"Coluna '{col}' não encontrada na planilha consolidada.")

df_novo_pad['chave'] = gerar_chave(df_novo_pad, colunas_chave)
df_existente_pad['chave'] = gerar_chave(df_existente_pad, colunas_chave)

chaves_existentes = set(df_existente_pad['chave'])
df_para_inserir = df_novo[df_novo_pad['chave'].apply(lambda x: x not in chaves_existentes)]

# =========================
# Inserção segura
# =========================

if not df_para_inserir.empty:
    print(f"🔄 Inserindo {len(df_para_inserir)} novos registros...")
    ultima_linha = ws.range("A" + str(ws.cells.last_cell.row)).end("up").row + 1
    ws.range(f"A{ultima_linha}").value = df_para_inserir.values.tolist()
else:
    print("✅ Nenhum registro novo encontrado. Excel já está atualizado.")

# Registro de atualização na aba de log
if "LogAtualizacao" not in [sheet.name for sheet in wb.sheets]:
    wb.sheets.add("LogAtualizacao")
log_sheet = wb.sheets["LogAtualizacao"]

if log_sheet.range("A1").value is None:
    log_sheet.range("A1").value = ["Data/Hora", "Registros Inseridos", "Arquivo Origem"]

data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
quantidade = len(df_para_inserir)
arquivo_nome = os.path.basename(arquivo_origem)

ultima_linha_log = log_sheet.range("A" + str(log_sheet.cells.last_cell.row)).end("up").row + 1
log_sheet.range(f"A{ultima_linha_log}").value = [data_hora, quantidade, arquivo_nome]

wb.save()
wb.close()
app_excel.quit()
