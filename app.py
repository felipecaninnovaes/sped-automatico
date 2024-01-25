import datetime
from time import sleep

import openpyxl
import pyautogui
import pyperclip

# 552,653   (Obrigações do ICMS)
# 303,723   (+)
# 636,375   (Codigo ST) 999
# 678,399   (Valor ST) Da Planilha
# 1008,397  (Data de Vencimento) da Planilha
# 612,424   (Codigo da receita) 100099
# 823,519   (Mes de Referencia) (Mes da apuração)
# 758,583	(Salvar)

Estado = "PR"
Codigo_ST = "999"
Codigo_Receita = "100099"
Mes_Apuração = "12/2023"


# Entrar na planilha
workbook = openpyxl.load_workbook("planilha.xlsx")
sheet_estado = workbook[Estado]
# Copiar informação de um campo e colar no seu campo correspondente

sleep(4)

# Obrigações do ICMS
pyautogui.click(552, 653, duration=1)

for linha in sheet_estado.iter_rows(min_row=2):
    # +
    pyautogui.click(303, 723, duration=1)

    # Codigo ST 999
    pyperclip.copy(Codigo_ST)
    pyautogui.click(636, 375, duration=1)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")

    # Valor ST
    valor = linha[39].value
    valor = str(valor).replace(".", ",")
    print(valor)
    pyperclip.copy(valor)
    pyautogui.click(678, 399, duration=1)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")

    # Data de Vencimento
    data = linha[1].value
    if isinstance(data, datetime.datetime):
        data = data.strftime("%d%m%Y")
    pyperclip.copy(data)
    pyautogui.click(1008, 397, duration=1)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")

    # Codigo da Receita
    pyperclip.copy(Codigo_Receita)
    pyautogui.click(612, 424, duration=1)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")

    # Mes de Apuração
    pyperclip.copy(Mes_Apuração)
    pyautogui.click(823, 519, duration=1)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")

    # Botão concluir
    pyautogui.click(758, 583, duration=1)
