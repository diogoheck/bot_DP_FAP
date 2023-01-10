# ROBO FAP: PASSOS - 
import pyautogui
from time import sleep
from openpyxl import load_workbook



def carregar_planilha():
    dic_fap = {}
    pasta_fap = load_workbook('R:\\Compartilhado\\DP\\fap\\FAP.xlsx')
    planilha_fap = pasta_fap['FAP']
    for linha in planilha_fap.values:
        # print(type(linha[0]), type(linha[1]))
        dic_fap[str(linha[0])] = str(linha[1])
    return dic_fap


def fazer_lctos_unico_fap(empresas):
    # UNICO FOLHA
    # CADASTROS
    sleep(5)
    pyautogui.click(227,33, duration=2)
    # EMPRESAS
    pyautogui.move(0, 20, duration=2)
    sleep(2)
    pyautogui.click()
    sleep(10)
    for empresa, fap in empresas.items():
        # CÓDIGO EMPRESA*
        pyautogui.typewrite(empresa)
        sleep(2)
        pyautogui.press('enter')
        sleep(10)
        # FOLHA
        pyautogui.click(509,659, duration=2)
        sleep(2)
        # FAP
        pyautogui.doubleClick(506,377, duration=2)
        sleep(2)
        # VALOR FAP*
        pyautogui.typewrite(fap)
        sleep(2)
        pyautogui.press('enter')
        sleep(2)
        # SALVAR COMPET 01/2023
        pyautogui.click(263,125, duration=2)
        sleep(20)
        pyautogui.click(204,179, duration=2)
        sleep(2)
    pyautogui.alert('Finalizado com sucesso')



if __name__ == '__main__':
    dicionario =  carregar_planilha()
    fazer_lctos_unico_fap(dicionario)
    