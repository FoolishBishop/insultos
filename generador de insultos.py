import random
import openpyxl
from colorama import Fore

book = openpyxl.load_workbook(r"C:\Users\saral\PycharmProjects\pythonProject\Programas noobs\insultos\proyecto.xlsx")

sheet = book.active


# aca agarro del excel las cosas y defino el coso aleatorio
def lista_adj():
    excel_adj = sheet["A"]
    adj_l = []
    for word in excel_adj:
        adj_l.append(word.value)
    adj_l.pop()
    return adj_l


def lista_insult():
    excel_insult = sheet["B"]
    insult_l = []
    for word in excel_insult:
        insult_l.append(word.value)
    insult_l = list(set(insult_l))
    insult_l.pop(None)
    return insult_l


def lista_noun():
    excel_noun = sheet["C"]
    noun_l = []
    for word in excel_noun:
        noun_l.append(word.value)
    noun_l = list(set(noun_l))
    noun_l.pop(None)
    return noun_l


def insult_english():
    adj_selected = random.choice(lista_adj())
    ins_selected = random.choice(lista_insult())
    noun_selected = random.choice(lista_noun())
    the_insult = f"result: {adj_selected} {ins_selected} {noun_selected}"
    return Fore.RED + the_insult


# aca defino que es lo que quiero mostrar en consola
walltext_bienvenido = "Welcome, this is an insult generator."
WIns_en = "Generating your random insult..."
options_en = "\n What do you want to do now?"
options2_en = Fore.GREEN + "run / quit \n"

print(walltext_bienvenido)
while True:
    print(Fore.RESET + options_en)
    answer = input(options2_en)
    if answer == "run":
        print(Fore.RESET + WIns_en)
        print(insult_english())
    elif answer == "quit":
        break
    else:
        print("??????")
