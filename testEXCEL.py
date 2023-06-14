import openpyxl

book = openpyxl.load_workbook(r"/Programas noobs/insultos/proyecto.xlsx")
sheet = book.active


def lista_adj():
    excel_adj = sheet["A"]
    adj = []
    for word in excel_adj:
        adj.append(word.value)
    adj.pop()
    return adj


def lista_insult():
    excel_insult = sheet["B"]
    insult = []
    for word in excel_insult:
        insult.append(word.value)
    insult.pop()
    return insult


def lista_noun():
    excel_noun = sheet["C"]
    noun = []
    for word in excel_noun:
        noun.append(word.value)
    noun.pop()
    return noun
