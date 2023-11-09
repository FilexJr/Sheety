import tkinter
import tkinter.ttk
from tkinter import ttk

import xlsxwriter

#Start up
window = tkinter.Tk()
window.title("Sheety")
window.geometry("400x300+300+120")
title = tkinter.Label(window, text = "Izaberi predmet:").pack()

def make_ORT2():
    workbook = xlsxwriter.Workbook('Test.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})

    header_format = workbook.add_format(
        {
            "bold": 1,
            "border": 1,
            "text_wrap": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "gray",
        }
    )

    cell_format = workbook.add_format(
        {
            "border": 1,
            "text_wrap": 1,
            "bg_color": "#CCCCCC"
        }
    )

    # Prva tabela
    worksheet.set_column("A:A", 5)
    worksheet.set_column("B:B", 30)
    worksheet.set_column("G:G", 12)
    worksheet.set_row(0, 40)

    worksheet.conditional_format('A2:H8', {'type': 'cell',
                                           'criteria': '>=',
                                           'value': 0, 'format': cell_format})

    worksheet.write('A1', '#', header_format)
    worksheet.write('B1', 'Адресе у меморији са којих је учитана инструкција', header_format)
    worksheet.write('C1', 'IR31..24', header_format)
    worksheet.write('D1', 'IR23..16', header_format)
    worksheet.write('E1', 'IR15..8', header_format)
    worksheet.write('F1', 'IR7..0', header_format)
    worksheet.write('G1', 'Инструкција', header_format)
    worksheet.write('H1', 'PC', header_format)

    # Druga tabela
    worksheet.set_row(11, 40)

    for i in range(12, 19):
        worksheet.merge_range(i, 2, i, 3, '')
        worksheet.merge_range(i, 4, i, 5, '')
        worksheet.merge_range(i, 6, i, 7, '')

    worksheet.conditional_format('A13:H19', {'type': 'cell',
                                             'criteria': '>=',
                                             'value': 0, 'format': cell_format})

    worksheet.write('A12', '#', header_format)
    worksheet.write('B12', 'Адресе у меморији са којих је учитана адреса операнда', header_format)
    worksheet.merge_range('C12:D12', 'Адресе у меморији са којих је учитан операнд', header_format)
    worksheet.merge_range('E12:F12', 'Операнд', header_format)
    worksheet.merge_range('G12:H12', 'Нови садржај регистара опште намене', header_format)

    # Treca tabela
    worksheet.set_row(22, 40)

    for i in range(23, 30):
        worksheet.merge_range(i, 7, i, 9, '')

    worksheet.conditional_format('A24:J30', {'type': 'cell',
                                             'criteria': '>=',
                                             'value': 0, 'format': cell_format})

    worksheet.write('A23', '#', header_format)
    worksheet.write('B23', 'Меморијске адресе којима се приступа у овој фази', header_format)
    worksheet.write('C23', 'N', header_format)
    worksheet.write('D23', 'Z', header_format)
    worksheet.write('E23', 'V', header_format)
    worksheet.write('F23', 'C', header_format)
    worksheet.write('G23', 'Акумулатор', header_format)
    worksheet.merge_range('H23:J23', 'Нови садржај регистара и меморијских локација који су промењени у овој фази',
                          header_format)

    workbook.close()
def make_AOR1_K2():
    workbook = xlsxwriter.Workbook('Test2.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})

    header_format = workbook.add_format(
        {
            "bold": 1,
            "border": 1,
            "text_wrap": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "gray",
        }
    )

    cell_format = workbook.add_format(
        {
            "border": 1,
            "text_wrap": 1,
            "bg_color": "#CCCCCC"
        }
    )

    worksheet.set_column("A:A", 14)

    if option == 1:
        # Prva tabela
        worksheet.conditional_format('A2:K9', {'type': 'cell',
                                               'criteria': '>=',
                                               'value': 0, 'format': cell_format})

        worksheet.write('A1', 'Виртуелна адреса', header_format)
        worksheet.write('B1', 'Тип', header_format)
        worksheet.write('C1', 'User', header_format)
        worksheet.write('D1', 'Segment', header_format)
        worksheet.write('E1', 'Page', header_format)
        worksheet.write('F1', 'Word', header_format)
        worksheet.write('G1', 'Tag', header_format)
        worksheet.write('H1', 'Entry', header_format)
        worksheet.write('I1', 'Коментар', header_format)
        worksheet.write('J1', 'Block', header_format)
        worksheet.write('K1', 'Физичка адреса', header_format)

    workbook.close()
def make_an_excel_table():
    if subject.get() == 'ORT2':
        make_ORT2()
    elif subject.get() == 'AOR1 K2':
        make_AOR1_K2()

subject = tkinter.StringVar()
combobox_choose_table = ttk.Combobox(window, textvariable=subject)
combobox_choose_table['values'] = ('ORT2', 'AOR1 K2')
combobox_choose_table['state'] = 'readonly'
combobox_choose_table.pack()

button_generate_table = tkinter.Button(window, text = "Generiši tabelu", command = make_an_excel_table).pack()

option = tkinter.IntVar()
radiobutton_prva_opcija = tkinter.Radiobutton(window, text="Prva opcija", variable=option, value=1).pack()
radiobutton_druga_opcija = tkinter.Radiobutton(window, text="Druga opcija", variable=option, value=2).pack()


window.mainloop()


