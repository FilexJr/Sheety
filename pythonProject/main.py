import tkinter
import tkinter.ttk
from tkinter import ttk

import xlsxwriter

#Start up
window = tkinter.Tk()
window.title("Sheety")
window.geometry("400x300+300+120")

def make_ORT2():
    workbook = xlsxwriter.Workbook('ORT2.xlsx')
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

    format_vcenter = workbook.add_format()
    format_vcenter.set_align("center", )
    worksheet.set_column('A:XFD', None, format_vcenter)
    workbook.close()

def make_AOR1_K2():
    workbook = xlsxwriter.Workbook('AOR1 K2.xlsx')
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
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#CCCCCC"
        }
    )

    worksheet.set_column("A:A", 14)
    worksheet.set_column("C:C", 12)
    worksheet.set_column("M:M", 12)
    worksheet.set_column("O:O", 12)
    worksheet.set_column("P:P", 12)
    worksheet.set_column("R:R", 20)

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

    #Druga tabela
    worksheet.conditional_format('B14:G39', {'type': 'cell',
                                           'criteria': '>=',
                                           'value': 0, 'format': cell_format})

    worksheet.write('B13', 'Takt', header_format)
    worksheet.write('C13', 'Addres Bus', header_format)
    worksheet.write('D13', 'Ddata Bus', header_format)
    worksheet.write('E13', 'rd', header_format)
    worksheet.write('F13', 'wr', header_format)
    worksheet.write('G13', 'ack', header_format)

    #Treca Tabela
    worksheet.conditional_format('K14:R39', {'type': 'cell',
                                             'criteria': '>=',
                                             'value': 0, 'format': cell_format})

    worksheet.write('K13', 'Приспео', header_format)
    worksheet.write('L13', 'Уређај', header_format)
    worksheet.write('M13', 'Адреса', header_format)
    worksheet.write('N13', 'Податак', header_format)
    worksheet.write('O13', 'Операција', header_format)
    worksheet.write('P13', 'Обрађено', header_format)
    worksheet.write('Q13', 'Ack', header_format)
    worksheet.write('R13', 'Нови захтеви', header_format)

    format_vcenter = workbook.add_format()
    format_vcenter.set_align("center",)
    worksheet.set_column('A:XFD', None, format_vcenter)
    workbook.close()

def make_an_excel_table():
    if subject.get() == 'ORT2':
        make_ORT2()
    elif subject.get() == 'AOR1 K2':
        make_AOR1_K2()

def check_for_different_options(event):
    if subject.get() == 'AOR1 K2':
        combobox_choose_option['values'] = ()
    else:
        combobox_choose_option['values'] = ()
        combobox_choose_option.set('')

title = tkinter.Label(window, text = "DOBAR DAN!!!").pack()

subject_frame = tkinter.Frame(window)
subject_frame.pack()

tkinter.Label(subject_frame, text = "Izaberi predmet: ").pack(side='left', padx=5)
subject = tkinter.StringVar()
combobox_choose_subject = ttk.Combobox(subject_frame, textvariable=subject)
combobox_choose_subject['values'] = ('ORT2', 'AOR1 K2')
combobox_choose_subject['state'] = 'readonly'
combobox_choose_subject.pack(side='right')
combobox_choose_subject.bind('<<ComboboxSelected>>', check_for_different_options)

options_frame = tkinter.Frame(window)
options_frame.pack()

tkinter.Label(options_frame, text = "Izaberi dodatne opcije(ako postoje):").pack(side='left', padx=5)
option = tkinter.StringVar()
combobox_choose_option = ttk.Combobox(options_frame, textvariable=option)
combobox_choose_option['values'] = ()
combobox_choose_option['state'] = 'readonly'
combobox_choose_option.pack(side='right')

button_generate_table = tkinter.Button(window, text = "Generiši tabelu", command = make_an_excel_table).pack()

window.mainloop()


