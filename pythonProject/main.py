import tkinter
import tkinter.ttk
import xlsxwriter

#Start up
window = tkinter.Tk()
window.title("Sheety")
window.geometry("400x300+300+120")
title = tkinter.Label(window, text = "ORT2: Ispitna tabela").pack()

def make_an_excel_table():
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

    #Prva tabela
    worksheet.set_column("A:A", 10)
    worksheet.set_column("B:B", 30)
    worksheet.set_column("G:G", 12)
    worksheet.set_row(0, 36)

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

    #Druga tabela



    workbook.close()


button_generate_tabel = tkinter.Button(window, text = "Generiši tabelu", command = make_an_excel_table()).pack()

window.mainloop()


