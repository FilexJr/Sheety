import tkinter
import tkinter.ttk
import xlsxwriter

#Start up
window = tkinter.Tk()
window.title("Sheety")
window.geometry("400x300+300+120")
title = tkinter.Label(window, text = "AOR1: Prvi Kolokvijum").pack()

def make_an_excel_table():
    workbook = xlsxwriter.Workbook('hello.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'Hello world')

    workbook.close()


button_generate_tabel = tkinter.Button(window, text = "Generiši tabelu", command = make_an_excel_table()).pack()

var = tkinter.IntVar(window)
checkbutton_asociativno = tkinter.Radiobutton(window, text = "Asociativnost", variable = var, value = 1).pack()
checkbutton_direktno = tkinter.Radiobutton(window, text = "Direktno", variable = var, value = 2).pack()
checkbutton_setasociativno = tkinter.Radiobutton(window, text = "Set-Asociativnost", variable = var, value = 3).pack()


window.mainloop()


