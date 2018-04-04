import tkinter as tk
import xlrd
import random

window = tk.Tk()
window.title('山东理工大学学生抽签系统')
window.geometry('800x500')

label_1 = tk.Label(window,text='SDUT  自动化1303班',font=('黑体',25),width=20,heigh=2)
label_1.place(x=200,y=30,anchor='nw')

var = tk.StringVar()
var.set('ARE YOU READY')
label_2 = tk.Label(window,fg='red',bg='lightyellow',textvariable=var,font=('黑体',20),width=15,heigh=2)
label_2.place(x=265,y=150,anchor='nw')

on_click = False
def Click():
    global on_click,num,nrows,table
    if on_click == False:
        on_click = True
        button.config(text='停  止',fg='red')
        file = xlrd.open_workbook('mingdan.xlsx')
        table = file.sheets()[0]
        nrows = table.nrows
        num = random.randint(0,nrows)
        name = table.row_values(num)
        label_2.config(fg='blue')
        var.set(name)
    else:
        on_click = False
        button.config(text='开  始',fg='blue')

    def loop():
        global num, nrows
        if on_click == False:
            return
        num += 1
        if num == nrows:
            num = 0
        name = table.row_values(num)
        var.set(name)
        label_2.after(100, loop)
    loop()

button = tk.Button(window,text='开  始',fg='blue',font=('黑体',15),width=15,heigh=2,command=Click)
button.place(x=290,y=300,anchor='nw')
window.mainloop()









