from tkinter import *
import runpy
from random import choice
import string
from xlrd import open_workbook
import openpyxl
from tkinter import messagebox

def getcode():
	key = ''
	length = 5
	for i in range(length):
		# print(string.)
		key += choice(string.digits)
	return key


def close():
	add.destroy()
	runpy.run_path('Stocks manager.py')

def write():
	work = open_workbook("_details.xlsx")
	sheet = work.sheet_by_index(0)
	for i in range(10000):
		try:
			val = sheet.cell_value(i,0)
			if len(val) == 0:
				break
		except:
			break
	workbook = openpyxl.load_workbook('_details.xlsx')
	# print(workbook['Details'])
	sheet = workbook['Details']

	cd = code.get()
	name = item.get()
	pr = price.get()
	amt = stock.get()

	if len(cd) == 0 or len(name) == 0 or len(pr) == 0 or len(amt) == 0:
		messagebox.showinfo('Fill all fields','All fields must be filled')
		return
	try:
		amt = int(amt)
	except:
		messagebox.showinfo('Invalid','Inventory stocks must be a number')
		return
	try:
		pr = float(pr)
	except:
		messagebox.showinfo('Invalid','Price must be number')
		return

	sheet['A' + str(i+1)] = cd
	sheet['B' + str(i+1)] = name
	sheet['C' + str(i+1)] = pr
	sheet['D' + str(i+1)] = int(amt)

	workbook.save('_details.xlsx')
	close()

add = Tk()
add.configure(bg = '#ff4d4d')


WIDTH = add.winfo_screenwidth()
HEIGHT = add.winfo_screenheight()
width = 500
height = 500

add.geometry('%dx%d+%d+%d'%(width,height,WIDTH//2 - width // 2, HEIGHT //2 - height // 2))
add.title('Add new item')
add.iconbitmap(r'images\\icon.ico')

code = StringVar(value = getcode())
item = StringVar()
price = StringVar()
stock = StringVar()

Label(add,text = "Item Code:", bg = "#ff4d4d", font = ('times',15)).grid(row = 1, column = 1, padx = 20, pady = 40)
Entry(add, textvariable = code).grid(row = 1, column = 2, padx = 5, ipadx = 50)

Label(add,text = "Item Name:", bg = "#ff4d4d", font = ('times',15)).grid(row = 2, column = 1, padx = 20, pady = 25)
Entry(add, textvariable = item).grid(row = 2, column = 2, padx = 5, ipadx = 50)

Label(add,text = "Price (each):", bg = "#ff4d4d", font = ('times',15)).grid(row = 3, column = 1, padx = 20, pady = 25)
Entry(add, textvariable = price).grid(row = 3, column = 2, padx = 5, ipadx = 50)

Label(add,text = "Available stock:", bg = "#ff4d4d", font = ('times',15)).grid(row = 4, column = 1, padx = 20, pady = 25)
Entry(add, textvariable = stock).grid(row = 4, column = 2, padx = 5, ipadx = 50)


frame = Frame(add, relief = SUNKEN, bd = 3, bg = "#FF1A1A")

frame.place(x = 20, y = 350 , width = 460, height = 100)
Button(frame,text = "Submit", bg = "#ff4d4d", command = write).grid(row = 5 , column = 2, pady = 35, ipadx = 30, padx = 55)
Button(frame, text = "Cancel" , bg = "#ff4d4d" ,  command = close).grid(row = 5 , column = 3, pady = 35, ipadx = 30, padx = 40)

add.mainloop()