from tkinter import *
import runpy
from random import choice
import string
from xlrd import open_workbook
import openpyxl
from tkinter import messagebox

global i

def find_data(index,find):
	work = open_workbook("_details.xlsx")
	sheet = work.sheet_by_index(0)
	for i in range(10000):
		try:
			val = sheet.cell_value(i,index)
			if val == find:
				code.set(sheet.cell_value(i,0))
				item.set(sheet.cell_value(i,1))
				price.set(sheet.cell_value(i,2))
				stock.set(sheet.cell_value(i,3))
				return i
		except Exception as e:
			print(e)
			continue
	code.set("Not Found")
	item.set("Not Found")
	price.set("Not Found")
	stock.set("Not Found")

	return 0




def search():
	global i

	CODE = code.get()
	ITEM = item.get()
	PRICE = price.get()
	STOCK = stock.get()

	if len(CODE) > 0:
		i = find_data(0,CODE)

	elif len(ITEM) > 0:
		i = find_data(1,ITEM)

	elif len(PRICE) > 0:
		i = find_data(2,PRICE)

	elif len(STOCK) >= 0:
		i = find_data(3,STOCK)

	

def close():
	add.destroy()
	runpy.run_path('Stocks manager.py')

def write():
	global i
	work = open_workbook("_details.xlsx")
	sheet = work.sheet_by_index(0)
	cd = code.get()
	name = item.get()
	pr = price.get()
	amt = stock.get()
	if len(cd) == 0 or len(name) == 0 or len(pr) == 0 or len(amt) == 0:
		messagebox.showinfo('Fill all fields','All fields must be filled')
		return
	try:
		amt = int(float(amt))
	except Exception as e:
		print(e)
		messagebox.showinfo('Invalid','Inventory stocks must be a number')
		return
	try:
		pr = float(pr)
	except:
		messagebox.showinfo('Invalid','Price must be number')
		return
	workbook = openpyxl.load_workbook('_details.xlsx')
	# print(workbook['Details'])
	sheet1 = workbook['Details']
	
	sheet1['A' + str(i+1)] = cd
	sheet1['B' + str(i+1)] = name
	sheet1['C' + str(i+1)] = pr
	sheet1['D' + str(i+1)] = amt
	workbook.save('_details.xlsx')
	close()
		

	

add = Tk()
add.configure(bg = '#ff4d4d')


WIDTH = add.winfo_screenwidth()
HEIGHT = add.winfo_screenheight()
width = 500
height = 500

add.geometry('%dx%d+%d+%d'%(width,height,WIDTH//2 - width // 2, HEIGHT //2 - height // 2))
add.title("modify items")
add.iconbitmap(r'images\\icon.ico')

code = StringVar()
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
Button(frame,text = "Search", bg = "#ff4d4d", command = search).grid(row = 5 , column = 1, pady = 35, ipadx = 20, padx = 30)
Button(frame,text = "Modify", bg = "#ff4d4d", command = write).grid(row = 5 , column = 2, pady = 35, ipadx = 20, padx = 30)
Button(frame, text = "Cancel" , bg = "#ff4d4d" ,  command = close).grid(row = 5 , column = 3, pady = 35, ipadx = 30, padx = 20)

add.mainloop()