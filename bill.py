from tkinter import *
from xlrd import open_workbook

from time import localtime,strftime
import threading
from tkinter import messagebox
import random
import openpyxl
import runpy

RED = '#ff4d4d'


def new():
	global inv,total

	total = 0
	Label(frame, text = "Total = "+ str(total),bg = 'white').grid(row = 1, column = 4, padx = 10)

	inv = randomize()
	text.delete("1.0" , "end-1c")
	line = "-" * 103
	bill = "\t\t                      Bill" + "\t\t\t\t Invoice No:" + str(inv) + "\n"
	bill += line + '\n'
	bill += " " * 5 + "Item code" + " " * 11 + "Item name" + " " * 11 + "Price" + " " * 15 + "Quantity" + " " * 9 + "Amount" + "\n"
	bill += line
	text.insert(END, bill)

def randomize():
	inv = ''
	lst = [1,2,3,4,5,6,7,8,9,0,]
	alp = ['A','B','C']

	for i in range(2):
		inv += str(random.choice(alp))
	for i in range(3):
		inv += str(random.choice(lst))

	return inv


def close():
	root.destroy()
	file = runpy.run_path('Stocks manager.py')



def write():
	bill = text.get("1.0", "end-1c")
	filename = 'bills\\' + inv + '.txt'
	file = open(filename,"w+")
	file.write(bill)
	file.close()
	work = open_workbook('bills\\__bills__.xlsx')
	sheet = work.sheet_by_index(0)
	for i in range(10000):
		try:
			val = sheet.cell_value(i,0)
		except:
			break

	workb = openpyxl.load_workbook('bills\\__bills__.xlsx')
	sheet1 = workb['bills']

	sheet1['A' +str(i + 1)] = int(i)
	sheet1['B' +str(i + 1)] = time_string
	sheet1['C' +str(i + 1)] = inv
	sheet1['D' +str(i + 1)] = total
	workb.save('bills\\__bills__.xlsx')


def add():
	global total,flag
	name = choice.get()
	q = qty.get()

	if name == "none":
		messagebox.showinfo('Name error','Name cannot be none', icon = 'warning')
		return
	try:
		q = int(q)
	except:
		q = 0
	if q == 0:
		messagebox.showinfo('Qty error','Invalid Qty', icon = 'warning')
		return
	work = open_workbook('_details.xlsx')
	sheet = work.sheet_by_index(0)
	workb = openpyxl.load_workbook('_details.xlsx')
	sheet1 = workb['Details']
	for i in range(10000):
		try:
			val = sheet.cell_value(i,1)
			if val == name:
				break
		except:
			continue
	code = sheet.cell_value(i,0)
	nme = sheet.cell_value(i,1)
	pr = sheet.cell_value(i,2)
	stock = int(sheet.cell_value(i,3))
	qty1 = str(int(q))
	pr = str(float(pr))
	amount = str(int(qty1) * float(pr))
	remaining = stock - int(qty1)
	if remaining < 0 :
		messagebox.showinfo('NA','Stocks are not available')
		flag = False
		return
	sheet1['D' + str(i + 1)] = remaining
	total += float(amount)
	string = " " * 5 + code + " " * (25 - len(code)) + nme + " " * (25 - len(nme)) + pr + " " * (20 - len(pr)) + qty1 + " " * (25 - len(qty1))+amount  +'\n'
	text.insert(END,string)
	# print(text.get("1.0", "end-1c"))
	Label(frame, text = "Total = "+ str(total),bg = 'white').grid(row = 1, column = 4, padx = 10)
	workb.save('_details.xlsx')

def find():
	global found,choices
	work = open_workbook('_details.xlsx')
	sheet = work.sheet_by_index(0)
	prevfound = []
	while flag:
		try:
			name = item.get()
			found = []
			if len(name) > 0:

				for i in range(1,1000):
					try:
						val = sheet.cell_value(i,1)
						if name.lower() in val.lower():
							found.append(val)
					except:
						continue
			if len(found) > 0 and prevfound != found:
				prevfound = found
				choices.grid_forget()
				choices = OptionMenu(root,choice,'none',*found)
				choices.grid(row = 2, column = 2, ipadx = 50)
				choice.set(found[0])
				root.update()
				print(found)
			if len(name) == 0:
				choice.set('none')
				root.update()
		except:
			pass

found = []
total = 0
named_tuple = localtime()
time_string = strftime("%d/%m/%Y", named_tuple)
print(time_string)

root = Tk()
root.iconbitmap(r'images\\icon.ico')
item = StringVar()
choice = StringVar(value = 'none')
qty = StringVar()

WIDTH = root.winfo_screenwidth()
HEIGHT = root.winfo_screenheight()
flag = True

root.configure(bg = RED)
root.geometry('%dx%d+%d+%d'%(550,500,WIDTH//2 - 275,HEIGHT//2 - 250))
root.title('Bill')

frame = Frame(root,relief = SUNKEN, bd = 4, bg = '#FF1A1A')
frame.place(x = 10 , y = 400 , width = 530, height = 80)

Label(root, text = "Item Name:", bg = RED, font = ('times',15)).grid(row = 1, column = 1, padx = 10, pady = 20)
Entry(root,textvariable = item,font = ('times',15)).grid(row = 1, column = 2, padx = 5, ipadx = 20,ipady = 5)

Label(root,text = "Choose Item:", bg = RED, font = ('times',15)).grid(row = 2, column = 1, padx = 5, pady = 20)
choices = OptionMenu(root,choice,'none',*found)
choices.grid(row = 2, column = 2, padx = 0, ipadx = 50)

Label(root, text = "Qty:" , bg = RED, font = ('times',15)).grid(row = 2, column = 3)
Entry(root, textvariable = qty).grid(row = 2, column = 4, ipadx = 1)

Button(root, text = "Add", bg = '#FF1A1A', command = add).grid(row = 3, column = 2, ipadx = 75)

frame2 = Frame(root, relief = RAISED, bd = 4, bg = RED)
frame2.place(x = 10, y = 190, width = 530 , height = 200)
text = Text(frame2,font = ('calibri',12,'bold'))
text.place(x = 0, y= 0, width = 520, height = 190)

inv = ''

inv = randomize()

line = "-" * 103
bill = "\t\t                      bill" + "\t\t\t\t Invoice No:" + str(inv) + "\n"
bill += line + '\n'
bill += " " * 5 + "Item code" + " " * 11 + "Item name" + " " * 11 + "Price" + " " * 15 + "Quantity" + " " * 9 + "Amount" + "\n"
bill += line
text.insert(END, bill)

add = Button(frame, text = "Print", bg = RED, command = write)
add.grid(row = 1, column = 1, padx = 25, pady = 20, ipadx = 30)

back = Button(frame, text = "Back",command = close, bg = RED)
back.grid(row = 1, column = 2, padx = 20 , ipadx = 30)

new = Button(frame, text = "New", bg = RED, command = new)
new.grid(row = 1, column = 3, padx = 25, pady = 20, ipadx = 30)


thread = threading.Thread(target=find)
thread.start()

root.mainloop()