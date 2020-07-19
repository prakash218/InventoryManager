from tkinter import *
import openpyxl
from xlrd import open_workbook
import runpy
from PIL import ImageTk, Image
from tkinter import messagebox
from win10toast import ToastNotifier

hr = ToastNotifier()

RED = '#ff4d4d'
GRAY = '#9f9f9f'
# GRAY = 'white'
WHITE = 'white'

BUTTONCOLOR = '#FF1A1A'

size = 14


def close():
	root.destroy()
	file = runpy.run_path('Start.py')

def billit():
	root.destroy()
	file = runpy.run_path('bill.py')

def addnew():
	root.destroy()
	file = runpy.run_path('add.py')

def search():
	root.destroy()
	file = runpy.run_path('search.py')

def modify():
	root.destroy()
	file = runpy.run_path('modify.py')

def delete():
	root.destroy()
	file = runpy.run_path('delete.py')

def changeadd():
	try:
		root.destroy()
	except:
		pass
	file = runpy.run_path('get_name.py')

def company():

	try:
		with open('files\\company.txt','r') as file:
			f = file.readlines()
			return f[0].split('\n')[0]
		
	except Exception as e:
		print(e)
		changeadd()
def test():
	workb = open_workbook('_details.xlsx')
	sheet = workb.sheet_by_index(0)

	a = 0
	for i in range(1000):
		try:
			Label(frame1,text = sheet.cell_value(i,0) , bg = WHITE, font = ('times',size)).grid(row = a,column = 2,padx = 10)
			Label(frame1,text = sheet.cell_value(i,1) , bg = WHITE,font = ('times',size)).grid(row = a,column = 3,padx = 10)
			if i >= 2:
				Label(frame1,text = str(a - 1),font = ('times',size), bg = WHITE).grid(row = a,column = 1,padx = 10)
				Label(frame1,text = str(float(sheet.cell_value(i,2))),font = ('times',size), bg = WHITE).grid(row = a,column = 4,padx = 10)
				Label(frame1,text = str(int(sheet.cell_value(i,3))),font = ('times',size) , bg = WHITE).grid(row = a,column = 5,padx = 10)
				
			else:
				Label(frame1,text = sheet.cell_value(i,2),font = ('times',size), bg = WHITE).grid(row = a,column = 4,padx = 10)
				Label(frame1,text = sheet.cell_value(i,3),font = ('times',size) , bg = WHITE).grid(row = a,column = 5,padx = 10)
			a+=1
			if sheet.cell_value(i,3) <= 10:
				try:
					hr.show_toast("PMG tech",'item ' + sheet.cell_value(i,1) + ' is less than 10 in number',threaded = True,icon_path = 'images\\icon.ico')
					# messagebox.showinfo('deficient stocks','item ' + sheet.cell_value(i,1) + ' is less than 10 in number')
				except Exception as e:
					print(e)
		except Exception as e:
			continue

		canvas.configure(scrollregion=canvas.bbox('all'), yscrollcommand=scroll_y.set)
		canvas.update_idletasks()
#--Tkinter windo--------------------------------------------------



comp = company()
root = Tk()

WIDTH = root.winfo_screenwidth()
HEIGHT = root.winfo_screenheight()

root.title(company() + "--PMG Tech")
root.geometry('%dx%d+%d+%d'%(WIDTH,HEIGHT,0,0))
root.configure(bg = RED)
root.iconbitmap(r'images\\icon.ico')



Label(root,text = comp , font = ('helvetica',40),  bg = RED).grid(row = 1,padx = WIDTH //4 - len(comp)//2, pady = 20)

appname = "--Inventory manager and bill generator "
Label(root,text = appname , font = ('helvetica',20),  bg = RED).grid(row = 2,padx = WIDTH //4 - len(comp)//2, pady = 20)



#---frame for canvas----------------------------------------------
frame = Frame(root,relief = 'raised',bd = 10, bg = RED)
frame.place(x = 20,y = 200,height = 300,width = WIDTH - 300)

canvas = Canvas(frame,bg = WHITE)
scroll_y = Scrollbar(frame, orient="vertical", command=canvas.yview)
scroll_x = Scrollbar(frame, orient="horizontal",command= canvas.xview)

#main frame=======================================================
frame1 = Frame(canvas,relief = RAISED,bg =WHITE)

canvas.create_window(0, 0, anchor='nw', window=frame1)
canvas.update_idletasks()
canvas.configure(scrollregion=canvas.bbox('all'), yscrollcommand=scroll_y.set)
canvas.configure(scrollregion=canvas.bbox('all'), xscrollcommand=scroll_x.set)
canvas.pack(fill='both', expand=True, side='left')


#buttons frame==============================================================
buttonFrame = Frame(root,bg = BUTTONCOLOR, relief = SUNKEN , bd = 5)
buttonFrame2 = Frame(root,bg = BUTTONCOLOR, relief = SUNKEN , bd = 5)

buttonFrame.place(x = 15, y = 540, width = WIDTH - 30, height = 150)
buttonFrame2.place(x = WIDTH - 270, y = 390, height = 80, width = 255)
scroll_y.pack(fill='y', side='right')


img = Image.open("images\\logo.png")
img = img.resize((245,170) ,  Image.ANTIALIAS)
img = ImageTk.PhotoImage(img)


#buttons================================================

add = Button(buttonFrame, text = "Add Item", bg = RED, command = addnew)
add.grid(row = 1, column = 1, padx = 50, pady = 50, ipadx = 50)

bill = Button(buttonFrame, text = "Bill", bg = RED, command = billit)
bill.grid(row = 1, column = 5, padx = 50, pady = 20, ipadx = 50)


search1= Button(buttonFrame, text = "Search invoice", bg = RED, command = search)
search1.grid(row = 1, column = 4, padx = 50, pady = 20, ipadx = 50)

modify1 = Button(buttonFrame, text = "Modify Item", bg = RED, command = modify)
modify1.grid(row = 1, column = 2, padx = 50, pady = 20, ipadx = 50)

delete1 = Button(buttonFrame, text = "Delete Item", bg = RED, command = delete)
delete1.grid(row = 1, column = 3, padx = 50, pady = 20, ipadx = 50)

buttonFrame3 = Frame(root, relief = SUNKEN , bd = 3, bg = RED)
buttonFrame3.place(x = WIDTH - 270, y = 200, height = 180, width = 255)
Label(buttonFrame3,image = img ).grid()

edit = Button(buttonFrame2, text = "Edit Company", bg = RED, command = changeadd)
edit.grid(row = 2 , column = 1, padx = 20, pady = 20)

back = Button(buttonFrame2, text = 'Back' , bg = RED, command = close)
back.grid(row = 2, column = 2, padx = 10, ipadx = 30)


test()


root.mainloop()