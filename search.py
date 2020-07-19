from tkinter import *
from xlrd import open_workbook
import sys
from tkinter import messagebox
from os import path
import runpy



lblue = '#FF1A1A'

global root

def submit(num = 5):
	# scrollbar = Scrollbar(search)
	# scrollbar.grid(row = 0, column = 1,stick = 'ns')
	# text = 'invoice number'+ '\t' +'invoice date'+ '\t' +'PO Date'+ '\t' +'PO Number'+ '\t' +'Company name'+ '\t' +'Amount'+ '\t' +'Received'
	global root
	try:
		root.destroy()
	except:
		a = 1
	

	text = ''
	inno = "Serial no:\t\t"
	indate = "Invoce Date:\t\t"
	datepo = "Invoice no:\t\t"
	po = "Amount:"

	s = searchfor.get()
	if len(s) == 0:
		messagebox.showinfo("Inavlid","Please enter a valid Search term");
		return 
	opt = choice.get()

	loc = ("bills\\__bills__.xlsx")
	workb = open_workbook(loc)
	sheet1 = workb.sheet_by_index(0)
	rcd = ''
	if opt == 'invoice number':
		try:
			s = s
		except:
			messagebox.showinfo("Inavlid","Please enter a valid Invoice Number");
			return

		for i in range(10000):
			try:
				inv = sheet1.cell_value(i,0)
				date = sheet1.cell_value(i,1)
				podate = sheet1.cell_value(i,2)
				ponum = sheet1.cell_value(i,3)
				
				
				if podate == s:
					amt = 'Rs.'+str(round(ponum,2))
					
					# text = 'invoice number'+ '\t' +'invoice date'+ '\t' +'PO Date'+ '\t' +'PO Number'+ '\t' +'Company name'+ '\t' +'Amount'+ '\t' +'Received'
					inno += '\n' + str(int(inv)) + '\t\t'
					indate += '\n' + str(date) + '\t\t'
					datepo += '\n' + str(podate) + '\t\t'
					po += '\n' +str(ponum)
					
					print(str(int(inv)) + '\t' + str(date) + '\t' + str(podate) + '\t' + str(ponum) + '\t' + str(comp) + '\t' +str(amt) + '\t' +str(rcd))
					break
			except Exception as e:
				print(e,i)
				continue
	
	if opt == 'Date':
		try:
			s = s
		except:
			messagebox.showinfo("Inavlid","Please enter a valid Invoice Number");
			return

		for i in range(10000):
			try:
				inv = sheet1.cell_value(i,0)
				date = sheet1.cell_value(i,1)
				podate = sheet1.cell_value(i,2)
				ponum = sheet1.cell_value(i,3)
				
				
				if date == s:
					amt = 'Rs.'+str(round(ponum,2))
					
					# text = 'invoice number'+ '\t' +'invoice date'+ '\t' +'PO Date'+ '\t' +'PO Number'+ '\t' +'Company name'+ '\t' +'Amount'+ '\t' +'Received'
					inno += '\n' + str(int(inv)) + '\t\t'
					indate += '\n' + str(date) + '\t\t'
					datepo += '\n' + str(podate) + '\t\t'
					po += '\n' +str(ponum)
					
					print(str(int(inv)) + '\t' + str(date) + '\t' + str(podate) + '\t' + str(ponum) + '\t' + str(comp) + '\t' +str(amt) + '\t' +str(rcd))
					break
			except Exception as e:
				print(e,i)
				continue
	if opt == 'Sl.no':
		try:
			s = s
		except:
			messagebox.showinfo("Inavlid","Please enter a valid Invoice Number");
			return

		for i in range(10000):
			try:

				inv = sheet1.cell_value(i,0)
				date = sheet1.cell_value(i,1)
				podate = sheet1.cell_value(i,2)
				ponum = sheet1.cell_value(i,3)
				print(inv,date,podate,ponum)
				
				if float(inv) == float(s):
					amt = 'Rs.'+str(round(ponum,2))
					
					# text = 'invoice number'+ '\t' +'invoice date'+ '\t' +'PO Date'+ '\t' +'PO Number'+ '\t' +'Company name'+ '\t' +'Amount'+ '\t' +'Received'
					inno += '\n' + str(int(inv)) + '\t\t'
					indate += '\n' + str(date) + '\t\t'
					datepo += '\n' + str(podate) + '\t\t'
					po += '\n' +str(ponum)
					
					print(str(int(inv)) + '\t' + str(date) + '\t' + str(podate) + '\t' + str(ponum) + '\t' + str(comp) + '\t' +str(amt) + '\t' +str(rcd))
					break
			except Exception as e:
				continue
	root = Tk()
	root.geometry('580x400+100+100')
	root.title('Search results')
	root.iconbitmap(r'images\\icon.ico')
	canvas = Canvas(root,bg = RED)
	scroll_y = Scrollbar(root, orient="vertical", command=canvas.yview)
	frame1 = Frame(canvas,relief = RAISED, bg = lblue)
	label = Label(frame1,text = inno,bg = RED)
	label.grid(row = 5,column =1)
	label = Label(frame1,text = indate,bg = RED)
	label.grid(row = 5,column =2)
	label = Label(frame1,text = datepo,bg = RED)
	label.grid(row = 5,column =3)
	label = Label(frame1,text = po,bg = RED)
	label.grid(row = 5,column =4)
	# label = Label(frame1,text = cmpy,bg = lblue)
	# label.grid(row = 5,column =5)
	# label = Label(frame1,text = amount,bg = lblue)
	# label.grid(row = 5,column =6)
	# label = Label(frame1,text = recd,bg = lblue)
	# label.grid(row = 5,column =7)
	canvas.create_window(0, 0, anchor='nw', window=frame1)
# make sure everything is displayed before configuring the scrollregion
	canvas.update_idletasks()
	canvas.configure(scrollregion=canvas.bbox('all'), yscrollcommand=scroll_y.set)
	canvas.pack(fill='both', expand=True, side='left')
	scroll_y.pack(fill='y', side='left')


def close():
	search.destroy()
	file_globals = runpy.run_path('Stocks manager.py')

RED = '#ff4d4d'
blue = RED

search = Tk()
search.title("Search")
search.iconbitmap(r'images\\icon.ico')
searchfor = StringVar()
choice = StringVar()
WIDTH = search.winfo_screenwidth()
HEIGHT = search.winfo_screenheight()
search.geometry('400x200+%d+%d'%(WIDTH//2 - 200, HEIGHT// 2 - 100 ))
search.bind('<Return>',submit)
search.configure(bg = RED)



options = ["invoice number","Date","Sl.no"]
choice.set(options[0])
frame = Frame(search,bg = lblue,relief = RAISED,bd = 10)
frame.grid(row = 3 ,padx = 25,column = 8, pady = 20)

entry = Entry(frame,textvariable = searchfor).grid(row = 1, column = 2,padx = 10)

label = Label(frame,text = "Search:").grid(row = 1,column = 1,padx = 25)

label = Label(frame,text = "Search By:").grid(row = 2,column = 1,padx = 25,pady = 10)

menu = OptionMenu(frame,choice,*options).grid(row = 2, column = 2, padx = 25, pady = 10)

submitbutton = Button(frame,text = "Search", bg = blue,activebackground = lblue, command = submit).grid(row = 3, column = 1,padx = 10)

cancel = Button(frame,text = "Cancel", bg = blue,activebackground = lblue, command = close).grid(row = 3, column = 2,padx = 10)


search.mainloop()