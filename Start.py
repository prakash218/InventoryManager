from tkinter import *
import runpy
from PIL import ImageTk, Image
from xlrd import open_workbook
from tkinter import messagebox
from win10toast import ToastNotifier

hr = ToastNotifier()


root = Tk()
root.iconbitmap(r'images\\icon.ico')
WIDTH = root.winfo_screenwidth()
HEIGHT = root.winfo_screenheight()

def test():
	workb = open_workbook('_details.xlsx')
	sheet = workb.sheet_by_index(0)

	for i in range(1000):
		try:
			
			if sheet.cell_value(i,3) <= 10:
				try:
					hr.show_toast("PMG tech",'item ' + sheet.cell_value(i,1) + ' is less than 10 in number',threaded = True,icon_path = 'images\\icon.ico')
					# messagebox.showinfo('deficient stocks','item ' + sheet.cell_value(i,1) + ' is less than 10 in number')
				except Exception as e:
					print(e)
		except Exception as e:
			continue


def startApp():
	root.destroy()
	file = runpy.run_path('Stocks manager.py')

def close():
	ans = messagebox.askquestion('Are you sure?','Do you want exit the application',icon = 'question')
	# print(ans)
	if ans == 'yes':
		root.destroy()


appname = "Inventory manager and bill generator "
root.geometry('500x500+%d+%d'%(WIDTH//2 - 250 ,HEIGHT//2 - 250))
root.title(appname)

img = Image.open("images\\logo.png")
img = img.resize((278,177) ,  Image.ANTIALIAS)
img = ImageTk.PhotoImage(img)

root.protocol("WM_DELETE_WINDOW", close)
test()

Label(root,text = appname , font = ('helvetica',15),  bg = '#ff4d4d').grid(row = 2,padx = 90, pady = 20)
hack = "Created for Bridge hacks hackathon\n Created by\nPrakash, Girish and Irfan"
Label(root,text = hack , font = ('helvetica',12),  bg = '#ff4d4d').grid(row = 3,padx = 90, pady = 370)


start = Button(root,text = "Start" , command = startApp,bg = '#FF1A1A')
start.place(x = 250 - 50, y = 350, width = 75)

closeb = Button(root, text = "Exit", command = close,bg = '#FF1A1A')
closeb.place(x = 250 - 50 , y = 400, width = 75)

frame = Frame(root,relief = SUNKEN, bd = 10, bg = '#FF1A1A')

frame.place(x = 100 , y = 75, width = 300, height = 200)
Label(frame,image = img ).grid()
root.configure(bg = '#ff4d4d')

root.mainloop()