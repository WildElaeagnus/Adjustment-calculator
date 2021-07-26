# pick a file and its path will become a fillename var
from tkinter import *
from tkinter import filedialog
import os

class file_browser_():
	"""docstring for file_browser_class
    pick a file and its path will become a filename var"""
	# path to file
	filename = None
	curr_directory = os.getcwd()
	def __init__(self, ):
		super().__init__()
		
	# a file explorer

	def file_browser_(self, ):

		# Function for opening the file explorer window
		def browseFiles():
			self.filename = filedialog.askopenfilename(initialdir = self.curr_directory,
												title = "Select a File",
												filetypes = (("DPT files",
																"*.xlsx*"),
															("all files",
																"*.*")))
			
			# Change label contents
			label_file_explorer.configure(text="File Opened: "+str(self.filename))

		def pepe():
			print(self.filename)
			window.destroy()
			return()
		def exit():
			window.destroy()
			return()	

		print(self.filename)

		# Create the root window
		window = Tk()

		# Set window title
		window.title('File Explorer')

		# Set window size
		window.geometry("700x100")

		#Set window background color
		window.config(background = "white")

		# Create a File Explorer label
		label_file_explorer = Label(window,
									text = "File Explorer for .xlsx",
									width = 100, height = 4,
									fg = "blue")

			
		button_explore = Button(window,
								text = "Browse Files",
								command = browseFiles)


		button_open = Button(window,
								text = "Open an excel file",
								command = pepe)


		button_exit = Button(window,
							text = "Exit",
							command = exit)

		label_file_explorer.grid(column = 1, row = 2, columnspan=1000)

		button_explore.grid(column = 1, row = 1)

		button_open.grid(column = 2, row = 1)

		button_exit.grid(column = 3,row = 1)

		# Let the window wait for any events
		window.mainloop()
		

if __name__ == "__main__":
	hello = file_browser_()
	hello.file_browser_()