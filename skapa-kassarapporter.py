from openpyxl import *
from os import mkdir
import tkinter as tk
from tkinter import messagebox


def loadFile():  # Laddar in mallen som skall användas för kassarapporterna
	try:  # Försöker ladda in filen
		wb = load_workbook(filename='Ny excel mall.xlsm', read_only=False, keep_vba=True)
		return wb
	except:  # Om det inte finns en sådan fil varnas användaren att filen saknas
		messagebox.showerror("Error", "Hittar inte 'Ny excel mall.xlsm' i mappen")
		return -1

def generateFiles(track, y):  # Skapar en mapp och 12 filer enligt namnkonventionen
	wb = loadFile()  # Laddar mallen
	if wb == -1: # Hoppar över skapandet om mallen inte hittas
		pass
	else:
		pretext = " Kassakontroll " + y + " " + track + " "
		months = ["Januari", "Februari", "Mars", "April", "Maj", "Juni", "Juli", "Augusti", "September", "Oktober", "November", "December"]
		folder = "Kassakontroll " + str(y) + "-" + track

		mkdir(folder)  # Skapar en mapp med namn enligt ovan
		
		for i in range(0,12):  # Skapar varje fil för de enskilda månaderna
			name = folder + "/" + str(i+1) + pretext + months[i] + ".xlsm"
			wb.save(name)
			progress.config(text="Förlopp: " + str(int(i*(100/12))) + " %")  # Uppdaterar förloppet
		progress.config(text="Förlopp: 100 %")  # Uppdaterar förloppet

def GBG():  # Genererar en mapp med kassarapporter för Göteborg
	generateFiles("GBG", year.get())


def KNG():  # Genererar en mapp med kassarapporter för Kungälv
	generateFiles("KNG", year.get())

def on_closing():  # Stänger fönstret utan att orsaka totalkaos
	window.destroy()
	
	
# Skapar ett fönster
window = tk.Tk()

# Lägger till en textruta
text = tk.Label(text="Ange vilket år rapporterna skall skapas för:", width=45)
text.setvar(value="2023")
text.pack()
#text.setvar("2012")

# Lägger till en textinmatningsruta
year = tk.Entry(width=20)
year.pack()

# Lägger till två knappar
gbg = tk.Button(text="GBG", width=20, height=2, bg="#B0B0B0", fg="black", command=GBG)
kng = tk.Button(text="KNG", width=20, height=2, bg="#B0B0B0", fg="black", command=KNG)
gbg.pack(side=tk.LEFT)
kng.pack(side=tk.RIGHT)

# Lägger till förloppsindikator
progress = tk.Label(text="Förlopp: --- %")
progress.pack(side=tk.BOTTOM)

# Lägger till ett protokoll för att hantera stängning av fönstret
window.protocol("WM_DELETE_WINDOW", on_closing)

# Startar loopen som sköter fönsterlogiken
window.mainloop()
