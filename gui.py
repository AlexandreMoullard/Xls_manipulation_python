from tkinter import *
from tkinter.ttk import *
from PIL import Image, ImageTk
from tkinter import filedialog
import os
import file_maker as fm

class OptionsWindow():
    def __init__(self, dirBancDeTest=None):
        self.running       = False
        self.root          = None
        self.dirBancDeTest = dirBancDeTest

    def show(self):
        if self.running:
            self.root.focus_force()
        else:
            root      = Tk()
            self.root = root

            root.title("Selection de fichier de production")
            root.resizable(width=False, height=False)
            root.iconbitmap('srett.ico')

            style = Style()
            style.configure('.', font=('Helvetica', 16), background="#3d4f6f")
            style.configure("BW.TLabel", foreground="white", background="#3d4f6f", padding=5, anchor=W, width=40)
            style.configure("Directory.TLabel", foreground="black", background="#BBBBBB", padding=5, anchor=W, width=40)
            style.configure("Error.TLabel", foreground="white", background="#3d4f6f")
            
            self.dirPv               = StringVar()
            self.resultTextBdt0Var   = StringVar()
            self.dirBdt1Var          = StringVar()
            self.resultTextBdt1Var   = StringVar()
            self.dirBdt2Var          = StringVar()
            self.resultTextBdt2Var   = StringVar()
            self.dirExp              = StringVar()
            self.resultTextBdt3Var   = StringVar()
            self.first_PN            = IntVar()
            self.text_variable_PN    = StringVar()
            
            #Champ PVAI (sert de template)
            master = Frame(root, padding=5)
            master.pack()

            Label(master, text="PV à remplir (.xlsx)", style="BW.TLabel").grid(row=0, column=0)
            Label(master, textvariable=self.dirPv, style="Directory.TLabel").grid(row=1, column=0)
            Button(master, text="...", command=lambda: self.selectXlsx(1)).grid(row=1, column=1)
            Label(master, textvariable=self.resultTextBdt0Var, style="Error.TLabel").grid(row=2, column=0, columnspan=2)
            
            #Champ banc de test étape 1
            Label(master, text="Banc de test etape 1 (.csv)", style="BW.TLabel").grid(row=3, column=0)
            Label(master, textvariable=self.dirBdt1Var, style="Directory.TLabel").grid(row=4, column=0)
            Button(master, text="...", command=lambda: self.selectCSV(1)).grid(row=4, column=1)
            Label(master, textvariable=self.resultTextBdt1Var, style="Error.TLabel").grid(row=5, column=0, columnspan=2)

            #Champ banc de test étape 2
            Label(master, text="Banc de test etape 2 (.csv)", style="BW.TLabel").grid(row=6, column=0)
            Label(master, textvariable=self.dirBdt2Var, style="Directory.TLabel").grid(row=7, column=0)
            Button(master, text="...", command= lambda: self.selectCSV(2)).grid(row=7, column=1)
            Label(master, textvariable=self.resultTextBdt2Var, style="Error.TLabel").grid(row=8, column=0, columnspan=2)

            #Champ export acceptation
            Label(master, text="Export banc acceptation (.csv)", style="BW.TLabel").grid(row=9, column=0)
            Label(master, textvariable=self.dirExp, style="Directory.TLabel").grid(row=10, column=0)
            Button(master, text="...", command=lambda: self.selectCSV(3)).grid(row=10, column=1)
            Label(master, textvariable=self.resultTextBdt3Var, style="Error.TLabel").grid(row=11, column=0, columnspan=2)

            #Champ du premier PN de génération de fichier
            Label(master, text="Premier PN de génération de fichier", style="BW.TLabel").grid(row=12, column=0)
            first_PN = Entry(master, style="Directory.TLabel")
            first_PN.grid(row=12, column=1)
            Label(master, textvariable=self.text_variable_PN, style="Error.TLabel").grid(row=13, column=0, columnspan=2)

            Button(master, text="Valider", command=lambda: self.validate(self.dirPv.get(), self.dirBdt1Var.get(), self.dirBdt2Var.get(), self.dirExp.get(), first_PN.get())).grid(row=14, column=0, columnspan=2)

            if self.dirBancDeTest:
                self.dirBdt1Var.set(self.dirBancDeTest)
            
            self.running = True
            master.mainloop()
            self.running = False
            return


    def selectCSV(self, step):
        dirname = filedialog.askopenfilename(parent=self.root,initialdir="os.getcwd()",title='Please select a csv file',filetypes=[('csv files','.csv'),('all files','.*')])
        if isinstance(dirname, str):
            if os.name == 'nt':
                dirname = dirname.replace('/', '\\') #Windows compatibility

            if step == 1:
                self.dirBdt1Var.set(dirname)
            elif step == 2:
                self.dirBdt2Var.set(dirname)
            elif step == 3:
                self.dirExp.set(dirname)

    def selectXlsx(self, step):
        dirname = filedialog.askopenfilename(parent=self.root,initialdir="os.getcwd()",title='Please select a xlsx file',filetypes=[('xlsx files','.xlsx'),('all files','.*')])
        if isinstance(dirname, str):
            if os.name == 'nt':
                dirname = dirname.replace('/', '\\') #Windows compatibility

            if step == 1:
                self.dirPv.set(dirname)
            elif step == 2:
                self.dirExp.set(dirname)


    def validate(self, file0, file1, file2, file3, PN):
        args = (file0, file1, file2, file3) 
        textBdtVar = (self.resultTextBdt0Var, self.resultTextBdt1Var, self.resultTextBdt2Var, self.resultTextBdt3Var)
        i = 0; validation_counter = 0

        for file in args:
            if file and os.path.exists(file):
                self.file = file
                validation_counter += 1
                textBdtVar[i].set('')
            else:
                self.file = None
                textBdtVar[i].set('Please select a valid file')
            i += 1
        
        patern = r"(\d{7})(\.)(\d{3})"
        other_patern = r"\d{7}"
        wrong_patern = r"(\d{7})(\.)"
        if re.match(patern, PN) or (re.match(other_patern, PN) and not re.match(wrong_patern, PN)):
            validation_counter += 1
            self.text_variable_PN.set('')
        else:
            self.text_variable_PN.set('Unrecognised part number, format must be XXXXXXX or XXXXXXX.XXX')
        
        if validation_counter == len(args)+ 1 : # all files ok is lengh of args + first PN check (so +1) 
        	self.root.destroy()
        	fm.file_gen(file0, file1, file2, file3, PN)
           
if __name__ == "__main__":
    OptionsWindow().show()
    
        
        
