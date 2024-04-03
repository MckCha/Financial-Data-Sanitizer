from tkinter.font import Font
import backend
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES

class createApplication():
    def __init__(self,root):
        self.root = root
        self.fileCounter = 1
        self.preset()
        self.buildSubmit()
        self.createLabel("Green = Download Successful (Located in Downloads)", "white")
        
    def preset(self):
        root.configure(bg="#141414", bd=10, relief=tk.SUNKEN)
        self.root.geometry("550x300")
        self.root.title("PayBank Excel Cleaner")
        bold_font = Font(family="Comic Sans MS", size=18, weight="bold")
        self.label = tk.Label(self.root, text="Enter desired file name",font=bold_font, fg="white", bg="#141414")
        self.label.pack(padx=30, pady=10)
        root.drop_target_register(DND_FILES)
        root.dnd_bind('<<Drop>>', self.dragDrop)
    
    def buildSubmit(self):
        self.submitFrame = tk.Frame(root, bg="#141414")
        self.entry = tk.Entry(self.submitFrame, 
                              width=15,
                              justify="center",
                              font=("Comic Sans MS",18),
                              relief=tk.RAISED,
                              bd=4,
                              bg="white")
        self.submitButton = tk.Button(self.submitFrame,
                                      text="Submit",
                                      command=self.getUserInput,
                                      font=("Comic Sans MS",11),
                                      bg="#555555", fg="white",
                                      padx= 20,
                                      relief=tk.RAISED,
                                      bd=5)
        self.entry.grid(row=0,column=0)
        self.submitButton.grid(row=0,column=1, padx=10)
        self.submitFrame.pack()

    def dragDrop(self, event):
        # add try except in here (Going to change color when there is no object)
        try:
            self.fileLabel.configure(bg="green")
            file = event.data
            cleaned_path = file.replace('{', '').replace('}', '')
            self.file_path = cleaned_path 
            self.label.config(text=f"Recieved {self.fileCounter} file(s). ", foreground="white")
            self.fileCounter += 1
            # Extract Excel Data and Save New File
            userSheet = backend.Excel(accountNumber,self.file_path, self.userInput)
            userSheet.applySettings()
            userSheet.extractData()
            userSheet.saveExcel()
            self.userInput = None
            self.fileLabel = None
        except Exception:
            self.label.config(text="Enter file name / Invalid File Format", foreground="red")

    def getUserInput(self):
        if self.entry.get():
            self.userInput = self.entry.get() + ".xlsx"
            self.buildResult()
            self.entry.delete(0, tk.END)
            print(self.userInput)
        else:
            self.label.config(text="Enter file name / Check File Format", foreground="red")

    def buildResult(self):
        self.resultFrame = tk.Frame(root)
        self.resultFrame.pack(side=tk.LEFT,padx=10, pady=10, anchor=tk.N)
        self.fileLabel = tk.Label(self.resultFrame, text=self.userInput, bg="#141414", fg="white", wraplength=480, anchor="w", justify="left")
        self.fileLabel.grid(row=0,column=0)
    
    def createLabel(self, createdLabel, color):
        label = tk.Label(self.root, text=createdLabel, bg="#141414", fg=color,pady=15, font=("Comic Sans MS",11))
        label.pack()

# Create Window
root = TkinterDnD.Tk()
window = createApplication(root)

accountNumber = "123456789123456"

root.mainloop()