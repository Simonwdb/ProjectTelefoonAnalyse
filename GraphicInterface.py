import FileManagement
import ProcessData
import tkinter as tk
from SETTINGS import SETTINGS
from tkinter.filedialog import askopenfilenames
from tkinter.messagebox import askquestion, showwarning

'''
TODO:       Er moet nog een entry veld worden gemaakt, waarin de maand aangegeven kan worden.
            Vervolgens moeten de gekozen files vanuit de list gevonden worden en toegewezen worden
            aan de correcte variabelen.
            Daarna moet het output_file dat gevonden wordt in de self.all_files, geupdate worden
            met de nieuwe maand. Daarna zou start_program() vanuit MainFile aangeroepen kunnen worden.
UPDATE:     07/06: Bovenstaande is verwerkt en werkt naar behoren. 
'''


class GUI:

    def __init__(self):
        self.root = tk.Tk()
        self.root.title('Telefoon Analyse')
        self.root.columnconfigure([0, 1, 2], minsize=20)
        self.root.rowconfigure([0, 1, 2, 3], minsize=20)
        self.all_files = None
        self.output_file = None
        self.field = None

        self.create_info_label()
        self.make_entry_field()
        self.make_load_button()
        self.make_start_button()

    def run(self):
        self.root.mainloop()

    def load_file(self):
        files = askopenfilenames(parent=self.root, title='Choose files')
        msgbox = askquestion('Add files', 'Add extra files?', icon='question')
        return list(files), msgbox

    def open_excel_file(self):
        files, msgbox = self.load_file()
        self.all_files = files
        while msgbox == 'yes':
            files_2, msgbox = self.load_file()
            for i in files_2:
                self.all_files.append(i)
        if len(self.all_files) != 4:
            showwarning('Waarschuwing', 'Er zijn geen 4 bestanden geselecteerd.'
                                        '\nAantal nu geselecteerd: {0}'.format(len(self.all_files)))

    def create_info_label(self):
        lbl = tk.Label(self.root, text='Fill in month:')
        lbl.grid(row=0, column=0, padx=5, pady=5)

    def make_entry_field(self):
        self.field = tk.Entry(self.root)
        self.field.grid(row=0, column=1, padx=5, pady=5)

    def make_load_button(self):
        load_button = tk.Button(master=self.root, text='Load Excel files', command=self.open_excel_file)
        load_button.grid(row=1, column=1, padx=5, pady=5)

    def make_start_button(self):
        start_button = tk.Button(master=self.root, text='Start program', command=self.start_program)
        start_button.grid(row=2, column=1, padx=5, pady=5)

    def make_finished_button(self, finished_correctly):
        if finished_correctly:
            color = 'green'
            text = 'Program is finished'
        else:
            color = 'red'
            text = 'Program finished not correctly'
        finished_button = tk.Button(master=self.root, text=text, bg=color, command=self.start_program)
        finished_button.grid(row=3, column=1, padx=5, pady=5)

    def start_program(self):
        month = self.field.get()

        cs_overall_file = FileManagement.get_file_name(SETTINGS().name_of_files[0], self.all_files)
        cs_employee_file = FileManagement.get_file_name(SETTINGS().name_of_files[1], self.all_files)
        answer_services_file = FileManagement.get_file_name(SETTINGS().name_of_files[2], self.all_files)
        output_file = FileManagement.get_file_name(SETTINGS().name_of_files[3], self.all_files)

        test = [cs_overall_file, cs_employee_file, answer_services_file, output_file, month]
        ProcessData.start_program(test)
        self.make_finished_button(True)

