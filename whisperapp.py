# pip install whisper
# install pytorch gpu version, default is cpu version
# need to download ffmpeg
# python-docx installed, --pip install python-docx


import customtkinter as ctk
from customtkinter import filedialog as fd
import os
import docx
import torch
import whisper

device = "mps"
torch_dtype = torch.float32
model_id = "distil-whisper/distil-large-v2"

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("dark-blue")

class TranscribeGUI(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.files = tuple()
        self.dir = ''
        self.check_var1 = ctk.IntVar(value=0)
        self.check_var2 = ctk.IntVar(value=0)
        self.message = 'Using OpenAI Whisper distil-large-v2.\n\n!!! Please check if you have the correct saving directory, and also if your files are in correct audio format.'

        # set window
        screen_width = self.winfo_screenwidth()  # Width of the screen
        screen_height = self.winfo_screenheight()  # Height of the screen
        width, height = 810, 1000
        x = (screen_width / 2) - (width / 2)   # x of the top-left starting point of the window
        y = (screen_height / 2) - (height / 2)   # y of the top-left starting point of the window
        self.geometry('%dx%d+%d+%d' % (width, height, x, y))
        self.resizable(False, False)
        self.title("Transcribe by Whisper")

        # result display info
        self.label1 = ctk.CTkLabel(self, text="Browse the direct result:", font=("Times", 30, "bold"), justify="center")
        self.label1.pack(pady=10, padx=10)

        # result display area
        self.textbox_result = ctk.CTkTextbox(self, width=900, height=600, font=("Helvetica", 20, "bold"))
        self.textbox_result.pack(pady=20, padx=10)
        self.textbox_result.insert("1.0", self.message)
        self.textbox_result.configure(state="disable")

        # buttons
        self.button1 = ctk.CTkButton(self, text="Select Files", text_color=("gray10", "#DCE4EE"),
                                     font=("Times", 20, "bold"), command=self.select_file)
        self.button1.pack(padx=30, pady=20)

        self.file_selected = ctk.CTkTextbox(self, width=900, height=30, font=("Helvetica", 20, "bold"))
        self.file_selected.pack(padx=30, pady=0) 
        self.file_selected.configure(state="disable")

        self.button2 = ctk.CTkButton(self, text="Select Saving Dir", text_color=("gray10", "#DCE4EE"),
                                     font=("Times", 20, "bold"), command=self.select_dir)
        self.button2.pack(padx=30, pady=20)

        self.dir_selected = ctk.CTkTextbox(self, width=900, height=30, font=("Helvetica", 20, "bold"))
        self.dir_selected.pack(padx=30, pady=0) 
        self.dir_selected.configure(state="disable")

        self.button3 = ctk.CTkButton(self, text="Transcribe and Save", text_color=("gray10", "#DCE4EE"),
                                     font=("Times", 30, "bold"), command=self.transcribe)
        self.button3.pack(padx=50, pady=20, side=ctk.LEFT)

        # checkboxes for saving formats

        self.ckbx1 = ctk.CTkCheckBox(self, text=".docx", text_color=("gray10", "#DCE4EE"),
                                     font=("Times", 20, "bold"), variable=self.check_var1, onvalue=1, offvalue=0, command=self.checkbox)
        self.ckbx1.pack(padx=0, pady=20, side=ctk.RIGHT, fill="x")

        self.ckbx2 = ctk.CTkCheckBox(self, text=".txt", text_color=("gray10", "#DCE4EE"),
                                     font=("Times", 20, "bold"), variable=self.check_var2, onvalue=1, offvalue=0, command=self.checkbox)
        self.ckbx2.pack(padx=0, pady=20, side=ctk.RIGHT, fill="x")
        
    @staticmethod
    def parse_name_from_path(path):
        path = path[::-1]
        for ch in path:
            if ch == ".":
                # remove the file extension, take the rest for further stripping
                path = path[path.index(ch)+1:]
                break
        for ch in path:
            if ch == "/":
                # separate the file name and the directory
                path = path[:path.index(ch)]
                break
        name = path[::-1]
        return name

    def checkbox(self):
        return self.check_var1.get(), self.check_var2.get()

    def select_file(self):
        self.files = fd.askopenfilenames()
        self.file_selected.configure(state="normal")
        self.file_selected.delete("1.0", "end")
        self.file_selected.insert("1.0", str(self.files))
        self.file_selected.configure(state="disable")
    
    def select_dir(self):
        self.dir = fd.askdirectory()
        self.dir_selected.configure(state="normal")
        self.dir_selected.delete("1.0", "end")
        self.dir_selected.insert("1.0", str(self.dir))
        self.dir_selected.configure(state="disable")

    def run_model(self, audio_path):
            model = whisper.load_model("tiny")
            transcript = model.transcribe(audio_path)

            return transcript['text']
    
    def transcribe(self):
        output = ''
        if not os.path.isdir(self.dir):
            self.textbox_result.configure(state="normal")
            self.textbox_result.delete("1.0", "end")
            self.textbox_result.insert("1.0", self.message + "\n\n!!! Saving directory does not exist.")
            self.textbox_result.configure(state="disable")

        else:
            output = self.message + "\n" + "="*20
            for path in self.files:
                name = self.parse_name_from_path(path)
                transcription = self.run_model(path)

                v1, v2 = self.checkbox()

                if v1 == 1: # doc
                    doc = docx.Document()
                    doc.add_paragraph(transcription)
                    doc.save(self.dir + f"/Transcript_{name}.docx")

                if v2 == 1: # txt
                    with open(self.dir + f"/Transcript_{name}.txt", "w", encoding="utf-8") as f:
                        f.write(transcription)

                output = output + "\n\n" + transcription
            
            self.textbox_result.configure(state="normal")
            self.textbox_result.delete("1.0", "end")
            self.textbox_result.insert("1.0", output)
            self.textbox_result.configure(state="disable")

            
        
app = TranscribeGUI()
app.mainloop()



