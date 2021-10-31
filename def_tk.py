# import tkinter as tk
# from tkinter import filedialog, messagebox, ttk
# from tkinter.constants import ACTIVE
# import pandas as pd
# import numpy as np
# import pyttsx3
# import os
# import shutil
# import time

# from datetime import date, datetime
# from openpyxl import load_workbook
# import time


# # a = tk.Tk()
# # a.geometry("900x600+15+15")
# # # selectmode='multiple'
# # box2 = tk.Listbox(a)
# # box2.insert(1, 'aaa')
# # box2.insert(1, 'bbb')
# # box2.insert(1, 'ccc')


# # def selected_item():
# #     for i in box2.curselection():
# #         v = box2.get(i)
# #         print(v)


# # box2.place(height=200, width=200, rely=0.43, relx=0.57)


# # btn = tk.Button(a, text='Print Selected', command=selected_item)
# # btn.pack(side='bottom')
# # a.mainloop()


# # def browse_button():
# #     # Allow user to select a directory and store it in global var
# #     # called folder_path
# #     today = date.today()

# #     # folder_result = "{}/resutl_{}".format(lbl1["text"], today)
# #     global folder_path
# #     filename = filedialog.askdirectory()
# #     lbl1["text"] = filename
# #     folder_result = "{}/resutl_{}".format(lbl1["text"], today)
# #     print(filename)
# #     print(folder_result)


# # root = tk.Tk()
# # root.geometry("300x300")
# # # folder_path = tk.StringVar()
# # lbl1 = tk.Label(master=root, text="")
# # lbl1.grid(row=0, column=1)
# # button2 = tk.Button(text="Browse", command=browse_button)
# # button2.grid(row=0, column=3)


# # root.mainloop()

# # now = datetime.now()

# # print(now)

# # dt_string = datetime.now().strftime("%d-%m-%Y_%Hh-%Mmin-%Ss")
# # print(dt_string)


from tkinter import *
import webbrowser


class MyApp:

    def __init__(self):
        self.window = Tk()
        self.window.title("My Application")
        self.window.geometry("720x480")
        self.window.minsize(480, 360)
        self.window.iconbitmap("TotalEnergies.ico")
        self.window.config(background='#41B77F')

        # initialization des composants
        self.frame = Frame(self.window, bg='#41B77F')

        # creation des composants
        self.create_widgets()

        # empaquetage
        self.frame.pack(expand=YES)

    def create_widgets(self):
        self.create_title()
        self.create_subtitle()
        self.create_youtube_button()

    def create_title(self):
        label_title = Label(self.frame, text="Bienvenue sur l'application", font=("Courrier", 40), bg='#41B77F',
                            fg='white')
        label_title.pack()

    def create_subtitle(self):
        label_subtitle = Label(self.frame, text="Hey salut à tous c'est Graven", font=("Courrier", 25), bg='#41B77F',
                               fg='white')
        label_subtitle.pack()

    def create_youtube_button(self):
        yt_button = Button(self.frame, text="Ouvrir Youtube", font=("Courrier", 25), bg='white', fg='#41B77F',
                           command=self.open_graven_channel)
        yt_button.pack(pady=25, fill=X)

    def open_graven_channel(self):
        webbrowser.open_new("http://youtube.com/gravenilvectuto")


# afficher
app = MyApp()
app.window.mainloop()
