# -*- coding: utf-8 -*-
"""
Created on Tue Feb 23 15:02:31 2021

@author: Markus Sandberg
"""
import kivy
from kivy.app import App
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from tkinter.filedialog import askopenfilename
import openpyxl
from pandas import DataFrame
import tkinter as tk

tiedot = []

class MainApp(App):

    def build(self):
        layout = GridLayout(cols=2)
        self.a = TextInput(text='Sarakkeita', multiline=False)
        self.b = TextInput(text='Rivejä', multiline=False)
        self.c = TextInput(text='Tiedostoja', multiline=False)
        submit = Button(text='Submit', on_press=self.submit)
        self.d = TextInput(text='Testi', multiline=True)
        layout.add_widget(self.a)
        layout.add_widget(self.b)
        layout.add_widget(self.c)
        layout.add_widget(submit)
        layout.add_widget(self.d)
        return layout

    def submit(self, obj):
        N = int(float(str(self.a.text)))
        X = int(float(str(self.b.text)))
        v1 = int(float(str(self.c.text)))
        v = range(v1)

        for i in v:
            root = tk.Tk()  # Poistetaan tkinter aloitusikkuna
            root.withdraw()
            filename = askopenfilename()  # Pyydetään käyttäjää valitsemaan vertailtava tiedosto
            wb1 = openpyxl.load_workbook(filename)  # Tehdään tiedostosta openpyxl workbook
            Sheet1 = wb1.active  # Muutetaan workbook worksheetiksi, ja poistetaan X ylintä riviä
            S = Sheet1.max_column  # Etsitään tiedoston viimeinen sarake
            L = S - N
            T = N-1
            Sheet1.delete_rows(1, X)
            Sheet1.delete_cols(1, T)  # Poistetaan turhat sarakkeet kahdessa osassa, tutkittava sarake on sarake numero N
            Sheet1.delete_cols(2, L)
            df1 = DataFrame(Sheet1.values)  # Muutetaan käsitelty worksheet pandas dataframeksi
            ar1 = df1.to_numpy()  # Pandas dataframe muutetaan numpyn arrayksi
            ar1 = ar1.flatten()  # Arraysta "litistetään" yksiulotteinen vertailua varten
            l1 = ar1.tolist()  # Array muutetaan listaksi vertailua varten
            tiedot.extend(l1)

        tuplat = set([x for x in tiedot if tiedot.count(x) > 1])  # Etsitään "tiedot listasta kaksoiskappaleet"
        bb = list(tuplat)
        g = str(bb)
        self.d.text = g


MainApp().run()
