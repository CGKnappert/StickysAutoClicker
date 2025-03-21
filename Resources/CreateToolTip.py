import tkinter as tk
from tkinter import ttk
from tkinter import Toplevel

class CreateToolTip(object):
    def __init__(self, widget, root, text='widget info'):
        self.waittime = 500     #miliseconds
        self.wraplength = 180   #pixels
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None
        self.root = root

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.waittime, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        x = y = 0
        # x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + self.widget.winfo_width()
        y += self.widget.winfo_rooty() + self.widget.winfo_height()
        # creates a toplevel window
        self.tw = Toplevel(self.root)
        self.tw.wm_attributes("-topmost", 1)
        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = ttk.Label(self.tw, text=self.text, justify='left',
                       background="#000000", relief='solid', borderwidth=1,
                       wraplength = self.wraplength)
        label.pack(ipadx=8, ipady=8)

    def hidetip(self):
        tw = self.tw
        self.tw= None
        if tw:
            tw.destroy()