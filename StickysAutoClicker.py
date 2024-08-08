import os
import ctypes
from ctypes import windll, c_char_p, c_int, byref, sizeof
errorCode = ctypes.windll.shcore.SetProcessDpiAwareness(0)
import time
import csv
import mouse
import mss
import pyautogui
import threading
import win32process
import win32gui
import psutil
import screeninfo
import datetime
from pathlib import Path
from pynput import keyboard
from pynput.keyboard import Key, Listener
from os.path import exists

import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import simpledialog
from tkinter.messagebox import askyesno
# import TKinterModernThemes as TKMT
# import sv_ttk
import tksvg

# from PIL import ImageGrab
import configparser as cp


pyautogui.FAILSAFE = False
pyautogui.PAUSE = False

# TODO: Export crash reports
# TODO: Make prettier

"""
Sticky's Autoclicker Documentation
This is an autoclicker with fancier functionality than I have been able to find online and have always desired.

The backbone of this GUI is tkinter using a grid.
The imports used are for:
- mouse import for finding mouse position outside of the tkinter window
- mss for understanding the arrangement of multiple monitors and their sizes
- PIL etc. for allowing multiple monitors when using locteonscreen 
- pyautogui for moving mouse and clicking
- threading for starting the clicking loop without the window becoming unresponsive
- csv for import and export functionality
- time for sleeping for delays
- os for finding filepath
- tkinter.simpledialog for asking for new macro name
- tkinter.filedialog for importing macroscd 
- pynput for listening to keyboard for emergency exit combo Ctrl + Shift + 1

build with "python -m PyInstaller --onefile --noconsole --icon=StickysAutoClicker\StickyHeadIconAll.ico  --collect-all="tksvg" StickysAutoClicker\StickysAutoClicker.py"
"""

USABILITY_NOTES = ("\n"
                  " - Selecting any row of a macro will auto-populate the X and Y positions, delay, action and comment fields with the values of that row, overwriting anything previously entered.\n"
                  " - Shift + LCtrl + ` will set loops to 0 and stop the autoclicker immediately.\n"
                  " - Shift + LCtrl + Tab will pause the autoclicker and pressing that key combination again will start the autoclicker where it left off.\n"
                  "           Remember you might need to focus back on the application to be clicked or key pressed before starting again.\n"
                  " - Overwrite will set all selected rows to the current X, Y, delay, action and comment values.\n"
                  " - Macros are exported as csv files that are written to folder '\StickyAutoClicker\Macros' in your Documents folder.\n"
                  " - Exported macro files are kept up to date with each edit created.\n"
                  " - Import expects a comma separated file with no headers and only five columns: X, Y, Action, Delay and Comment.\n"
                  " - Action will recognize the main three mouse clicks as well as any keyboard action including Shift + keys (doesn't work with Alt or Ctrl + keys).\n"
                  " - Key presses do not move cursor, so X and Y positions do not matter.\n"
                  " - The Start From Selected Row option in Settings will cause the macro to start from the first highlighted row in the macro or from the start if nothing is highlighted.\n"
                  " - Action has two escape characters, ! and #, that will allow the user to continue typing rather than overwriting the action key.\n"
                  "           This is to allow for calling another macro (!) or finding an image (#)\n"
                  " - Typing !macroName into action and adding that row will make the row look for another macro with a name matching what follows the ! and execute that macro.\n"
                  "           Delay serves another purpose when used with a macro action and will repeat that macro for the amount of times in the Delay column.\n"
                  "           Macros do not need to be in a tab to be called by another macro. The csv in the \StickyAutoClicker\Macros folder will be used if it exists.\n"
                  " - Typing #imageName into action and adding that row will make a macro look for a .png image with that name in the \StickyAutoClicker\Images folder and move the cursor to a found image and left click.\n"
                  "           Delay will serve as the confidence percentage level when used with a find image action.\n"
                  "           Confidence of 100 will find an exact match and confidence of 1 will find the first roughly similar image.\n"
                  "           Finding an image will try 5 times with a .1 second delay if not found and click the image once found.\n"
                  "           If image is not found then loop will end and next loop will start. This is to prevent the rest of the loop from going awry because the expected image was not found.\n"
                  " - Action also allows underscore _ as a special character that will indicate the following key(s) should be pressed and held for the set amount of time in the Delay field.\n"
                  "           Note that for these rows the Delay no longer delays after the key is held.\n"
                  "           You can hold multiple keys at a time by continuing to type into the Action field once a _ has been entered. A | will delineate the different keys to be held.\n"
                  "           Typing _ into action can also reset the field to remove keys you do not want pressed since backspace doesn't work.\n"
                  " - The Record functionality will begin entering rows to the end of the current macro tab reading key presses and delays.\n"
                  "           This functionality is quite accurate (for python) and can be quite efficien especially when paired with manual edits to shorten or remove unnessecary delays or actions.\n"
                  "           For playback of recordings it is highly recommended to use the Busy Wait option in the Settings menu.\n"
                  "           Busy Wait allows the delay to be accurate to around one millisecond where as non-Busy Wait is accurate to around 10-15 ms.\n"
                  "           The downside of Busy Wait and why it shouldn't always be used is that it incurs heavy CPU usage and can be felt by users and other programs.")

FILE_PATH = os.path.join(os.path.expanduser(r'~\Documents'), r'StickysAutoClicker')

global_monitor_left = 0
global_monitor_top = 0
global_recording = False

PAUSE_COMBO = ['Key.ctrl_l', 'Key.shift', 'Key.tab']
EXIT_COMBO = [['Key.ctrl_l', 'Key.shift', '\'`\''], ['Key.ctrl_l', 'Key.shift', '\'~\''], ['Key.ctrl_l', 'Key.shift', '<192>']]
ALL_COMBO = ['Key.ctrl_l', 'Key.shift', 'Key.tab', '\'`\'', '\'~\'', '<192>']

# Standard main window size on open
NORMAL_SIZE = (750, 481)

# Notebook row colors
EVEN_COLOR = '#080808'
ODD_COLOR = '#676767'
SELECTED_COLOR = '#3fb4ea'
RUNNING_COLOR = '#ff5d12'
SELECTED_AND_RUNNING_COLOR_BG = 'white'
SELECTED_AND_RUNNING_COLOR_FG = 'black'


# define global monitors
with mss.mss() as sct:
    if sct.monitors:
        global_monitor_left = sct.monitors[0].get('left')
        global_monitor_top = sct.monitors[0].get('top')

    
class Titlebar():
    # Class for addingg titlebars after overrideredirect removes them
    # Needed for greater control of buttons and theme
    def __init__(self, root, pack, parent, icon, title_text, minimize, maximize, close, help):
        self.root = root
        self.parent = parent
        root.minimized = False # only to know if root is minimized
        root.maximized = False # only to know if root is maximized

        # Create a parent for the titlebar
        self.title_bar = ttk.Frame(root, height=10)
        self.helpWindow = None

        
        # Pack the title bar window
        if pack:
            self.title_bar.pack(fill=tk.X)
        else:
            self.title_bar.columnconfigure(0, weight=10)
            self.title_bar.columnconfigure(1, weight=10)
            self.title_bar.columnconfigure(2, weight=10)
            self.title_bar.columnconfigure(3, weight=10)
            self.title_bar.columnconfigure(4, weight=10)
            self.title_bar.columnconfigure(5, weight=10)
            self.title_bar.grid(row=0, column=0, columnspan=6, sticky='ew')

        # Create the title bar buttons
        buttonPos = 5
        if close:
            self.close_button = ttk.Button(self.title_bar, text='  ×  ', command=parent.onClose, takefocus=False)
            if pack:
                self.close_button.pack(side=tk.RIGHT, padx=(0, 10), pady=(5, 0))
            else:
                self.title_bar.columnconfigure(buttonPos, minsize=45)
                self.close_button.grid(row=0, column=buttonPos, padx=(0, 10), pady=(5, 0), sticky='e')
                buttonPos -= 1
        else:
            self.title_bar.grid(row=0, column=0, columnspan=6)
        if maximize:
            self.expand_button = ttk.Button(self.title_bar, text=' 🗖 ', command=self.maximize_window, takefocus=False)
            if pack:
                self.expand_button.pack(side=tk.RIGHT, padx=(0, 5), pady=(5, 0))
            else:
                self.title_bar.columnconfigure(buttonPos, minsize=45)
                self.expand_button.grid(row=0, column=buttonPos, padx=(0, 5), pady=(5, 0), sticky='e')
                buttonPos -= 1
        if minimize:
            self.minimize_button = ttk.Button(self.title_bar, text=' 🗕 ',command=self.minimize_window, takefocus=False)
            if pack:
                self.minimize_button.pack(side=tk.RIGHT, padx=(0, 5), pady=(5, 0))
            else:
                self.title_bar.columnconfigure(buttonPos, minsize=45)
                self.minimize_button.grid(row=0, column=buttonPos, padx=(0, 5), pady=(5, 0), sticky='e')
                buttonPos -= 1
        if help:
            self.helpButton = ttk.Button(self.title_bar, text="Help", command=self.showHelp, takefocus=False)
            if pack:
                self.helpButton.pack(side=tk.RIGHT, padx=(0, 5), pady=(5, 0))
            else:
                self.title_bar.columnconfigure(buttonPos, minsize=45)
                self.helpButton.grid(row=0, column=buttonPos, padx=(0, 5), pady=(5, 0), sticky='e')
                buttonPos -= 1

        while buttonPos> 0:
            self.title_bar.columnconfigure(buttonPos, minsize=50)
            buttonPos -= 1
        
        if icon != None:
            # Create the title bar icon
            self.title_bar_icon = ttk.Label(self.title_bar, image=icon)
            if pack:
                self.title_bar_icon.pack(side=tk.LEFT, padx=(10,0), pady=(5, 0), fill=tk.Y)
            else:
                self.title_bar_icon.grid(row=0, column=0, padx=(10,0), pady=(5, 0), sticky='w')

        # Create the title bar title
        self.title_bar_title = ttk.Label(self.title_bar, text=title_text)
        self.title_bar_title.configure(font=("Arial bold", 14))
        if pack:
            self.title_bar_title.pack(side=tk.LEFT, padx=(10,0))
        else:
            self.title_bar_title.grid(row=0, column=1, padx=(10,0), sticky='w')

        # Bind events for moving the title bar
        self.title_bar.bind('<Button-1>', self.move_window_bindings)
        self.title_bar_title.bind('<Button-1>', self.move_window_bindings)
        if icon != None:
            self.title_bar_icon.bind('<Button-1>', self.move_window_bindings)
        self.move_window_bindings(status=True)

        # Set up the window for minimizing functionality
        self.root.bind("<FocusIn>", self.deminimizeEvent)
        self.root.after(10, lambda: self.set_appwindow(self.root))

    # Window manager functions
    def minimize_window(self):
        self.root.attributes("-alpha",0)
        self.root.minimized = True

    def deminimizeEvent(self, event):
        self.deminimize()

    def deminimize(self):
        # root.focus() 
        self.root.attributes("-alpha",1)
        if self.root.minimized == True:
            self.root.minimized = False                              

    def maximize_window(self):
        if self.root.maximized == False:
            self.root.normal_size = self.root.geometry()
            self.expand_button.config(text=" 🗗 ")

            monitors = screeninfo.get_monitors()
            max = False
            for m in reversed(monitors):
                if m.x <= self.root.winfo_x() <= m.width + m.x - 1 and m.y <= self.root.winfo_y() <= m.height + m.y - 1:
                    # print(f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}+{m.x}+{m.y}")
                    self.root.geometry(f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}+{m.x}+{m.y}")
                    max = True
            if not max:
                self.root.geometry(f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}+{monitors[0].x}+{monitors[0].y}")
        else:
            self.expand_button.config(text=" 🗖 ")
            self.root.geometry(self.root.normal_size)

        self.root.maximized = not self.root.maximized

    def set_appwindow(self, mainWindow):
        global hwnd
        hwnd = windll.user32.GetParent(mainWindow.winfo_id())
        mainWindow.wm_withdraw()
        mainWindow.after(10, lambda: mainWindow.wm_deiconify())

    def get_pos(self, event):
        global xwin, ywin
        xwin = event.x_root
        ywin = event.y_root

    def move_window(self, event):
        global xwin, ywin
        self.root.geometry(f'+{event.x_root + self.root.winfo_x() - xwin}+{event.y_root + self.root.winfo_y() - ywin}')
        xwin = event.x_root
        ywin = event.y_root

    def move_window_bindings(self, *args, status=True):
        if self.root.maximized:
            self.maximize_window()
        if status == True:
            self.title_bar.bind("<B1-Motion>", self.move_window)
            self.title_bar.bind("<Button-1>", self.get_pos)
            self.title_bar_title.bind("<B1-Motion>", self.move_window)
            self.title_bar_title.bind("<Button-1>", self.get_pos)
            try:
                self.title_bar_icon.bind("<B1-Motion>", self.move_window)
                self.title_bar_icon.bind("<Button-1>", self.get_pos)
            except:
                pass
            
    def showHelp(self):
        if self.parent.helpWindow is None:
            self.helpWindow = helpWindow(self.root, self.parent)
            self.parent.helpWindow = self.helpWindow
            try:
                if self.parent.stayOnTop.get() == 0:
                    self.helpWindow.helpWindow.attributes("-topmost", False)
                else:
                    self.helpWindow.helpWindow.attributes("-topmost", True)
            except:
                pass
        else:
            self.helpWindow.onClose()
            self.helpWindow = None



class helpWindow(ttk.Frame):
    # window to display help text that informs user of available functionality
    def __init__(self, root, parent):
        super().__init__(root)
        self.root = root
        self.parent = parent
        
        try:
            if self.state() == "normal":
                self.onClose()
            else:
                self.helpWindow.wm_state("normal")
        except:
            self.helpWindow = Toplevel(self.root)
            self.helpWindow.overrideredirect(True)
            config = cp.ConfigParser()
            config.read(os.path.join(FILE_PATH, r'config.ini'))
            if config.has_option("Position", "helpx") and config.has_option("Position", "helpy"):
                self.helpWindow.geometry("+" + config.get("Position", "helpx") + "+" + config.get("Position", "helpy"))
            self.titlebar = Titlebar(self.helpWindow, True, self, STICKY_ICON, "Sticky's Autoclicker Help", True, False, True, False)
            self.helpWindow.resizable(height=None, width=None)
            self.helpWindow.wm_title("Sticky's Autoclicker Help")
            self.helpWindow.protocol("WM_DELETE_WINDOW", self.onClose)
            self.helpWindow.focus_force()

            self.helpLabel = ttk.Label(self.helpWindow, text=USABILITY_NOTES, justify=LEFT)
            self.helpLabel.pack(side=tk.BOTTOM, padx=15, pady=15)

            self.titleLabel = ttk.Label(self.helpWindow, text="Usability Notes", font=("Arial bold", 14), justify=CENTER)
            self.titleLabel.pack(side=tk.BOTTOM)

    def onClose(self):
        config = cp.ConfigParser()
        config.read(os.path.join(FILE_PATH, 'config.ini'))
        config.set('Position', 'helpx', str(self.helpWindow.winfo_rootx()))
        config.set('Position', 'helpy', str(self.helpWindow.winfo_rooty()))
        with open(os.path.join(FILE_PATH, r'config.ini'), 'w') as configfile:
            config.write(configfile)
        self.helpWindow.destroy()
        self.helpWindow = None
        self.parent.helpWindow = None
        self.root.focus()


class settingsWindow(ttk.Frame):
    # window to display options and handle changing of options
    def __init__(self, root, parent):
        super().__init__(root)
        self.root = root
        self.parent = parent
        
        try:
            if self.state() == "normal":
                self.onClose()
            else:
                self.settingsWindow.wm_state("normal")
        except:
            self.settingsWindow = Toplevel(self)
            self.settingsWindow.overrideredirect(True)
            config = cp.ConfigParser()
            config.read(os.path.join(FILE_PATH, r'config.ini'))
            if config.has_option("Position", "settingsx") and config.has_option("Position", "settingsy"):
                self.settingsWindow.geometry("+" + config.get("Position", "settingsx") + "+" + config.get("Position", "settingsy"))
            self.titlebar = Titlebar(self.settingsWindow, False, self, STICKY_ICON, "Settings", True, False, True, False)
            self.settingsWindow.resizable(height=None, width=None)
            self.settingsWindow.wm_title("Sticky's Autoclicker Settings")
            self.settingsWindow.protocol("WM_DELETE_WINDOW", self.onClose)
            self.settingsWindow.iconphoto(False, STICKY_ICON)
            self.settingsWindow.focus_force()

            self.settingsFrame = ttk.LabelFrame(self.settingsWindow)
            self.settingsFrame.grid(row=1, column=0, rowspan=2, columnspan=6, padx=(10, 10), pady=(0, 10), sticky="nsew")

            # Rows 0 and 1\
            self.busyLabel = ttk.Label(self.settingsFrame, text="Use Busy Wait")
            self.busyLabel.grid(row=1, column=0, columnspan=2, padx=10, sticky="nsew")
            self.busyButton = ttk.Checkbutton(self.settingsFrame, variable=parent.busyWait, onvalue=1, offvalue=0, command=parent.toggleBusy)
            self.busyButton.grid(row=2, column=0, padx=10, pady=10)

            self.windowLabel = ttk.Label(self.settingsFrame, text="Application Selector")
            self.windowLabel.grid(row=1, column=2, sticky='s', padx=10)
            self.windowButton = ttk.Button(self.settingsFrame, text="Find App", command=parent.windowFinder, takefocus=False)
            self.windowButton.grid(row=2, column=2, sticky='n', padx=10, pady=10)

            self.hiddenLabel = ttk.Label(self.settingsFrame, text="Use Hidden Mode")
            self.hiddenLabel.grid(row=1, column=4, sticky='s', padx=10)
            self.hiddenButton = ttk.Checkbutton(self.settingsFrame, variable=parent.hiddenMode, onvalue=1, offvalue=0, command=parent.toggleHidden)
            self.hiddenButton.grid(row=2, column=4, sticky='n', padx=10, pady=10)

            # Rows 3 and 4
            self.startFromSelectedLabel = ttk.Label(self.settingsFrame, text="Start From Selected Row")
            self.startFromSelectedLabel.grid(row=3, column=0, padx=10)
            self.startFromSelectedButton = ttk.Checkbutton(self.settingsFrame, variable=parent.startFromSelected, onvalue=1, offvalue=0, command=parent.toggleStartFromSelected)
            self.startFromSelectedButton.grid(row=4, column=0, padx=10, pady=10)

            self.stayOnTopLabel = ttk.Label(self.settingsFrame, text="Stay On Top")
            self.stayOnTopLabel.grid(row=3, column=2, padx=10)
            self.stayOnTopButton = ttk.Checkbutton(self.settingsFrame, variable=parent.stayOnTop, onvalue=1, offvalue=0, command=parent.toggleStayOnTop)
            self.stayOnTopButton.grid(row=4, column=2, padx=10, pady=10)

            self.loopsByMacroLabel = ttk.Label(self.settingsFrame, text="Loops By Macro")
            self.loopsByMacroLabel.grid(row=3, column=4, padx=10)
            self.loopsByMacroButton = ttk.Checkbutton(self.settingsFrame, variable=parent.loopsByMacro, onvalue=1, offvalue=0, command=parent.toggleLoopsByMacro)
            self.loopsByMacroButton.grid(row=4, column=4, padx=10, pady=10)

    def onClose(self):
        # print("closesettings")
        config = cp.ConfigParser()
        config.read(os.path.join(FILE_PATH, 'config.ini'))
        config.set('Position', 'settingsx', str(self.settingsWindow.winfo_rootx()))
        config.set('Position', 'settingsy', str(self.settingsWindow.winfo_rooty()))
        with open(os.path.join(FILE_PATH, r'config.ini'), 'w') as configfile:
            config.write(configfile)
        self.settingsWindow.destroy()
        self.parent.settingsWindow = None
        self.settingsWindow = None
        self.root.focus()


class logWindow(ttk.Frame):
    # window for viewing click history and errors
    def __init__(self, root, parent, text):
        super().__init__(root)
        self.root = root
        self.parent = parent
        
        try:
            if self.state() == "normal":
                self.onClose()
            else:
                self.logWindow.wm_state("normal")
        except:
            self.logWindow = Toplevel(self)
            config = cp.ConfigParser()
            config.read(os.path.join(FILE_PATH, r'config.ini'))
            if config.has_option("Position", "logx") and config.has_option("Position", "logy"):
                self.logWindow.geometry("480x600" + "+" + config.get("Position", "logx") + "+" + config.get("Position", "logy"))
            self.logWindow.overrideredirect(True)
            self.titlebar = Titlebar(self.logWindow, True, self, STICKY_ICON, "Sticky's Autoclicker Log", True, True, True, False)
            self.logWindow.resizable(height=None, width=None)
            self.logWindow.wm_title("Sticky's Autoclicker Log")
            self.logWindow.protocol("WM_DELETE_WINDOW", self.onClose)
            self.logWindow.focus_force()

            self.frame = ttk.Frame(self.logWindow)
            self.frame.pack(side="left", ipadx=10, fill="both", expand=True) 
            self.frame.config(width=470, height=480)
            self.text = tk.Text(self.frame)

            # self.horizontalTabScroll = ttk.Scrollbar(self.frame, orient="horizontal", command=self.text.xview)
            # self.horizontalTabScroll.pack(side='bottom', fill=X, anchor="s")
            self.verticalTabScroll = ttk.Scrollbar(self.frame, orient='vertical', command=self.text.yview)
            self.verticalTabScroll.pack(side=RIGHT, fill=Y, anchor="e")

            self.text.configure(yscrollcommand=self.verticalTabScroll.set)
            self.text.pack(side="left", padx=(5, 0), pady=(5, 0), fill="both", expand=True)  
            self.text.insert('1.0', text)

            # sizegrip could not lit above either scrollbar despite the config being essentially the same as on the main windows notebook
            # self.grip = ttk.Sizegrip()
            # self.grip.place(relx=1.0, rely=1.0, anchor="se")
            # self.grip.lift(self.text)
            # self.grip.bind("<B1-Motion>", self.moveMouseButton)

    def updateText(self, text):
        self.text.delete("1.0", END)
        self.text.insert('1.0', text)


    def moveMouseButton(self, e):
        x1=self.root.winfo_pointerx()
        y1=self.root.winfo_pointery()
        x0=self.root.winfo_rootx()
        y0=self.root.winfo_rooty()

    def onClose(self):
        # print("closelog")
        config = cp.ConfigParser()
        config.read(os.path.join(FILE_PATH, 'config.ini'))
        config.set('Position', 'logx', str(self.logWindow.winfo_rootx()))
        config.set('Position', 'logy', str(self.logWindow.winfo_rooty()))
        with open(os.path.join(FILE_PATH, r'config.ini'), 'w') as configfile:
            config.write(configfile)
        self.logWindow.destroy()
        self.parent.logWindow = None
        self.logWindow = None
        self.root.focus()



class treeviewNotebook(ttk.Frame):
    # class with notebook with tabs of treeviews
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        # init buttons and labels
        self.initElements()
        # init notebook and treeview
        self.initTab()
        # load settings from config file or creaete config file and necessary folders
        self.loadSettings()

        self.parent.title('Fancy Autoclicker')
        self.parent.geometry(str(NORMAL_SIZE[0]) + "x" + str(NORMAL_SIZE[1]))

        self.parent.protocol("WM_DELETE_WINDOW", self.onClose)

        self.parent.update_idletasks()

    def initElements(self):
        self.notebook = None
        self.notebookFrame = None
        self.recorder = None
        self.window = None
        self.titlebar = Titlebar(self.parent, True, self, STICKY_ICON, "Sticky's Autoclicker", True, False, True, True)

        # Array for currently pressed keys so they can be unpressed when stopping clicking
        self.currPressed = []
        self.currTab = None
        self.runningRows = []
        self.macroLoops = {}
        self.busyWait = tk.IntVar()
        self.hiddenMode = tk.IntVar()
        self.startFromSelected = tk.IntVar()
        self.stayOnTop = tk.IntVar()
        self.loopsByMacro = tk.IntVar()
        self.selectedApp = ''
        self.keysListener = False
        self.pauseEvent = None
        self.monitorKeysPressed = set()
        self.clickLog = []

        # window references for knowing if window is already created or not
        self.settingsWindow = None
        self.helpWindow = None
        self.logWindow = None

        # store clicking thread for reference, needed to pass event.set to loop stops while waiting.
        # setting event lets the sleep of the clicking thread get interrupted and stop the thread
        # Otherwise starting clicking while thread is sleeping makes two threads clicking and this reference would be overwritten
        self.activeThread = None
        # allow treeview array to be referenced for macro looping and export
        self.treeTabs = {}
        self.treeView = None
        self.priorDraggedTab = -1

        self.vcmd = (self.parent.register(self.checkNumerical), '%S')

        self.actionVar = StringVar()
        self.actionVar.trace_add('write', self.cleanseActionEntry)
        self.delayVar = StringVar()
        self.delayVar.set("1")
        # self.delayVar.trace_add("write", self.globalDelayCallback)
        self.globalLoops = IntVar()
        self.globalLoops.set(1)
        self.loopsEntry = StringVar()
        # self.loopsEntry.trace_add("write", self.globalLoopsCallback)
        self.loopsEntry.set("1")
        self.loopsLeft = IntVar()
        self.loopsLeft.trace_add("write", self.updateLoops)
        self.loops = IntVar()
        self.loops.set(0)
        self.x = IntVar()
        self.x.set(0)
        self.y = IntVar()
        self.y.set(0)
        
        # Row weights
        self.rowconfigure(3, weight=20)
        self.rowconfigure(4, weight=20)
        self.rowconfigure(5, weight=20)
        self.rowconfigure(6, weight=20)
        self.rowconfigure(7, weight=10)
        self.grid_rowconfigure(0, minsize=40)
        self.grid_rowconfigure(1, minsize=75)
        self.grid_rowconfigure(2, minsize=40)
        self.grid_rowconfigure(3, minsize=80)
        self.grid_rowconfigure(4, minsize=80)
        self.grid_rowconfigure(5, minsize=60)
        self.grid_rowconfigure(5, minsize=40)

        self.loopsFrame = ttk.LabelFrame(self, height=100)
        self.loopsFrame.grid(row=0, column=0, rowspan=3, columnspan=2, padx=(10, 10), sticky="ew")

        self.actionFrame = ttk.LabelFrame(self)
        self.actionFrame.grid(row=0, column=2, rowspan=3, columnspan=5, padx=(0, 10), sticky="ew")

        # Column 0
        # for reference: notebook.grid(row=3, column=2, columnspan=3, rowspan=5, sticky='')
        self.loopsFrame.columnconfigure(0, weight=10)
        self.grid_columnconfigure(0, minsize=100)
        self.clickLabel = ttk.Label(self.loopsFrame, text="Click Loops", anchor='center')
        self.clickLabel.grid(row=0, column=0, columnspan=2, padx=10, pady=(2, 0), sticky='nsew')
        self.clickLabel.configure(font=("Arial", 11))
        self.loopEntry = ttk.Spinbox(self.loopsFrame, width=0, justify="center", from_=1, to=9999, increment=1, textvariable=self.loopsEntry, validate='key', validatecommand=self.vcmd)
        self.loopEntry.grid(row=1, column=0, padx=10, pady=(5,0), columnspan=2, sticky='nsew')
        self.loopEntry.bind("<KeyRelease>", self.loopEntryKey)

        s = ttk.Style()
        s.configure('S.TButton', font=("Arial", 15))

        self.LoopsLeftLabel = ttk.Label(self.loopsFrame, text="Loops Left")
        self.LoopsLeftLabel.grid(row=2, column=0, sticky='n', padx=(10,0), pady=(20,0))
        self.LoopsLeftLabel.configure(font=("Arial", 11))
        self.clicksLeftLabel = ttk.Label(self.loopsFrame, textvariable=self.loopsLeft)
        self.clicksLeftLabel.grid(row=2, column=0, sticky='s', padx=(15,0), pady=(45,10))
        self.clicksLeftLabel.configure(font=("Arial", 11))
        self.startButton = ttk.Button(self, text="Start Clicking", command=self.threadStartClicking, takefocus=False, style='S.TButton')
        self.startButton.grid(row=3, column=0, columnspan=2, padx=10, pady=(10, 10), sticky="nsew")
        self.stopButton = ttk.Button(self, text="Stop Clicking", command=self.stopClicking, takefocus=False, style='S.TButton')
        self.stopButton.grid(row=4, column=0, columnspan=2, padx=10, pady=(5, 10), sticky="nsew")

        self.importMacroButton = ttk.Button(self, text="Import", command=self.importMacro, takefocus=False)
        self.importMacroButton.grid(row=5, column=0, sticky='nsew', padx=10, pady=(0, 10))
        self.logButton = ttk.Button(self, text="Log", command=self.openLogWindow, takefocus=False)
        self.logButton.grid(row=5, column=1, sticky='nsew', padx=10, pady=(0, 10))
        self.settingsButton = ttk.Button(self, text="Settings", command=self.openSettingsWindow, takefocus=False)
        self.settingsButton.grid(row=6, column=0, columnspan=1, padx=10, sticky='nsew')
        self.recordButton = ttk.Button(self, text="Record", command=self.toggleRecording, takefocus=False)
        self.recordButton.grid(row=6, column=1, columnspan=1, padx=10, sticky='nsew')

        # Column 1
        self.loopsFrame.columnconfigure(1, weight=10)
        self.grid_columnconfigure(1, minsize=100)
        self.loopsLabel = ttk.Label(self.loopsFrame, text="Total Loops")
        self.loopsLabel.grid(row=2, column=1, sticky='n', padx=(10,0), pady=(20,0))
        self.loopsLabel.configure(font=("Arial", 11))
        self.loops2Label = ttk.Label(self.loopsFrame, textvariable=self.loops)
        self.loops2Label.grid(row=2, column=1, sticky='', pady=(45,10))
        self.loops2Label.configure(font=("Arial", 11))

        # Column 2
        self.columnconfigure(2, weight=60)
        self.insertPositionButton = ttk.Button(self.actionFrame, text="Insert Position", command=self.addRow, takefocus=False)
        self.insertPositionButton.grid(row=0, column=2, padx=(10, 10), sticky='ew')
        self.getCursorButton = ttk.Button(self.actionFrame, text=" Choose Position", command=self.getCursorPosition, takefocus=False)
        self.getCursorButton.grid(row=1, column=2, sticky='ew', padx=(10, 10))
        self.editRowButton = ttk.Button(self.actionFrame, text="Overwrite Row(s)", command=self.overwriteRows, takefocus=False)
        self.editRowButton.grid(row=2, column=2, sticky='ew', padx=(10, 10), pady=(5, 10))

        # Column 3
        self.actionFrame.columnconfigure(3, weight=60)
        self.actionFrame.grid_columnconfigure(3, minsize=105)
        self.xPosTitleLabel = ttk.Label(self.actionFrame, text="X position")
        self.xPosTitleLabel.configure(font=("Arial", 13))
        self.xPosTitleLabel.grid(row=0, column=3, padx=(0, 0), pady=(15, 0), sticky='n')
        self.xPosLabel = ttk.Label(self.actionFrame, textvariable=self.x)
        self.xPosLabel.grid(row=1, column=3, padx=(10, 0), pady=(0, 0), sticky='n')
        self.yPosTitleLabel = ttk.Label(self.actionFrame, text="Y Position")
        self.yPosTitleLabel.configure(font=("Arial", 13))
        self.yPosTitleLabel.grid(row=1, column=3, padx=(0, 0), pady=(25, 0), sticky='n')
        self.yPosLabel = ttk.Label(self.actionFrame, textvariable=self.y)
        self.yPosLabel.grid(row=2, column=3, padx=(10, 0), pady=(10, 0), sticky='n')

        # Column 4 & 5
        # self.columnconfigure(4, weight=100)
        self.actionFrame.columnconfigure(4, weight=10)
        self.actionFrame.columnconfigure(5, weight=100)
        self.actionFrame.grid_columnconfigure(4, minsize=50)
        self.delayLabel = ttk.Label(self.actionFrame, text="Delay (ms)", font=("Arial", 13))
        self.delayLabel.grid(row=0, column=4, padx=(25, 0), pady=0, sticky="w")
        self.delayEntry = ttk.Entry(self.actionFrame, justify="right", validate='key', textvariable=self.delayVar, validatecommand=self.vcmd)
        # self.delayEntry.insert(0, 0)
        self.delayEntry.grid(row=0, column=5, padx=(10, 0), pady=5, sticky='w')
        self.delayEntry.bind("<KeyRelease>", self.delayEntryKey)

        self.actionLabel = ttk.Label(self.actionFrame, text="Action", font=("Arial", 13))
        self.actionLabel.grid(row=1, column=4, padx=(25, 0), pady=5, sticky="w")
        self.actionEntry = ttk.Entry(self.actionFrame, justify="center", textvariable=self.actionVar)
        self.actionEntry.insert(0, 'M1')
        self.actionEntry.grid(row=1, column=5, columnspan=8, padx=(10, 10), pady=5, sticky='ew')

        # bind this entry to all keyboard and mouse actions
        self.actionEntry.bind("<Key>", self.actionPopulate)
        self.actionEntry.bind("<KeyRelease>", self.actionRelease)
        self.actionEntry.bind('<Return>', self.actionPopulate, add='+')
        self.actionEntry.bind('<KeyRelease-Return>', self.actionRelease, add='+')
        self.actionEntry.bind('<Escape>', self.actionPopulate, add='+')
        self.actionEntry.bind('<KeyRelease-Escape>', self.actionRelease, add='+')
        self.actionEntry.bind('<Button-1>', self.actionPopulate, add='+')
        self.actionEntry.bind('<ButtonRelease-1>', self.actionRelease, add='+')
        self.actionEntry.bind('<Button-2>', self.actionPopulate, add='+')
        self.actionEntry.bind('<ButtonRelease-2>', self.actionRelease, add='+')
        self.actionEntry.bind('<Button-3>', self.actionPopulate, add='+')
        self.actionEntry.bind('<ButtonRelease-3>', self.actionRelease, add='+')
        self.actionEntry.bind("<MouseWheel>", self.actionPopulate, add='+')
        self.actionEntry.bind("<<Paste>>", self.actionPaste)

        self.commentLabel = ttk.Label(self.actionFrame, text="Comment", font=("Arial", 13))
        self.commentLabel.grid(row=2, column=4, sticky="w", padx=(25,0))
        self.commentEntry = ttk.Entry(self.actionFrame, justify="right")
        self.commentEntry.insert(0, '')
        self.commentEntry.grid(row=2, column=5, columnspan=2, padx=(10, 10), pady=5, sticky='ew')

        # Make horizontal scrollbar normal sized
        self.rowconfigure(8, weight=2, minsize=25)

        self.delayToolTip = CreateToolTip(self.delayEntry, self,
                                     "Delay in milliseconds to take place after click. If macro is specified this will be loop amount of that macro. If Action starts with underscore this will be the amount of time to hold listed keys.")
        self.actionToolTip = CreateToolTip(self.actionEntry, self,
                                      "Action to occur before delay. Accepts all mouse buttons and keystrokes. Type !macroname to call another macro with delay amount as loops or #filename to find an image with delay amounts as certainty.")



        self.stopButton.config(state=DISABLED)

        self.rCM = rightClickMenu(self)
        self.previouslySelectedTab = None
        self.previouslySelectedRow = None


    def initTab(self):
        self.canvas = Canvas(self, highlightthickness=0)
        self.canvas.config(height=0)
        self.frame = ttk.Frame(self.canvas)
        self.frame.grid(row=3, column=2, columnspan=8, sticky="ew")

        self.treeFrame = ttk.Frame(self)
        self.treeFrame.grid(row=3, column=2, columnspan=8, rowspan=8, sticky="nsew", pady=(50, 18))
        self.treeView = ttk.Treeview(self.treeFrame, columns=("Step", "X", "Y", "Action", "Delay", "Comment"))
        self.treeView['show'] = 'headings'

        self.horizontalTabScroll = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.horizontalTabScroll.grid(row=8, column=2, columnspan=3, rowspan=1, padx=(0, 5), sticky="nsew")

        self.canvas.configure(xscrollcommand=self.horizontalTabScroll.set)
        self.canvas.grid(row=3, column=2, columnspan=3, sticky="nsew", pady=(10, 0))
        self.canvas.create_window((0, 0), window=self.frame, anchor="nw", tags="frame")

        self.frame.bind("<Configure>", self.frame_configure)
        self.frame.bind("<Button-1>", self.unselectTab)

        self.notebook = ttk.Notebook(self.frame, style="Custom.TNotebook")
        self.tabFrame = ttk.Frame(self.notebook)
        label = Label(self.tabFrame, image = RUNNING_IMAGE)
        label.pack()

        self.notebook.grid(row=3, column=2, columnspan=8, rowspan=1, sticky="ew")
        self.notebook.bind('<<NotebookTabChanged>>', self.tabRefresh)
        self.notebook.bind('<B1-Motion>', self.dragTab)

        self.verticalTabScroll = ttk.Scrollbar(self.treeFrame, orient='vertical', command=self.treeView.yview)
        self.verticalTabScroll.pack(side=RIGHT, fill=Y, anchor="e")
        self.treeView.configure(yscrollcommand=self.verticalTabScroll.set)

        # self.treeView['columns'] = ("Step", "X", "Y", "Action", "Delay", "Comment")

        self.treeView.column('Step', anchor='c', width=35, stretch=NO)
        self.treeView.column('X', anchor='c', width=40, stretch=NO)
        self.treeView.column('Y', anchor='c', width=40, stretch=NO)
        self.treeView.column('Action', anchor='center', width=120, stretch=YES)
        self.treeView.column('Delay', anchor='center', width=85, stretch=NO)
        self.treeView.column('Comment', anchor='center', minwidth=125, stretch=YES)

        self.treeView.heading('Step', text='Step')
        self.treeView.heading('X', text='X')
        self.treeView.heading('Y', text='Y')
        self.treeView.heading('Action', text='Action')
        self.treeView.heading('Delay', text='Delay/Repeat')
        self.treeView.heading('Comment', text='Comment')

        self.treeView.pack(fill=BOTH, expand=True)
        self.treeView.tag_configure('oddrow', background=ODD_COLOR)
        self.treeView.tag_configure('evenrow', background=EVEN_COLOR)
        self.treeView.tag_configure('selected', background=SELECTED_COLOR)
        self.treeView.tag_configure('running', background=RUNNING_COLOR)
        self.treeView.tag_configure('selectedandrunning', background=SELECTED_AND_RUNNING_COLOR_BG, foreground=SELECTED_AND_RUNNING_COLOR_FG)

        self.treeView.bind("<Button-3>", self.rCM.showRightClickMenu)
        self.treeView.bind("<Button-1>", self.tagSelectionClear)
        self.treeView.bind("<ButtonRelease-1>", self.selectRow)

        # sizegrip for resizing window
        self.grip = ttk.Sizegrip()
        self.grip.place(relx=1.0, rely=1.0, anchor="se")
        self.grip.lift(self.horizontalTabScroll)
        self.grip.bind("<B1-Motion>", self.moveMouseButton)


    def moveMouseButton(self, e):
        x1 = self.parent.winfo_pointerx()
        y1 = self.parent.winfo_pointery()
        x0 = self.parent.winfo_rootx()
        y0 = self.parent.winfo_rooty()

        self.parent.geometry("%sx%s" % ((x1-x0),(y1-y0)))

    

    def loopEntryKey(self, event):
        # print(event)
        # print(self.loopsByMacro.get())
        try:
            if (self.loopEntry.get() == "" or int(self.loopEntry.get()) < 1):
                self.loopEntry.set(1)
            elif int(self.loopEntry.get()) > 9999:
                self.loopEntry.set(9999)
        except:
            self.loopEntry.set(1)

        if self.loopsByMacro.get() == 0:
            self.globalLoops.set(int(self.loopsEntry.get()))
        else:
            try:
                self.macroLoops[str(self.notebook.tab(self.notebook.select(), "text"))] = str(self.loopsEntry.get())
                macroLoopsString = ""
                for tab, loops in self.macroLoops.items():
                    macroLoopsString = macroLoopsString + "|" + str(tab) + ":" + str(loops)
                
                config = cp.ConfigParser()
                config.read(os.path.join(FILE_PATH, r'config.ini'))
                config.set('Tabs', 'tabloops', str(macroLoopsString)[1:])
                with open(os.path.join(FILE_PATH, r'config.ini'), 'w') as configfile:
                    config.write(configfile)
            except:
                pass

    
    def delayEntryKey(self, event):
        try:
            if (self.delayVar.get() == "" or int(self.delayVar.get()) < 0):
                self.delayVar.set(0)
        except:
            self.delayVar.set(0)
        # cleanse
        self.delayVar.set(str(int(self.delayVar.get())))
        

    def loadSettings(self):
        # Make StickysAutoClicker folder in windows user's documents folder
        if not os.path.exists(FILE_PATH):
            os.mkdir(FILE_PATH)

        # Make Macros folder in StickysAutoClicker folder
        if not os.path.exists(os.path.join(FILE_PATH, r'Macros')):
            path = os.path.join(FILE_PATH, "Macros")
            os.mkdir(path)

        # Make Images folder in StickysAutoClicker folder
        if not os.path.exists(os.path.join(FILE_PATH, r'Images')):
            path = os.path.join(os.path.join(FILE_PATH, ''), "Images")
            os.mkdir(path)

        # Make config file in StickysAutoClicker folder
        if os.path.exists(os.path.join(FILE_PATH, r'config.ini')):
            config = cp.ConfigParser()
            config.read(os.path.join(FILE_PATH, r'config.ini'))
            updateConfig = False
            # print(os.path.join(FILE_PATH, r'config.ini'))
            
            try:
                # Position x and y save main windows coordinates
                if config.has_option("Position", "x") and config.has_option("Position", "y"):
                    inMonitor = False
                    monitors = screeninfo.get_monitors()
                    for m in reversed(monitors):
                        if m.x <= int(config.get("Position", "x")) <= m.width + m.x - 1 and m.y <= int(config.get("Position", "y")) <= m.height + m.y - 1:
                            self.parent.geometry(str(NORMAL_SIZE[0]) + "x" + str(NORMAL_SIZE[1]) + "+" + config.get("Position", "x") + "+" + config.get("Position", "y"))
                            inMonitor = True
                    if not inMonitor:
                        self.parent.geometry(f"{NORMAL_SIZE[0]}x{NORMAL_SIZE[1]}+100+100")
                else:
                    config.set('Position', 'x', str(100))
                    config.set('Position', 'y', str(100))
                    self.parent.geometry(str(NORMAL_SIZE[0]) + "x" + str(NORMAL_SIZE[1]) + "+100+100")
                    updateConfig = True
            except cp.NoSectionError:
                config.add_section('Position')
                config.set('Position', 'x', str(100))
                config.set('Position', 'y', str(100))
                updateConfig = True

            try:
                # Position logx and logy save log windows coordinates
                if not config.has_option("Position", "logx") or not config.has_option("Position", "logy"):
                    config.set('Position', 'logx', str(10))
                    config.set('Position', 'logy', str(10))
                    updateConfig = True
            except cp.NoSectionError:
                config.add_section('Position')
                config.set('Position', 'logx', str(10))
                config.set('Position', 'logy', str(10))
                updateConfig = True

            try:
                # Position helpx and helpy save help windows coordinates
                if not config.has_option("Position", "helpx") or not config.has_option("Position", "helpy"):
                    config.set('Position', 'helpx', str(10))
                    config.set('Position', 'helpy', str(10))
                    updateConfig = True
            except cp.NoSectionError:
                config.add_section('Position')
                config.set('Position', 'helpx', str(10))
                config.set('Position', 'helpy', str(10))
                updateConfig = True

            try:
                # Position settingsx and settingsy save settings windows coordinates
                if not config.has_option("Position", "settingsx") or not config.has_option("Position", "settingsy"):
                    config.set('Position', 'settingsx', str(10))
                    config.set('Position', 'settingsy', str(10))
                    updateConfig = True
            except cp.NoSectionError:
                config.add_section('Position')
                config.set('Position', 'settingsx', str(10))
                config.set('Position', 'settingsy', str(10))
                updateConfig = True

            try:
                if config.has_option("Settings", "busywait"):
                    self.busyWait.set(config.get("Settings", "busywait"))
                else:
                    self.busyWait.set(0)
                    config.set('Settings', 'busywait', str(self.busyWait.get()))
                    updateConfig = True

                if config.has_option("Settings", "startfromselected"):
                    self.startFromSelected.set(config.get("Settings", "startfromselected"))
                else:
                    self.startFromSelected.set(0)
                    config.set('Settings', 'startfromselected', str(self.startFromSelected.get()))
                    updateConfig = True

                if config.has_option("Settings", "stayontop"):
                    self.stayOnTop.set(int(config.get("Settings", "stayontop")))
                    if self.stayOnTop.get() == 0:
                        self.parent.attributes("-topmost", False)
                    else:
                        self.parent.attributes("-topmost", True)
                else:
                    self.stayOnTop.set(1)
                    config.set('Settings', 'stayontop', '1')
                    updateConfig = True

                if config.has_option("Settings", "hiddenmode"):
                    self.hiddenMode.set(config.get("Settings", "hiddenmode"))
                else:
                    self.hiddenMode.set(0)
                    config.set('Settings', 'hiddenmode', str(self.hiddenMode.get()))
                    updateConfig = True

                if config.has_option("Settings", "selectedapp"):
                    self.selectedApp = config.get("Settings", "selectedapp")
                else:
                    self.selectedApp = ""
                    config.set('Settings', 'selectedapp',  str(self.selectedApp))
                    updateConfig = True

                if config.has_option("Settings", "loops"):
                    try:
                        self.loopsEntry.set(str(config.get("Settings", "loops")))
                        self.globalLoops.set(str(config.get("Settings", "loops")))
                    except:
                        self.loopsEntry.set("1")
                        self.globalLoops.set("1")
                else:
                    self.loopsEntry.set("1")
                    self.globalLoops.set(1)
                    config.set('Settings', 'loops',  str(self.loopEntry.get()))
                    updateConfig = True

                if config.has_option("Settings", "loopsbymacro"):
                    self.loopsByMacro.set(config.get("Settings", "loopsbymacro"))
                else:
                    self.loopsByMacro.set(0)
                    config.set('Settings', 'loopsbymacro',  str(0))
                    updateConfig = True
            except cp.NoSectionError:
                config.add_section('Settings')
                self.busyWait.set(0)
                config.set('Settings', 'busywait', str(self.busyWait.get()))
                self.startFromSelected.set(0)
                config.set('Settings', 'startfromselected', str(self.startFromSelected.get()))
                self.stayOnTop.set(1)
                config.set('Settings', 'stayontop', str(self.stayOnTop.get()))
                self.hiddenMode.set(0)
                config.set('Settings', 'hiddenmode', str(self.hiddenMode.get()))
                self.selectedApp = ""
                config.set('Settings', 'selectedapp',  str(self.selectedApp))
                self.loopEntry.insert(0, 1)
                config.set('Settings', 'loops',  str(self.loopEntry.get()))
                updateConfig = True

            try:
                if config.has_option("Tabs", "opentabs"):
                    openTabs = config.get("Tabs", "opentabs").split("|")
                    if len(openTabs) == 1 and openTabs[0] =='':
                        # config file is blank, open macro1 like old times
                        self.notebook.add(self.tabFrame, text='macro1')
                        self.treeTabs['macro1'] = self.notebook.index(self.notebook.select())
                        config['Tabs'] = {'opentabs': str('|'.join(self.treeTabs.keys()))}
                        updateConfig = True
                    else:
                        for tab in openTabs:
                            self.addTab(tab)
                        # after opening all tabs from config file, select first tab
                        self.notebook.select(self.treeTabs[openTabs[0]])
                else:
                    # config file doesnt exist, open macro1 like old times
                    self.notebook.add(self.tabFrame, text='macro1')
                    self.treeTabs['macro1'] = self.notebook.index(self.notebook.select())
                    config.set('Tabs', 'opentabs', str('|'.join(self.treeTabs.keys())))
                    updateConfig = True
                if not config.has_option("Tabs", "tabloops"):
                    config.set('Tabs', 'tabloops', str(str('|' + str(self.loopEntry.get()))*len(self.treeTabs.keys()))[1:])
                    updateConfig = True
                else:
                    try:
                        for macro in config.get("Tabs", "tabloops").split("|"):
                            self.macroLoops[macro[0:macro.index(":")]] = macro[macro.index(":")+1:]
                    except:
                        config.set('Tabs', 'tabloops', str(str('|' + str(self.loopEntry.get()))*len(self.treeTabs.keys()))[1:])
                        updateConfig = True
            except cp.NoSectionError:
                config.add_section('Tabs')
                self.notebook.add(self.tabFrame, text='macro1')
                self.treeTabs['macro1'] = self.notebook.index(self.notebook.select())
                config['Tabs'] = {'opentabs': str('|'.join(self.treeTabs.keys()))}
                config['Settings'] = {'busyWait': int(self.busyWait.get()), 'startFromSelected': int(self.startFromSelected.get()), 'stayontop': int(self.stayOnTop.get()),  'selectedApp': str(self.selectedApp), 'hiddenMode': int(self.hiddenMode.get()), 'loops': int(self.loopEntry.get())}
                self.loopsByMacro.set(0)
                config.set('Settings', 'loopsbymacro',  str(0))
                self.startFromSelected.set(0)
                self.stayOnTop.set(1)
                if self.stayOnTop.get() == 0:
                    self.parent.attributes("-topmost", False)
                else:
                    self.parent.attributes("-topmost", True)
                self.hiddenMode.set(0)
                self.selectedApp = ""
                self.loopEntry.insert(0, 1)
                config.set('Tabs', 'tabloops', '1')
                updateConfig = True
                config.add_section('Position')
                config.set('Position', 'x', str(100))
                config.set('Position', 'y', str(100))
                inMonitor = False
                monitors = screeninfo.get_monitors()
                for m in reversed(monitors):
                    if m.x <= int(config.get("Position", "x")) <= m.width + m.x - 1 and m.y <= int(config.get("Position", "y")) <= m.height + m.y - 1:
                        self.parent.geometry(str(NORMAL_SIZE[0]) + "x" + str(NORMAL_SIZE[1]) + "+" + config.get("Position", "x") + "+" + config.get("Position", "y"))
                        inMonitor = True
                if not inMonitor:
                    self.parent.geometry(f"{NORMAL_SIZE[0]}x{NORMAL_SIZE[1]}+100+100")
                config.set('Position', 'logx', str(10))
                config.set('Position', 'logy', str(10))
                config.set('Position', 'helpx', str(10))
                config.set('Position', 'helpy', str(10))
                config.set('Position', 'settingsx', str(10))
                config.set('Position', 'settingsy', str(10))

            if updateConfig:
                with open(os.path.join(FILE_PATH, r'config.ini'), 'w') as configfile:
                    config.write(configfile)

        else:
            self.notebook.add(self.tabFrame, text='macro1')
            self.treeTabs['macro1'] = self.notebook.index(self.notebook.select())
            self.startFromSelected.set(0)
            self.stayOnTop.set(1)
            if self.stayOnTop.get() == 0:
                self.parent.attributes("-topmost", False)
            else:
                self.parent.attributes("-topmost", True)
            self.hiddenMode.set(0)
            self.selectedApp = ""
            self.loopEntry.insert(0, 1)

            config = cp.ConfigParser()
            config['Settings'] = {'busyWait': int(self.busyWait.get()), 'startFromSelected': int(self.startFromSelected.get()), 'stayontop': int(self.stayOnTop.get()),  'selectedApp': str(self.selectedApp), 'hiddenMode': int(self.hiddenMode.get()), 'loops': int(self.loopEntry.get())}
            self.loopsByMacro.set(0)
            config.set('Settings', 'loopsbymacro',  str(0))
            config['Tabs'] = {'openTabs': str('|'.join(self.treeTabs.keys()))}
            config.set('Tabs', 'tabloops', '1')
            config.add_section('Position')
            config.set('Position', 'x', str(100))
            config.set('Position', 'y', str(100))
            inMonitor = False
            monitors = screeninfo.get_monitors()
            for m in reversed(monitors):
                if m.x <= int(config.get("Position", "x")) <= m.width + m.x - 1 and m.y <= int(config.get("Position", "y")) <= m.height + m.y - 1:
                    self.parent.geometry(str(NORMAL_SIZE[0]) + "x" + str(NORMAL_SIZE[1]) + "+" + config.get("Position", "x") + "+" + config.get("Position", "y"))
                    inMonitor = True
            if not inMonitor:
                self.parent.geometry(f"{NORMAL_SIZE[0]}x{NORMAL_SIZE[1]}+100+100")
            config.set('Position', 'logx', str(10))
            config.set('Position', 'logy', str(10))
            config.set('Position', 'helpx', str(10))
            config.set('Position', 'helpy', str(10))
            config.set('Position', 'settingsx', str(10))
            config.set('Position', 'settingsy', str(10))

            with open(os.path.join(FILE_PATH, r'config.ini'), 'w') as configfile:
                config.write(configfile)
        
        self.after(200, self.scrollLeft)

    # Start clicking helper to multi thread click loop, otherwise sleep will make windows want to kill
    def threadStartClicking(self):
        # save current macro to csv, easier to read, helps update retention
        self.exportMacro()

        self.startButton.config(state=DISABLED)
        self.stopButton.config(state=NORMAL)
        self.clickLog = []

        try:
            self.runningTab = self.notebook.select()
            self.notebook.tab(self.notebook.select(), image=RUNNING_IMAGE, compound="right", padding ={0, 0, 0, 0})
            currTab = self.notebook.tab(self.notebook.select(), "text")
            clickArray = list(csv.reader(open(FILE_PATH + r'\Macros' + '\\' + str(
                currTab.replace(r'/', r'-').replace(r'\\', r'-').replace(
                    r'*', r'-').replace(r'?', r'-').replace(r'[', r'-').replace(r']', r'-').replace(r'<', r'-').replace(
                    r'>', r'-').replace(r'|', r'-')) + '.csv', mode='r')))

            # set loops counter
            self.loops.set(int(self.loopEntry.get()))
            self.loopsLeft.set(self.loopEntry.get())
            self.update()

            # create event and give as attribute to thread object, this can be referenced in loop and prevent sleeping of thread
            # using wait instead of sleep lets thread stop when Stop Clicking is clicked while waiting allow the thread to be quickly killed
            # because if Start Clicking is clicked before thread is killed this will overwrite saved thread and prevent setting event with intent to kill thread
            threadFlag = threading.Event()
            self.pauseEvent = threading.Event()
            self.pauseEvent.set()
            self.activeThread = threading.Thread(target=self.startClicking, args=(currTab, self.busyWait.get() == 1, clickArray, int(self.loopEntry.get()), 0))
            self.activeThread.threadFlag = threadFlag
            self.activeThread.start()
        except Exception as e:
            # print(e)
            self.logAction("", depth, e)
            # hmmmm
            print("Error: unable to start clicking thread")

        # Emergency Exit key combo that will stop auto clicking in case mouse is moving to fast to click stop
        try:
            self.monitorThread = threading.Thread(target=self.monitorKeys, args=())
            self.monitorThread.threadFlag = self.activeThread.threadFlag
            self.monitorThread.start()
        except Exception as e:
            # print(e)
            self.logAction("", depth, e)
            # hmmmmm
            print("Error: unable to start exit monitoring thread")




    def on_press(self, key):
        if str(key) in ALL_COMBO:
            # print("Press: ", key)
            if str(key) not in self.monitorKeysPressed:
                self.monitorKeysPressed.add(str(key))
            if all(x in self.monitorKeysPressed for x in PAUSE_COMBO):
                # print("PAUSE COMBO PRESSED")
                self.logAction("", 0, "Pause Combo Pressed.")
                self.monitorKeysPressed = set()
                self.togglePause()
            if any(all(x in self.monitorKeysPressed for x in z) for z in EXIT_COMBO):
                # print("EXIT COMBO PRESSED")
                self.logAction("", 0, "Exit Combo Pressed.")
                self.monitorKeysPressed = set()
                self.stopClicking()


    def on_release(self, key):
        if str(key) in ALL_COMBO:
            # print("Release: ", key)
            self.monitorKeysPressed.discard(str(key))


    def monitorKeys(self):
        with keyboard.Listener(on_press=self.on_press, on_release=self.on_release) as listener:
            self.keysListener = listener
            listener.join()
            

    # Start Clicking button will loop through active notebook tab's treeview and move mouse, call macro, search for image and click or type keystrokes with specified delays
    # Intentionally uses busy wait for most accurate delays
    def startClicking(self, macroName, busy, clickArray, loopsParam, depth):
        try:
            self.update()
            # print("startClicking ", macroName)

            if loopsParam == 0 or self.loopsLeft.get() == 0: return

            depthDelineator = " | "
            clickTime = 0
            prevDelay = 0
            blnDelay = False
            firstLoop = True
            selection = self.treeView.selection()
            startFromSelected = self.startFromSelected.get()
            if len(selection) > 0:
                selectedRow = self.treeView.item(selection[0]).get("values")[0]
            else:
                selectedRow = 0
            if depth > 0 or startFromSelected == 0 or selectedRow == 1 or selectedRow == 0:
                self.runningRows.append((macroName, 1))

            # check Loopsleft as well to make sure Stop button wasn't pressed since this doesn't ues a global for loop count
            while loopsParam > 0 and self.loopsLeft.get() > 0 or self.activeThread.threadFlag.is_set():
                # print("New loop")
                self.logAction("", depth, "Begin " + macroName + ", looping " + str(loopsParam) + " times.")
                intRow = 0
                startTime = time.time()
                skipped = 0

                for row in clickArray:
                    # print(row)
                    if loopsParam == 0 or self.loopsLeft.get() == 0: return
                    intRow += 1
                    # When start from selected row setting is true then find highlighted row(s) and skip to from first selected row
                    # Only for first loop and first macro, not for subsequent loops nor macros called by the starting macro
                    if firstLoop and depth == 0 and startFromSelected == 1 and selectedRow > 1 and intRow < selectedRow:
                        # skipped = 1
                        continue
                    # else:
                        # self.updateRunningRow()

                    # check row to see if its still holding keys
                    if len(self.currPressed) > 0 and (row[2] == "" or row[2][0] == '_'):
                        # stop time if next action is not hold
                        toBePressed = []
                        if row[2] != "" and row[2][0] == '_':
                            toBePressed = row[2][1:].split('|')

                        # key hold and release is action
                        # print('PreRelease', (prevDelay / 1000), time.time() - clickTime)
                        if busy:
                            while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                pass
                        else:
                            while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                return
                        blnDelay = False
                        # print('PreReleasePostWait', (prevDelay / 1000), time.time() - clickTime)
                        # Release all keys not pressed in next step
                        for key in self.currPressed:
                            if key not in toBePressed:
                                if key == 'M1':
                                    pyautogui.mouseUp(button='left')
                                elif key == 'M3':
                                    pyautogui.mouseUp(button='right')
                                elif key == 'M2':
                                    pyautogui.mouseUp(button='middle')
                                elif key == 'space':
                                    pyautogui.keyUp(' ')
                                elif key == 'tab':
                                    pyautogui.keyUp('\t')
                                else:
                                    pyautogui.keyUp(key)
                                try:
                                    self.currPressed.remove(key)
                                except:
                                    pass
                                # print("Release " + str(key) + " after " + str(time.time() - clickTime))
                                # print("Release " + str(key) + " after total " + str(time.time() - startTime))
                                # print("Release1 " + str(key) + " after " + str(int(prevDelay) / 1000) + " with time diff: " + str(time.time() - clickTime - int(prevDelay) / 1000))
                        # if this row is also a press then start timer as hols is continuing
                        # clickTime = time.time()
                    elif len(self.currPressed) > 0:
                        # print('PreRelease', str(pressed), (int(row[3]) / 1000), time.time() - clickTime)
                        if busy:
                            while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                pass
                        else:
                            while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                return
                        # print('PreReleaseWait', (int(row[3]) / 1000), time.time() - clickTime)
                        # Release all keys as next step is not pressing
                        for key in self.currPressed:
                            if key == 'M1':
                                pyautogui.mouseUp(button='left')
                            elif key == 'M3':
                                pyautogui.mouseUp(button='right')
                            elif key == 'M2':
                                pyautogui.mouseUp(button='middle')
                            elif key == 'space':
                                pyautogui.keyUp(' ')
                            elif key == 'tab':
                                pyautogui.keyUp('\t')
                            else:
                                pyautogui.keyUp(key)
                            try:
                                self.currPressed.remove(key)
                            except:
                                pass
                                # print("Release " + str(key) + " after " + str(time.time() - clickTime))
                            # print("Release " + str(key) + " after total " + str(time.time() - startTime))
                            # print("Release2 " + str(key) + " after " + str(int(prevDelay) / 1000) + " with time diff: " + str(time.time() - clickTime - int(prevDelay) / 1000))
                        # if this row is not a press then wait to start timer until next click
                        clickTime = time.time()

                    if loopsParam == 0 or self.loopsLeft.get() == 0: break

                    # Empty must be first because other references to first character of Action will error
                    if row[2] == "":
                        # Nothing is just a pause
                        # print("blank", (int(row[3]) / 1000), time.time() - clickTime)
                        self.logAction(row[4], depth, "Pause for " + str(row[3]) + " ms.")
                        if blnDelay:
                            if busy:
                                while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                    pass
                            else:
                                while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                    return

                        clickTime = time.time()
                    # Moved _ hold action as close top as possible as it's accuracy is most important
                    # Action is starts with an underscore (_), hold the key for the given amount of time
                    elif row[2][0] == '_' and len(row[2]) > 1:
                        # loop until end of string of keys to press
                        toPress = row[2][1:].split('|')

                        # wait prior row amount before pressing new keys
                        if blnDelay:
                            if busy:
                                while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                    pass
                            else:
                                while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                    return

                        for key in toPress:
                            # Do not press if already pressed
                            if key not in self.currPressed:
                                if key in ['M1', 'M2', 'M3']:
                                    if int(row[0]) != 0 or int(row[1]) != 0:
                                        # mouse click is action
                                        if key == 'M1':
                                            pyautogui.mouseDown(int(row[0]), int(row[1]), button='left')
                                            self.currPressed.append(key)
                                        elif key == 'M3':
                                            pyautogui.moveTo(int(row[0]), int(row[1]))
                                            pyautogui.mouseDown(button='right')
                                            self.currPressed.append(key)
                                        else:
                                            pyautogui.mouseDown(int(row[0]), int(row[1]), button='middle')
                                            self.currPressed.append(key)
                                    else:
                                        # mouse click is action without position, do not move, just click
                                        if key == 'M1':
                                            pyautogui.mouseDown(button='left')
                                            self.currPressed.append(key)
                                        elif key == 'M3':
                                            pyautogui.mouseDown(button='right')
                                            self.currPressed.append(key)
                                        else:
                                            pyautogui.mouseDown(button='middle')
                                            self.currPressed.append(key)
                                else:
                                    # key press is action
                                    if key == 'space':
                                        pyautogui.keyDown(' ')
                                        self.currPressed.append(key)
                                    elif key == 'tab':
                                        pyautogui.keyDown('\t')
                                        self.currPressed.append(key)
                                    else:
                                        pyautogui.keyDown(key)
                                        self.currPressed.append(key)
                                # print("Press " + str(key) + " at " + str(time.time() - clickTime))
                        if len(toPress) > 2:
                            self.logAction(row[4], depth, "Hold " + ", ".join(toPress[0:len(toPress) - 2])  + " and " + str(toPress[0]) + " for " + str(row[3]) + " ms.")
                        elif len(toPress) > 1:
                            self.logAction(row[4], depth, "Hold " + str(toPress[0]) + " and " + str(toPress[1]) + " for " + str(row[3]) + " ms.")
                        else:
                            self.logAction(row[4], depth, "Hold " + str(toPress[0]) + " for " + str(row[3]) + " ms.")
                        clickTime = time.time()

                    elif row[2] in ['M1', 'M2', 'M3']:
                        if int(row[0]) != 0 or int(row[1]) != 0:
                            # mouse click is action
                            if row[2] == 'M1':
                                if blnDelay:
                                    if busy:
                                        while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                            pass
                                    else:
                                        while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                            return

                                self.pauseEvent.wait()
                                pyautogui.click(int(row[0]), int(row[1]), button='left')
                                clickTime = time.time()

                            elif row[2] == 'M3':
                                if blnDelay:
                                    if busy:
                                        while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                            pass
                                    else:
                                        while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                            return

                                pyautogui.moveTo(int(row[0]), int(row[1]))
                                pyautogui.mouseDown(button='right')
                                time.sleep(.1)  # TODO remove forced delay?
                                pyautogui.mouseUp(button='right')
                                clickTime = time.time()
                            else:
                                if blnDelay:
                                    if busy:
                                        while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                            pass
                                    else:
                                        while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                            return

                                pyautogui.click(int(row[0]), int(row[1]), button='middle')
                                clickTime = time.time()
                        else:
                            # mouse click is action without position, do not move, just click
                            if row[2] == 'M1':
                                if blnDelay:
                                    if busy:
                                        while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                            pass
                                    else:
                                        while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                            return

                                pyautogui.click(button='left')
                                clickTime = time.time()
                            elif row[2] == 'M3':
                                if blnDelay:
                                    if busy:
                                        while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                            pass
                                    else:
                                        while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                            return

                                pyautogui.mouseDown(button='right')
                                time.sleep(.1)
                                pyautogui.mouseUp(button='right')
                                clickTime = time.time()
                            else:
                                if blnDelay:
                                    if busy:
                                        while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                            pass
                                    else:
                                        while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                            return

                                pyautogui.click(button='middle')
                                clickTime = time.time()
                        self.logAction(row[4], depth, "Wait " + str(row[3]) + " ms and then press " + str(row[2]) + ".")
                        # print("Press " + str(row[2]) + " at " + str(time.time() - startTime))

                    # Action is #string, find image in Images folder with string name
                    elif row[2][0] == '#' and len(row[2]) > 1:
                        # delay/100 is confidence
                        confidence = row[3]
                        position = 0
                        # If Find image ends in ? then do not end loop if image not found
                        optional = (row[2][len(row[2]) - 1] == '?')
                        if optional:
                            image = str(row[2][1:len(row[2]) - 1]) 
                        else:
                            image = str(row[2][1:len(row[2])])
                        # confidence must be a percentile
                        if not (100 >= int(confidence) > 0):
                            confidence = 80
                        # print(FILE_PATH + r'\Images' + '\\' + str(row[2][1:len(row[2])]) + '.png')
                        if blnDelay:
                            if busy:
                                while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                    pass
                            else:
                                while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                    return
                        try:
                            # Confidence specified, use Delay as confidence percentile
                            position = pyautogui.locateCenterOnScreen(FILE_PATH + r'\Images' + '\\' + image + '.png', confidence=int(confidence) / 100)
                        except (IOError, pyautogui.ImageNotFoundException) as error:
                            self.logAction(row[4], depth, "Exception caught: " + str(error))
                            pass

                        if position:
                            self.logAction(row[4], depth, image + " image found at coords: " + str(position[0] + global_monitor_left) + ' : ' + str(position[1] + global_monitor_top) + ".")

                            pyautogui.click(position[0], position[1], button='left')
                            clickTime = time.time()
                        else:
                            if not optional:
                                self.logAction(row[4], depth, image + " image not found and not optional. Quitting loops.")
                                break
                            else:
                                self.logAction(row[4], depth, image + " image not found but optional. Continuing loops.")
                        
                    # Action is !string, run macro with string name for Delay amount of times
                    elif row[2][0] == '!' and len(row[2]) > 1:

                        if  '!MScrl' in row[2]:
                            # print(int(row[2][row[2].index('(') + 1:row[2].index(')')]))
                            self.logAction(row[4], depth, "Scrolling " + str(abs(int(row[2][row[2].index('(') + 1:row[2].index(')')]))) + " units " + (" up." if int(row[2][row[2].index('(') + 1:row[2].index(')')]) > 0 else " down."))
                            if busy:
                                while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                    pass
                            else:
                                while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                    return
                            pyautogui.scroll(int(row[2][row[2].index('(') + 1:row[2].index(')')]) * 120)
                            clickTime = time.time()
                        elif  '!Paste' in row[2]:
                            # print(str(row[2][row[2].index('(') + 1:row[2].index(')')]))
                            self.logAction(row[4], depth, "Pasted " + str(row[2][row[2].index('(') + 1:row[2].index(')')]) + ".")
                            if busy:
                                while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                    pass
                            else:
                                while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                    return
                            pyautogui.typewrite(str(row[2][row[2].index('(') + 1:row[2].index(')')]))
                        else:
                            # macro is action, repeat for amount in delay
                            if exists(FILE_PATH + r'\Macros' + '\\' + row[2][1:len(row[2])] + '.csv'):
                                arrayParam = list(csv.reader(
                                    open(FILE_PATH + r'\Macros' + '\\' + row[2][1:len(row[2])] + '.csv',
                                        mode='r')))
                                # print("file: ", FILE_PATH + r'\Macros' + '\\' + row[2][1:len(row[2])] + '.csv')
                                if busy:
                                    while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                        pass
                                else:
                                    while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                        return
                                    
                                self.logAction(row[4], depth, "Running " + str(row[2][1:len(row[2])]) + " macro " + str(row[3]) + " times.")
                                self.startClicking(row[2][1:len(row[2])], busy, arrayParam, int(row[3]), depth + 1)
                            else:
                                self.logAction(row[4], depth, "Macro " + str(row[2][1:len(row[2])]) + " not found in " + str(FILE_PATH + r'\Macros') + ".")
                    else:
                        # print(row[2])
                        # key press is action
                        if row[2] != 'space':
                            if blnDelay:
                                if busy:
                                    while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                        pass
                                else:
                                    while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                        return

                            pyautogui.press(row[2])
                            clickTime = time.time()
                        elif row[2] == 'space':
                            if blnDelay:
                                if busy:
                                    while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                        pass
                                else:
                                    while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                        return

                            pyautogui.press(' ')
                            clickTime = time.time()
                        if row[2] == 'tab':
                            if blnDelay:
                                if busy:
                                    while (time.time() < clickTime + (int(prevDelay) / 1000) or not self.pauseEvent.wait()) and not self.activeThread.threadFlag.is_set():
                                        pass
                                else:
                                    while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime) or not self.pauseEvent.wait() or self.activeThread.threadFlag.is_set():
                                        return

                            pyautogui.press('\t')
                            clickTime = time.time()
                        else:
                            clickTime = time.time()
                        self.logAction(row[4], depth, "Wait " + str(row[3]) + " ms and then press " + str(row[2]) + ".")
                        # print("Press " + str(row[2]) + " at " + str(time.time() - startTime))

                    if len(self.runningRows) > 0:
                        # self.removeRunningRow(self.runningRows[depth])
                        self.runningRows[depth] = (macroName, intRow)
                    else:
                        self.runningRows.append((macroName, intRow))
                    if macroName == self.currTab:
                        self.updateRunningRow()

                    if loopsParam == 0 or self.loopsLeft.get() == 0: break

                    prevDelay = int(row[3])
                    blnDelay = row[2] == "" or row[2][0] != "_"
                    
                # break outer loop only if inner loop breaks for loop stop
                if loopsParam == 0 or self.loopsLeft.get() == 0:
                    break
                # decrement loop count param, also decrement main loop counter if main loop
                if loopsParam > 0: 
                    loopsParam = loopsParam - 1
                    
                if depth == 0 and self.loopsLeft.get() > 0: 
                    self.loopsLeft.set(self.loopsLeft.get() - 1)
                firstLoop = False
                continue

            while self.activeThread.threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                return
            # Release all keys that are pressed so as to not leave pressed
            # Necessary to be outside of loop because last row in macro will expect next row to release right prior to next press.
            if len(self.currPressed) > 0:
                self.logAction("", depth, "Release all pressed keys: " + str(self.currPressed))
            for pressedKey in self.currPressed:
                if pressedKey != 'space':
                    pyautogui.keyUp(pressedKey)
                elif pressedKey == 'M1':
                    pyautogui.mouseUp(button='left')
                elif pressedKey == 'M3':
                    pyautogui.mouseUp(button='right')
                elif pressedKey == 'M2':
                    pyautogui.mouseUp(button='middle')
                elif pressedKey == 'space':
                    pyautogui.keyUp(' ')
                elif pressedKey == 'tab':
                    pyautogui.keyUp('\t')
                try:
                    self.currPressed.remove(key)
                except:
                    pass
                # print("Release3 " + str(pressedKey) + " after " + str(int(prevDelay) / 1000) + " with time diff: " + str(time.time() - clickTime - int(prevDelay) / 1000))
                # print("Release " + str(pressedKey) + " after " + str(time.time() - clickTime))

            if self.loopsLeft.get() > 0 and depth == 0: self.loopsLeft.set(self.loopsLeft.get() - 1)
            self.update()
            self.removeRunningRow(self.runningRows[depth])
            self.runningRows.pop()
            if self.loopsLeft.get() == 0 and depth == 0:
                self.startButton.config(state=NORMAL)
                self.stopButton.config(state=DISABLED)
                self.keysListener.stop()
                self.runningRows = []
                try:
                    self.notebook.tab(self.runningTab, image=NONE_IMAGE, compound="center")
                except:
                    pass
            self.logAction("", depth, "Exiting clicking.")
            # print("Exit Clicking")
        except Exception as e:
            self.logAction("", depth, "Exception caught: " + str(e))
            return


    def stopClicking(self):
        # print("Stop")
        self.logAction("", 0, "Stopping clicking.")
        self.loopsLeft.set(0)
        self.startButton.config(state=NORMAL)
        self.stopButton.config(state=DISABLED)

        try:
            self.keysListener.stop()
            self.activeThread.threadFlag.set()
            self.activeThread.join(1)
        except:
            pass

        try:
            self.pauseEvent.set()
        except:
            pass

        self.runningRows = []
        self.monitorKeysPressed = set()
        self.reorderRows()
        self.tagSelection()
        try:
            self.notebook.tab(self.runningTab, image=NONE_IMAGE, compound="center")
        except:
            pass

        # Release all keys that are pressed so as to not leave pressed
        for pressedKey in self.currPressed:
            if pressedKey != 'space':
                pyautogui.keyUp(pressedKey)
            elif pressedKey == 'M1':
                pyautogui.mouseUp(button='left')
            elif pressedKey == 'M3':
                pyautogui.mouseUp(button='right')
            elif pressedKey == 'M2':
                pyautogui.mouseUp(button='middle')
            elif pressedKey == 'space':
                pyautogui.keyUp(' ')
            elif pressedKey == 'tab':
                pyautogui.keyUp('\t')
            # print("Release " + str(pressedKey))
        self.currPressed = []


    def logAction(self, comment, depth, string):
        if comment != "":
            self.clickLog.append(str(datetime.datetime.now()) + ': ' + str(" | "*depth) + '"' + str(comment) + '" ' + str(string))
        else:
            self.clickLog.append(str(datetime.datetime.now()) + ': ' + str(" | "*depth) + str(string))
            
        if self.logWindow is not None:
            self.logWindow.updateText("\n".join(self.clickLog))


    def logError(self, error):
        self.clickLog.append("Error thrown! \n" + str(error))


    def tabRefresh(self, event):
        # print('tabRefresh')
        # clear treeview of old macro
        for item in self.treeView.get_children():
            self.treeView.delete(item)

        self.currTab = event.widget.tab('current')['text']

        # import csv file into macro where filename will be macro name, is it best to import only exported macros
        filename = os.path.join(FILE_PATH + r'\Macros', self.currTab + '.csv')
        fileExists = exists(filename)

        if fileExists:
            with open(filename, 'r') as csvFile:
                csvReader = csv.reader(csvFile)
                if not filename:
                    self.addTab(os.path.splitext(os.path.basename(self.currTab))[0])
                step = 1
                for line in csvReader:
                    # backwards compatibility to before there were comments
                    if len(line) > 4:
                        self.addRowWithParams(line[0], line[1], line[2], line[3], line[4])
                    elif line[3]:
                        self.addRowWithParams(line[0], line[1], line[2], line[3], "")
                    step += 1
        self.reorderRows()
        self.xPosLabel.focus()
        self.scrollUp()
        
        if self.loopsByMacro.get() == 1:
            try:
                if self.currTab in self.macroLoops:
                    self.loopsEntry.set(self.macroLoops[self.currTab])
                else:
                    self.loopsEntry.set(self.globalLoops.get())
            except:
                pass
        else:
            self.loopsEntry.set(self.globalLoops.get())


    def dragTab(self, event):
        try:
            # get tab that is under cursor
            index = self.notebook.index(f"@{event.x},{event.y}")
            # to prevent 
            if index != self.priorDraggedTab:
                self.priorDraggedTab = self.notebook.index(self.notebook.select())
                self.notebook.insert(index, child=self.notebook.select())

                config = cp.ConfigParser()
                config.read(os.path.join(FILE_PATH, r'config.ini'))
                config.set('Tabs', 'openTabs', str('|'.join([self.notebook.tab(i, option="text") for i in self.notebook.tabs()])))
                with open(os.path.join(FILE_PATH, 'config.ini'), 'w') as configfile:
                    config.write(configfile)
        except tk.TclError:
            pass

    def frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))


    def unselectTab(self, event):
        self.xPosLabel.focus()


    def addTab(self, name):
        self.tabFrame = ttk.Frame(self.notebook, takefocus=0)
        self.notebook.add(self.tabFrame, text=name)
        self.notebook.select(self.tabFrame)
        self.treeTabs[name] = self.notebook.index(self.notebook.select())

        try:
            config = cp.ConfigParser()
            config.read(os.path.join(FILE_PATH, r'config.ini'))
            config.set('Tabs', 'openTabs', str('|'.join(self.treeTabs.keys())))
            with open(os.path.join(FILE_PATH, 'config.ini'), 'w') as configfile:
                config.write(configfile)
        except:
            pass
        
        self.after(200, self.scrollRight)

    
    def scrollLeft(self):
        self.canvas.xview_moveto(0.0)


    def scrollRight(self):
        self.canvas.xview_moveto(1.0)

    
    def scrollUp(self):
        self.treeView.yview_moveto(0.0)

    
    def scrollDown(self):
        self.treeView.yview_moveto(1.0)


    def toggleRecording(self):
        if self.recorder is not None:
            self.recorder.stopRecordingThread()
            self.monitorKeysPressed = set()
            del self.recorder
            self.recorder = None
        else:
            # currently think supporting clicking while recording is a bad idea
            self.stopClicking()
            self.startButton.config(state=DISABLED)
            # self.recordButton.configure(style="RecordButton")

            self.recorder = Recorder(self)
            self.recorder.startRecordingThread()


    def togglePause(self): 
        if self.pauseEvent.is_set():
            self.pauseEvent.clear()
        else:
            self.pauseEvent.set()


    def getOrigin(self):
        # unhook mouse listener after getting mouse click, set to X, Y where mouse was clicked
        try:
            (a, b) = pyautogui.position()
            self.x.set(a)
            self.y.set(b)
        finally:
            mouse.unhook_all()


    def getMouseMove(self, event):
        # return position from mouse click from anywhere on screen, set to X, Y entry values as mouse moves\
        (a, b) = pyautogui.position()
        self.x.set(a)
        self.y.set(b)


    def getCursorPosition(self):
        # listen to mouse for next click
        mouse.hook(self.getMouseMove)
        mouse.on_click(self.getOrigin)

    def addRow(self):
        # from Insert position button, adds row to bottom current treeview using entry values
        if len(self.treeView.get_children()) % 2 == 0:
            self.treeView.insert(parent='', index='end', iid=None, text=len(self.treeView.get_children()) + 1, values=(len(self.treeView.get_children()) + 1, self.x.get(), self.y.get(), self.actionEntry.get(), int(self.delayEntry.get()), self.commentEntry.get()), tags='evenrow')
        else:
            self.treeView.insert(parent='', index='end', iid=None, text=len(self.treeView.get_children()) + 1, values=(len(self.treeView.get_children()) + 1, self.x.get(), self.y.get(), self.actionEntry.get(), int(self.delayEntry.get()), self.commentEntry.get()), tags='oddrow')
        self.exportMacro()
        self.scrollDown()


    def addRowWithParams(self, xParam, yParam, keyParam, delayParam, commentParam):
        # for import to populate new treeview
        if len(self.treeView.get_children()) % 2 == 0:
            self.treeView.insert(parent='', index='end', iid=None, text=len(self.treeView.get_children()) + 1, values=(len(self.treeView.get_children()) + 1, xParam, yParam, keyParam, delayParam, commentParam), tags='evenrow')
        else:
            self.treeView.insert(parent='', index='end', iid=None, text=len(self.treeView.get_children()) + 1, values=(len(self.treeView.get_children()) + 1, xParam, yParam, keyParam, delayParam, commentParam), tags='oddrow')
        self.exportMacro()
        self.scrollDown()


    def selectRow(self, event):
        # this event populates the entries with the values of the selected row for easy editing/copying
        # must use global variables to tell if new row selected or just clicking whitespace in treeview, do not set values if not changing row selection
        selectedRow = self.treeView.focus()
        self.tagSelection()
        # print("selectRow")
        selectedValues = self.treeView.item(selectedRow, 'values')

        if len(selectedValues) > 0 and (self.previouslySelectedTab != self.notebook.index(
                self.notebook.select()) or self.previouslySelectedRow != selectedRow):
            self.actionEntry.delete(0, 'end')
            self.actionEntry.insert(0, selectedValues[3])
            self.delayVar.set(selectedValues[4])
            self.commentEntry.delete(0, 'end')
            self.commentEntry.insert(0, selectedValues[5])
            self.x.set(selectedValues[1])
            self.y.set(selectedValues[2])
            self.previouslySelectedTab = self.notebook.index(self.notebook.select())
            self.previouslySelectedRow = selectedRow
            self.update()


    def removeRunningRow(self, tupleRow):
        print("Remove " + str(tupleRow))
        rows = self.treeView.get_children()
        if tupleRow[0] == self.notebook.tab(self.notebook.select(), "text"):
            print(self.treeView.item(rows[tupleRow[1] - 1])['tags'])
            if tupleRow[1] - 1 % 2 == 0:
                if self.treeView.item(rows[tupleRow[1] - 1])['tags'][0] == 'selected' or self.treeView.item(rows[tupleRow[1] - 1])['tags'][0] == 'selectedandrunning':
                    self.treeView.item(rows[tupleRow[1] - 1], tags='selected')
                else:
                    self.treeView.item(rows[tupleRow[1] - 1], tags='evenrow')
            else:
                if self.treeView.item(rows[tupleRow[1] - 1])['tags'][0] == 'selected' or self.treeView.item(rows[tupleRow[1] - 1])['tags'][0] == 'selectedandrunning':
                    self.treeView.item(rows[tupleRow[1] - 1], tags='selected')
                else:
                    self.treeView.item(rows[tupleRow[1] - 1], tags='oddrow')

    def reorderRows(self):
        rows = self.treeView.get_children()
        i = 1
        # print("reorderRows")
        # overwrite all rows back to even n odd
        for row in rows:
            if i % 2 == 0:
                self.treeView.item(row, text=i, tags='oddrow')
                self.treeView.set(row, "Step", i)
            else:
                self.treeView.item(row, text=i, tags='evenrow')
                self.treeView.set(row, "Step", i)
            i += 1


    def updateRunningRow(self):
        # print("updateRunningRow")
        # print(self.runningRows)
        rows = self.treeView.get_children()
        for tuple in self.runningRows:
            # If running row in current tab
            # print(tuple)
            # print(self.treeView.item(rows[tuple[1] - 1])['tags'])
            if tuple[0] == self.notebook.tab(self.notebook.select(), "text"):
                if self.treeView.item(rows[tuple[1] - 1])['tags'][0] == 'selected' or self.treeView.item(rows[tuple[1] - 1])['tags'][0] == 'selectedandrunning':
                    self.treeView.item(rows[tuple[1] - 1], tags='selectedandrunning')
                else:
                    self.treeView.item(rows[tuple[1] - 1], tags='running')
                # Update prior row tag back to even n odd
                if tuple[1] > 1:
                    if tuple[1] % 2 == 0:
                        if self.treeView.item(rows[tuple[1] - 2])['tags'][0] == 'selected' or self.treeView.item(rows[tuple[1] - 2])['tags'][0] == 'selectedandrunning':
                            self.treeView.item(rows[tuple[1] - 2], tags='selected')
                        else:
                            self.treeView.item(rows[tuple[1] - 2], tags='evenrow')
                    else:
                        if self.treeView.item(rows[tuple[1] - 2])['tags'][0] == 'selected' or self.treeView.item(rows[tuple[1] - 2])['tags'][0] == 'selectedandrunning':
                            self.treeView.item(rows[tuple[1] - 2], tags='selected')
                        else:
                            self.treeView.item(rows[tuple[1] - 2], tags='oddrow')
                # If first row then update last row back to even or odd only if more than one row in macro
                elif len(rows) > 1:
                        # If running is first row in macro then update last row in macro as that might have been last running row before loop restarts
                    if (len(rows) - 1) % 2 == 0:
                        if self.treeView.item(rows[len(rows) - 1])['tags'][0] == 'selected' or self.treeView.item(rows[tuple[1] - 2])['tags'][0] == 'selectedandrunning':
                            self.treeView.item(rows[len(rows) - 1], tags='selected')
                        else:
                            self.treeView.item(rows[len(rows) - 1], tags='evenrow')
                    else:
                        if self.treeView.item(rows[len(rows) - 1])['tags'][0] == 'selected' or self.treeView.item(rows[tuple[1] - 2])['tags'][0] == 'selectedandrunning':
                            self.treeView.item(rows[len(rows) - 1], tags='selected')
                        else:
                            self.treeView.item(rows[len(rows) - 1], tags='oddrow')


    def tagSelectionClear(self, event):
        # print("tagSelectionClear")
        # print(self.treeView.selection())
        rows = self.treeView.get_children()
        if self.treeView.selection() != ():
            for row in self.treeView.selection():
                if (self.treeView.item(row)['values'][0] - 1) % 2 == 0:
                    if self.treeView.item(row)['tags'] == 'selectedandrunning':
                        self.treeView.item(row, tags='selected')
                    else:
                        self.treeView.item(row, tags='evenrow')
                else:
                    if self.treeView.item(row)['tags'] == 'selectedandrunning':
                        self.treeView.item(row, tags='selected')
                    else:
                        self.treeView.item(row, tags='oddrow')


    def tagSelection(self):
        # print("tagSelection")
        for row in self.treeView.selection():
            if not self.treeView.item(row)['tags'][0] == 'running':
                self.treeView.item(row, tag='selected')
            else:
                self.treeView.item(row, tag='selectedandrunning')


    def importMacro(self):
        # import csv file into macro where filename will be macro name, is it best to import only exported macros
        filename = filedialog.askopenfilename(initialdir=FILE_PATH + r'\Macros', title="Select a .csv file",
                                            filetypes=(("csv files", "*.csv"),))
        answer = True
        found = 0

        # look for tab with same name as imported macro, will overwrite that tab if imported
        # this will ensure each macro has a unique name so that when a macro calls another macro there is no confusion over which macro should be called
        for i in range(len(self.notebook.tabs())):
            if self.notebook.tab(i, 'text') == os.path.splitext(os.path.basename(filename))[0]:
                answer = askyesno("Overwrite macro?", "Are you sure you want to overwrite current " +
                                os.path.splitext(os.path.basename(filename))[0])
                found = i + 1
                break

        # only open if file exists and overwrite true (answer defaults to true in case it is not asked)
        if filename and answer:
            with open(filename, 'r') as csvFile:
                csvReader = csv.reader(csvFile)
                if not found:
                    self.addTab(os.path.splitext(os.path.basename(filename))[0])


    def exportMacro(self):
        # save as csv file with name of file as macro name
        savePath = FILE_PATH + r'\Macros'
        filename = str(
            self.notebook.tab(self.notebook.select(), 'text').replace(r'/', r'-').replace(r'\\', r'-').replace(r'*', r'-').replace(
                r'?', r'-').replace(r'[', r'-').replace(r']', r'-').replace(r'<', r'-').replace(r'>', r'-').replace(r'|', r'-'))

        # print('export; ', filename)
        with open(os.path.join(savePath, str(filename + '.csv')), 'w', newline='') as newMacro:
            csvWriter = csv.writer(newMacro, delimiter=',')
            children = self.treeView.get_children()

            for child in children:
                childValues = self.treeView.item(child, 'values')
                # Do not include the Step column
                csvWriter.writerow(childValues[1:])


    def actionRelease(self, event):
        # check if key already entered into hold list and remove it duplicate
        if str(self.actionEntry.get())[0:1] == '_' and '|' in str(self.actionEntry.get()):
            # at least two entries exist
            keys = str(self.actionEntry.get())[1:str(self.actionEntry.get()).rfind('|')].split('|')
            recentKey = str(self.actionEntry.get())[str(self.actionEntry.get()).rfind('|') + 1:]

            if recentKey in keys:
                cleanActionEntry = str(self.actionEntry.get())[0:str(self.actionEntry.get()).rfind('|')]
                self.actionEntry.delete(0, END)
                self.actionEntry.insert(0, cleanActionEntry)


    def actionPopulate(self, event):
        self.actionEntry.icursor(len(self.actionEntry.get()))
        # _ is special character that should reset the self.actionEntry field to prevent all keys from adding to self.actionEntry
        if str(event.char) == '_' or (event.keysym == 'BackSpace' and str(self.actionEntry.get())[0:1] == '_'):
            # Clear self.actionEntry field
            self.actionEntry.delete(0, END)
            # Then allow _ to be typed

        # !, #, and _ is special character that allows typing of action instead of instating setting action to each key press
        if str(self.actionEntry.get())[0:1] != '!' and str(self.actionEntry.get())[0:1] != '#' and str(self.actionEntry.get())[0:1] != '*' and str(self.actionEntry.get())[0:1] != '_':
            # need to use different properties for getting key press for letters vs whitespace/ctrl/shift/alt
            # != ?? to exclude mouse button as their char and keysym are not empty but are edqual to ??
            if event.keysym == 'Escape' and event.char != '??':
                # special case for Escape because it has a char and might otherwise act like a letter but won't fill in the
                # box with 'Escape'
                self.actionEntry.delete(0, END)
                self.actionEntry.insert(0, event.keysym)
            elif str(event.char).strip() and event.char != '??':
                # clear entry before new char is entered
                if 96 <= event.keycode <= 105:
                    # Append NUM if keystroke comes from numpad
                    self.actionEntry.delete(0, END)
                    self.actionEntry.insert(0, 'NUM')
                else:
                    self.actionEntry.delete(0, END)
            elif event.keysym and event.char != '??':
                # clear entry and enter key string
                self.actionEntry.delete(0, END)
                self.actionEntry.insert(0, event.keysym)
            else:
                # clear entry and enter Mouse event
                if event.num != '??': 
                    self.actionEntry.delete(0, END)
                    self.actionEntry.insert(0, 'M' + str(event.num))

            # else key event is duplicate
        # _ is special character that needs to be followed by a key stroke
        if str(self.actionEntry.get())[0:1] == '_':
            if event.keysym == 'Escape' and event.char != '??':
                # special case for Escape because it has a char and might otherwise act like a letter but won't fill in the
                # box with 'Escape'
                if len(self.actionEntry.get()) > 1:
                    self.actionEntry.insert(len(self.actionEntry.get()), '|' + event.keysym)
                else:
                    self.actionEntry.insert(len(self.actionEntry.get()), event.keysym)

            elif str(event.char).strip() and event.char != '??':
                # clear entry before new char is entered
                if 96 <= event.keycode <= 105:
                    # Append NUM if keystroke comes from numpad
                    if len(self.actionEntry.get()) > 1:
                        self.actionEntry.insert(len(self.actionEntry.get()), '|NUM')
                    else:
                        self.actionEntry.insert(len(self.actionEntry.get()), 'NUM')
                elif len(self.actionEntry.get()) > 1:
                    # normal character
                    self.actionEntry.insert(len(self.actionEntry.get()), '|')

            elif event.keysym and event.char != '??':
                if event.keysym == 'space':
                    # delete the ' ' that was just typed as this uses 'space' as ' '
                    self.actionEntry.delete(len(self.actionEntry.get()) - 1)
                # clear entry and enter key string
                if len(self.actionEntry.get()) > 1:
                    self.actionEntry.insert(len(self.actionEntry.get()), '|' + event.keysym)
                elif event.keysym not in str(self.actionEntry.get()[:-1]):
                    self.actionEntry.insert(len(self.actionEntry.get()), event.keysym)
            else:
                # clear entry and enter Mouse event
                if len(self.actionEntry.get()) > 1:
                    self.actionEntry.insert(len(self.actionEntry.get()), '|M' + str(event.num))
                else:
                    self.actionEntry.insert(len(self.actionEntry.get()), 'M' + str(event.num))
            # else event is already entered, do not allow repeats in self.actionEntry
            
        # if str(self.actionEntry.get())[0:1] == '!':
        if event.delta == 120:
            if '!MScrlUp(' in self.actionEntry.get():
                try:
                    currScroll = int(self.actionEntry.get()[self.actionEntry.get().index('(') + 1:self.actionEntry.get().index(')')])
                    self.actionEntry.delete(0, END)
                    self.actionEntry.insert(0, '!MScrlUp(' + str(int((event.delta / 120) + currScroll)) + ')')
                except:
                    self.actionEntry.delete(0, END)
                    self.actionEntry.insert(0, '!MScrlUp(' + str(int(event.delta / 120)) + ')')
            else:
                self.actionEntry.delete(0, END)
                self.actionEntry.insert(0, '!MScrlUp(' + str(int(event.delta / 120)) + ')')
        elif event.delta == -120:
            if '!MScrlDown(' in self.actionEntry.get():
                try:
                    currScroll = int(self.actionEntry.get()[self.actionEntry.get().index('(') + 1:self.actionEntry.get().index(')')])
                    self.actionEntry.delete(0, END)
                    self.actionEntry.insert(0, '!MScrlDown(' + str(int((event.delta / 120) + currScroll)) + ')')
                except:
                    self.actionEntry.delete(0, END)
                    self.actionEntry.insert(0, '!MScrlDown(' + str(int(event.delta / 120)) + ')')
            else:
                self.actionEntry.delete(0, END)
                self.actionEntry.insert(0, '!MScrlDown(' + str(int(event.delta / 120)) + ')')


    def actionPaste(self, event):
        # clear actionEntry
        self.actionEntry.delete(0, END)
        # apply paste function first
        self.actionEntry.insert(0, '!Paste(')
        self.after(10, self.actionPasteClose)

    def actionPasteClose(self):
        # let paste apply and then add closing
        self.actionEntry.insert(END, ')')
    
    
    def overwriteRows(self):  # what row and column was clicked on
        rows = self.treeView.selection()
        for row in rows:
            # Edit everything except Step
            self.treeView.set(row, "X", self.x.get())
            self.treeView.set(row, "Y", self.y.get())
            self.treeView.set(row, "Action", self.actionEntry.get())
            self.treeView.set(row, "Delay", self.delayEntry.get())
            self.treeView.set(row, "Comment", self.commentEntry.get())
        self.exportMacro()


    def overwriteRow(self, rowNum, xPos, yPos, action, delay, comment):
        rows = self.treeView.get_children()
        i = 0

        for row in rows:
            if i == rowNum:
                self.treeView.item(row, values=(rowNum, xPos, yPos, action, delay, comment))
                break
            i += 1


    def closeTab(self):
        self.exportMacro()
        if len(self.notebook.tabs()) > 1:
            self.treeTabs.pop(self.notebook.tab(self.notebook.select(), "text"))
            self.notebook.forget("current")

            config = cp.ConfigParser()
            config.read(os.path.join(FILE_PATH, 'config.ini'))
            config.set('Tabs', 'openTabs',  str('|'.join(self.treeTabs)))
            with open(os.path.join(FILE_PATH, 'config.ini'), 'w') as configfile:
                config.write(configfile)


    def toggleBusy(self):
        try:
            config = cp.ConfigParser()
            config.read(os.path.join(FILE_PATH, 'config.ini'))
            config.set('Settings', 'busyWait', str(self.busyWait.get()))
            with open(os.path.join(FILE_PATH, 'config.ini'), 'w') as configfile:
                config.write(configfile)
        except:
            pass


    def toggleHidden(self):
        try:
            config = cp.ConfigParser()
            config.read(os.path.join(FILE_PATH, 'config.ini'))
            config.set('Settings', 'hiddenMode', str(self.hiddenMode.get()))
            with open(os.path.join(FILE_PATH, 'config.ini'), 'w') as configfile:
                config.write(configfile)
        except:
            pass


    def toggleStartFromSelected(self):
        try:
            config = cp.ConfigParser()
            config.read(os.path.join(FILE_PATH, 'config.ini'))
            config.set('Settings', 'startfromselected', str(self.startFromSelected.get()))
            with open(os.path.join(FILE_PATH, 'config.ini'), 'w') as configfile:
                config.write(configfile)
        except:
            pass


    def toggleLoopsByMacro(self):
        if self.loopsByMacro.get() == 1:
            try:
                if self.currTab in self.macroLoops:
                    self.loopEntry.set(self.macroLoops[self.currTab])
                else:
                    self.loopEntry.set(self.globalLoops.get())
                    # print("set tabs to  " + self.macroLoops[self.currTab])
            except:
                pass
        else: 
            self.loopEntry.set(self.globalLoops.get())
        try:
            config = cp.ConfigParser()
            config.read(os.path.join(FILE_PATH, 'config.ini'))
            config.set('Settings', 'loopsbymacro', str(self.loopsByMacro.get()))
            with open(os.path.join(FILE_PATH, 'config.ini'), 'w') as configfile:
                config.write(configfile)
        except:
            pass


    def toggleStayOnTop(self):
        try:
            if self.stayOnTop.get() == 0:
                self.parent.attributes("-topmost", False)
            else:
                self.parent.attributes("-topmost", True)
            config = cp.ConfigParser()
            config.read(os.path.join(FILE_PATH, 'config.ini'))
            config.set('Settings', 'stayontop', str(self.stayOnTop.get()))
            with open(os.path.join(FILE_PATH, 'config.ini'), 'w') as configfile:
                config.write(configfile)
            config = cp.ConfigParser()
            config.read(os.path.join(FILE_PATH, 'config.ini'))
            config.set('Settings', 'stayontop', str(self.stayOnTop.get()))
            with open(os.path.join(FILE_PATH, 'config.ini'), 'w') as configfile:
                config.write(configfile)
        except:
            pass
        try:
            if self.settingsWindow is not None:
                if self.stayOnTop.get() == 0:
                    self.settingsWindow.settingsWindow.attributes("-topmost", False)
                else:
                    self.settingsWindow.settingsWindow.attributes("-topmost", True)
        except:
            pass
        try:
            if self.logWindow != None:
                if self.stayOnTop.get() == 0:
                    self.logWindow.logWindow.attributes("-topmost", False)
                else:
                    self.logWindow.logWindow.attributes("-topmost", True)
        except:
            pass
        try:
            if self.helpWindow != None:
                if self.stayOnTop.get() == 0:
                    self.helpWindow.helpWindow.attributes("-topmost", False)
                else:
                    self.helpWindow.helpWindow.attributes("-topmost", True)
        except:
            pass


    def openLogWindow(self):
        if self.logWindow is None:
            self.logWindow = logWindow(self.parent, self, "\n".join(self.clickLog))
            try:
                if self.logWindow is not None:
                    if self.stayOnTop.get() == 0:
                        self.logWindow.logWindow.attributes("-topmost", False)
                    else:
                        self.logWindow.logWindow.attributes("-topmost", True)
            except:
                pass
        else:
            self.logWindow.onClose()
            self.logWindow = None


    def openSettingsWindow(self):
        if self.settingsWindow is None:
            self.settingsWindow = settingsWindow(self.parent, self)
            try:
                if self.settingsWindow is not None:
                    if self.stayOnTop.get() == 0:
                        self.settingsWindow.settingsWindow.attributes("-topmost", False)
                    else:
                        self.settingsWindow.settingsWindow.attributes("-topmost", True)
            except:
                pass
        else:
            self.settingsWindow.onClose()
            self.settingsWindow = None


    def windowFinder(self):
        self.titlebar.minimize_window()
        if self.settingsWindow is not None:
            self.settingsWindow.titlebar.minimize_window()
        if self.helpWindow is not None:
            self.helpWindow.titlebar.minimize_window()
        if self.logWindow is not None:
            self.logWindow.titlebar.minimize_window()
        mouse.on_click(self.getWindow)


    def getWindow(self):
        # unhook mouse listener after getting window that was clicked
        try:
            self.window = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())
            print(psutil.Process(self.window[-1]).name())
        finally:
            self.titlebar.deminimize()
            mouse.unhook_all()

        config = cp.ConfigParser()
        config.read(os.path.join(FILE_PATH, 'config.ini'))
        config.set('Settings', 'selectedApp', str(self.selectedApp))
        with open(os.path.join(FILE_PATH, 'config.ini'), 'w') as configfile:
            config.write(configfile)


    def onClose(self):
        self.stopClicking()
        if self.recorder is not None: self.recorder.stopRecordingThread()

        config = cp.ConfigParser()
        config.read(os.path.join(FILE_PATH, 'config.ini'))
        config.set('Position', 'x', str(self.parent.winfo_rootx()))
        config.set('Position', 'y', str(self.parent.winfo_rooty()))
        with open(os.path.join(FILE_PATH, 'config.ini'), 'w') as configfile:
            config.write(configfile)
        self.destroy()
        self.parent.destroy()

                
    def checkNumerical(self, key):
        return (key in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'] or key.isnumeric())


    def cleanseActionEntry(self, a, b, c):
        action = self.actionEntry.get()
        action = action.replace(' ', '')
        self.actionEntry.delete(0, END)
        self.actionEntry.insert(0, action)


    def updateLoops(self, var, index, mode):
        try:
            config = cp.ConfigParser()
            config.read(os.path.join(FILE_PATH, 'config.ini'))
            config.set('Settings', 'loops', str(self.loopEntry.get()))
            with open(os.path.join(FILE_PATH, 'config.ini'), 'w') as configfile:
                config.write(configfile)
        except:
            pass



class rightClickMenu():
    def __init__(self, tree):
        self.rightClickMenu = Menu(tree, tearoff=0)
        self.rightClickMenu.add_command(label="Move up", command=self.moveUp)
        self.rightClickMenu.add_command(label="Move down", command=self.moveDown)
        self.rightClickMenu.add_command(label="Remove", command=self.removeRow)
        self.rightClickMenu.add_command(label="Select All", command=self.selectAll)
        self.rightClickMenu.add_separator()
        self.rightClickMenu.add_command(label="New Macro", command=self.newMacro)
        self.rightClickMenu.add_command(label="Close Macro", command=self.closeTab)
        self.tree = tree

    def showRightClickMenu(self, event):
        try:
            self.rightClickMenu.tk_popup(event.x_root + 0, event.y_root + 0, 0)
        finally:
            self.rightClickMenu.grab_release()

    def newMacro(self):
        name = simpledialog.askstring("Input", "New Macro Name", parent=self.tree)
        if str(name).strip() and name:
            self.tree.addTab(str(name).strip())

    def moveUp(self):
        selectedRows = self.tree.treeView.selection()
        for row in selectedRows:
            self.tree.treeView.move(row, self.tree.treeView.parent(row), self.tree.treeView.index(row) - 1)
        self.tree.reorderRows()
        self.tree.tagSelection()
        self.tree.exportMacro()

    def moveDown(self):
        selectedRows = self.tree.treeView.selection()
        for row in reversed(selectedRows):
            self.tree.treeView.move(row, self.tree.treeView.parent(row), self.tree.treeView.index(row) + 1)
        self.tree.reorderRows()
        self.tree.tagSelection()
        self.tree.exportMacro()

    def removeRow(self):
        selectedRows = self.tree.treeView.selection()
        for row in selectedRows:
            self.tree.treeView.delete(row)
        self.tree.reorderRows()
        self.tree.tagSelection()
        self.tree.exportMacro()

    def selectAll(self):
        for row in self.tree.treeView.get_children():
            self.tree.treeView.selection_add(row)
            self.tree.treeView.item(row, tag='selected')

    def closeTab(self):
        self.tree.closeTab()


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



class Recorder:
    # Constants for key conversion
    # pynput needed for hooking
    # pyautogui better for
    _keyDict = {
        'Key.shift': 'Shift_L',
        'Key.shift_r': 'Shift_R',
        str(keyboard.Key.ctrl_l): 'Control_L',
        str(keyboard.Key.ctrl_r): 'Control_R',
        str(keyboard.Key.tab): 'Tab',
        str(keyboard.Key.caps_lock): 'Caps_Lock',
        str(keyboard.KeyCode.from_vk(96)): 'NUM0',
        str(keyboard.KeyCode.from_vk(97)): 'NUM1',
        str(keyboard.KeyCode.from_vk(98)): 'NUM2',
        str(keyboard.KeyCode.from_vk(99)): 'NUM3',
        str(keyboard.KeyCode.from_vk(100)): 'NUM4',
        str(keyboard.KeyCode.from_vk(101)): 'NUM5',
        str(keyboard.KeyCode.from_vk(102)): 'NUM6',
        str(keyboard.KeyCode.from_vk(103)): 'NUM7',
        str(keyboard.KeyCode.from_vk(104)): 'NUM8',
        str(keyboard.KeyCode.from_vk(105)): 'NUM9',
        'Key.enter': 'Return',
        'Key.space': 'space',
        'Key.f1': 'F1',
        'Key.f2': 'F2',
        'Key.f3': 'F3',
        'Key.f4': 'F4',
        'Key.f5': 'F5',
        'Key.f6': 'F6',
        'Key.f7': 'F7',
        'Key.f8': 'F8',
        'Key.f9': 'F9',
        'Key.f10': 'F10',
        'Key.f11': 'F11',
        'Key.f12': 'F12',
        'Key.scroll_lock': 'Scroll_Lock',
    }

    def __init__(self, tree):
        # print(self._keyDict)
        self.start = None
        self.startPress = None
        self.thread = None
        self.pressed = []
        self.keycode = keyboard.KeyCode
        self.recording = False
        self.lastRow = []

        self.tree = tree

    def startRecordingThread(self):
        self.recording = True

        # create new thread for recording so as to not disturb the autoclicker window
        threadRecordingFlag = threading.Event()
        self.thread = threading.Thread(target=self.record)
        self.thread.threadFlag = threadRecordingFlag
        self.thread.start()

    def stopRecordingThread(self):
        # print("stopRecordingThread")
        self.tree.startButton.config(state=NORMAL)
        self.tree.recordButton.configure(style="")

        if self.thread:
            self.thread.threadFlag.set()
            self.thread.join(1)
            self.listener.stop
 
        self.recording = False
        self.tree.exportMacro()

    def record(self):
        # start listening
        self.tree.recordButton.configure(style="Accent.TButton")
        # self.listener = Listener(on_press=self.__recordPress, on_release=self.__recordRelease)
        with Listener(on_press=self.__recordPress, on_release=self.__recordRelease) as listener:
            try:
                self.listener = listener
                self.listener.join()
            except Exception as ex:
                print('{0} was pressed'.format(ex.args[0]))
        # self.listener.start()

    def __recordPress(self, key):
        # log time ASAP for accuracy
        tempTime = time.time()
        # if self.startPress is not None:
        #     print(int((time.time() - self.startPress) * 1000))

        if self.thread.threadFlag.is_set():
            return False

        # if key.char exists use that, else translate to pyautogui keys
        try:
            if key.char is not None:
                key = key.char
            else:
                try:
                    key = self._keyDict[str(key)]
                except KeyError:
                    key = str(key)
        except AttributeError:
            try:
                key = self._keyDict[str(key)]
            except KeyError:
                key = str(key)

        # ignore if key already pressed
        if key in self.pressed:
            return

        # if startTime is set then this is not first key press in recording
        # must fill in wait time between

        # only add row on press if more than one key now being pressed, thus ending the prior action
        if len(self.pressed) > 0:
            # print("New press: " + self.pressed)
            # other keys are already pressed
            # if int((time.time() - self.startPress) * 1000) < 1:
            # key was pressed close to prior press, merge with prior
            # p = 0 # TODO edit prior row adding this time to that delay and skip this new press
            # else:
            self.__addRow(0, 0, "_" + "|".join(self.pressed), int((time.time() - self.startPress) * 1000))
        elif self.startPress is not None:
            # if startTime is set then this is not first key press in recording
            # must fill in wait time between last press and new press

            if self.lastRow[2][0] == '_':
                # last row was hold so add row to account for delay
                self.__addRow(0, 0, "", int((time.time() - self.startPress) * 1000))
            else:
                # last row was not hold so edit last row to change delay
                # print("change: " + str(len(self.treeView.get_children())))
                # print(int((time.time() - self.startPress) * 1000))
                self.__changeRow(len(self.treeView.get_children()) - 1, 0, 0, self.lastRow[2], int((time.time() - self.startPress) * 1000))

        # key is being pressed, add to array and log time
        self.startPress = tempTime
        self.pressed.append(str(key))
        # print("{0}Down ".format(str(format(key))))

    def __recordRelease(self, key):
        # check how long key was pressed ASAP
        pressTime = int((time.time() - self.startPress) * 1000)
        # start new timer for next action's delay
        self.startPress = time.time()

        if self.thread.threadFlag.is_set():
            return False

        # if key.char exists use that, else translate to pyautogui keys
        try:
            if key.char is not None:
                key = key.char
            else:
                try:
                    key = self._keyDict[str(key)]
                except KeyError:
                    key = str(key)
        except AttributeError:
            try:
                key = self._keyDict[str(key)]
            except KeyError:
                key = str(key)

        i = 0
        for keyPressed in self.pressed:
            if key == keyPressed:
                # if pressTime < 35:
                #     # Short press, consider it not a hold
                #     self.__addRow(0, 0, key, pressTime)
                # else:
                # Longer press, consider it a hold
                self.__addRow(0, 0, "_" + "|".join(self.pressed), pressTime)
                self.pressed.pop(i)
            i += 1

        # print("{0} Release".format(key))

    def __addRow(self, x, y, key, delay):
        self.lastRow = [x, y, key, delay]
        self.tree.addRowWithParams(x, y, key, delay, "")

    def __changeRow(self, row, x, y, key, delay):
        self.lastRow = [x, y, key, delay]
        self.tree.overwriteRow(row, x, y, key, delay, "")


def main():
    root = tk.Tk()
    root.tk.call('source', r'C:\Users\cgkna\Documents\GitHub\StickysAutoClicker\StickysAutoClicker\sun-valley.tcl')
    root.tk.call("set_theme", "dark")

    global STICKY_ICON
    STICKY_ICON = PhotoImage(data="""iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAav0lEQVR4nM2baZBlR3Xnf5l517e/qldrV/Wm3iQ1jSQkRgqEjIVYjG1ghsVsihFgYNhhiGExwyAWeVg8BsGAzRgckmMIYwwOC9lWgMEDZpdBC2hrdbda6r22V2+97y65zIeSCAESoquLifl/qS91zz3n906ezDyZV/D/kQ79YCAbgd8K+9k+uTyIgnIjbxf62OpK+57dV221v4l3ijM1cPy6G2ZLQe1ZyvDvCESgm6VkEAQ3rGT6m+c/48L017Vz0/W3Tz6+OvaOscS8TB5bGKM79GQY50UUpcNq9M1TZvDWvW+66L4z9fcXdUYA7v3o9f9h3p/9sMjCrb4Jpav4iK0T0tRMerzf++bBfvEfn/aSJyw+lp0/fdufB2dVWzduOtm5fMYoL233qWpFIBVRtYYbq9lsx2x7MN146z291S9c8con6DPx++Hy1vPQl979/mBvefNVzY7/MXn8eEkUPtYKVKOEbfeRY0E01YyuWC6ya4BXPZY93U3euXz3rU/JVkfeUuCj8IhzS8UPmCsyGqNMuiJvTbXO/5yuNceAT6zH70eSXM9DW1qzVzUz/9rw8KiUrQ5J+x3y4TLJqaMUDxxDLCZ4Cx2vOui+6PoPffHiX2XrP736I9N5Z/U1WbcT9IuMPCvQ6YiRLVg2Q452l0iyBG+Qwz0ngpZR19xy3R3PXV+4v6zTBvD1T3x2+0w4/b5yW0aBUTihSLOUhaTHkc4yw6Vleg88gG23mfPCyoRnX/krHUj0LjFMxiQS4TTapihnccKgnWVgCwYmR2qDvO8k8W2HK9uN/NTtn7nlvPWH/bD3n+4Dc7J07VRRmvYzQeAHlMOYyPcprKGdDVhIV2kPVhh0VqA3pG7tsz79vmsrj2ZPZEUg8tSzxQjjUhKXkaOxWKR1FLrgZHeFY6snWW4fY7h/P96td89uNvKGH37mx/vOLPzTrAHf/OC1V022xTO7hw6RLPYZC0sgBU5JAt/H0x4Dm7Lc7nJscIyJrM/4zsnJ8ci/Avj7R7IpjR4TaSLzQQdjLYVUCK+gGpUInI9F0tUj+p2MqgpoBCUm0dSCcPNZ52798v4vHblk9/M3L//GAXz29e9ulIr4fekDHa97ZIUCS5oP6eQjrLEYp+nkffouZWG4SqEKopPHuLh0idfYUn/OowHwdfEEL0ul73k4NNppunqAKSw1FRPLCBQYLH1TMEiHdLKEmSJncqy2YxT6nweesV4Av/YQmKpNfKB/78rsyeMrdIqcrstYyruMdEYhDEY4jILjwxVGsiDBUGo0WT6+SMWLL/701Z9+xHfFo/TiknPsOu88xs/ZSb/u0fcKOsWQgUkp0GANwlrA4KwmVYbcpWRHjzKT6yvu+8K9b/mNAvj4q/9oX3pi+JJiOfEGeUbuCqzWOCuwQjAyBb1ixOJwlQGalQD2XXope3/rYpJUE59MttdK/hMfyXaUpbviuE5zfo753Xu46ClPpbJjllU/Z1EPSDE4Y5CAA6QSOM+jawray4uYQ/fLCSnfc/t1P5n7jQGQxr1Nr4xqRZqDEniBj5QS3w+IlI/nBKkwnNBD0kjx5CuezvTcVoznEzUaLN93zKs5+YgFyx+NKrVKDedJIi+iETW4YO+FzJx3Dp3YsZx3GUiNFgLhezhfge/jwpCoUsb1EuJ+MtaoVD76GwHwgZe9OfIG9vKwEF4YR3hBgPMVMgoohAMJyhO0TcKSLLj4qZczOTaJlCFO+lQnWgxWO1Ks9C76RdtvfMbrdgRF4cVxiAg8LCCcJPAr7N62h02P28lCkNmhLMgDsJ6EwAPfw/qKdjJE5wa3OqAVhs/+1v/+6QUbDqAZ1y6bL01MV4KIqFQirpYJHgQRhiEiUKQ+LKU9zt33OGbHpvCdAgRCKMJajdxo3HDwS/N2qLOKs4ZCOlKrQUqclFgHsfPZMXcW4eyEPWJ6eSY12hM4T2GkoADKzTGMrxBOETgRbZmceP2GA5CFeLJMnRRKIKQAKYniiEatRuQHlMsV+r5h6Dt27d6NcwKDQUiLAqySEPpEnpz9RduBcRIhrfR8rDEIY3BYEAaLI3AhO886W+pGnJy0/ST3wHoKpxTC80ApomYVKQOklXIsiF/0zb89dFoLpMcEYAv5xCCIZaVapdkco1qtUI7LRH7EZGuCII44knTZtvccfOWBs1hrsUbj0GsBxQFpmtb+5uqrN//82x1OWKw1hNJDCsAUWGfAOgRQjWI5Wa3Tj8SRoWc1gY+KI7xKCXwfEZcg8nC6oFItlaYnx1+7oQCatfqFcVRmamKSWr1OqVohimKaYy1qrXH8RoVekTI9M4N0IIx5cI9pEdIitKZRq2OGeVQvN3c93Lb2lcUZCQZrDYXVSOdwzuGsQxoDOseMchmUK99bdunBrBSgYx+vWqHZmsFrjkGzjgtCcIKJ8fglt36384gzzmkD+Py7/sfv7dm8rbZlbo7xiSnGWy1qzQa1iQmiRhMd+JzM+zicDTwP5ywGwDmwBmM0VjqUUghtEEb/nGNCKYuzGGsY6RztLMZYnHU4Ac5aPCSD7iCoVBqf7yr1mkVRLKtKTKs1RdQcx5VjTBjg/AirDc2mX5lrxdd+41tHar8OgF+5EpwZm/6Dmfq01xBVlJTINFtLU6fAWXwXcejQdxMh8ciLABVgpEMai8BhhMP6PiKOyBaHckp5T/7aJ/+sNVjsdUbDjMP3HKkgpHVIPM/HWAcCpBDw4AwjNCg/kqpc9d79V+//lw+98L3vmKzWP6VKQcR0E1EOUaUYF/m4EuhhTrPpP3FLOH4N8MbHAvCoGfDX13y2ddbWc35vYnwGWasiWy3cRIui2YCZKZge51jWy48mq7dFUdAzaYYAnBQ4KXFCIlBIqxiP6szFU3JHY9elU97ETXtnd/x4d2vTj8+f2/a5kl8JZCooaY8SPjiwzqGtxQLKU1TGxqSsVj2Ad37xfX95b3/hmrtWjuRDUkw5glaEGFfImsRrBMhQyOlm/Lqbf3Ds6evOgJ3z25/fKrcaTiv8VhMRBTirUdohpKS9cNTefPAntxTKfADc55xea9lZZ3HCYZwlKCRutU/QTdmzbR++P10Za0QXkoy4/+h+fnpwyVrryfaJFdLBLcS1mIltM8hahHNgjEE7Q7laliv97s98e8v1/+2Dn3zth3ZmP7Ev2nbWOcF0MUNYi/HrJQQKYaEUStmMwjcDX1sXACnUc8J6ExlHawuUQCA8HzfQdNtL9lu3fuO2xXTpVYVMjwhlbWEtBRaJJMgdveVVFg6fotLJmZ3ajK8FspORL3a49+57uefAITr9RAoLKRa7OiDt9ChWVpnfNkcwN8HIV5gixRcWmWcXAv/0kH/twL12f//Ygfad7VeOHyzP1ksVL6yU8MIA6QfoOMxP6PQxe4iPCmBqYnIfDtJuH2cydJ7S7bftyqnF5P4Th79ztH30jW+57sMHAT74Oy9jlCVUi4IIj/aBY5w6+ABVEbBpep7x6UkA3KkOo5VlThw4zGAwoOo8jHCMnCZzlsBYgqGGxYTQrmImQrISFHlKEanw4f6999p3JcAHr/2jD/+vjs4vriTDrV7hTSTZaHvox+204Nahzr+wbgBHjx1+1/2r+99ns6xVmFGeFcmy1uktg1zfuLTS/ru3XP+hn3V8HSAyQ5Q7Th44xMkjpwiRnL17D2dt2oIsDBiL7SV4I7M2XXqK0EgyZyg5RW4NOTkniozB6iJ7Sh6NUwW64lFklsIXj9gWf/Mfv2MR+MpjBXraAC7+w2f91efe+qF/Ula3NKYzyIb5W/78I+1H+t9iZG1qR9yzsp9ed4iSinN37GTP2WcjUrtW0UcajME6h9MWz0kQkkCAcALPSZx0aClZynN2VGJmt29nc7VBs93mnxfuydYb5LoAALzyY+9cBh6z22I06XI+xAIuDJlSIWfPzEM3weUOk2o8DSYr6A/6COsoCQ+jJMJqtAMnFBUXEpoCoRXlVDI+vxWxfZ4gGckn3BNNbFTQD9e62uK/KH+sfv9oJHb4qZW+hrlSFb+XIiRQWEgyhPSxWpMMEyQQSoF1DisVClDGoJF40ifzclRucYs9TLVNWIuYCIN17fcfSxsCINfmrsDIK0pWEBpBzYXYQU7P5iihENoiPcEwSzB2rQZEvk+hNQaHcQ4pJdJajBQIqSiSIWaxjackuqaYrQe/9vL2dLQhAJwnrRbWKo2sGp+lfgetC/zC0ojK1IRP4qDjclKdEQYBRoLG4QFKCBzgKYmxlth5uCxjuHCK/ORxVquOycc3Nuw06OHaEAAq9ALZS6SvLVZYhtZhkoLIKfomI/B8fCTWWYQTa0ODtT0CzuEAaw1SOiSghaNrEr5w13fYn6ww02jwktbFNfeJj+wSb3r7vRvh80PaGABZcU61QA7tiH9tP0BJhWwpj7OzMs24dTjtyB0IKZBIpJBou7brQzhwDikdzlkkDjB8Y+kAX+seQjfKPNWFpMOi5MrVK4ANBbCuo7FflDV6LnKSoYLelio3h6v8bfsublj4KZnQ4AxOOhwW4yzaOqyxOAdrvz8I55AInICTWYd/6N7HsicRQQk53QKvEukiXHf7+9F0xgDe/uLX7g20bAgnMcbRaI3jYp+BNNzcPcr3+kcwcu3XXtsjgsVROIt2BmMc2jkKKSikw+G4bbhEc+dWauUqLisoAp/YryL9+cuyv7xu12M6dRo6YwAqp6YKUTJIHIJSHFOtVRmbHONpV/57fth5gAQDFnLjKCxoIRnhGAhInWOoNYM8J9cGYyydUR/Pc5Qmy5TKipX2AsZTYFo1IyffuhGBP6QzBuAL/8LQeIFEUhE+wciwd+t25mYmueOWH7GQ9zhZDOhZzcBoMufAKYT28ExA5hQFiqG2LBc5pyhokzM4foJqUTAdBuwcSVp+CDqTIqq96OZPXt/YiOBhA4pgYL3dfqG90ChCVbL3HDpMMRVKKTTxsCDxQ5bRlJxBIBFxSG1+HsIQV4qIMdjBgOX77kf3uxR5xtnxFLsDQ1VETAYVnrP1QqQoI6ngpa7RrEz8IfAnGxD/mV+RueZ33nh3uWt21bSHz1pHeKXoym6eIJBUg5jQSiIZ4AUxuy97ErI+hlUBQnmMigyT57hsxH23/IjixHE8BJeddTbbpjdBXMUvlVHNGOdZRDmgHQ4W7+4cOPfJb371ug9FH9IZDYH3v+xt5yhjJyVInJXOWhk4IWe9JrvKm9hZnmFG1qirMkL5yMkmSTmmp6DRqrNr6zR7ts0ifEEHR2nLPK5WRviKbDgkXV7GdTqwtEJ66Cj66Ar66AJNI1u7p6a//JPrr9tzpgDUeh+87g0fvqBY7H5KDvXOUCOUcXgClIDQCSIUoRA4a8iEo7x5E629eylCwergKLK8QDTVpise4NADd5HnmnqjRaUxTrq6Sksp8mTI4soSw+4qytMsLt6Pj8MVVlTKY/NjY7MvfvXzft//2N/+9bfXG8dp14CvfujT02Ny7LX33nHfVaKXz/oaqYxFybV0UkiUsEi51u/XwqIrEc2dWzCBR1l61Gotup2T3H1kQFEUVMMqzfI0mfbo10qM7dhKvtJjub1KkuYEheTo0WWa9TFcb0CoBLXQl2XUWD8f/T7wx//PAPzo326dq+alq+yqnhOukLJQSOeAB1d2AEqCs4DFSkFlagoX+AQ2J0BQ8iLq3ma6nS55ljNeHsdqD2s1SmjcWJXBMME5i8aS5jlZltPNNXHSZS7axKn9i4z3p7g3XfjOeoNfF4DC447OqJ+UnSI0AuXWzvakE4QPtrXd2h+0cFjnKDUra8EpUJ7FSYEvfOqiTColgZBkoiBSkEiJBlalgTwhsAbpBM4JhvmQjklITkq6yYCVu36a9qbifzsTAKddBK/+m79IRyb9Si7S3NoCi8ZhcE5jncEJg3UG8+BJj3EaEQi01GjPYnSBLHJ0MSKxGStuwMj0cXZE5oZIlyHR5NIwzEcUdu2ozAqHwVHkBfcfOcKR7qJe8vIfdEbJI948+XW1rnVALvVHR35xHlo+JbQy8A1gJUJ4CMfagYhzaGswArI8IbQRRWHIdYryPZIAgt0zNEtzdE4u4h5YxhWWvCjIjSEzI7JiRBiEOGcpcDgk1hQULrddsrtMXHr5J//PF/MzAbCuWeD7B+9MnrTnnK/nirPBbJPGKM9ZpHAIZ9EYnDM4AbkpYLxCWApQOgc0OQXVvVuw58xiSiH+VIP82BJ0E4ZFhjaaQbeDa3fxpcA4TWYLMp3Rtxl93yzlsXjVtd+98bYzCX7dAAC+ffCuwYW7z/mixpSwZh/WhGspb3HYB/f3CuscedkjrsUYmyGweNaRCoOabJBnKVmni9l/FD1KSPIEk+e0Ty3g9RIUkOYpI5MxsCmrvmkPK967PvHDm758psHDBqwEAd7xlGc/W2XFfy/lYkesVeAZTSh8FD6ZM/TrIdNn70T4gtAX1KRCBgFiagwnFdmgj1jtk2YZvVFKbzhiZf8DlIYZsSfQRUEfZ0e+O5KWvCs/9eN/OaPK/3CtOwMeru/ev3//RTt3fsVKVI7Zqa2JCmNFbnJyUvpmiC0JhMoR2pI5S2E0ttMl766S9jok2YBe2mc1WWWpc4p8cQlP54x0Rk+6pBfLrwxD9/rP3PKtmzfC54e0IRnwkF5x1aujWpp+VXWSS30tpMtyMBkJkNbqiEaNzUrQ8BxYgRIhwmoKa8itY5CnLKwu0V1Zxh9ppIwYejGmWvnCl35w44s30teHtCEtsYc0tWP+mb4Ul/mlCBX66DRHFpZAe9x67ymOrIy4m5TtomBWGUJ/CNaSFgWjomCYaUaFIxjfzsj3WAA6QZWg0ajxgxs30tWfaUOGwEN6zgt+V6dZ9+npcLVlij7ODYkiQ71R5vhSm+UiYKAiTiSGpd6AMIxJVEDHWRKlKI3PUW3uoK2qHMSwGpTIS2OoSrxlan7rB0/tP+Oi/0vaUAD//I9fa+/aNHHncCSvWk4Vy+0hBw+f5NBCm3YiKFyItAZjNYkxpDKkMVEiNW2G6RL9wTId53E40ySeRHkhngxwhVaejG4/eef37tlIf2GDa8AbnveivRXlf9ukQaO8ZyflmXHSSLDYHXLLHYfJXBnrLMnKEmbQZXaqwaW7a5RCyPM+o5UjHOsa7hu26OUK6wX45RqlSgNf6sVWxZv5+49fvaHfDm1YDXjL5c+f9Dvi836vaFRHOa57gOBJitq5m6lPbuLmO4+gmjWssUijcTYniELGGk22bFY43yc91WNiYUh+os4dxxK8cpV4fBJZqpClyWQ/GV4OfH2jfIYzBPCpqz/CA7cc2BH2syu9QfGGIBNjXi7X+gLtDHO8DUFAeH4DTxoCCnpZgjIjlBAMegNkVFD1ekzUDaujIZGKOE6Fg4sjECmhlyKlwOoh5tTilR9+6X+p1Tx7XyT0bS+/7tozBrDuGvD2p7xS5guD16hu8Y9RV/92I1NxxSpiBB4OaSU4STjXImxWUb4ltCuU6DAeG2Ym61jl2LX7cUw3I6LYsdoRrI7qnOw16PYHlHyHHSzTP3GQusuY7uSPv8BW/+C3d533inK1fvbvnXfBj75w87e6j+3to+u0a8DVL3gBtlO/VA7MW8tp+FxfF9IzhkCAJyy+s+AcufTtKPBk7XmXoAMI9m4jxyGN4s4Dh/nx7fdwfKGNcJoXPOMiznvcdvpJm5/ccpSvfu8n5AWUmyW2z82xb99eLtq9lfTbtyFuup3L95xP5ezNdFR6/92nDv3dQn/5vc//n+8frAfAr50Bb3vmy73LNl9wjhpV3ldK1Z/WRvJxJW1EIEA6h4fDl7B2m1ZgPU9oJ+j1ujTmpwkDj8p4g6/98Ba+9dMD9FLQBoQQ3HHnYY4cPoQ1lpu+eRuCCl5YI9Wwon361Qa7Nk/x5LO2c/xfbyXqjBjXIeH0fGNmdv6SMAxe+NInX/GNJ+07a+mGH/7wtAA8Zg34z899RUMZ+Qpv6K5spMG+KBeyrCWxA4GmoEAoCVY/+FEDSKUohMAZhz65SnHbYcpWYYzjty65iFRKDtx7kiKI0Dh0PKLZHOPxtV3ct8dw4NQyBHXC0MPzFTVh2Ts3R9RJ8LRjadSldeIUY3EVN15l56Y9O9q19q0K+2bgz08HwK/MgM+/5t1XlF30N6W2vXJ84E1XCyHKVlISoIRBKIdyDk85hDA/G09aKvpKYnxF0ymiMKI+MU2aF5hkyK75GfyyYLCyyMiGeGHIBdt3cFYqmLtgF3ecOIFWJUqhx1RoedK+7TxR+MRRmePfv52s1yHKCiJt8bsjpFF41leT01NPf+lvX/aEF152yd3X/cvXF9YN4Efv+bN973nGiz87R+O/hstmVi/0ibUgFAIfhy/AokG4NQPuwa7Q2nUphkpAEJIDM9UG5fEmJrVEYw3iSoxeaTO3c465+UkG/QGDwSqXPn4fmwaa5lQdWalyz6lFJmohz7rsQrZ2DdsoU46rTKYeq0sLdLorjIqcwGm8YYoaaaTz1PjElj2lSvzSP7jsifZ1L37m9z9zw1fdaQH48tv++NIdYuIfJpf98+NTmdddWCUpUpDgIwmEACxi7ZQTgaEwOYXJMdZghSKPQoSTeNJjvDVOFJfwlI8/zBG5pTI1iVzoU2vWOPfcrYSxYOfkJqrLGbbQzO/bgglSLt+3i/OEj7/co/m4c4iXLZXmGIGQ3H1wPzrNKVi7VF0pwLUTlIgp11ths157qtW69NKnXXLzX9z09Ue9YPVzAP7kyjddurc2f+N8W42zPKDX69JJ++SmwBMC5RyGtT6focChsbYgcQWZyRBKYoKYoXMYBxP1BvXxFjKK8D0f5Ss8o3D9gnK5TJgbdG/Etm2bKBlJ3AY/l8SNMpvGKkSH24zbiF1bd9DIa3hagjWUalU6J47RWVmg0GuZZ4qcUCrsYIgbasKRL2rNySeV4+Cpr/3dK7708RtvfMQPuX+uCFZz71X143lDZpAMhnRGQ1Jr8KTE4SiExpoCYzXWaZwzjIqUxBUMdALCw1iQYY1ASqqNOmEUWRWF2lNeIDyFVCFOKopujsoU4xM1sqUcl2YErH1vkO4/QT3WDGo1ji902J7X8eamEdLhhERFEeeefx4n7jvAQCfYoSHJU9ppQi2sUu61adYnrVxuyfJs6YmioV/Ho5wd/F8RJJ9Psj/+dwAAAABJRU5ErkJggg==""")
    STICKY_ICON = STICKY_ICON.subsample(2)

    global RUNNING_IMAGE
    RUNNING_IMAGE = PhotoImage(data="""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/GMzLdsAADv2aVRYdFhNTDpjb20uYWRvYmUueG1wAAAAAAA8P3hwYWNrZXQgYmVnaW49Iu+7vyIgaWQ9Ilc1TTBNcENlaGlIenJlU3pOVGN6a2M5ZCI/Pgo8eDp4bXBtZXRhIHhtbG5zOng9ImFkb2JlOm5zOm1ldGEvIiB4OnhtcHRrPSJBZG9iZSBYTVAgQ29yZSA5LjAtYzAwMCA3OS4xNzFjMjdmLCAyMDIyLzA4LzE2LTE4OjAyOjQzICAgICAgICAiPgogICA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPgogICAgICA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIgogICAgICAgICAgICB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iCiAgICAgICAgICAgIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyIKICAgICAgICAgICAgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iCiAgICAgICAgICAgIHhtbG5zOnN0RXZ0PSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VFdmVudCMiCiAgICAgICAgICAgIHhtbG5zOnBob3Rvc2hvcD0iaHR0cDovL25zLmFkb2JlLmNvbS9waG90b3Nob3AvMS4wLyIKICAgICAgICAgICAgeG1sbnM6dGlmZj0iaHR0cDovL25zLmFkb2JlLmNvbS90aWZmLzEuMC8iCiAgICAgICAgICAgIHhtbG5zOmV4aWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20vZXhpZi8xLjAvIj4KICAgICAgICAgPHhtcDpDcmVhdG9yVG9vbD5BZG9iZSBQaG90b3Nob3AgRWxlbWVudHMgMjQuMCAoV2luZG93cyk8L3htcDpDcmVhdG9yVG9vbD4KICAgICAgICAgPHhtcDpDcmVhdGVEYXRlPjIwMjQtMDctMTBUMjI6MjM6MzgtMDc6MDA8L3htcDpDcmVhdGVEYXRlPgogICAgICAgICA8eG1wOk1ldGFkYXRhRGF0ZT4yMDI0LTA3LTEwVDIyOjI3OjEwLTA3OjAwPC94bXA6TWV0YWRhdGFEYXRlPgogICAgICAgICA8eG1wOk1vZGlmeURhdGU+MjAyNC0wNy0xMFQyMjoyNzoxMC0wNzowMDwveG1wOk1vZGlmeURhdGU+CiAgICAgICAgIDxkYzpmb3JtYXQ+aW1hZ2UvcG5nPC9kYzpmb3JtYXQ+CiAgICAgICAgIDx4bXBNTTpJbnN0YW5jZUlEPnhtcC5paWQ6ZDRjNzI4MzktNGJkYy03ZTRiLWFkODctYTA1YTViMGFlZGJhPC94bXBNTTpJbnN0YW5jZUlEPgogICAgICAgICA8eG1wTU06RG9jdW1lbnRJRD54bXAuZGlkOjc3YzYxZmIyLTQ4ZmItY2M0OC1iMTFhLTJhNGJkMjE5NjBmMDwveG1wTU06RG9jdW1lbnRJRD4KICAgICAgICAgPHhtcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD54bXAuZGlkOjc3YzYxZmIyLTQ4ZmItY2M0OC1iMTFhLTJhNGJkMjE5NjBmMDwveG1wTU06T3JpZ2luYWxEb2N1bWVudElEPgogICAgICAgICA8eG1wTU06SGlzdG9yeT4KICAgICAgICAgICAgPHJkZjpTZXE+CiAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6YWN0aW9uPmNyZWF0ZWQ8L3N0RXZ0OmFjdGlvbj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0Omluc3RhbmNlSUQ+eG1wLmlpZDo3N2M2MWZiMi00OGZiLWNjNDgtYjExYS0yYTRiZDIxOTYwZjA8L3N0RXZ0Omluc3RhbmNlSUQ+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDp3aGVuPjIwMjQtMDctMTBUMjI6MjM6MzgtMDc6MDA8L3N0RXZ0OndoZW4+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDpzb2Z0d2FyZUFnZW50PkFkb2JlIFBob3Rvc2hvcCBFbGVtZW50cyAyNC4wIChXaW5kb3dzKTwvc3RFdnQ6c29mdHdhcmVBZ2VudD4KICAgICAgICAgICAgICAgPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6YWN0aW9uPnNhdmVkPC9zdEV2dDphY3Rpb24+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDppbnN0YW5jZUlEPnhtcC5paWQ6YTZmNDU5ZjAtZWMyYS0wMzQ2LTkzMjgtMTUyMGE3NTkxZWZjPC9zdEV2dDppbnN0YW5jZUlEPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6d2hlbj4yMDI0LTA3LTEwVDIyOjI0OjU4LTA3OjAwPC9zdEV2dDp3aGVuPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6c29mdHdhcmVBZ2VudD5BZG9iZSBQaG90b3Nob3AgRWxlbWVudHMgMjQuMCAoV2luZG93cyk8L3N0RXZ0OnNvZnR3YXJlQWdlbnQ+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDpjaGFuZ2VkPi88L3N0RXZ0OmNoYW5nZWQ+CiAgICAgICAgICAgICAgIDwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OmFjdGlvbj5zYXZlZDwvc3RFdnQ6YWN0aW9uPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6aW5zdGFuY2VJRD54bXAuaWlkOmQ0YzcyODM5LTRiZGMtN2U0Yi1hZDg3LWEwNWE1YjBhZWRiYTwvc3RFdnQ6aW5zdGFuY2VJRD4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OndoZW4+MjAyNC0wNy0xMFQyMjoyNzoxMC0wNzowMDwvc3RFdnQ6d2hlbj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0OnNvZnR3YXJlQWdlbnQ+QWRvYmUgUGhvdG9zaG9wIEVsZW1lbnRzIDI0LjAgKFdpbmRvd3MpPC9zdEV2dDpzb2Z0d2FyZUFnZW50PgogICAgICAgICAgICAgICAgICA8c3RFdnQ6Y2hhbmdlZD4vPC9zdEV2dDpjaGFuZ2VkPgogICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgPC9yZGY6U2VxPgogICAgICAgICA8L3htcE1NOkhpc3Rvcnk+CiAgICAgICAgIDxwaG90b3Nob3A6Q29sb3JNb2RlPjM8L3Bob3Rvc2hvcDpDb2xvck1vZGU+CiAgICAgICAgIDxwaG90b3Nob3A6SUNDUHJvZmlsZT5zUkdCIElFQzYxOTY2LTIuMTwvcGhvdG9zaG9wOklDQ1Byb2ZpbGU+CiAgICAgICAgIDx0aWZmOk9yaWVudGF0aW9uPjE8L3RpZmY6T3JpZW50YXRpb24+CiAgICAgICAgIDx0aWZmOlhSZXNvbHV0aW9uPjMwMDAwMDAvMTAwMDA8L3RpZmY6WFJlc29sdXRpb24+CiAgICAgICAgIDx0aWZmOllSZXNvbHV0aW9uPjMwMDAwMDAvMTAwMDA8L3RpZmY6WVJlc29sdXRpb24+CiAgICAgICAgIDx0aWZmOlJlc29sdXRpb25Vbml0PjI8L3RpZmY6UmVzb2x1dGlvblVuaXQ+CiAgICAgICAgIDxleGlmOkNvbG9yU3BhY2U+MTwvZXhpZjpDb2xvclNwYWNlPgogICAgICAgICA8ZXhpZjpQaXhlbFhEaW1lbnNpb24+NTwvZXhpZjpQaXhlbFhEaW1lbnNpb24+CiAgICAgICAgIDxleGlmOlBpeGVsWURpbWVuc2lvbj4xNTwvZXhpZjpQaXhlbFlEaW1lbnNpb24+CiAgICAgIDwvcmRmOkRlc2NyaXB0aW9uPgogICA8L3JkZjpSREY+CjwveDp4bXBtZXRhPgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgIAo8P3hwYWNrZXQgZW5kPSJ3Ij8+yXzqMQAAACBjSFJNAAB6JQAAgIMAAPn/AACA6AAAdTAAAOpgAAA6lwAAF2+XqZnUAAAAOklEQVR42rSPsQ0AIAzDnIoH6f+/hAEkEOoI3uwpkW1ImY3kzhkACAr+x5iqeqcPTwHQlrx+pKtpDAAy8xIDQSyqLwAAAABJRU5ErkJggg==""")
    
    global NONE_IMAGE
    NONE_IMAGE = PhotoImage(data="""""")

    root.overrideredirect(True)
    root.minsize(NORMAL_SIZE[0], NORMAL_SIZE[1])

    Tn = treeviewNotebook(root)

    Tn.pack(expand=True, fill="both")
    root.update_idletasks()
    root.withdraw()

    root.mainloop()


if __name__ == "__main__":
    main()