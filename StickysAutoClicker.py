import os
import time
import csv
import mouse
import mss
import pyautogui
import threading
import win32process
import win32gui
import psutil
from pynput import keyboard
from pynput.keyboard import Key, Listener
from os.path import exists

import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import simpledialog
from tkinter.messagebox import askyesno

from PIL import ImageGrab
from functools import partial
import configparser as cp

ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)
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

build with "pyinstaller --onefile --noconsole --icon=StickyHeadIcon.ico StickysAutoClicker.py"
"""

usabilityNotes = ("                 Usability Notes\n\n"
                  " - Selecting any row of a macro will auto-removeulate the X and Y positions, delay and action fields with the values of that row, overwriting anything previously entered.\n"
                  " - Shift + Ctrl + ` will set loops to 0 and stop the autoclicker immediately.\n"
                  " - Overwrite will set all selected rows to the current X, Y, delay and action values.\n"
                  " - Macros are exported as csv files that are writes to folder '\Macros' in exe directory.\n"
                  " - Exported Files are kept up to date with each edit created.\n"
                  " - Import expects a comma separated file with no headers and only four columns: X, Y, Action, Delay\n"
                  " - Macros do not need to be in a tab to be called by another macro. The csv in the Macros folder will be used, so be sure that is up to date.\n"
                  " - Action will recognize the main three mouse clicks as well as any keyboard action including Shift + keys (doesn't work with Alt or Ctrl + keys)\n"
                  " - Key presses do not move cursor, so X and Y positions do not matter.\n"
                  " - Action has two escape characters, ! and #, that will allow the user to continue typing rather than overwriting the action key.\n"
                  "           This is to allow for calling another macro (!) or finding an image (#)\n"
                  " - Action also allows underscore _ as a special character that will indicate the following key should be pressed and held for the set amount of time in the Delay field.\n"
                  "           Note that for these rows the Delay no longer delays after the key is held. \n"
                  " - Typing !macroName into action and adding that row will make a macro look for another tab in Sticky's Autoclicker with the tab name macroName and execute that macro.\n"
                  "           Delay serves another purpose when used with a macro action (!macroName) and will call that macro for the delay amount of times\n"
                  " - Typing #imageName into action and adding that row will make a macro look for a .png image with that name in the \Images folder and move the cursor to a found image and left click.\n"
                  "           Delay serves even another purpose when used with a find image action (#imageName) and serve as the confidence percentage level.\n"
                  "           Confidence of 100 will find an exact match and confidence of 1 will find the first roughly similar image.\n"
                  "           If image is not found then loop will end and next loop will start.\n"
                  " - Finding an image will try 5 times with a .1 second delay if not found and click once found.\n")

root = Tk()
global_monitor_left = 0
global_recording = False

# define global monitors
with mss.mss() as sct:
    if sct.monitors:
        global_monitor_left = sct.monitors[0].get('left')


class treeviewNotebook:
    notebook = None
    recorder = None
    window = None

    # Array for currently pressed keys so they can be unpressed when stopping clicking
    currPressed = []
    busyWait = tk.IntVar()
    hiddenMode = tk.IntVar()
    startFromSelected = tk.IntVar()
    selectedApp = ''

    settingsWindow = None
    helpWindow = None

    # store clicking thread for reference, needed to pass event.set to loop stops while waiting.
    # setting event lets the sleep of the clicking thread get interrupted and stop the thread
    # Otherwise starting clicking while thread is sleeping makes two threads clicking and this reference would be overwritten
    activeThread = None
    # allow treeview array to be referenced for macro looping and export
    treeTabs = []
    treeView = None

    # class with notebook with tabs of treeviews
    def __init__(self):
        self.initElements()
        self.initTab()
        self.loadSettings()

    def initElements(self):
        # Column 0
        # for reference: notebook.grid(row=3, column=2, columnspan=3, rowspan=5, sticky='')
        root.columnconfigure(0, weight=6)
        self.welcomeLabel = Label(root, text="Sticky's Autoclicker", font=("Arial bold", 14), bg=bgColor)
        self.welcomeLabel.grid(row=0, column=0, padx=5, pady=0, sticky='n', columnspan=2)
        self.helpButton = Button(root, text="Help", bg=bgButton, command=showHelp, borderwidth=2)
        self.helpButton.grid(row=0, column=0, padx=20, sticky='s', columnspan=2)
        self.madeByLabel = Label(root, text="Made by Sticky", bg=bgColor)
        self.madeByLabel.grid(row=0, column=0, padx=20, sticky='', columnspan=2)
        self.clickLabel = Label(root, text="Click Loops", font=("Arial", 10), bg=bgColor)
        self.clickLabel.grid(row=1, column=0, sticky='e', pady=10)
        self.LoopsLeftLabel = Label(root, text="Loops Left", font=("Arial", 11), bg=bgColor)
        self.LoopsLeftLabel.grid(row=2, column=0, sticky='s')
        self.clicksLeftLabel = Label(root, textvariable=loopsLeft, bg=bgColor)
        self.clicksLeftLabel.grid(row=3, column=0, sticky='n')
        self.startButton = Button(root, text="Start Clicking", font=("Arial", 15), command=threadStartClicking, padx=10, pady=10, borderwidth=6, bg=bgButton)
        self.startButton.grid(row=4, column=0, columnspan=2, sticky="n")
        self.stopButton = Button(root, text="Stop Clicking", font=("Arial", 15), command=stopClicking, padx=0, pady=10, borderwidth=6, bg=bgButton)
        self.stopButton.grid(row=5, column=0, columnspan=2, sticky="n")

        self.importMacroButton = Button(root, text="Import Macro", bg=bgButton, command=importMacro, borderwidth=2)
        self.importMacroButton.grid(row=6, column=0, columnspan=2, sticky='n')
        self.settingsButton = Button(root, text="Settings", bg=bgButton, command=showSettings, borderwidth=2)
        self.settingsButton.grid(row=7, column=0, columnspan=1, sticky='n')
        self.recordButton = Button(root, text="Record", bg=bgButton, command=self.toggleRecording, borderwidth=2)
        self.recordButton.grid(row=7, column=1, columnspan=1, sticky='n')

        # Column 1
        root.columnconfigure(1, weight=4)
        loopEntry.grid(row=1, column=1, sticky='', pady=10)
        self.loopsLabel = Label(root, text="Total Loops", font=("Arial", 11), bg=bgColor)
        self.loopsLabel.grid(row=2, column=1, sticky='s')
        self.loops2Label = Label(root, textvariable=loops, bg=bgColor)
        self.loops2Label.grid(row=3, column=1, sticky='n')

        # Column 2
        root.columnconfigure(2, weight=10)
        self.insertPositionButton = Button(root, text="  Insert Position   ", bg=bgButton, command=addRow, padx=0, pady=0,
                                           borderwidth=2)
        self.insertPositionButton.grid(row=0, column=2, pady=6, sticky='n')
        self.getCursorButton = Button(root, text=" Choose Position ", bg=bgButton, command=getCursorPosition, borderwidth=2)
        self.getCursorButton.grid(row=0, column=2, sticky='s')
        self.editRowButton = Button(root, text="Overwrite Row(s)", bg=bgButton, command=overwriteRows, borderwidth=2)
        self.editRowButton.grid(row=1, column=2, sticky='')

        # Column 3
        root.columnconfigure(3, weight=10)
        self.xPosTitleLabel = Label(root, text="X position", font=("Arial", 11), bg=bgColor)
        self.xPosTitleLabel.grid(row=0, column=3, pady=0, sticky='nw')
        self.xPosLabel = Label(root, textvariable=x, bg=bgColor)
        self.xPosLabel.grid(row=0, column=3, padx=35, pady=25, sticky='w')
        self.yPosTitleLabel = Label(root, text="Y Position", font=("Arial", 11), bg=bgColor)
        self.yPosTitleLabel.grid(row=0, column=3, padx=0, pady=0, sticky='sw')
        self.yPosLabel = Label(root, textvariable=y, bg=bgColor)
        self.yPosLabel.grid(row=1, column=3, padx=35, pady=0, sticky='nw')

        # Column 4
        root.columnconfigure(4, weight=10)
        self.timeLabel = Label(root, text="Delay (ms)", font=("Arial", 10), bg=bgColor)
        self.timeLabel.grid(row=0, column=4, padx=0, pady=0, sticky="nw")
        delayEntry.grid(row=0, column=4, padx=0, pady=5, sticky='ne')

        self.actionLabel = Label(root, text="Action", font=("Arial", 10), bg=bgColor)
        self.actionLabel.grid(row=0, column=4, padx=0, sticky="sw")
        actionEntry.grid(row=0, column=4, padx=0, pady=5, sticky='se')

        # bind this entry to all keyboard and mouse actions
        actionEntry.bind("<Key>", actionPopulate)
        actionEntry.bind("<KeyRelease>", actionRelease)
        actionEntry.bind('<Return>', actionPopulate, add='+')
        actionEntry.bind('<KeyRelease-Return>', actionRelease, add='+')
        actionEntry.bind('<Escape>', actionPopulate, add='+')
        actionEntry.bind('<KeyRelease-Escape>', actionRelease, add='+')
        actionEntry.bind('<Button-1>', actionPopulate, add='+')
        actionEntry.bind('<ButtonRelease-1>', actionRelease, add='+')
        actionEntry.bind('<Button-2>', actionPopulate, add='+')
        actionEntry.bind('<ButtonRelease-2>', actionRelease, add='+')
        actionEntry.bind('<Button-3>', actionPopulate, add='+')
        actionEntry.bind('<ButtonRelease-3>', actionRelease, add='+')

        root.protocol("WM_DELETE_WINDOW", destroy)

        self.commentLabel = Label(root, text="Comment", font=("Arial", 10), bg=bgColor)
        self.commentLabel.grid(row=1, column=4, padx=0, sticky="w")
        commentEntry.grid(row=1, column=4, padx=0, pady=5, ipadx=20, sticky='e')

        # Column 5
        root.columnconfigure(5, weight=10)

        # Row weights
        root.rowconfigure(0, weight=10)
        root.rowconfigure(1, weight=10)
        root.rowconfigure(2, weight=10)
        root.rowconfigure(3, weight=10)
        root.rowconfigure(4, weight=10)
        root.rowconfigure(5, weight=10)
        root.rowconfigure(6, weight=10)
        root.rowconfigure(7, weight=10)
        # Make horizontal scrollbar normal sized
        root.rowconfigure(8, weight=0)

        delayToolTip = CreateToolTip(delayEntry,
                                     "Delay in milliseconds to take place after click. If macro is specified this will be loop amount of that macro.")
        actionToolTip = CreateToolTip(actionEntry,
                                      "Action to occur before delay. Accepts all mouse buttons and keystrokes. Type !macroname to call macro in another tab with delay amount as loops.")

        self.stopButton.config(state=DISABLED)

    def initTab(self):
        self.canvas = tk.Canvas(root, bg=bgColor, highlightthickness=0)
        self.canvas.config(height=0)
        self.frame = tk.Frame(self.canvas, bg=bgColor)
        self.frame.grid(row=2, column=2, columnspan=8, sticky="s")

        self.horizontalTabScroll = tk.Scrollbar(root, orient="horizontal", command=self.canvas.xview)
        self.horizontalTabScroll.grid(row=8, column=2, columnspan=4, rowspan=1, sticky="nsew")

        self.canvas.configure(xscrollcommand=self.horizontalTabScroll.set)
        self.canvas.grid(row=2, column=2, columnspan=4, sticky="nsew")
        self.canvas.create_window((0, 0), window=self.frame, anchor="nw", tags="frame")

        self.frame.bind("<Configure>", self.frame_configure)

        self.notebook = ttk.Notebook(self.frame)
        self.tabFrame = tk.Frame(self.notebook, bg=bgColor)
        self.notebook.add(self.tabFrame, text='macro1')
        self.notebook.grid(row=2, column=2, columnspan=8, rowspan=1)
        self.notebook.bind('<<NotebookTabChanged>>', self.tabRefresh)

        self.treeFrame = tk.Frame(root, bg=bgColor)
        self.treeFrame.grid(row=2, column=2, columnspan=6, rowspan=6, sticky="nsew", pady=(20, 0))
        # treeviewStyle = ttk.Style()
        # treeviewStyle.theme_use("default")
        # treeviewStyle.configure("treeviewStyle.style", background="white")
        # treeviewStyle.map("treeviewStyle.style")
        self.treeView = ttk.Treeview(self.treeFrame, selectmode="extended") #, style="treeviewStyle.style")
        self.treeView['show'] = 'headings'
        self.treeTabs.append('macro1')

        self.verticalTabScroll = Scrollbar(self.treeFrame, orient='vertical', command=self.treeView.yview)
        self.verticalTabScroll.pack(side=RIGHT, fill=Y)
        self.treeView.configure(yscrollcommand=self.verticalTabScroll.set)

        self.treeView['columns'] = ("Step", "X", "Y", "Action", "Delay", "Comment")

        self.treeView.column('Step', anchor='c', width=35, stretch=False)
        self.treeView.column('X', anchor='center', minwidth=60, stretch=False)
        self.treeView.column('Y', anchor='center', width=60, stretch=False)
        self.treeView.column('Action', anchor='center', width=100, stretch=False)
        self.treeView.column('Delay', anchor='center', width=80, stretch=False)
        self.treeView.column('Comment', anchor='center', width=120, stretch=False)

        self.treeView.heading('Step', text='Step')
        self.treeView.heading('X', text='X')
        self.treeView.heading('Y', text='Y')
        self.treeView.heading('Action', text='Action')
        self.treeView.heading('Delay', text='Delay/Repeat')
        self.treeView.heading('Comment', text='Comment')
        self.treeView.pack(fill=BOTH, expand=True)
        # self.treeView.grid(row=0, column=0, columnspan=4, rowspan=4)
        self.treeView.tag_configure('oddrow', background="red")
        self.treeView.tag_configure('evenrow', background="white")

        self.treeView.bind("<Button-3>", rCM.showRightClickMenu)
        self.treeView.bind("<ButtonRelease-1>", selectRow)

    def loadSettings(self):
        if os.path.exists(os.path.expanduser('~/Documents/StickysAutoClicker/config.ini')):
            config = cp.ConfigParser()
            config.read(os.path.expanduser('~/Documents/StickysAutoClicker/config.ini'))
            print(config.sections())
            print(os.path.expanduser('~/Documents/StickysAutoClicker/config.ini'))
        else:
            config = cp.ConfigParser()
            config['Settings'] = {'busyWait': int(self.busyWait.get()), 'startFromSelected': int(self.startFromSelected.get()),  'selectedApp': str(self.selectedApp), 'hiddenMode': int(self.hiddenMode.get())}
            config['Tabs'] = {'openTabs': str('|'.join(self.treeTabs))}
            with open(os.path.expanduser('~/Documents/StickysAutoClicker/config.ini'), 'w') as configfile:
                config.write(configfile)


    def tabRefresh(self, event):
        print('tabRefresh')
        # clear treeview of old macro
        for item in self.treeView.get_children():
            self.treeView.delete(item)

        tab = event.widget.tab('current')['text']

        # import csv file into macro where filename will be macro name, is it best to import only exported macros
        filename = os.path.join(os.getcwd() + r'\Macros', tab + '.csv')
        fileExists = exists(filename)

        if fileExists:
            with open(filename, 'r') as csvFile:
                csvReader = csv.reader(csvFile)
                if not filename:
                    tN.addTab(os.path.splitext(os.path.basename(tab))[0])
                step = 1
                for line in csvReader:
                    # backwards compatibility to before there were comments
                    if len(line) > 4:
                        addRowWithParams(line[0], line[1], line[2], line[3], line[4])
                    elif line[3]:
                        addRowWithParams(line[0], line[1], line[2], line[3], "")
                    step += 1

    def frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def addTab(self, name):
        self.tabFrame = ttk.Frame(self.notebook)
        self.tabFrame.config(height=320, width=440)

        self.treeTabs.append(name)
        tabCount = len(self.treeTabs)

        config = cp.ConfigParser()
        config['Tabs'] = {'openTabs': str('|'.join(self.treeTabs))}
        with open(os.path.expanduser('~/Documents/StickysAutoClicker/config.ini'), 'w') as configfile:
            config.write(configfile)

        self.notebook.add(self.tabFrame, text=name)
        self.notebook.select(self.tabFrame)

    def getNotebook(self):
        return self.notebook

    def toggleRecording(self):
        if self.recorder is not None:
            self.recorder.stopRecordingThread()
            del self.recorder
            self.recorder = None
        else:
            self.recorder = Recorder()
            self.recorder.startRecordingThread()


class rightClickMenu():
    def __init__(self):
        self.rightClickMenu = Menu(root, tearoff=0)
        self.rightClickMenu.add_command(label="Move up", command=self.moveUp)
        self.rightClickMenu.add_command(label="Move down", command=self.moveDown)
        self.rightClickMenu.add_command(label="Remove", command=self.removeRow)
        self.rightClickMenu.add_command(label="Select All", command=self.selectAll)
        self.rightClickMenu.add_separator()
        self.rightClickMenu.add_command(label="New Macro", command=self.newMacro)
        self.rightClickMenu.add_command(label="Close Macro", command=self.closeTab)

    def showRightClickMenu(self, event):
        try:
            self.rightClickMenu.tk_popup(event.x_root + 50, event.y_root + 10, 0)
        finally:
            self.rightClickMenu.grab_release()

    def newMacro(self):
        name = simpledialog.askstring("Input", "New Macro Name", parent=root)
        if str(name).strip() and name:
            tN.addTab(str(name).strip())

    def moveUp(self):
        selectedRows = tN.treeView.selection()
        for row in selectedRows:
            tN.treeView.move(row, tN.treeView.parent(row), tN.treeView.index(row) - 1)
        reorderRows(tN.treeView)
        exportMacro()

    def moveDown(self):
        selectedRows = tN.treeView.selection()
        for row in reversed(selectedRows):
            tN.treeView.move(row, tN.treeView.parent(row), tN.treeView.index(row) + 1)
        reorderRows(tN.treeView)
        exportMacro()

    def removeRow(self):
        selectedRows = tN.treeView.selection()
        for row in selectedRows:
            tN.treeView.delete(row)
        reorderRows(tN.treeView)
        exportMacro()

    def selectAll(self):
        for row in tN.treeView.get_children():
            tN.treeView.selection_add(row)

    def closeTab(self):
        exportMacro()
        if len(tN.getNotebook().tabs()) > 1:
            del tN.treeTabs[tN.getNotebook().index(tN.getNotebook().select())]
            tN.getNotebook().forget("current")
            config = cp.ConfigParser()
            config['Tabs'] = {'openTabs': str('|'.join(self.treeTabs))}
            with open(os.path.expanduser('~/Documents/StickysAutoClicker/config.ini'), 'w') as configfile:
                config.write(configfile)


class CreateToolTip(object):
    # create a tooltip for a given widget
    def __init__(self, widget, text='widget info'):
        self.waittime = 500  # miliseconds
        self.wraplength = 180  # pixels
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None

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
        xPos = yPos = 0
        xPos, yPos, cx, cy = self.widget.bbox("insert")
        xPos += self.widget.winfo_rootx() + 25
        yPos += self.widget.winfo_rooty() + 25
        # creates a toplevel window
        self.tw = Toplevel(root)

        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (xPos, yPos))
        self.tw.wm_attributes("-topmost", 1)
        label = ttk.Label(self.tw, text=self.text, justify='left',
                          background="#ffffff", relief='solid', borderwidth=1,
                          wraplength=self.wraplength)
        label.pack(ipadx=1)
        root.update()

    def hidetip(self):
        tw = self.tw
        self.tw = None
        if tw:
            tw.destroy()


bgColor = None  # 'CadetBlue1'   '#9ac6e3'
fgColor = None  # '#759fba'
bgButton = None  # '#e0f3ff'


root.title('Fancy Autoclicker')
root.geometry("680x385")
# force window on top
root.attributes("-topmost", True)
root.configure(bg=bgColor)


def checkNumerical(key):
    if key in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'] or key.isnumeric():
        return True
    else:
        return False


def resizeNotebook(self):
    # attempt to scale grid items to resized window
    winWidth = root.grid_bbox(column=2, row=3, col2=7, row2=8)[2] - 18
    winWeight = 14

    tN.treeView.column('#0', width=int((winWidth / winWeight)))  # 1
    tN.treeView.column('X', width=int((winWidth / winWeight) * 1.5))  # 2
    tN.treeView.column('Y', width=int((winWidth / winWeight) * 1.5))  # 2
    tN.treeView.column('Action', width=int((winWidth / winWeight) * 3))  # 3
    tN.treeView.column('Delay', width=int((winWidth / winWeight) * 3))  # 3
    tN.treeView.column('Comment', width=int((winWidth / winWeight) * 4))  # 4


def cleanseActionEntry(a, b, c):
    action = actionEntry.get()
    action = action.replace(' ', '')
    actionEntry.delete(0, END)
    actionEntry.insert(0, action)


vcmd = (root.register(checkNumerical), '%S')
loopEntry = Entry(root, width=5, justify="right", validate='key', vcmd=vcmd)
loopEntry.insert(0, 1)
delayEntry = Entry(root, width=12, justify="right", validate='key', vcmd=vcmd)
delayEntry.insert(0, 0)
actionVar = StringVar()
actionEntry = Entry(root, width=10, justify="center", textvariable=actionVar)
actionEntry.insert(0, 'M1')
actionVar.trace('w', cleanseActionEntry)
commentEntry = Entry(root, width=10, justify="right")
commentEntry.insert(0, '')
rCM = rightClickMenu()
loopsLeft = IntVar()
loopsLeft.set(0)
loops = IntVar()
loops.set(0)
x = IntVar()
x.set(0)
y = IntVar()
y.set(0)
previouslySelectedTab = None
previouslySelectedRow = None


# Start clicking helper to multi thread click loop, otherwise sleep will make windows want to kill
def threadStartClicking():
    # save current macro to csv, easier to read, helps update retention
    exportMacro()

    tN.startButton.config(state=DISABLED)
    tN.stopButton.config(state=NORMAL)

    try:
        clickArray = list(csv.reader(open(os.getcwd() + r'\Macros' + '\\' + str(
            tN.getNotebook().tab(tN.getNotebook().select(), 'text').replace(r'/', r'-').replace(r'\\', r'-').replace(
                r'*', r'-').replace(r'?', r'-').replace(r'[', r'-').replace(r']', r'-').replace(r'<', r'-').replace(
                r'>', r'-').replace(r'|', r'-')) + '.csv', mode='r')))

        # set loops counter
        loops.set(int(loopEntry.get()))
        loopsLeft.set(loopEntry.get())
        root.update()

        # create event and give as attribute to thread object, this can be referenced in loop and prevent sleeping of thread
        # using wait instead of sleep lets thread stop when Stop Clicking is clicked while waiting allow the thread to be quickly killed
        # because if Start Clicking is clicked before thread is killed this will overwrite saved thread and prevent setting event with intent to kill thread
        threadFlag = threading.Event()
        if tN.busyWait == 1:
            tN.activeThread = threading.Thread(target=startClickingBusy,
                                           args=(clickArray, int(loopEntry.get()), True))
        else:
            tN.activeThread = threading.Thread(target=startClicking,
                                           args=(clickArray, threadFlag, int(loopEntry.get()), True))
        tN.activeThread.threadFlag = threadFlag
        tN.activeThread.start()
    except:
        # hmmmm
        print("Error: unable to start clicking thread")

    # Emergency Exit key combo that will stop auto clicking in case mouse is moving to fast to click stop
    try:
        sE = threading.Thread(target=monitorExit, args=())
        sE.start()
    except:
        # hmmmmm
        print("Error: unable to start exit monitoring  thread")


def monitorExit():
    def for_canonical(f):
        return lambda k: f(l.canonical(k))
    # TODO make work even if other keys are pressed
    hotkey = keyboard.HotKey(
        keyboard.HotKey.parse('<ctrl>+<Shift>+`'),
        stopClicking)

    with keyboard.Listener(
            on_press=for_canonical(hotkey.press)) as l:
        l.join()


# Start Clicking button will loop through active notebook tab's treeview and move mouse, call macro, search for image and click or type keystrokes with specified delays
# Intentionally uses busy wait for most accurate delays
def startClickingBusy(clickArray, loopsParam, mainLoop):
    # loop though param treeview, for param loops
    # needed so as to not mess with LoopsLeft from outer loop
    root.update()
    startTime = time.time()

    if loopsParam == 0 or loopsLeft.get() == 0: return

    clickTime = 0
    prevDelay = 0
    blnDelay = False
    firstLoop = True
    selection = tN.treeView.selection()
    startFromSelected = tN.startFromSelected.get()
    if len(selection) > 0:
        selectedRow = tN.treeView.item(selection[0]).get("values")[0]
    else:
        selectedRow = 0

    # check Loopsleft as well to make sure Stop button wasn't pressed since this doesn't ues a global for loop count
    while loopsParam > 0 and loopsLeft.get() > 0:
        startTime = time.time()

        print("New loop")
        for row in clickArray:
            if loopsParam == 0 or loopsLeft.get() == 0: return

            # When start from selected row setting is true then find highlighted row(s) and skip to from first selected row
            # Only for first loop and first macro, not for subsequent loops nor macros called by the starting macro
            if firstLoop and mainLoop and startFromSelected == 1 and selectedRow > 0 and intLoop < selectedRow:
                continue

            # check row to see if its still holding keys
            if len(tN.currPressed) > 0 and (row[2] == "" or row[2][0] == '_'):
                # stop time if next action is not hold
                toBePressed = []
                if row[2][0] == '_':
                    toBePressed = row[2][1:].split('|')

                # key hold and release is action
                # print('PreRelease', (prevDelay / 1000), time.time() - clickTime)
                while time.time() < clickTime + (int(prevDelay) / 1000):
                    pass
                blnDelay = False
                # print('PreReleasePostWait', (prevDelay / 1000), time.time() - clickTime)
                # Release all keys not pressed in next step
                for key in tN.currPressed:
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
                            tN.currPressed.remove(key)
                        except:
                            pass
                        # print("Release " + str(key) + " after " + str(time.time() - clickTime))
                        print("Release " + str(key) + " after total " + str(time.time() - startTime))
                        print("Release1 " + str(key) + " after " + str(int(prevDelay) / 1000) + " with time diff: " + str(time.time() - clickTime - int(prevDelay) / 1000))
                # if this row is also a press then start timer as hols is continuing
                # clickTime = time.time()
            elif len(tN.currPressed) > 0:
                # print('PreRelease', str(pressed), (int(row[3]) / 1000), time.time() - clickTime)
                # if blnDelay:
                while time.time() < clickTime + (int(prevDelay) / 1000):
                    pass
                # blnDelay = False
                # print('PreReleaseWait', (int(row[3]) / 1000), time.time() - clickTime)
                # Release all keys as next step is not pressing
                for key in tN.currPressed:
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
                        tN.currPressed.remove(key)
                    except:
                        pass
                        # print("Release " + str(key) + " after " + str(time.time() - clickTime))
                    print("Release " + str(key) + " after total " + str(time.time() - startTime))
                    print("Release2 " + str(key) + " after " + str(int(prevDelay) / 1000) + " with time diff: " + str(time.time() - clickTime - int(prevDelay) / 1000))
                # if this row is not a press then wait to start timer until next click
                clickTime = time.time()

            # Empty must be first because other references to first character of Action will error
            if row[2] == "":
                # Nothing is just a pause
                print("blank", (int(row[3]) / 1000), time.time() - clickTime)
                if blnDelay:
                    while time.time() < clickTime + (int(prevDelay) / 1000):
                        pass
                clickTime = time.time()
            # Moved _ hold action as close top as possible as it's accuracy is most important
            # Action is starts with an underscore (_), hold the key for the given amount of time
            elif row[2][0] == '_' and len(row[2]) > 1:
                # loop until end of string of keys to press
                toPress = row[2][1:].split('|')

                # wait prior row amount before pressing new keys
                if blnDelay:
                    while time.time() < clickTime + (int(prevDelay) / 1000):
                        pass

                for key in toPress:
                    # Do not press if already pressed
                    if key not in tN.currPressed:
                        if key in ['M1', 'M2', 'M3']:
                            if int(row[0]) != 0 or int(row[1]) != 0:
                                # mouse click is action
                                if key == 'M1':
                                    pyautogui.mouseDown(int(row[0]), int(row[1]), button='left')
                                    tN.currPressed.append(key)
                                elif key == 'M3':
                                    pyautogui.moveTo(int(row[0]), int(row[1]))
                                    pyautogui.mouseDown(button='right')
                                    tN.currPressed.append(key)
                                else:
                                    pyautogui.mouseDown(int(row[0]), int(row[1]), button='middle')
                                    tN.currPressed.append(key)
                            else:
                                # mouse click is action without position, do not move, just click
                                if key == 'M1':
                                    pyautogui.mouseDown(button='left')
                                    tN.currPressed.append(key)
                                elif key == 'M3':
                                    pyautogui.mouseDown(button='right')
                                    tN.currPressed.append(key)
                                else:
                                    pyautogui.mouseDown(button='middle')
                                    tN.currPressed.append(key)
                        else:
                            # key press is action
                            if key == 'space':
                                pyautogui.keyDown(' ')
                                tN.currPressed.append(key)
                            elif key == 'tab':
                                pyautogui.keyDown('\t')
                                tN.currPressed.append(key)
                            else:
                                pyautogui.keyDown(key)
                                tN.currPressed.append(key)
                        print("Press " + str(key) + " at " + str(time.time() - startTime))
                clickTime = time.time()

            elif row[2] in ['M1', 'M2', 'M3']:
                if int(row[0]) != 0 or int(row[1]) != 0:
                    # mouse click is action
                    if row[2] == 'M1':
                        if blnDelay:
                            while time.time() < clickTime + (int(prevDelay) / 1000):
                                pass
                        pyautogui.click(int(row[0]), int(row[1]), button='left')
                        clickTime = time.time()

                    elif row[2] == 'M3':
                        if blnDelay:
                            while time.time() < clickTime + (int(prevDelay) / 1000):
                                pass
                        pyautogui.moveTo(int(row[0]), int(row[1]))
                        pyautogui.mouseDown(button='right')
                        time.sleep(.1) # TODO remove forced delay?
                        pyautogui.mouseUp(button='right')
                        clickTime = time.time()
                    else:
                        if blnDelay:
                            while time.time() < clickTime + (int(prevDelay) / 1000):
                                pass
                        pyautogui.click(int(row[0]), int(row[1]), button='middle')
                        clickTime = time.time()
                    print("Press " + str(row[2]) + " at " + str(time.time() - startTime))
                else:
                    # mouse click is action without position, do not move, just click
                    if row[2] == 'M1':
                        if blnDelay:
                            while time.time() < clickTime + (int(prevDelay) / 1000):
                                pass
                        pyautogui.click(button='left')
                        clickTime = time.time()
                    elif row[2] == 'M3':
                        if blnDelay:
                            while time.time() < clickTime + (int(prevDelay) / 1000):
                                pass
                        pyautogui.mouseDown(button='right')
                        time.sleep(.1)
                        pyautogui.mouseUp(button='right')
                        clickTime = time.time()
                    else:
                        if blnDelay:
                            while time.time() < clickTime + (int(prevDelay) / 1000):
                                pass
                        pyautogui.click(button='middle')
                        clickTime = time.time()

            # Action is #string, find image in Images folder with string name
            elif row[2][0] == '#' and len(row[2]) > 1:
                # delay/100 is confidence
                confidence = row[3]
                position = 0
                # confidence must be a percentile
                if 100 >= int(confidence) > 0:
                    for a in range(5):
                        try:
                            # print(os.getcwd() + r'\Images' + '\\' + str(row[2][1:len(row[2])]) + '.png')
                            # Confidence specified, use Delay as confidence percentile
                            position = pyautogui.locateCenterOnScreen(os.getcwd() + r'\Images' + '\\' + str(
                                row[2][1:len(row[2])]) + '.png', confidence=int(confidence) / 100)
                        except IOError:
                            pass
                        if position: break
                    if position:
                        # print("Found image at: ", position[0], ' : ', position[1])
                        if blnDelay:
                            while time.time() < clickTime + (int(prevDelay) / 1000):
                                pass
                        pyautogui.click(position[0] + global_monitor_left, position[1], button='left')
                        clickTime = time.time()
                    else:
                        break
                else:
                    # confidence could not be determined, use default
                    for a in range(5):
                        try:
                            # print(os.getcwd() + r'\Images' + '\\' + str(row[2][1:len(row[2])]) + '.png')
                            position = pyautogui.locateCenterOnScreen(os.getcwd() + r'\Images' + '\\' + str(
                                row[2][1:len(row[2])]) + '.png',
                                                                      confidence=int(confidence) / 100)
                        except IOError:
                            pass
                        if position: break
                    if position:
                        # print("Found image at: ", position[0], ' : ', position[1])
                        if blnDelay:
                            while time.time() < clickTime + (int(prevDelay) / 1000):
                                pass
                        pyautogui.click(position[0] + global_monitor_left, position[1], button='left')
                        clickTime = time.time()
                    else:
                        break

            # Action is !string, run macro with string name for Delay amount of times
            elif row[2][0] == '!' and len(row[2]) > 1:
                # macro is action, repeat for amount in delay
                if exists(os.getcwd() + r'\Macros' + '\\' + row[2][1:len(row[2])] + '.csv'):
                    arrayParam = list(csv.reader(
                        open(os.getcwd() + r'\Macros' + '\\' + row[2][1:len(row[2])] + '.csv',
                             mode='r')))
                    print("file: ", os.getcwd() + r'\Macros' + '\\' + row[2][1:len(row[2])] + '.csv')
                    startClickingBusy(arrayParam, int(row[3]), False)
            else:
                # print(row[2])
                # key press is action
                if row[2] != 'space':
                    if blnDelay:
                        while time.time() < clickTime + (int(prevDelay) / 1000):
                            pass
                    pyautogui.press(row[2])
                    clickTime = time.time()
                elif row[2] == 'space':
                    if blnDelay:
                        while time.time() < clickTime + (int(prevDelay) / 1000):
                            pass
                    pyautogui.press(' ')
                    clickTime = time.time()
                if row[2] == 'tab':
                    if blnDelay:
                        while time.time() < clickTime + (int(prevDelay) / 1000):
                            pass
                    pyautogui.press('\t')
                    clickTime = time.time()
                else:
                    clickTime = time.time()

            if loopsParam == 0 or loopsLeft.get() == 0: return
            # Only sleep if row is not macro, image finder, or key hold
            # if row[2][0] != '!' and row[2][0] != '#' and (len(row[2]) == 1 or row[2][0] != '_'):
            #     print(int(row[3]) / 1000)

            prevDelay = int(row[3])
            blnDelay = row[2] == "" or row[2][0] != "_"

        # decrement loop count param, also decrement main loop counter if main loop
        if loopsParam > 0: loopsParam = loopsParam - 1
        if mainLoop and loopsLeft.get() > 0: loopsLeft.set(loopsLeft.get() - 1)
    # if blnDelay:
    while time.time() < clickTime + (int(prevDelay) / 1000):
        pass
    # Release all keys that are pressed so as to not leave pressed
    # Necessary to be outside of loop because last row in macro will expect next row to release right prior to next press.
    for pressedKey in tN.currPressed:
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
        print("Release3 " + str(pressedKey) + " after " + str(int(prevDelay) / 1000) + " with time diff: " + str(time.time() - clickTime - int(prevDelay) / 1000))
        # print("Release " + str(pressedKey) + " after " + str(time.time() - clickTime))

    root.update()
    if loopsLeft.get() == 0:
        tN.startButton.config(state=NORMAL)
        tN.stopButton.config(state=DISABLED)


# Start Clicking button will loop through active notebook tab's treeview and move mouse, call macro, search for image and click or type keystrokes with specified delays
# Intentionally uses busy wait for most accurate delays
def startClicking(clickArray, threadFlag, loopsParam, mainLoop):
    root.update()

    if loopsParam == 0 or loopsLeft.get() == 0: return

    clickTime = 0
    prevDelay = 0
    blnDelay = False
    firstLoop = True
    selection = tN.treeView.selection()
    startFromSelected = tN.startFromSelected.get()
    if len(selection) > 0:
        selectedRow = tN.treeView.item(selection[0]).get("values")[0]
    else:
        selectedRow = 0

    # check Loopsleft as well to make sure Stop button wasn't pressed since this doesn't ues a global for loop count
    while loopsParam > 0 and loopsLeft.get() > 0:
        print("New loop")

        intLoop = 0
        startTime = time.time()

        for row in clickArray:
            if loopsParam == 0 or loopsLeft.get() == 0: return
            intLoop += 1

            # When start from selected row setting is true then find highlighted row(s) and skip to from first selected row
            # Only for first loop and first macro, not for subsequent loops nor macros called by the starting macro
            if firstLoop and mainLoop and startFromSelected == 1 and selectedRow > 0 and intLoop < selectedRow:
                continue

            # check row to see if its still holding keys
            if len(tN.currPressed) > 0 and (row[2] == "" or row[2][0] == '_'):
                # stop time if next action is not hold
                toBePressed = []
                if row[2] != "" and row[2][0] == '_':
                    toBePressed = row[2][1:].split('|')

                # key hold and release is action
                # print('PreRelease', (prevDelay / 1000), time.time() - clickTime)
                while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                    return
                blnDelay = False
                # print('PreReleasePostWait', (prevDelay / 1000), time.time() - clickTime)
                # Release all keys not pressed in next step
                for key in tN.currPressed:
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
                            tN.currPressed.remove(key)
                        except:
                            pass
                        # print("Release " + str(key) + " after " + str(time.time() - clickTime))
                        print("Release " + str(key) + " after total " + str(time.time() - clickTime))
                        print("Release1 " + str(key) + " after " + str(int(prevDelay) / 1000) + " with time diff: " + str(time.time() - clickTime - int(prevDelay) / 1000))
                # if this row is also a press then start timer as hols is continuing
                # clickTime = time.time()
            elif len(tN.currPressed) > 0:
                # print('PreRelease', str(pressed), (int(row[3]) / 1000), time.time() - clickTime)
                # if blnDelay:
                while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                    return
                # blnDelay = False
                # print('PreReleaseWait', (int(row[3]) / 1000), time.time() - clickTime)
                # Release all keys as next step is not pressing
                for key in tN.currPressed:
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
                        tN.currPressed.remove(key)
                    except:
                        pass
                        # print("Release " + str(key) + " after " + str(time.time() - clickTime))
                    print("Release " + str(key) + " after total " + str(time.time() - clickTime))
                    print("Release2 " + str(key) + " after " + str(int(prevDelay) / 1000) + " with time diff: " + str(time.time() - clickTime - int(prevDelay) / 1000))
                # if this row is not a press then wait to start timer until next click
                clickTime = time.time()

            # Empty must be first because other references to first character of Action will error
            if row[2] == "":
                # Nothing is just a pause
                print("blank", (int(row[3]) / 1000), time.time() - clickTime)
                if blnDelay:
                    while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                        return
                clickTime = time.time()
            # Moved _ hold action as close top as possible as it's accuracy is most important
            # Action is starts with an underscore (_), hold the key for the given amount of time
            elif row[2][0] == '_' and len(row[2]) > 1:
                # loop until end of string of keys to press
                toPress = row[2][1:].split('|')

                # wait prior row amount before pressing new keys
                if blnDelay:
                    while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                        return

                for key in toPress:
                    # Do not press if already pressed
                    if key not in tN.currPressed:
                        if key in ['M1', 'M2', 'M3']:
                            if int(row[0]) != 0 or int(row[1]) != 0:
                                # mouse click is action
                                if key == 'M1':
                                    pyautogui.mouseDown(int(row[0]), int(row[1]), button='left')
                                    tN.currPressed.append(key)
                                elif key == 'M3':
                                    pyautogui.moveTo(int(row[0]), int(row[1]))
                                    pyautogui.mouseDown(button='right')
                                    tN.currPressed.append(key)
                                else:
                                    pyautogui.mouseDown(int(row[0]), int(row[1]), button='middle')
                                    tN.currPressed.append(key)
                            else:
                                # mouse click is action without position, do not move, just click
                                if key == 'M1':
                                    pyautogui.mouseDown(button='left')
                                    tN.currPressed.append(key)
                                elif key == 'M3':
                                    pyautogui.mouseDown(button='right')
                                    tN.currPressed.append(key)
                                else:
                                    pyautogui.mouseDown(button='middle')
                                    tN.currPressed.append(key)
                        else:
                            # key press is action
                            if key == 'space':
                                pyautogui.keyDown(' ')
                                tN.currPressed.append(key)
                            elif key == 'tab':
                                pyautogui.keyDown('\t')
                                tN.currPressed.append(key)
                            else:
                                pyautogui.keyDown(key)
                                tN.currPressed.append(key)
                        print("Press " + str(key) + " at " + str(time.time() - clickTime))
                clickTime = time.time()

            elif row[2] in ['M1', 'M2', 'M3']:
                if int(row[0]) != 0 or int(row[1]) != 0:
                    # mouse click is action
                    if row[2] == 'M1':
                        if blnDelay:
                            while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                                return
                        pyautogui.click(int(row[0]), int(row[1]), button='left')
                        clickTime = time.time()

                    elif row[2] == 'M3':
                        if blnDelay:
                            while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                                return
                        pyautogui.moveTo(int(row[0]), int(row[1]))
                        pyautogui.mouseDown(button='right')
                        time.sleep(.1)  # TODO remove forced delay?
                        pyautogui.mouseUp(button='right')
                        clickTime = time.time()
                    else:
                        if blnDelay:
                            while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                                return
                        pyautogui.click(int(row[0]), int(row[1]), button='middle')
                        clickTime = time.time()
                    print("Press " + str(row[2]) + " at " + str(time.time() - clickTime))
                else:
                    # mouse click is action without position, do not move, just click
                    if row[2] == 'M1':
                        if blnDelay:
                            while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                                return
                        pyautogui.click(button='left')
                        clickTime = time.time()
                    elif row[2] == 'M3':
                        if blnDelay:
                            while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                                return
                        pyautogui.mouseDown(button='right')
                        time.sleep(.1)
                        pyautogui.mouseUp(button='right')
                        clickTime = time.time()
                    else:
                        if blnDelay:
                            while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                                return
                        pyautogui.click(button='middle')
                        clickTime = time.time()

            # Action is #string, find image in Images folder with string name
            elif row[2][0] == '#' and len(row[2]) > 1:
                # delay/100 is confidence
                confidence = row[3]
                position = 0
                # confidence must be a percentile
                if 100 >= int(confidence) > 0:
                    for a in range(5):
                        try:
                            # print(os.getcwd() + r'\Images' + '\\' + str(row[2][1:len(row[2])]) + '.png')
                            # Confidence specified, use Delay as confidence percentile
                            position = pyautogui.locateCenterOnScreen(os.getcwd() + r'\Images' + '\\' + str(
                                row[2][1:len(row[2])]) + '.png', confidence=int(confidence) / 100)
                        except IOError:
                            pass
                        if position: break
                    if position:
                        # print("Found image at: ", position[0], ' : ', position[1])
                        if blnDelay:
                            while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                                return
                        pyautogui.click(position[0] + global_monitor_left, position[1], button='left')
                        clickTime = time.time()
                    else:
                        break
                else:
                    # confidence could not be determined, use default
                    for a in range(5):
                        try:
                            # print(os.getcwd() + r'\Images' + '\\' + str(row[2][1:len(row[2])]) + '.png')
                            position = pyautogui.locateCenterOnScreen(os.getcwd() + r'\Images' + '\\' + str(
                                row[2][1:len(row[2])]) + '.png',
                                                                      confidence=int(confidence) / 100)
                        except IOError:
                            pass
                        if position: break
                    if position:
                        # print("Found image at: ", position[0], ' : ', position[1])
                        if blnDelay:
                            while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                                return
                        pyautogui.click(position[0] + global_monitor_left, position[1], button='left')
                        clickTime = time.time()
                    else:
                        break

            # Action is !string, run macro with string name for Delay amount of times
            elif row[2][0] == '!' and len(row[2]) > 1:
                # macro is action, repeat for amount in delay
                if exists(os.getcwd() + r'\Macros' + '\\' + row[2][1:len(row[2])] + '.csv'):
                    arrayParam = list(csv.reader(
                        open(os.getcwd() + r'\Macros' + '\\' + row[2][1:len(row[2])] + '.csv',
                             mode='r')))
                    print("file: ", os.getcwd() + r'\Macros' + '\\' + row[2][1:len(row[2])] + '.csv')
                    startClicking(arrayParam, threadFlag, int(row[3]), False)
            else:
                # print(row[2])
                # key press is action
                if row[2] != 'space':
                    if blnDelay:
                        while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                            return
                    pyautogui.press(row[2])
                    clickTime = time.time()
                elif row[2] == 'space':
                    if blnDelay:
                        while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                            return
                    pyautogui.press(' ')
                    clickTime = time.time()
                if row[2] == 'tab':
                    if blnDelay:
                        while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
                            return
                    pyautogui.press('\t')
                    clickTime = time.time()
                else:
                    clickTime = time.time()
                print("Press " + str(row[2]) + " at " + str(time.time() - startTime))

            if loopsParam == 0 or loopsLeft.get() == 0: return
            # Only sleep if row is not macro, image finder, or key hold
            # if row[2][0] != '!' and row[2][0] != '#' and (len(row[2]) == 1 or row[2][0] != '_'):
            #     print(int(row[3]) / 1000)

            prevDelay = int(row[3])
            blnDelay = row[2] == "" or row[2][0] != "_"

        # decrement loop count param, also decrement main loop counter if main loop
        if loopsParam > 0: loopsParam = loopsParam - 1
        if mainLoop and loopsLeft.get() > 0: loopsLeft.set(loopsLeft.get() - 1)
        firstLoop = False
    # if blnDelay:
    while threadFlag.wait((int(prevDelay) / 1000) - time.time() + clickTime):
        return
    # Release all keys that are pressed so as to not leave pressed
    # Necessary to be outside of loop because last row in macro will expect next row to release right prior to next press.
    for pressedKey in tN.currPressed:
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
        print("Release3 " + str(pressedKey) + " after " + str(int(prevDelay) / 1000) + " with time diff: " + str(time.time() - clickTime - int(prevDelay) / 1000))
        # print("Release " + str(pressedKey) + " after " + str(time.time() - clickTime))

    root.update()
    if loopsLeft.get() == 0:
        tN.startButton.config(state=NORMAL)
        tN.stopButton.config(state=DISABLED)


def stopClicking():
    # Reset clicker back to off state
    if tN.activeThread:
        tN.activeThread.threadFlag.set()
    loopsLeft.set(0)

    # Release all keys that are pressed so as to not leave pressed
    for pressedKey in tN.currPressed:
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
        print("Release " + str(pressedKey))

    tN.currPressed = []

    tN.startButton.config(state=NORMAL)
    tN.stopButton.config(state=DISABLED)


def getOrigin():
    # unhook mouse listener after getting mouse click, set to X, Y where mouse was clicked
    try:
        (a, b) = pyautogui.position()
        x.set(a)
        y.set(b)
    finally:
        mouse.unhook_all()


def getMouseMove(event):
    # return position from mouse click from anywhere on screen, set to X, Y entry values as mouse moves\
    (a, b) = pyautogui.position()
    x.set(a)
    y.set(b)


def getCursorPosition():
    # listen to mouse for next click
    mouse.hook(getMouseMove)
    mouse.on_click(getOrigin)


def addRow():
    # from Insert position button, adds row to bottom current treeview using entry values
    if len(tN.treeView.get_children()) % 2 == 0:
        tN.treeView.insert(parent='', index='end', iid=None, text=len(tN.treeView.get_children()) + 1, values=(len(tN.treeView.get_children()) + 1, x.get(), y.get(), actionEntry.get(), int(delayEntry.get()), commentEntry.get()), tags='evenrow')
    else:
        tN.treeView.insert(parent='', index='end', iid=None, text=len(tN.treeView.get_children()) + 1, values=(len(tN.treeView.get_children()) + 1, x.get(), y.get(), actionEntry.get(), int(delayEntry.get()), commentEntry.get()), tags='oddrow')
    exportMacro()


def addRowWithParams(xParam, yParam, keyParam, delayParam, commentParam):
    # for import to populate new treeview
    if len(tN.treeView.get_children()) % 2 == 0:
        tN.treeView.insert(parent='', index='end', iid=None, text=len(tN.treeView.get_children()) + 1, values=(len(tN.treeView.get_children()) + 1, xParam, yParam, keyParam, delayParam, commentParam), tags='evenrow')
    else:
        tN.treeView.insert(parent='', index='end', iid=None, text=len(tN.treeView.get_children()) + 1, values=(len(tN.treeView.get_children()) + 1, xParam, yParam, keyParam, delayParam, commentParam), tags='oddrow')


def showRightClickMenu(event):
    # show option menu when right clicking treeview of rows
    try:
        rightClickMenu.tk_removeup(event.x_root, event.y_root)
    finally:
        rightClickMenu.grab_release()


def selectRow(event):
    # this event populates the entries with the values of the selected row for easy editing/copying
    # must use global variables to tell if new row selected or just clicking whitespace in treeview, do not set values if not changing row selection
    global previouslySelectedTab
    global previouslySelectedRow

    selectedRow = tN.treeView.focus()
    selectedValues = tN.treeView.item(selectedRow, 'values')

    if len(selectedValues) > 0 and (previouslySelectedTab != tN.getNotebook().index(
            tN.getNotebook().select()) or previouslySelectedRow != selectedRow):
        actionEntry.delete(0, 'end')
        actionEntry.insert(0, selectedValues[3])
        delayEntry.delete(0, 'end')
        delayEntry.insert(0, selectedValues[4])
        commentEntry.delete(0, 'end')
        commentEntry.insert(0, selectedValues[5])
        x.set(selectedValues[1])
        y.set(selectedValues[2])
        previouslySelectedTab = tN.getNotebook().index(tN.getNotebook().select())
        previouslySelectedRow = selectedRow

        root.update()


def reorderRows(treeView):
    rows = treeView.get_children()
    i = 1
    for row in rows:
        if i % 2 == 0:
            treeView.item(row, text=i, tags='oddrow')
            treeView.set(row, "Step", i)
        else:
            treeView.item(row, text=i, tags='evenrow')
            treeView.set(row, "Step", i)
        i += 1


def importMacro():
    # import csv file into macro where filename will be macro name, is it best to import only exported macros
    filename = filedialog.askopenfilename(initialdir=os.getcwd() + r'\Macros', title="Select a .csv file",
                                          filetypes=(("csv files", "*.csv"),))
    answer = True
    found = 0

    # look for tab with same name as imported macro, will overwrite that tab if imported
    # this will ensure each macro has a unique name so that when a macro calls another macro there is no confusion over which macro should be called
    for i in range(len(tN.getNotebook().tabs())):
        if tN.getNotebook().tab(i, 'text') == os.path.splitext(os.path.basename(filename))[0]:
            answer = askyesno("Overwrite macro?", "Are you sure you want to overwrite current " +
                              os.path.splitext(os.path.basename(filename))[0])
            found = i + 1
            break

    # only open if file exists and overwrite true (answer defaults to true in case it is not asked)
    if filename and answer:
        with open(filename, 'r') as csvFile:
            csvReader = csv.reader(csvFile)
            if not found:
                tN.addTab(os.path.splitext(os.path.basename(filename))[0])
            # else:
            #     tN.notebook.select(found - 1)
            #     tN.getTabTree(found - 1).delete(*tN.getTabTree(found - 1).get_children())
            # for line in csvReader:
            #     print(line)
            #     print(len(line))
            #     #backwards compatibility to before there were comments
            #     if len(line) > 4:
            #         addRowWithParams(line[0], line[1], line[2], line[3], line[4])
            #     if line[3]:
            #         addRowWithParams(line[0], line[1], line[2], line[3], "")


def exportMacro():
    print('export')
    # save as csv file with name of file as macro name
    savePath = os.getcwd() + r'\Macros'
    filename = str(
        tN.getNotebook().tab(tN.getNotebook().select(), 'text').replace(r'/', r'-').replace(r'\\', r'-').replace(r'*',
                                                                                                                 r'-').replace(
            r'?', r'-').replace(r'[', r'-').replace(r']', r'-').replace(r'<', r'-').replace(r'>', r'-').replace(r'|',
                                                                                                                r'-'))

    print('export; ', filename)
    with open(os.path.join(savePath, str(filename + '.csv')), 'w', newline='') as newMacro:
        csvWriter = csv.writer(newMacro, delimiter=',')
        children = tN.treeView.get_children()

        for child in children:
            childValues = tN.treeView.item(child, 'values')
            # Do not include the Step column
            csvWriter.writerow(childValues[1:])


def actionRelease(event):
    # check if key already entered into hold list and remove it duplicate
    if str(actionEntry.get())[0:1] == '_' and '|' in str(actionEntry.get()):
        # at least two entries exist
        keys = str(actionEntry.get())[1:str(actionEntry.get()).rfind('|')].split('|')
        recentKey = str(actionEntry.get())[str(actionEntry.get()).rfind('|') + 1:]

        if recentKey in keys:
            cleanActionEntry = str(actionEntry.get())[0:str(actionEntry.get()).rfind('|')]
            actionEntry.delete(0, END)
            actionEntry.insert(0, cleanActionEntry)


def actionPopulate(event):
    # print(event)
    # print(event.keysym)
    actionEntry.icursor(len(actionEntry.get()))
    # _ is special character that should reset the actionEntry field to prevent all keys from adding to actionEntry
    if str(event.char) == '_':
        # Clear actionEntry field
        actionEntry.delete(0, END)
        # Then allow _ to be typed

    # !, #, and _ is special character that allows typing of action instead of instating setting action to each key press
    if str(actionEntry.get())[0:1] != '!' and str(actionEntry.get())[0:1] != '#' and str(actionEntry.get())[0:1] != '_':
        # need to use different properties for getting key press for letters vs whitespace/ctrl/shift/alt
        # != ?? to exclude mouse button as their char and keysym are not empty but are equal to ??
        if event.keysym == 'Escape' and event.char != '??':
            # special case for Escape because it has a char and might otherwise act like a letter but won't fill in the
            # box with 'Escape'
            actionEntry.delete(0, END)
            actionEntry.insert(0, event.keysym)
        elif str(event.char).strip() and event.char != '??':
            # clear entry before new char is entered
            print(event.char)
            if 96 <= event.keycode <= 105:
                # Append NUM if keystroke comes from numpad
                actionEntry.delete(0, END)
                actionEntry.insert(0, 'NUM')
            else:
                actionEntry.delete(0, END)
        elif event.keysym and event.char != '??':
            # clear entry and enter key string
            actionEntry.delete(0, END)
            actionEntry.insert(0, event.keysym)
        else:
            # clear entry and enter Mouse event
            actionEntry.delete(0, END)
            actionEntry.insert(0, 'M' + str(event.num))
        # else key event is duplicate
    # _ is special character that needs to be followed by a key stroke
    if str(actionEntry.get())[0:1] == '_':
        if event.keysym == 'Escape' and event.char != '??':
            # special case for Escape because it has a char and might otherwise act like a letter but won't fill in the
            # box with 'Escape'
            if len(actionEntry.get()) > 1:
                actionEntry.insert(len(actionEntry.get()), '|' + event.keysym)
            else:
                actionEntry.insert(len(actionEntry.get()), event.keysym)

        elif str(event.char).strip() and event.char != '??':
            # clear entry before new char is entered
            if 96 <= event.keycode <= 105:
                # Append NUM if keystroke comes from numpad
                if len(actionEntry.get()) > 1:
                    actionEntry.insert(len(actionEntry.get()), '|NUM')
                else:
                    actionEntry.insert(len(actionEntry.get()), 'NUM')
            elif len(actionEntry.get()) > 1:
                # normal character
                actionEntry.insert(len(actionEntry.get()), '|')

        elif event.keysym and event.char != '??':
            if event.keysym == 'space':
                # delete the ' ' that was just typed as this uses 'space' as ' '
                actionEntry.delete(len(actionEntry.get()) - 1)
            # clear entry and enter key string
            if len(actionEntry.get()) > 1:
                actionEntry.insert(len(actionEntry.get()), '|' + event.keysym)
            elif event.keysym not in str(actionEntry.get()[:-1]):
                actionEntry.insert(len(actionEntry.get()), event.keysym)
        else:
            # clear entry and enter Mouse event
            if len(actionEntry.get()) > 1:
                actionEntry.insert(len(actionEntry.get()), '|M' + str(event.num))
            else:
                actionEntry.insert(len(actionEntry.get()), 'M' + str(event.num))
        # else event is already entered, do not allow repeats in actionEntry


def destroy():
    if root.messagebox.askokcancel("Quit", "Do you really wish to quit?"):
        root.destroy()


def overwriteRows():  # what row and column was clicked on
    rows = tN.treeView.selection()
    for row in rows:
        # Edit everything except Step
        tN.treeView.set(row, "X", x.get())
        tN.treeView.set(row, "Y", y.get())
        tN.treeView.set(row, "Action", actionEntry.get())
        tN.treeView.set(row, "Delay", delayEntry.get())
        tN.treeView.set(row, "Comment", commentEntry.get())
        # tN.treeView.item(row, values=(x.get(), y.get(), .get(), delayEntry.get(), commentEntry.get()))
    exportMacro()


def overwriteRow(rowNum, xPos, yPos, action, delay, comment):
    print("change row")
    print(rowNum, xPos, yPos, action, delay, comment)
    rows = tN.treeView.get_children()
    print(rows)
    i = 0

    for row in rows:
        print(i)
        if i == rowNum:
            tN.treeView.item(row, values=(rowNum, xPos, yPos, action, delay, comment))
            print("changed!")
            break
        i += 1


def showHelp():
    if tN.helpWindow is not None:
        if tN.helpWindow.state() == "normal":
            closeHelp()
        else:
            tN.helpWindow.wm_state("normal")
    else:
        tN.helpWindow = Toplevel(root)
        tN.helpWindow.attributes("-topmost", True)
        tN.helpWindow.resizable(height=None, width=None)
        tN.helpWindow.wm_title("Sticky's Autoclicker Help")
        tN.helpWindow.protocol("WM_DELETE_WINDOW", closeHelp)

        helpLabel = Label(tN.helpWindow, text=usabilityNotes, justify=LEFT)
        helpLabel.grid(row=0, column=0)

        closeButton = ttk.Button(tN.helpWindow, text="Close", command=closeHelp)
        closeButton.grid(row=1, column=0)


def showSettings():
    if tN.settingsWindow is not None:
        if tN.settingsWindow.state() == "normal":
            saveSettings()
            closeSettings()
        else:
            tN.settingsWindow.wm_state("normal")
    else:
        tN.settingsWindow = Toplevel(root)
        tN.settingsWindow.attributes("-topmost", True)
        tN.settingsWindow.resizable(height=None, width=None)
        tN.settingsWindow.wm_title("Sticky's Autoclicker Settings")
        tN.settingsWindow.protocol("WM_DELETE_WINDOW", closeSettings)

        # Rows 0 and 1
        busyLabel = Label(tN.settingsWindow, text="Use Busy Wait", borderwidth=2)
        busyLabel.grid(row=0, column=0, padx=10)
        busyButton = Checkbutton(tN.settingsWindow, variable=tN.busyWait, onvalue=1, offvalue=0, command=toggleBusy, borderwidth=2)
        busyButton.grid(row=1, column=0, padx=10, pady=10)

        windowLabel = Label(tN.settingsWindow, text="Application Selector", borderwidth=2)
        windowLabel.grid(row=0, column=1, sticky='s', padx=10)
        windowButton = Button(tN.settingsWindow, text="Find App", bg=bgButton, command=windowFinder, borderwidth=2)
        windowButton.grid(row=1, column=1, sticky='n', padx=10, pady=10)

        hiddenLabel = Label(tN.settingsWindow, text="Use hidden mode", borderwidth=2)
        hiddenLabel.grid(row=0, column=2, sticky='s', padx=10)
        hiddenButton = Checkbutton(tN.settingsWindow, variable=tN.hiddenMode, onvalue=1, offvalue=0, command=toggleHidden)
        hiddenButton.grid(row=1, column=2, sticky='n', padx=10, pady=10)

        # Rows 3 and 4
        startFromSelectedLabel = Label(tN.settingsWindow, text="Start From Selected Row", borderwidth=2)
        startFromSelectedLabel.grid(row=3, column=0, padx=10)
        startFromSelectedButton = Checkbutton(tN.settingsWindow, variable=tN.startFromSelected, onvalue=1, offvalue=0, command=toggleStartFromSelected, borderwidth=2)
        startFromSelectedButton.grid(row=4, column=0, padx=10, pady=10)

        closeButton = ttk.Button(tN.settingsWindow, text="Close", command=closeSettings)
        closeButton.grid(row=5, column=1, padx=10, pady=10)
0

def closeSettings():
    tN.settingsWindow.destroy()
    tN.settingsWindow = None
    # config['Settings'] = {'busyWait': str(tN.busyWait.get()), 'startfromselected': str(tN.startFromSelected.get()), 'selectedApp': str(tN.selectedApp), 'hiddenMode': str(tN.hiddenMode.get())}


def toggleBusy():
    config = cp.ConfigParser()
    config.read(os.path.expanduser('~/Documents/StickysAutoClicker/config.ini'))
    config.set('Settings', 'busyWait', str(tN.busyWait.get()))
    with open(os.path.expanduser('~/Documents/StickysAutoClicker/config.ini'), 'w') as configfile:
        config.write(configfile)


def toggleHidden():
    config = cp.ConfigParser()
    config.read(os.path.expanduser('~/Documents/StickysAutoClicker/config.ini'))
    config.set('Settings', 'hiddenMode', str(tN.hiddenMode.get()))
    with open(os.path.expanduser('~/Documents/StickysAutoClicker/config.ini'), 'w') as configfile:
        config.write(configfile)


def toggleStartFromSelected():
    config = cp.ConfigParser()
    config.read(os.path.expanduser('~/Documents/StickysAutoClicker/config.ini'))
    config.set('Settings', 'startfromselected', str(tN.startFromSelected.get()))
    with open(os.path.expanduser('~/Documents/StickysAutoClicker/config.ini'), 'w') as configfile:
        config.write(configfile)


def closeHelp():
    tN.helpWindow.destroy()
    tN.helpWindow = None


def windowFinder():
    root.wm_state("iconic")
    if tN.settingsWindow is not None:
        tN.settingsWindow.wm_state("iconic")
    if tN.helpWindow is not None:
        tN.helpWindow.wm_state("iconic")
    mouse.on_click(getWindow)


def getWindow():
    # unhook mouse listener after getting window that was clicked
    try:
        tN.window = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())
        print(psutil.Process(tN.window[-1]).name())
    finally:
        root.wm_state("normal")
        mouse.unhook_all()

    config = cp.ConfigParser()
    config.read(os.path.expanduser('~/Documents/StickysAutoClicker/config.ini'))
    config.set('Settings', 'selectedApp', str(tN.selectedApp))
    with open(os.path.expanduser('~/Documents/StickysAutoClicker/config.ini'), 'w') as configfile:
        config.write(configfile)



def onClose():
    stopClicking()

    root.destroy()


class Recorder:
    start = None
    startPress = None
    thread = None
    pressed = []
    keycode = keyboard.KeyCode
    recording = False
    lastRow = []

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

    def __init__(self):
        print(self._keyDict)

    def startRecordingThread(self):
        # currently think supporting clicking while recording is a bad idea
        stopClicking()
        tN.startButton.config(state=DISABLED)
        tN.recordButton.configure(bg='red')

        # create thread to monitor for exit keycombo
        sE = threading.Thread(target=monitorExit, args=())
        sE.start()

        self.recording = True

        # create new thread for recording so as to not disturb the autoclicker window
        threadRecordingFlag = threading.Event()
        self.thread = threading.Thread(target=self.record)
        self.thread.threadFlag = threadRecordingFlag
        self.thread.start()

    def stopRecordingThread(self):
        tN.startButton.config(state=NORMAL)
        tN.recordButton.configure(bg=root.cget('bg'))

        if self.thread:
            self.thread.threadFlag.set()

        self.recording = False

        exportMacro()

    def record(self):
        # start listening
        with Listener(on_press=self.__recordPress, on_release=self.__recordRelease) as listener:
            try:
                listener.join()
            except Exception as ex:
                print('{0} was pressed'.format(ex.args[0]))

    def __recordPress(self, key):
        # log time ASAP for accuracyas
        tempTime = time.time()
        print(self.pressed)
        if self.startPress is not None:
            print(int((time.time() - self.startPress) * 1000))

        if self.thread.threadFlag.is_set():
            return False

        # if key.char exists use that, else translate to pyautogui keys
        try:
            if key.char is not None:
                key = key.char
                print(key.char)
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

            print(self.lastRow)
            if self.lastRow[2][0] == '_':
                # last row was hold so add row to account for delay
                self.__addRow(0, 0, "", int((time.time() - self.startPress) * 1000))
            else:
                # last row was not hold so edit last row to change delay
                print("change: " + str(len(tN.treeView.get_children())))
                print(int((time.time() - self.startPress) * 1000))
                self.__changeRow(len(tN.treeView.get_children()) - 1, 0, 0, self.lastRow[2], int((time.time() - self.startPress) * 1000))

        # key is being pressed, add to array and log time
        self.startPress = tempTime
        self.pressed.append(str(key))
        print("{0}Down ".format(str(format(key))))

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
                print(key.char)
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

        print("key is " + str(key))
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

        print("{0} Release".format(key))

    def __addRow(self, x, y, key, delay):
        self.lastRow = [x, y, key, delay]
        addRowWithParams(x, y, key, delay, "")

    def __changeRow(self, row, x, y, key, delay):
        self.lastRow = [x, y, key, delay]
        overwriteRow(row, x, y, key, delay, "")


tN = treeviewNotebook()

# bind after instantiation because resizing occurs while building window before all elements are even created
root.bind("<Configure>", resizeNotebook)

root.protocol("WM_DELETE_WINDOW", onClose)
root.update_idletasks()
root.mainloop()
