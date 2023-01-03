import os
import time
import csv
import mouse
import mss
from PIL import ImageGrab
from functools import partial
ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)
import pyautogui
pyautogui.FAILSAFE = False
import threading
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import simpledialog
from tkinter.messagebox import askyesno
from pynput import keyboard
from os.path import exists


# TODO: Export crash reports
# TODO: Make pretty

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
- tkinter.filedialog for importing macros
- pynput for listening to keyboard for emergency exit combo Ctrl + Shift + 1

build with "pyinstaller --onefile --noconsole StickysAutoClicker.py"
"""

usabilityNotes = ("                 Usability Notes\n\n"
                  " - Selecting any row of a macro will auto-populate the X and Y positions, delay and action fields with the values of that row, overwriting anything previously entered.\n"
                  " - Shift + Ctrl + ` will set loops to 0 and stop the autoclicker immediately."
                  " - Overwrite will set all selected rows to the current X, Y, delay and action values.\n"
                  " - Macros are exported as csv files that are writes to folder '\Macros' in exe directory.\n"
                  " - Exported Files are kept up to date with each edit created.\n"
                  " - Import expects a comma separated file with no headers and only four columns: X, Y, Action, Delay\n"
                  " - Macros do not need to be in a tab to be called by another macro. The csv in the Macros folder will be used, so be sure that is up to date.\n"
                  " - Action will recognize the main three mouse clicks as well as any keyboard action including Shift + keys (doesn't work with Alt or Ctrl + keys)\n"
                  " - Key presses do not move cursor, so X and Y positions do not matter.\n"
                  " - Action has two escape characters, ! and #, that will allow the user to continue typing rather than overwriting the action key.\n"
                  "           This is to allow for calling another macro (!) or finding an image (#)\n"
                  " - Typing !macroName into action and adding that row will make a macro look for another tab in Sticky's Autoclicker with the tab name macroName and execute that macro.\n"
                  "           Delay serves another purpose when used with a macro action (!macroName) and will call that macro for the delay amount of times\n"
                  " - Typing #imageName into action and adding that row will make a macro look for a .png image with that name in the \Images folder and move the cursor to a found image and left click.\n"
                  "           Delay serves even another purpose when used with a find image action (#imageName) and serve as the confidence percentage level.\n"
                  "           Confidence of 100 will find an exact match and confidence of 1 will find the first roughly similar image.\n"
                  "           If image is not found then loop will end and next loop will start.\n"
                  " - Finding an image will try 5 times with a .1 second delay if not found and click once found.\n")

root = Tk()
global_monitor_left = 0

#define global monitors
with mss.mss() as sct:
    if sct.monitors:
        global_monitor_left = sct.monitors[0].get('left')

class treeviewNotebook:
    notebook = None

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

    def initElements(self):
        # Column 0
        # for reference: notebook.grid(row=3, column=2, columnspan=3, rowspan=5, sticky='')
        root.columnconfigure(0, weight=6)
        self.welcomeLabel = Label(root, text="Sticky's Autoclicker", font=("Arial bold", 11), bg=bgColor)
        self.welcomeLabel.grid(row=0, column=0, padx=5, pady=0, sticky='n', columnspan=2)
        self.madeByLabel = Label(root, text="Made by Sticky", bg=bgColor)
        self.madeByLabel.grid(row=0, column=0, padx=20, sticky='', columnspan=2)
        self.clickLabel = Label(root, text="Click Loops", font=("Arial", 10), bg=bgColor)
        self.clickLabel.grid(row=1, column=0, sticky='e', pady=10)
        self.LoopsLeftLabel = Label(root, text="Loops Left", font=("Arial", 11), bg=bgColor)
        self.LoopsLeftLabel.grid(row=2, column=0, sticky='s')
        self.clicksLeftLabel = Label(root, textvariable=loopsLeft, bg=bgColor)
        self.clicksLeftLabel.grid(row=3, column=0, sticky='n')
        self.startButton = Button(root, text="Start Clicking", font=("Arial", 15), command=threadStartClicking, padx=10,
                             pady=10,
                             borderwidth=6, bg=bgButton)
        self.startButton.grid(row=4, column=0, columnspan=2, sticky="n")
        self.stopButton = Button(root, text="Stop Clicking", font=("Arial", 15), command=stopClicking, padx=0, pady=10,
                            borderwidth=6, bg=bgButton)
        self.stopButton.grid(row=5, column=0, columnspan=2, sticky="n")

        self.importMacroButton = Button(root, text="Import Macro", bg=bgButton, command=importMacro, borderwidth=2)
        self.importMacroButton.grid(row=6, column=0, sticky='n', columnspan=2)
        self.helpButton = Button(root, text="Help", bg=bgButton, command=showHelp, borderwidth=2)
        self.helpButton.grid(row=7, column=0, sticky='n', columnspan=2)

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
        actionEntry.bind('<Return>', actionPopulate, add='+')
        actionEntry.bind('<Button-1>', actionPopulate, add='+')
        actionEntry.bind('<Escape>', actionPopulate, add='+')
        actionEntry.bind('<Button-2>', actionPopulate, add='+')
        actionEntry.bind('<Button-3>', actionPopulate, add='+')

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

        # notebookStyle = ttk.Style()
        # notebookStyle.configure("TNotebook", background=bgColor)

        self.notebook = ttk.Notebook(self.frame)
        self.tabFrame = tk.Frame(self.notebook, bg=bgColor)
        self.notebook.add(self.tabFrame, text='macro1')
        self.notebook.grid(row=2, column=2, columnspan=8, rowspan=1)
        self.notebook.bind('<<NotebookTabChanged>>', self.tabRefresh)

        self.treeFrame = tk.Frame(root, bg=bgColor)
        self.treeFrame.grid(row=2, column=2, columnspan=6, rowspan=6, sticky="nsew", pady=(20,0))
        self.treeView = ttk.Treeview(self.treeFrame, selectmode="extended")
        self.treeTabs.append('macro1')

        self.verticalTabScroll = Scrollbar(self.treeFrame, orient='vertical', command=self.treeView.yview)
        self.verticalTabScroll.pack(side=RIGHT, fill=Y)
        self.treeView.configure(yscrollcommand=self.verticalTabScroll.set)

        self.treeView['columns'] = ("X", "Y", "Action", "Delay", "Comment")

        self.treeView.column('#0', anchor='center', width=35, stretch=False)
        self.treeView.column('X', anchor='center', minwidth=60, stretch=False)
        self.treeView.column('Y', anchor='center', width=60, stretch=False)
        self.treeView.column('Action', anchor='center', width=100, stretch=False)
        self.treeView.column('Delay', anchor='center', width=80, stretch=False)
        self.treeView.column('Comment', anchor='center', width=120, stretch=False)

        self.treeView.heading('#0', text='Step')
        self.treeView.heading('X', text='X')
        self.treeView.heading('Y', text='Y')
        self.treeView.heading('Action', text='Action')
        self.treeView.heading('Delay', text='Delay/Repeat')
        self.treeView.heading('Comment', text='Comment')
        self.treeView.pack(fill=BOTH, expand=True)
        # self.treeView.grid(row=0, column=0, columnspan=4, rowspan=4)
        self.treeView.tag_configure('oddrow', background="lightblue")
        self.treeView.tag_configure('evenrow', background="white")

        self.treeView.bind("<Button-3>", rCM.showRightClickMenu)
        self.treeView.bind("<ButtonRelease-1>", selectRow)

    def tabRefresh(self, event):
        print('tabRefresh')
        #clear treeview of old macro
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

                for line in csvReader:
                    #backwards compatability to before there were comments
                    if len(line) > 4:
                        addRowWithParams(line[0], line[1], line[2], line[3], line[4])
                    elif line[3]:
                        addRowWithParams(line[0], line[1], line[2], line[3], "")

        # root.update()

    def frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def addTab(self, name):
        self.tabFrame = ttk.Frame(self.notebook)
        self.tabFrame.config(height=320, width=440)

        self.treeTabs.append(name)
        tabCount = len(self.treeTabs)

        self.notebook.add(self.tabFrame, text=name)

        self.notebook.select(self.tabFrame)

    def getNotebook(self):
        return self.notebook


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


bgColor = None #'CadetBlue1'   '#9ac6e3'
fgColor = None  # '#759fba'
bgButton = None  # '#e0f3ff'

# set icon to cute bunny
try:
    root.iconbitmap(os.getcwd() + '\Resources\StickyHeadIcon.ico')
except Exception as e:
    print("no cute bun")

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

    tN.treeView.column('#0', width=int((winWidth / winWeight))) #1
    tN.treeView.column('X', width=int((winWidth / winWeight) * 1.5)) #2
    tN.treeView.column('Y', width=int((winWidth / winWeight) * 1.5)) #2
    tN.treeView.column('Action', width=int((winWidth / winWeight) * 3)) #3
    tN.treeView.column('Delay', width=int((winWidth / winWeight) * 3)) #3
    tN.treeView.column('Comment', width=int((winWidth / winWeight) * 4)) #4


vcmd = (root.register(checkNumerical), '%S')
loopEntry = Entry(root, width=5, justify="right", validate='key', vcmd=vcmd)
loopEntry.insert(0, 0)
delayEntry = Entry(root, width=12, justify="right", validate='key', vcmd=vcmd)
delayEntry.insert(0, 0)
actionEntry = Entry(root, width=10, justify="center")
actionEntry.insert(0, 'M1')
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
    #save current macro to csv, easier to read, helps update retention
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
        tN.activeThread = threading.Thread(target=startClicking,
                                           args=(clickArray, int(loopEntry.get()), True, threadFlag))
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

    hotkey = keyboard.HotKey(
        keyboard.HotKey.parse('<ctrl>+<Shift>+`'),
        stopClicking)

    with keyboard.Listener(
            on_press=for_canonical(hotkey.press)) as l:
        l.join()


# Start Clicking button will loop through active notebook tab's treeview and move mouse, call macro, search for image and click or type keystrokes with specified delays
def startClicking(clickArray, loopsParam, mainLoop, threadFlag):
    # loop though param treeview, for param loops
    # needed so as to not mess with LoopsLeft from outer loop
    root.update()

    # check Loopsleft as well to make sure Stop button wasn't pressed since this doesn't ues a global for loop count
    while loopsParam > 0 and loopsLeft.get() > 0:
        for row in range(len(clickArray)):
            if loopsParam == 0 or loopsLeft.get() == 0: return

            if clickArray[row][2] in ['M1', 'M2', 'M3']:
                if int(clickArray[row][0]) != 0 or int(clickArray[row][1]) != 0:
                    # mouse click is action
                    if clickArray[row][2] == 'M1':
                        pyautogui.click(int(clickArray[row][0]), int(clickArray[row][1]), button='left')
                    elif clickArray[row][2] == 'M3':
                        pyautogui.moveTo(int(clickArray[row][0]), int(clickArray[row][1]))
                        pyautogui.mouseDown(button='right')
                        time.sleep(.1)
                        pyautogui.mouseUp(button='right')
                    else:
                        pyautogui.click(int(clickArray[row][0]), int(clickArray[row][1]), button='middle')
                else:
                    # mouse click is action without position, do not move, just click
                    if clickArray[row][2] == 'M1':
                        pyautogui.click(button='left')
                    elif clickArray[row][2] == 'M3':
                        pyautogui.mouseDown(button='right')
                        time.sleep(.1)
                        pyautogui.mouseUp(button='right')
                    else:
                        pyautogui.click(button='middle')

            # Action is #string, find image in Images folder with string name
            elif clickArray[row][2][0] == '#' and len(clickArray[row][2]) > 1:
                # delay/100 is confidence
                confidence = clickArray[row][3]
                position = 0
                # confidence must be a percentile
                if 100 >= int(confidence) > 0:
                    for a in range(5):
                        try:
                            # print(os.getcwd() + r'\Images' + '\\' + str(clickArray[row][2][1:len(clickArray[row][2])]) + '.png')
                            # Confidence specified, use Delay as confidence percentile
                            position = pyautogui.locateCenterOnScreen(os.getcwd() + r'\Images' + '\\' + str(
                                clickArray[row][2][1:len(clickArray[row][2])]) + '.png',
                                                                      confidence=int(confidence) / 100)
                        except IOError:
                            pass
                        if position: break
                    if position:
                        # print("Found image at: ", position[0], ' : ', position[1])
                        pyautogui.click(position[0] + global_monitor_left, position[1], button='left')
                    else:
                        break
                else:
                    # confidence could not be determined, use default
                    for a in range(5):
                        try:
                            # print(os.getcwd() + r'\Images' + '\\' + str(clickArray[row][2][1:len(clickArray[row][2])]) + '.png')
                            position = pyautogui.locateCenterOnScreen(os.getcwd() + r'\Images' + '\\' + str(
                                clickArray[row][2][1:len(clickArray[row][2])]) + '.png',
                                                                      confidence=int(confidence) / 100)
                        except IOError:
                            pass
                        if position: break
                    if position:
                        # print("Found image at: ", position[0], ' : ', position[1])
                        pyautogui.click(position[0] + global_monitor_left, position[1], button='left')
                    else:
                        break

            # Action is !string, run macro with string name for Delay amount of times
            elif clickArray[row][2][0] == '!' and len(clickArray[row][2]) > 1:
                # macro is action, repeat for amount in delay
                if exists(os.getcwd() + r'\Macros' + '\\' + clickArray[row][2][1:len(clickArray[row][2])] + '.csv'):
                    arrayParam = list(csv.reader(
                        open(os.getcwd() + r'\Macros' + '\\' + clickArray[row][2][1:len(clickArray[row][2])] + '.csv',
                             mode='r')))
                    # print("file: ", os.getcwd() + r'\Macros' + '\\' + clickArray[row][2][1:len(clickArray[row][2])] + '.csv')
                    startClicking(arrayParam, int(clickArray[row][3]), False, threadFlag)
            else:
                # key press is action
                pyautogui.press(clickArray[row][2])

            if loopsParam == 0 or loopsLeft.get() == 0: return
            # Only sleep if row is not macro or image finder
            if clickArray[row][2][0] != '!' and clickArray[row][2][0] != '#' and len(clickArray[row][2]) > 1:
                while threadFlag.wait(int(clickArray[row][3]) / 1000):
                    return

        # decrement loop count param, also decrement main loop counter if main loop
        if loopsParam > 0: loopsParam = loopsParam - 1
        if mainLoop and loopsLeft.get() > 0: loopsLeft.set(loopsLeft.get() - 1)
        root.update()

    if loopsLeft.get() == 0:
        tN.startButton.config(state=NORMAL)
        tN.stopButton.config(state=DISABLED)


def stopClicking():
    loopsLeft.set(0)
    if tN.activeThread:
        tN.activeThread.threadFlag.set()
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
        tN.treeView.insert(parent='', index='end', iid=None,
                                                                                text=len(tN.treeView.get_children()) + 1,
                                                                                values=(
                                                                                x.get(), y.get(), actionEntry.get(),
                                                                                int(delayEntry.get()), commentEntry.get()), tags='evenrow')
    else:
        tN.treeView.insert(parent='', index='end', iid=None,
                                                                                text=len(tN.treeView.get_children()) + 1,
                                                                                values=(
                                                                                x.get(), y.get(), actionEntry.get(),
                                                                                int(delayEntry.get()), commentEntry.get()), tags='oddrow')
    exportMacro()


def addRowWithParams(xParam, yParam, keyParam, delayParam, commentParam):
    # for import to populate new treeview
    if len(tN.treeView.get_children()) % 2 == 0:
        tN.treeView.insert(parent='', index='end', iid=None,
                                                                                text=len(tN.treeView.get_children()) + 1,
                                                                                values=(
                                                                                xParam, yParam, keyParam, delayParam, commentParam),
                                                                                tags='evenrow')
    else:
        tN.treeView.insert(parent='', index='end', iid=None,
                                                                                text=len(tN.treeView.get_children()) + 1,
                                                                                values=(
                                                                                xParam, yParam, keyParam, delayParam, commentParam),
                                                                                tags='oddrow')
    # exportMacro()


def showRightClickMenu(event):
    # show option menu when right clicking treeview of rows
    try:
        rightClickMenu.tk_popup(event.x_root, event.y_root)
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
        actionEntry.insert(0, selectedValues[2])
        delayEntry.delete(0, 'end')
        delayEntry.insert(0, selectedValues[3])
        commentEntry.delete(0, 'end')
        commentEntry.insert(0, selectedValues[4])
        x.set(selectedValues[0])
        y.set(selectedValues[1])
        previouslySelectedTab = tN.getNotebook().index(tN.getNotebook().select())
        previouslySelectedRow = selectedRow

        root.update()


def reorderRows(treeView):
    rows = treeView.get_children()
    i = 1
    for row in rows:
        if i % 2 == 0:
            treeView.item(row, text=i, tags='oddrow')
        else:
            treeView.item(row, text=i, tags='evenrow')
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
            #     #backwards compatability to before there were comments
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

    print('export; ',  filename)
    with open(os.path.join(savePath, str(filename + '.csv')), 'w', newline='') as newMacro:
        csvWriter = csv.writer(newMacro, delimiter=',')
        children = tN.treeView.get_children()

        for child in children:
            childValues = tN.treeView.item(child, 'values')
            csvWriter.writerow(childValues)


def actionPopulate(event):
    # ! is special character that allows typing of action instead of instating setting action to each key press
    if str(actionEntry.get())[0:1] != '!' and str(actionEntry.get())[0:1] != '#':
        # need to use different properties for getting key press for letters vs whitespace/ctrl/shift/alt
        # != ?? to exclude mouse button as their char and keysym are not empty but are equal to ??
        print(event)
        if event.keysym == 'Escape' and event.char != '??':
            # special case for Escape because it has a char and might otherwise act like a letter but won't fill in the
            # box with 'Escape'
            actionEntry.delete(0, END)
            actionEntry.insert(0, event.keysym)
        elif str(event.char).strip() and event.char != '??':
            # clear entry before new char is entered
            actionEntry.delete(0, END)
        elif event.keysym and event.char != '??':
            # clear entry and enter key string
            actionEntry.delete(0, END)
            actionEntry.insert(0, event.keysym)
        else:
            # clear entry and enter key string
            actionEntry.delete(0, END)
            actionEntry.insert(0, 'M' + str(event.num))


def overwriteRows():  # what row and column was clicked on
    rows = tN.treeView.selection()
    for row in rows:
        tN.treeView.item(row, values=(
        x.get(), y.get(), actionEntry.get(), delayEntry.get(), commentEntry.get()))
    exportMacro()


def showHelp():
    helpWindow = Toplevel(root)
    helpWindow.wm_title("Sticky's Autoclicker Help")

    helpLabel = Label(helpWindow, text=usabilityNotes, justify=LEFT)
    helpLabel.grid(row=0, column=0)

    closeButton = ttk.Button(helpWindow, text="Close", command=helpWindow.destroy)
    closeButton.grid(row=1, column=0)


def onClose():
    stopClicking()

    root.destroy()


tN = treeviewNotebook()

# bind after instantiation because resizing occurs while building window before all elements are even created
root.bind("<Configure>", resizeNotebook)

root.protocol("WM_DELETE_WINDOW", onClose)
root.update_idletasks()
root.mainloop()
