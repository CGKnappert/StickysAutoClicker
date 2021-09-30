import os
import time
import csv
import mouse
import pyautogui
import threading
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import simpledialog
from tkinter.messagebox import askyesno

"""
Sticky's Autoclicker Documentation
This is an autoclicker with fancier functionality than I have been able to find online and have always desired.

The backbone of this GUI is tkinter using a grid.
The imports used are for:
 - mouse import for finding mouse position outside of the tkinter window
 - pyautogui for moving mouse and clicking
 - threading for starting the clicking loop without the window becoming unresponsive
 - csv for import and export functionality
 - time for sleeping for delays
 - os for finding filepath 
 - tkinter.simpledialog for asking for new macro name
 - tkinter.filedialog for importing macros
"""

usabilityNotes = ("                 Usability Notes\n\n"
" - Selecting any row of a macro will auto-populate the X and Y positions, delay and action fields with the values of that row, overwriting anything previously entered.\n"
" - Overwrite will set all selected rows to the current X, Y, delay and action values.\n"
" - Export writes to folder '\Macros' in exe directory.\n"
" - Export will overwrite any existing file of the matching name without prompt\n"
" - Import expects a comma separated file with no headers and only four columns: X, Y, Action, Delay\n"
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


class treeviewNotebook:
    notebook = None
    # allow treeview array to be referenced for macro looping
    treeTabs = []

    # class with notebook with tabs of treeviews
    def __init__(self, rootRoot):
        self.notebook = ttk.Notebook(rootRoot)
        self.tabFrame = ttk.Frame(self.notebook)
        self.tabScroll = Scrollbar(self.tabFrame)
        self.tabScroll.pack(side=RIGHT, fill=Y)
        self.notebook.add(self.tabFrame, text='macro1')
        self.notebook.grid(row=3, column=2, columnspan=3, rowspan=5)
        self.treeTabs.append(ttk.Treeview(self.tabFrame, yscrollcommand=self.tabScroll.set, selectmode="extended"))
        self.treeTabs[0]['columns'] = ("X", "Y", "Action", "Delay")

        self.treeTabs[0].column('#0', anchor='center', width=40)
        self.treeTabs[0].column('X', anchor='center', width=60)
        self.treeTabs[0].column('Y', anchor='center', width=60)
        self.treeTabs[0].column('Action', anchor='center', width=120)
        self.treeTabs[0].column('Delay', anchor='center', width=120)

        self.treeTabs[0].heading('#0', text='Step')
        self.treeTabs[0].heading('X', text='X')
        self.treeTabs[0].heading('Y', text='Y')
        self.treeTabs[0].heading('Action', text='Action')
        self.treeTabs[0].heading('Delay', text='Delay')
        self.treeTabs[0].pack(fill=BOTH, expand=TRUE)
        self.treeTabs[0].tag_configure('oddrow', background="lightblue")
        self.treeTabs[0].tag_configure('evenrow', background="white")

        self.treeTabs[0].bind("<Button-3>", rCM.showRightClickMenu)
        self.treeTabs[0].bind("<ButtonRelease-1>", selectRow)

    def addTab(self, name):
        self.tabFrame = ttk.Frame(self.notebook)
        tabCount = len(self.treeTabs)

        self.tabScroll = Scrollbar(self.tabFrame)
        self.tabScroll.pack(side=RIGHT, fill=Y)
        self.notebook.add(self.tabFrame, text=name)
        self.treeTabs.append(ttk.Treeview(self.tabFrame, yscrollcommand=self.tabScroll.set, selectmode="extended"))

        self.treeTabs[tabCount]['columns'] = ("X", "Y", "Action", "Delay")

        self.treeTabs[tabCount].column('#0', anchor='center', width=40)
        self.treeTabs[tabCount].column('X', anchor='center', width=60)
        self.treeTabs[tabCount].column('Y', anchor='center', width=60)
        self.treeTabs[tabCount].column('Action', anchor='center', width=120)
        self.treeTabs[tabCount].column('Delay', anchor='center', width=120)

        self.treeTabs[tabCount].heading('#0', text='Step')
        self.treeTabs[tabCount].heading('X', text='X')
        self.treeTabs[tabCount].heading('Y', text='Y')
        self.treeTabs[tabCount].heading('Action', text='Action')
        self.treeTabs[tabCount].heading('Delay', text='Delay/Repeat')
        self.treeTabs[tabCount].pack(fill=BOTH, expand=TRUE)

        self.treeTabs[tabCount].tag_configure('oddrow', background="lightblue")
        self.treeTabs[tabCount].tag_configure('evenrow', background="white")

        self.treeTabs[tabCount].bind("<Button-3>", rCM.showRightClickMenu)
        self.treeTabs[tabCount].bind("<ButtonRelease-1>", selectRow)

        self.notebook.select(self.tabFrame)

    def getTabTree(self, tabNum):
        return self.treeTabs[tabNum]

    def getNotebook(self):
        return self.notebook


class rightClickMenu():
    def __init__(self):
        self.rightClickMenu = Menu(root, tearoff=0)
        self.rightClickMenu.add_command(label="New Macro", command=self.newMacro)
        self.rightClickMenu.add_command(label="Close Macro", command=self.closeTab)
        self.rightClickMenu.add_separator()
        self.rightClickMenu.add_command(label="Move up", command=self.moveUp)
        self.rightClickMenu.add_command(label="Move down", command=self.moveDown)
        self.rightClickMenu.add_command(label="Remove", command=self.removeRow)
        self.rightClickMenu.add_command(label="Select All", command=self.selectAll)

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
        selectedRows = tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).selection()
        for row in selectedRows:
            tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).move(row, tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).parent(row), tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).index(row) - 1)
        reorderRows(tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())))

    def moveDown(self):
        selectedRows = tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).selection()
        for row in reversed(selectedRows):
            tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).move(row, tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).parent(row), tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).index(row) + 1)
        reorderRows(tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())))

    def removeRow(self):
        selectedRows = tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).selection()
        for row in selectedRows:
            tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).delete(row)
        reorderRows(tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())))

    def selectAll(self):
        for row in tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).get_children():
            tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).selection_add(row)

    def closeTab(self):
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


bgColor = None  # '#9ac6e3'
fgColor = None  # '#759fba'
bgButton = None  # '#e0f3ff'

# set icon to cute bunny
root.iconbitmap(os.getcwd() + '\Resources\StickyHeadIcon.ico')
root.title('Fancy Autoclicker')
root.geometry("640x420")
root.attributes("-topmost", True)
root.configure(bg=bgColor)


def checkNumerical(key):
    if key in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'] or key.isnumeric():
        return True
    else:
        return False


def resizeNotebook(self):
    tN.notebook.config(height=root.winfo_height() - 180, width=root.winfo_width() - 210)


vcmd = (root.register(checkNumerical), '%S')
loopEntry = Entry(root, width=5, justify="right", validate='key', vcmd=vcmd)
loopEntry.insert(0, 0)
delayEntry = Entry(root, width=12, justify="right", validate='key', vcmd=vcmd)
delayEntry.insert(0, 0)
actionEntry = Entry(root, width=10, justify="center")
actionEntry.insert(0, 'M1')
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
    try:
        treeView = tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select()))
        sC = threading.Thread(target=startClicking, args=(treeView,))
        sC.start()
    except:
        # if cant multithread might as well single thread
        print("Error: unable to start thread")
        startClicking(None)


# Start Clicking button will loop through active notebook tab's treeview and move mouse and click or type keystrokes with specified delays
def startClicking(treeView):
    # loop though current treeview
    loops.set(int(loopEntry.get()))
    loopsLeft.set(loopEntry.get())
    root.update()

    if treeView:
        children = treeView.get_children()
    else:
        treeView = tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select()))
        children = treeView.get_children()

    while loopsLeft.get() > 0:
        for child in children:
            childValues = treeView.item(child, 'values')
            if childValues[2] in ['M1', 'M2', 'M3']:
                if (int(childValues[0]) != 0 or int(childValues[1]) != 0):
                    # mouse click is action
                    if childValues[2] == 'M1':
                        pyautogui.click(int(childValues[0]), int(childValues[1]), button='left')
                    elif childValues[2] == 'M3':
                        pyautogui.click(int(childValues[0]), int(childValues[1]), button='right')
                    else:
                        pyautogui.click(int(childValues[0]), int(childValues[1]), button='middle')
                else:
                    # mouse click is action without position, do not move, just click
                    if childValues[2] == 'M1':
                        pyautogui.click(button='left')
                    elif childValues[2] == 'M3':
                        pyautogui.click(button='right')
                    else:
                        pyautogui.click(button='middle')

            # Action is #string, find image in Images folder with string name
            elif childValues[2][0] == '#' and len(childValues[2]) > 1:
                # delay/100 is confidence
                confidence = childValues[3]
                # confidence must be a percentile
                if int(confidence) <= 100 and int(confidence) > 0:
                    for a in range(5):
                        #Confidence specified, use Delay as confidence percentile
                        position = pyautogui.locateCenterOnScreen(os.getcwd() + r'\Images' + '\\' + str(childValues[2][1:len(childValues[2])]) + '.png', confidence=int(confidence) / 100)
                        if position: break
                    if position: pyautogui.click(position[0], position[1], button='left')
                    else: break
                else:
                    # confidence could not be determined, use default
                    for a in range(5):
                        position = pyautogui.locateCenterOnScreen(os.getcwd() + r'\Images' + '\\' + str(childValues[2][1:len(childValues[2])]) + '.png', confidence=int(confidence) / 100)
                        if position: break
                    if position: pyautogui.click(position[0], position[1], button='left')
                    else: break

            # Action is !string, run macro with string name for Delay amount of times
            elif childValues[2][0] == '!' and len(childValues[2]) > 1:

                # macro is action, repeat for amount in delay
                for i in range(len(tN.getNotebook().tabs())):
                    # tab with macro found
                    if tN.getNotebook().tab(i, 'text') == childValues[2][1:len(childValues[2])]:
                        startClickingChild(tN.getTabTree(i), int(childValues[3]))
                        break
            else:
                # key press is action
                pyautogui.press(childValues[2])

            if loopsLeft.get() == 0: break
            #Only sleep if row is not macro or image finder
            if childValues[2][0] != '!' and childValues[2][0] != '#' and len(childValues[2]) > 1:
                time.sleep(int(childValues[3]) / 1000)
                print(int(childValues[3]) / 1000)
        if loopsLeft.get() > 0: loopsLeft.set(loopsLeft.get() - 1)
        root.update()


# Start Clicking button will loop through active notebook tab's treeview and move mouse and click or type keystrokes with specified delays
def startClickingChild(treeView, loopsParam):
    # loop though param treeview, for param loops
    # needed so as to not mess with LoopsLeft from outer loop
    root.update()

    children = treeView.get_children()

    # check Loosleft as well to make sure Stop button wasn't pressed since this doesn't ues a global for loop count
    while loopsParam > 0 and loopsLeft.get() != 0:
        for child in children:
            childValues = treeView.item(child, 'values')
            if childValues[2] in ['M1', 'M2', 'M3']:
                if (int(childValues[0]) != 0 or int(childValues[1]) != 0):
                    # mouse click is action
                    if childValues[2] == 'M1':
                        pyautogui.click(int(childValues[0]), int(childValues[1]), button='left')
                    elif childValues[2] == 'M3':
                        pyautogui.click(int(childValues[0]), int(childValues[1]), button='right')
                    else:
                        pyautogui.click(int(childValues[0]), int(childValues[1]), button='middle')
                else:
                    # mouse click is action without position, do not move, just click
                    if childValues[2] == 'M1':
                        pyautogui.click(button='left')
                    elif childValues[2] == 'M3':
                        pyautogui.click(button='right')
                    else:
                        pyautogui.click(button='middle')

            # Action is #string, find image in Images folder with string name
            elif childValues[2][0] == '#' and len(childValues[2]) > 1:
                # delay/100 is confidence
                confidence = childValues[2][3]
                # confidence must be a percentile
                if int(confidence) <= 100 and int(confidence) > 0:
                    for a in range(5):
                        #Confidence specified, use Delay as confidence percentile
                        position = pyautogui.locateCenterOnScreen(os.getcwd() + r'\Images' + '\\' + str(childValues[2][1:len(childValues[2])]) + '.png', confidence=int(confidence) / 100)
                        if position: break
                    if position: pyautogui.click(position[0], position[1], button='left')
                    else: break
                else:
                    # confidence could not be determined, use default
                    for a in range(5):
                        position = pyautogui.locateCenterOnScreen(os.getcwd() + r'\Images' + '\\' + str(childValues[2][1:len(childValues[2])]) + '.png', confidence=int(confidence) / 100)
                        if position: break
                    if position: pyautogui.click(position[0], position[1], button='left')
                    else: break

            # Action is !string, run macro with string name for Delay amount of times
            elif childValues[2][0] == '!' and len(childValues[2]) > 1:
                # macro is action, repeat for amount in delay
                for i in range(len(tN.getNotebook().tabs())):
                    # tab with macro found
                    if tN.getNotebook().tab(i, 'text') == childValues[2][1:len(childValues[2])]:
                        startClickingChild(tN.getTabTree(i), int(childValues[3]))
                        break
            else:
                # key press is action
                pyautogui.press(childValues[2])

            if loopsParam == 0 or loopsLeft.get(): break
            #Only sleep if row is not macro or image finder
            if childValues[2][0] != '!' and childValues[2][0] != '#' and len(childValues[2]) > 1: time.sleep(int(childValues[3]) / 1000)
        if loopsParam > 0: loopsParam = loopsParam - 1
        root.update()


def stopClicking():
    loopsLeft.set(0)


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
    if len(tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).get_children()) % 2 == 0:
        tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).insert(parent='', index='end', iid=None, text=len(tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).get_children()) + 1, values=(x.get(), y.get(), actionEntry.get(), int(delayEntry.get())), tags='evenrow')
    else:
        tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).insert(parent='', index='end', iid=None, text=len(tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).get_children()) + 1, values=(x.get(), y.get(), actionEntry.get(), int(delayEntry.get())), tags='oddrow')


def addRowWithParams(xParam, yParam, keyParam, delayParam):
    # for import to populate new treeview
    if len(tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).get_children()) % 2 == 0:
        tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).insert(parent='', index='end', iid=None, text=len(tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).get_children()) + 1, values=(xParam, yParam, keyParam, delayParam), tags='evenrow')
    else:
        tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).insert(parent='', index='end', iid=None, text=len(tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).get_children()) + 1, values=(xParam, yParam, keyParam, delayParam), tags='oddrow')


def showRightClickMenu(event):
    # show option menu when right clicking treeview of rows
    try:
        rightClickMenu.tk_popup(event.x_root, event.y_root)
    finally:
        rightClickMenu.grab_release()


def selectRow(event):
    # this event popuates the entries with the values of the selected row for easy editing/copying
    #must use global variables to tell if new row selected or just clikcing whitespace in treeview, do not set values if not changing row selection
    global previouslySelectedTab
    global previouslySelectedRow

    selectedRow = tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).focus()
    selectedValues = tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).item(selectedRow, 'values')

    if len(selectedValues) > 0 and (previouslySelectedTab != tN.getNotebook().index(tN.getNotebook().select()) or previouslySelectedRow != selectedRow):
        delayEntry.delete(0, 'end')
        delayEntry.insert(0, selectedValues[3])
        actionEntry.delete(0, 'end')
        actionEntry.insert(0, selectedValues[2])
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
    filename = filedialog.askopenfilename(initialdir=os.getcwd() + r'\Macros', title="Select a .csv file", filetypes=(("csv files", "*.csv"),))
    answer = True
    found = 0

    # look for tab with same name as imported macro, will overwrite that tab if imported
    # this will ensure each macro has a unique name so that when a macro calls another macro there is no confusion over which macro should be called
    for i in range(len(tN.getNotebook().tabs())):
        if tN.getNotebook().tab(i, 'text') == os.path.splitext(os.path.basename(filename))[0]:
            answer = askyesno("Overwrite macro?", "Are you sure you want to overwrite current " + os.path.splitext(os.path.basename(filename))[0])
            found = i + 1
            break

    # only open if file exists and overwrite true (answer defaults to true in case it is not asked)
    if filename and answer:
        with open(filename, 'r') as csvFile:
            csvReader = csv.reader(csvFile)
            if not found:
                tN.addTab(os.path.splitext(os.path.basename(filename))[0])
            else:
                tN.notebook.select(found - 1)
                tN.getTabTree(found - 1).delete(*tN.getTabTree(found - 1).get_children())
            for line in csvReader:
                if line[3]:
                    addRowWithParams(line[0], line[1], line[2], line[3])


def exportMacro():
    # save as csv file with name of file as macro name
    savePath = os.getcwd() + r'\Macros'
    finename = str(tN.getNotebook().tab(tN.getNotebook().select(), 'text').replace(r'/', r'-').replace(r'\\', r'-').replace(r'*', r'-').replace(r'?', r'-').replace(r'[', r'-').replace(r']', r'-').replace(r'<', r'-').replace(r'>', r'-').replace(r'|', r'-'))

    with open(os.path.join(savePath, str(finename + '.csv')), 'w', newline='') as newMacro:
        csvWriter = csv.writer(newMacro, delimiter=',')
        children = tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).get_children()

        for child in children:
            childValues = tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).item(child, 'values')
            csvWriter.writerow(childValues)


def actionPopulate(event):
    # ! is special character that allows typing of action instead of instating setting action to each key press
    if str(actionEntry.get())[0:1] != '!' and str(actionEntry.get())[0:1] != '#':
        # need to use different properties for getting key press for letters vs whitespace/ctrl/shift/alt
        # != ?? to exclude mouse button as their char and keysym are
        if str(event.char).strip() and event.char != '??':
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
    rows = tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).selection()
    for row in rows:
        tN.getTabTree(tN.getNotebook().index(tN.getNotebook().select())).item(row, values=(x.get(), y.get(), actionEntry.get(), delayEntry.get()))

def showHelp():
    helpWindow = Toplevel(root)
    helpWindow.wm_title("Sticky's Autoclicker Help")

    helplabel = Label(helpWindow, text=usabilityNotes, justify=LEFT)
    helplabel.grid(row=0, column=0)

    closeButton = ttk.Button(helpWindow, text="Close", command=helpWindow.destroy)
    closeButton.grid(row=1, column=0)


# Column 0
# for reference: notebook.grid(row=3, column=2, columnspan=3, rowspan=5, sticky='')
welcomeLabel = Label(root, text="Sticky's Autoclicker", font=("Arial bold", 11), bg=bgColor).grid(row=0, column=0, padx=5, pady=0, sticky='nw', columnspan=2)
madeByLabel = Label(root, text="Made by Colin Knappert", bg=bgColor).grid(row=0, column=0, padx=20, sticky='w', columnspan=2)
clickLabel = Label(root, text="Click Loops", font=("Arial", 10), bg=bgColor).grid(row=1, column=0, sticky='e', pady=10)
LoopsLeftLabel = Label(root, text="Loops Left", font=("Arial", 11), bg=bgColor).grid(row=2, column=0, sticky='s')
clicksLeftLabel = Label(root, textvariable=loopsLeft, bg=bgColor).grid(row=3, column=0, sticky='n')
startButton = Button(root, text="Start Clicking", font=("Arial", 15), command=threadStartClicking, padx=0, pady=10, borderwidth=6, bg=bgButton).grid(row=4, column=0, columnspan=2, sticky="n")
stopButton = Button(root, text="Stop Clicking", font=("Arial", 15), command=stopClicking, padx=0, pady=10, borderwidth=6, bg=bgButton).grid(row=5, column=0, columnspan=2, sticky="n")

exportMacroButton = Button(root, text="Export Macro", bg=bgButton, command=exportMacro, borderwidth=2).grid(row=6, column=0, sticky='n')
importMacroButton = Button(root, text="Import Macro", bg=bgButton, command=importMacro, borderwidth=2).grid(row=6, column=1, sticky='n')
helpButton = Button(root, text="Help", bg=bgButton, command=showHelp, borderwidth=2).grid(row=7, column=0, sticky='n', columnspan=2)

# Column 1
loopEntry.grid(row=1, column=1, sticky='', pady=10)
loopsLabel = Label(root, text="Total Loops", font=("Arial", 11), bg=bgColor).grid(row=2, column=1, sticky='s')
loops2Label = Label(root, textvariable=loops, bg=bgColor).grid(row=3, column=1, sticky='n')

# Column 2
insertPositionButton = Button(root, text="  Insert Position   ", bg=bgButton, command=addRow, padx=0, pady=0, borderwidth=2).grid(row=0, column=2, pady=6, sticky='n')
getCursorButton = Button(root, text=" Choose Position ", bg=bgButton, command=getCursorPosition, borderwidth=2).grid(row=0, column=2, sticky='s')
editRowButton = Button(root, text="Overwrite Row(s)", bg=bgButton, command=overwriteRows, borderwidth=2).grid(row=1, column=2, sticky='')

# Column 3
xPosTitleLabel = Label(root, text="X position", font=("Arial", 11), bg=bgColor).grid(row=0, column=3, pady=0, sticky='nw')
xPosLabel = Label(root, textvariable=x, bg=bgColor).grid(row=0, column=3, padx=35, pady=25, sticky='w')
yPosTitleLabel = Label(root, text="Y Position", font=("Arial", 11), bg=bgColor).grid(row=0, column=3, padx=0, pady=0, sticky='sw')
yPosLabel = Label(root, textvariable=y, bg=bgColor).grid(row=1, column=3, padx=35, pady=0, sticky='nw')

# Column 4
timeLabel = Label(root, text="Delay (ms)", font=("Arial", 10), bg=bgColor).grid(row=0, column=4, padx=0, sticky="n")
delayEntry.grid(row=0, column=4, padx=0, sticky='')

actionLabel = Label(root, text="Action", font=("Arial", 10), bg=bgColor, pady=0).grid(row=0, column=4, padx=0, sticky="s")
actionEntry.grid(row=1, column=4, padx=0, pady=10, sticky='')
# kind this entry to all keyboard and mouse actions
actionEntry.bind("<Key>", actionPopulate)
actionEntry.bind('<Return>', actionPopulate, add='+')
actionEntry.bind('<Button-1>', actionPopulate, add='+')
actionEntry.bind('<Button-2>', actionPopulate, add='+')
actionEntry.bind('<Button-3>', actionPopulate, add='+')

delayToolTip = CreateToolTip(delayEntry, "Delay in milliseconds to take place after click. If macro is specified this will be loop amount of that macro.")
actionToolTip = CreateToolTip(actionEntry, "Action to occur before delay. Accepts all mouse buttons and keystrokes. Type !macroname to call macro in another tab with delay amount as loops.")


# mainloop and instantiation
tN = treeviewNotebook(root)
# bnd after instantiation because resizing occurs while building window before all elements are even created
root.bind("<Configure>", resizeNotebook)

root.update_idletasks()
root.mainloop()