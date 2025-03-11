import tkinter as tk
from tkinter import ttk, RIGHT, Y, END
from tkinter import Toplevel
import os
import sys
import ctypes
import screeninfo
from ctypes import windll
# config file support
import configparser as cp
# memory leak detection
import tracemalloc
import objgraph
from pympler import summary, muppy


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
                  "           If image is not found then loop will end and next loop will start. This is to prevent the rest of the loop from going awry because the expected image was not found.\n"
                  "           Ending the #imageName in a ? will allow the finding of an image to be optional and not exsit the macro if not found. This allows for repeat attempts or alternate paths to be taken.\n"
                  " - Action also allows underscore _ as a special character that will indicate the following key(s) should be pressed and held for the set amount of time in the Delay field.\n"
                  "           Note that for these rows the Delay no longer delays after the key is held.\n"
                  "           You can hold multiple keys at a time by continuing to type into the Action field once a _ has been entered. A | will delineate the different keys to be held.\n"
                  "           Typing _ into action can also reset the field to remove keys you do not want pressed since backspace doesn't work.\n"
                  " - The Record functionality will begin entering rows to the end of the current macro tab reading key presses and delays.\n"
                  "           This functionality is quite accurate (for python) and can be quite efficien especially when paired with manual edits to shorten or remove unnessecary delays or actions.\n"
                  "           For playback of recordings it is highly recommended to use the Busy Wait option in the Settings menu.\n"
                  "           Busy Wait allows the delay to be accurate to around one millisecond where as non-Busy Wait is accurate to around 10-15 ms.\n"
                  "           The downside of Busy Wait and why it shouldn't always be used is that it incurs heavy CPU usage and can be felt by users and other programs.")


class Logger():
    stdout = sys.stdout
    messages = ""

    def start(self): 
        sys.stdout = self

    def stop(self): 
        sys.stdout = self.stdout

    def write(self, text): 
        self.messages += text


class Titlebar():
    # Class for addingg titlebars after overrideredirect removes them
    # Needed for greater control of buttons and theme
    def __init__(self, root, pack, parent, icon, title_text, minimize, maximize, close, help, file_path):
        self.root = root
        self.parent = parent
        self.file_path = file_path
        self.icon = icon
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
            self.helpWindow = helpWindow(self.root, self.parent, self.file_path, self.icon)
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
    def __init__(self, root, parent, file_path, icon):
        super().__init__(root)
        self.root = root
        self.parent = parent
        self.file_path = file_path
        
        # create new window from parent
        self.helpWindow = Toplevel(self.root)
        self.helpWindow.overrideredirect(True)
        # Get previous position and set window to load in it
        config = cp.ConfigParser()
        config.read(os.path.join(self.file_path, r'config.ini'))
        if config.has_option("Position", "helpx") and config.has_option("Position", "helpy"):
            self.helpWindow.geometry("+" + config.get("Position", "helpx") + "+" + config.get("Position", "helpy"))

        self.titlebar = Titlebar(self.helpWindow, True, self, icon, "Sticky's Autoclicker Help", True, False, True, False, self.file_path)
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
        config.read(os.path.join(self.file_path, 'config.ini'))
        config.set('Position', 'helpx', str(self.helpWindow.winfo_rootx()))
        config.set('Position', 'helpy', str(self.helpWindow.winfo_rooty()))
        with open(os.path.join(self.file_path, r'config.ini'), 'w') as configfile:
            config.write(configfile)
        self.helpWindow.destroy()
        self.helpWindow = None
        self.parent.helpWindow = None
        self.root.focus()


class settingsWindow(ttk.Frame):
    # window to display options and handle changing of options
    def __init__(self, root, parent, file_path, icon):
        super().__init__(root)
        self.root = root
        self.parent = parent
        self.file_path = file_path
        
        # create new window from parent
        self.settingsWindow = Toplevel(self)
        self.settingsWindow.overrideredirect(True)
        # Get previous position and set window to load in it
        config = cp.ConfigParser()
        config.read(os.path.join(self.file_path, r'config.ini'))
        if config.has_option("Position", "settingsx") and config.has_option("Position", "settingsy"):
            self.settingsWindow.geometry("+" + config.get("Position", "settingsx") + "+" + config.get("Position", "settingsy"))

        self.titlebar = Titlebar(self.settingsWindow, False, self, icon, "Settings", True, False, True, False, self.file_path)
        self.settingsWindow.resizable(height=None, width=None)
        self.settingsWindow.wm_title("Sticky's Autoclicker Settings")
        self.settingsWindow.protocol("WM_DELETE_WINDOW", self.onClose)
        self.settingsWindow.iconphoto(False, icon)
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

        self.hiddenLabel = ttk.Label(self.settingsFrame, text="Developer Mode")
        self.hiddenLabel.grid(row=1, column=4, sticky='s', padx=10)
        self.hiddenButton = ttk.Checkbutton(self.settingsFrame, variable=parent.developerMode, onvalue=1, offvalue=0, command=parent.toggleDeveloper)
        self.hiddenButton.grid(row=2, column=4, sticky='n', padx=10, pady=10)

        # Undoing this setting for a while, perhaps until I figure out how python can send commands to asleep MS windows
        # self.hiddenLabel = ttk.Label(self.settingsFrame, text="Use Hidden Mode")
        # self.hiddenLabel.grid(row=1, column=4, sticky='s', padx=10)
        # self.hiddenButton = ttk.Checkbutton(self.settingsFrame, variable=parent.hiddenMode, onvalue=1, offvalue=0, command=parent.toggleHidden)
        # self.hiddenButton.grid(row=2, column=4, sticky='n', padx=10, pady=10)

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
        config.read(os.path.join(self.file_path, 'config.ini'))
        config.set('Position', 'settingsx', str(self.settingsWindow.winfo_rootx()))
        config.set('Position', 'settingsy', str(self.settingsWindow.winfo_rooty()))
        with open(os.path.join(self.file_path, r'config.ini'), 'w') as configfile:
            config.write(configfile)
        self.settingsWindow.destroy()
        self.parent.settingsWindow = None
        self.settingsWindow = None
        self.root.focus()


class logWindow(ttk.Frame):
    # window for viewing click history and errors
    def __init__(self, root, parent, text, file_path, icon):
        super().__init__(root)
        self.root = root
        self.parent = parent
        self.file_path = file_path
        
        # create new window from parent
        self.logWindow = Toplevel(self)
        # Get previous position and set window to load in it
        config = cp.ConfigParser()
        config.read(os.path.join(self.file_path, r'config.ini'))
        if config.has_option("Position", "logx") and config.has_option("Position", "logy"):
            self.logWindow.geometry("480x600" + "+" + config.get("Position", "logx") + "+" + config.get("Position", "logy"))

        self.logWindow.overrideredirect(True)
        self.titlebar = Titlebar(self.logWindow, True, self, icon, "Sticky's Autoclicker Log", True, True, True, False, self.file_path)
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
        self.text.insert("1.0", text)

        if parent.developerMode:
            # Developer mode will display memory usage in Log window 
            log = Logger()
            log.start()
            print("muppy top 100 summary:\n")
            all_objects = muppy.get_objects()
            sum_text = summary.summarize(all_objects)
            summary.print_(sum_text)    
            
            print("\n\n\nobjgraph top 100 stats:\n")
            objgraph.show_most_common_types()
            snapshot = tracemalloc.take_snapshot() 
            top_stats = snapshot.statistics('lineno') 
            
            log.stop()
            self.updateText(log.messages) 

        # sizegrip could not lift above either scrollbar despite the config being essentially the same as on the main windows notebook
        # self.grip = ttk.Sizegrip()
        # self.grip.place(relx=1.0, rely=1.0, anchor="se")
        # self.grip.lift(self.text)
        # self.grip.bind("<B1-Motion>", self.moveMouseButton)


    def updateText(self, text):
        self.text.delete("1.0", END)
        self.text.insert("1.0", text)


    def moveMouseButton(self, e):
        x1=self.root.winfo_pointerx()
        y1=self.root.winfo_pointery()
        x0=self.root.winfo_rootx()
        y0=self.root.winfo_rooty()

    def onClose(self):
        # print("closelog")
        config = cp.ConfigParser()
        config.read(os.path.join(self.file_path, 'config.ini'))
        config.set('Position', 'logx', str(self.logWindow.winfo_rootx()))
        config.set('Position', 'logy', str(self.logWindow.winfo_rooty()))
        with open(os.path.join(self.file_path, r'config.ini'), 'w') as configfile:
            config.write(configfile)
        self.logWindow.destroy()
        self.parent.logWindow = None
        self.logWindow = None
        self.root.focus()