from pynput import keyboard

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