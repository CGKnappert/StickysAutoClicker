import tkinter as tk
from tkinter import Menu

class RightClickMenu():
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