'''
This module is creates GUI windows, so user can select options.
'''

from Tkinter import *
from sys import exit

class MainWindow(object):
    '''
    This is the parent class for GUI windows. The purpose of window objects is to receive a lsit of options, visually-
    open a dialog, allow the user to select one or more options (depending of child class), and return said value(s).
    '''
    def __init__(self, options):
        self.master = Tk()
        self.master.title("Avi's Intuit Interchage File (.iif) creater")
        self.width = 100
        self.height = 800
        self.button_width = 150
        self.button_height = 40
        self.half_button_height = self.button_height / 2
        self.canvas = Canvas(self.master, width=self.width, height=self.height)
        self.options = list(options)
        self.selected_options = []
        self.create_buttons()
        self.canvas.pack()
        self.canvas.bind("<Button-1>", self.callback_left)
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.master.mainloop()

    def callback_left(self, event):
        '''
        Processes mouse input for left-click. Will be overrode by child-class method
        '''
        pass

    def create_buttons(self):
        '''
        Visually creates buttons. One per account
        '''
        i = 0
        self.options.sort()
        for option in self.options:
            y1 = i * self.button_height
            y2 =  i * self.button_height + self.button_height
            self.canvas.create_rectangle(0,y1,self.width,y2, fill='grey')
            self.canvas.create_text(50,y1 + self.half_button_height, text=option, fill="black")
            i += 1

    def visual_select_box(self,y_axis, text):
        y_axis_2 = y_axis + self.button_height
        self.canvas.create_rectangle(0,y_axis,self.width,y_axis_2, fill='green')
        self.canvas.create_text(50,y_axis + self.half_button_height, text=text, fill="black")

    def visual_deselect_box(self,y_axis, text):
        y_axis_2 = y_axis + self.button_height
        self.canvas.create_rectangle(0,y_axis,self.width,y_axis_2, fill='grey')
        self.canvas.create_text(50,y_axis + self.half_button_height, text=text, fill="black")

    def on_closing(self):
        '''
        This method is executed when window closes.
        '''
        if len(self.selected_options) == 0:
            print 'No options selected. Ending script...'
            exit()
        else:
            self.master.destroy()


class SingleSelectionWindow(MainWindow):
    '''
    This window class allows user to select only one option.
    '''
    def __init__(self, options):
        MainWindow.__init__(self, options)

    def callback_left(self, event):
        '''
        Processes mouse input for left-click.
        '''
        click_height = event.y
        i = 0
        for option in self.options:
            y1 = i * self.button_height
            y2 =  i * self.button_height + self.button_height
            if y2 >= click_height >= y1:
                self.visual_select_box(y1, option)
                self.selected_options.append(option)
            else:
                if option in self.selected_options:
                    self.selected_options.remove(option)
                self.visual_deselect_box(y1, option)
            i += 1


class MultiSelectionWindow(MainWindow):
    '''
    This window class allows user to select multiple options.
    '''
    def __init__(self, options):
        MainWindow.__init__(self, options)

    def callback_left(self, event):
        '''
        Processes mouse input for left-click.
        '''
        click_height = event.y
        i = 0
        for option in self.options:
            y1 = i * self.button_height
            y2 =  i * self.button_height + self.button_height
            if y2 >= click_height >= y1:
                if option in self.selected_options:
                    self.selected_options.remove(option)
                    self.visual_deselect_box(y1, option)
                else:
                    self.selected_options.append(option)
                    self.visual_select_box(y1, option)
            i += 1