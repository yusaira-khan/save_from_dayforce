import re

import pywin
import pywin.framework
import pywin.framework.interact

import win32api
import win32con as wc
import win32gui
import win32ui


class WindowFinder:
    """Class to find and make focus on a particular Native OS dialog/Window """
    def __init__ (self):
        self.handle = None

    def find_window(self, class_name, window_name = None):
        """Pass a window class name & window name directly if known to get the window """
        self.handle = win32gui.FindWindow(class_name, window_name)


    def _window_enum_callback(self, hwnd, wildcard):
        '''Call back func which checks each open window and matches the name of window using reg ex'''
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        """ This function takes a string as input and calls EnumWindows to enumerate through all open windows """
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)
    def set_foreground(self):
        """Get the focus on the desired open window"""
        win32gui.SetForegroundWindow(self.handle)

ia = pywin.framework.interact.edit.currentView
ia.SetFocus()
ia.AppendToPrompt(["a.pdf"])
ia.ProcessEnterEvent(None)


