#!/usr/bin/env python3
import tk
import tkinter

import common

def main():
    win = tkinter.Tk()
    app = common.MainWin(win)
    win.mainloop()
    try:
        win.destroy()
    except tkinter.TclError:
        pass

if __name__ == '__main__':
    main()

