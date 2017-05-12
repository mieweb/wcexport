#!/usr/bin/env python
import Tkinter as tk

import common

def main():
	win = tk.Tk()
	app = common.MainWin(win, False)
	win.mainloop()
	try:
		win.destroy()
	except tk.TclError:
		pass

if __name__ == '__main__':
	main()
