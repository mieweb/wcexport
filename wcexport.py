#!/usr/bin/env python3
import tk
import tkinter

import common
import argparse

def main():
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description="WebChart Export Utility")
    parser.add_argument("url", nargs="?", help="WebChart URL (optional)")
    parser.add_argument("username", nargs="?", help="WebChart username (optional)")
    args = parser.parse_args()

    win = tkinter.Tk()
    app = common.MainWin(win)

    # Set URL and username if provided
    if args.url:
        app.url.insert(0, args.url)
    if args.username:
        app.username.insert(0, args.username)

    win.mainloop()
    try:
        win.destroy()
    except tkinter.TclError:
        pass

if __name__ == '__main__':
    main()

