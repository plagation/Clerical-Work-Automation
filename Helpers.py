# -*- coding: utf-8 -*-
"""
Created on Fri Aug  6 17:18:39 2021

Contains a collection of helper modules for classes used in application

@author: Kyle Conrad
"""

import subprocess, sys, traceback, tkinter

from tkinter import filedialog
from pathlib import Path

class Helpers:
    
    class PowershellException(Exception):
        pass
    
    class Email:
        dirPath = str(Path.cwd())
        
        def main(path, *args):
            cmd = ["PowerShell", "-ExecutionPolicy", "Bypass", "-File", path, *args]
            print(cmd)
            ec = subprocess.call(cmd)
            
            if ec != 0:
                print("Powershell returned: {0:d}".format(ec))
                raise Helpers.PowershellException
                   
        def run(self, path):
            if __name__ == "__main__":
                print("Python {0:s} {1:d}bit on {2:s}\n".format(" ".join(item.strip() for item in sys.version.split("\n")), 64 if sys.maxsize > 0x100000000 else 32, sys.platform))
                self.main(path)
                print("\nDone.")
                
        def traceback_report(self):   
            path = self.dirPath + "\\TracebackEmail.ps1"
            print(path)
            Helpers.Email.main(path, "-traceback", traceback.format_exc())
            
    class FileBrowser:   
        def get_file():
            #create top-level window instance then hide it
            win = tkinter.Tk()
            win.withdraw()
            
            #note that int variable sets the transparency
            win.attributes('-alpha', 0)
            win.overrideredirect
            
            win.deiconify()
            win.lift()
            win.focus_force()
            
            folderPath = filedialog.askopenfilename(parent=win)
            
            win.destroy()
            return folderPath
        
    class Exceptions:
        def user_interrupt(driver):
            print("Interrupted by user")
            driver.quit()
            sys.exit(1)
        
        def unexpected_exception(msg):
            print("An unexpected exception occurred while attempting to " + msg + ":\n")
            print(traceback.print_exc())
            Helpers.Email().traceback_report()


    