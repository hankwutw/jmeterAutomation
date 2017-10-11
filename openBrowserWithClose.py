import webbrowser
import csv
import time
import os
import win32com.client

shell = win32com.client.Dispatch("WScript.Shell")
shell.SendKeys("^w") # CTRL+A may "select all" depending on which window's focused


path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
exampleFile = open('searchUrl_2.txt')
exampleReader = csv.reader(exampleFile)
exampleData = list(exampleReader)

for item in exampleData:
    webbrowser.get(path).open("https://www.momoshop.com.tw"+item[0],new=0)
    time.sleep(5)
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys("^w") # CTRL+A may "select all" depending on which window's focused

    #os.system("taskkill /im chrome.exe /f")
    #webbrowser.close()
   
