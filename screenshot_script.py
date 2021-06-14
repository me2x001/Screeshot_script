import os
import shutil
import keyboard
import pyautogui as py
from docx import Document
from docx.shared import Inches

n=1
screenshots=[]
toInsert=[]

def takeScreenshot():
    def fileName():
        global n
        name='Screenshot'
        name=name+str(n)+'.png'
        n+=1
        return name
    
    global screenshots
    ss=py.screenshot()
    fn=fileName()
    ss.save(fn)
    cwd=os.getcwd()
    path=cwd+'\\'+fn
    try:
        screenshots.append(path)
    except:
        print('Unable to take screenshot')
    else:
        print(fn+' is captured')
    # print(cwd)
    # print(screenshots)

def insert():
    global toInsert
    try:
        try:
            ss=screenshots[-1]
            fname=os.path.split(ss)
            if ss not in toInsert:
                toInsert.append(ss)
            else:
                raise Exception
        except:
            print(fname[1]+' is already Inserted')
        else:
            print(fname[1]+' is Inserted')
    except:
        print('Please take screenshot first.')
    
    # print(toInsert)
       
def makeDoc(fname,rPath):

    global toInsert
    # print(fname)
    
    document = Document()

    for ss in toInsert:
        document.add_picture(ss, width=Inches(6.5))
    
    try:
        document.save(rPath+'\\'+fname+'.docx')
    except:
        print('Error in saving file')
    else:
        print('Dcoument Generated')
        os.chdir(rPath)
        shutil.rmtree(rPath+'\\temp')   
        os._exit(0)

rPath=os.getcwd()
cwd=os.getcwd()
cwd=cwd+'\\temp'
try:
    os.mkdir(cwd)
except:
    print('Temp directory already exists please delete and continue')

os.chdir(cwd)
# print(cwd)

print("Press following keys to perform respective actions\n")
print('Ctrl+x       -   Take Screenshot')
print('Insert       -   To insert screenshot in doc')
print('Ctrl+Enter   -   Generate doc')
print("\nEnter File Name : ", end='')
fname=input()

keyboard.add_hotkey('ctrl+x', takeScreenshot)
keyboard.add_hotkey('insert', insert)
keyboard.add_hotkey('ctrl+enter', lambda: makeDoc(fname,rPath))

print("Press ESC to stop.")
keyboard.wait('esc')