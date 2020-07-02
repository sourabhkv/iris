import tkinter as tk
import wolframalpha
import win32com.client,os,random
import wikipedia,webbrowser,keyboard
import speech_recognition as sr
import time
r=sr.Recognizer()

#_____________________________________________________________________________________________________________________ 
def recogn():
    with sr.Microphone() as source:
        audio=r.listen(source)
        try:
            text=r.recognize_google(audio)
            return text
        except:
            return False
#_____________________________________________________________________________________________________________________ 
def opener():
    webbrowser.open('https://github.com/sourabhkv?tab=repositories')
#_____________________________________________________________________________________________________________________ 
def wiki(n):
    a=wikipedia.summary(n,sentences=2)
    return a
#_____________________________________________________________________________________________________________________ 
def say(n):
    speak=win32com.client.Dispatch("Sapi.SpVoice")
    speak.Voice=speak.GetVoices().Item(1)
    speak.Speak(n)
#_____________________________________________________________________________________________________________________     
def wolf(n):
    client = wolframalpha.Client('42XG9Q-L76KJ35T39')
    res =client.query(n)
    output =next(res.results).text
    a=str(output)
    return a
#_____________________________________________________________________________________________________________________ 
def assist(cmd):
    if cmd==False or cmd=="":
        say('error')
    else:
        try:
            p=0
            if "open" in cmd:
                cmd=cmd.lower()
                apps={'powerpoint':'C:/Program Files/Microsoft Office/root/Office16/POWERPNT.exe',
                      'control panel':'C:/Windows/System32/control.exe',
                      'word':'C:/Program Files/Microsoft Office/root/Office16/WINWORD.exe',
                      'excel':'C:/Program Files/Microsoft Office/root/Office16/EXCEL.exe',
                      'command prompt':'C:/Windows/system32/cmd.exe',
                      'notepad':'C:/Windows/system32/notepad.exe',
                      'paint':'C:\Windows\system32\mspaint.exe'}
                for i in apps:
                    if str(i) in cmd:
                        p=1
                        text3="opening"+str(i)
                        say(text3)
                        os.startfile(apps[i])
                        break
                if p==0:
                    if "explorer" in cmd:
                        say('opening fil explorer')
                        keyboard.press_and_release("win+e")
                        p=1
                    elif "settings" in cmd:
                        say("opening settings")
                        keyboard.press_and_release("win+i")
                        p=1
                    else:
                        say("Sorry, I can't find that")
                        p=1
            if "play" in cmd or "game" in cmd:
                url1="https://www.google.com/search?q=tic+tac+toe"
                url2="https://www.google.com/search?q=minesweeper"
                url3="https://www.google.com/search?q=play%20snake"
                url4="https://doodlecricket.github.io"
                url5='https://quickdraw.withgoogle.com/'
                url6="https://www.bing.com/fun/chess"
                r5=random.randint(0,5)
                url=[url1,url2,url3,url4,url5,url6]
                url=url[r5]
                say("OK let's play a game")
                webbrowser.open(url)
                p=1
            if "lock" in cmd and "windows" in cmd:
                keyboard.press("win")
                keyboard.press("l")
                time.sleep(0.15)
                keyboard.release("win")
                keyboard.release("l")
                p=1
            if "close" in cmd or "turn off"in cmd or "shutdown" in cmd and "computer" in cmd :
                os.system("shutdown -s")
                say("shutting down in  minute")
                p=1
            #____________________________________________________________________________________________________
            try:
                if p==0:
                    res="According to wolframalpha "+str(wolf(cmd))
                    if res!="":
                        p=1
                        result=""
                        if len(res)>80:
                            for i in range(0,len(res),80):
                                result=result+res[i:i+80]+"\n"
                        return result
            except:
                if p==0:
                    res="According to wikipedia "+str(wiki(cmd))
                    if res!=0:
                        p=1
                        result=""
                        if len(res)>80:
                            for i in range(0,len(res),80):
                                result=result+res[i:i+80]+"\n"
                        return result
        except:
            if p==0:
                link='https://www.google.com/search?q='+str(cmd)
                webbrowser.open(link)
#_____________________________________________________________________________________________________________________ 
def opener():
    webbrowser.open('https://github.com/sourabhkv?tab=repositories')
#_____________________________________________________________________________________________________________________ 
def assistrec():
    cmd=recogn()
    if cmd!=False:
        say("you said"+cmd)
        res2=assist(cmd)
        return res2
#_____________________________________________________________________________________________________________________ 
root= tk.Tk()
canvas1 = tk.Canvas(root, width = 600, height = 500,  relief = 'raised')
canvas1.pack()
label1 = tk.Label(root, text='IRIS')
label1.config(font=('algerian', 30))
canvas1.create_window(300, 45, window=label1)

label2 = tk.Label(root, text='Type your query:')
label2.config(font=('helvetica', 10))
canvas1.create_window(300, 100, window=label2)

entry1 = tk.Entry (root) 
canvas1.create_window(290, 140, window=entry1,height=25,width=200)
#_____________________________________________________________________________________________________________________ 
def getRoot ():
    
    x1 = entry1.get()
    res=assist(x1)
    
    label3 = tk.Label(root, text= 'Result:\n',font=('helvetica', 10))
    canvas1.create_window(300, 210, window=label3)
    
    label4 = tk.Label(root, text= res,font=('helvetica', 10))
    canvas1.create_window(300, 300, window=label4)
    say(res)
#_____________________________________________________________________________________________________________________ 

def getsp():
    
    res=assistrec()
    
    label3 = tk.Label(root, text= 'Result:\n',font=('helvetica', 10))
    canvas1.create_window(300, 210, window=label3)
    
    label4 = tk.Label(root, text= res,font=('helvetica', 10))
    canvas1.create_window(300, 300, window=label4)
    say(res)
#_____________________________________________________________________________________________________________________ 
button1 = tk.Button(text='Go', command=getRoot, bg='red', fg='white', font=('helvetica', 9, 'bold'))
canvas1.create_window(400, 140, window=button1)
root.title("IRIS")
button2 = tk.Button(text='Voice search', command=getsp, bg='orange', fg='white', font=('helvetica', 9, 'bold'))
canvas1.create_window(244, 168, window=button2)
button3 = tk.Button(text='Github repository', command=opener, bg='blue', fg='white', font=('helvetica', 9, 'bold'))
canvas1.create_window(348, 168, window=button3)
root.mainloop()
