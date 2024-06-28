import pyttsx3 
import speech_recognition as sr
import datetime
import sys
import os

engine=pyttsx3.init('sapi5')
voice=engine.getProperty('voices')
engine.setProperty("voice",voice[1].id)

#text to speech conversion process
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

#taking command from the user in the form at of query and generating the result in the form of speech and printing it in front of the user 
def takecommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....")
        r.pause_threshold=1
        audio=r.listen(source,timeout=5,phrase_time_limit=5)
    try:
        print("Recognizing.....")
        query=r.recognize_google(audio,language='en-in')
        print(f"user said: {query}")
    except Exception as e:
        speak("Say that query agin please")
        return "none"
    return query

#part of wishing the user as per the data and time and as per the requirement

def wish():
    hour=int(datetime.datetime.now().hour)
    
    if hour>0 and hour<=12:
        speak(f"Good Morning {user_name}")
    elif hour>12 and hour<18:
        speak(f"Good afternoon {user_name}")
    else:
        speak(f"Good evening {user_name}")
    speak("Please tell me how may I help you")
        
if __name__=="__main__":
    speak("Hello welcome to the world of Virtual assistant please tell how may I call you")
    user_name=input("enter your name:")
    wish();
    
    while True:
        
        query=takecommand().lower()
        
        #for opening the notepad and opearting the resources of it in order to perform the opeartion 
        if "open notepad" in query:
            apath="C:\\WINDOWS\\system32\\notepad.exe"
            from pywinauto import application
            app=application.Application()
            app.start(f"{apath}")
            dlg=app.window(title="Untitled - Notepad")
            speak("UNTITLED FILE IS CREATED SUCCESFULLY DO YOU WANT TO DO ANYTHING ELSE")
            while True:
                
                query=takecommand().lower()
                
                #opeartion for saving the contents of the file at specific location 
                
                if ("want to save the file" or "save the file" or "save this file" or "saving opeartion" or "want to save this file") in query:
                    app.Notepad.menu_select("File->SaveAs")
                    speak("Ok please write the name of file along with the extension with which you want to save it")
                    file1=input("enter the name of file:")
                    app.SaveAs.edit.set_edit_text(f"{file1}")
                    app.SaveAs.Save.click(double=True)
                    speak("file is created and saved succesfully in the desired folder")
                    #waiting for the user to close the opened dialog box 
                    app.SaveAs.wait_out('enabled')
                #process for creating a new file at specific location 
                
                elif ("create a new file" or "create a file" or "create file" or "want to create a new file" or "want to create a file") in query:
                    app.Notepad.menu_select("File->New")
                    app.New.click(double=True)
                    speak("Please enter the name of file for saving it")
                    app.Notepad.menu_select("File->SaveAs")
                    file1=input("enter the name of file for saving it:")
                    app.SaveAs.edit.set_edit_text(f"{file1}")
                    app.SaveAs.Save.click(double=True)
                    speak("New file is created sucessfully with the desired name in the respected folder")
                    #waiting for the user to close the dialog box 
                    app.SaveAs.wait_out('enabled')
                #process for knowing about the version of notepad and its contents related information 
                
                elif ("want to know about the version of notepad" or "notepad version" or "version of notepad" or "please tell about the version of notepad" or "open notepad version" or "notepad details") in query:
                    app.Notepad.menu_select("Help->About Notepad")
                    #waiting for the dialog box to become disappear 
                    app.AboutNotepad.wait_out('enabled')
                    
                #process for getting some hep from the user in the web format by searching for desired help
                elif ("want to get some help from the notepad" or "need some help" or "help" or "help regarding some activity") in query:
                    speak(f"ok {user_name}")
                    app.Notepad.menu_select("Help->View Help")
                    app.ViewHelp.wait_out('enabled')
                #process for zooming in or out the contents of the file 
                
                elif ("want to perform the zoom in opeartion" or "zoom in" or "zoom out" or "want to perform the zoom out opeartion" or "zoom in the content" or "zoom out the content") in query:
                    if ("want to perform the zoom in opeartion" or "zoom in" or "zoom in the contents") in query:
                        app.Notepad.menu_select("View->Zoom->Zoom Out")
                        app.ZoomOut.wait_out('enabled')
                    else:
                        app.Notepad.menu_select("View->Zoom->Zoom In")
                        app.ZoomIn.wait_out('enabled')
                
                #process for opening the desired file as per the requirement 
                
                elif ("want to open a file" or "open a file" or "open file") in query:
                    speak("Ok please choose your desired file according to its location")
                    app.Notepad.menu_select("File->Open")
                    app.Open.wait_out('enabled')
                
                #process for doing the desired opeartions like operating the font style and font size 
                
                elif ("want to do some font adjustments" or "font adjustments" or "font changes" or "font settings" or "want to do some font setings") in query:
                    speak("Ok please do the required settings by giving the desired input")
                    ch=input("enter want do you to perform:")
                    #process for updating the font style
                    
                    if ("change font style" or "font style adjustment" or "font style settings" or "font style related settings") in ch:
                        speak("Ok please do the required font style settings")
                        app.Notepad.menu_select("Format->Font->Font style")
                        app.Fontstyle.wait_out('enabled')
                    
                    #process for changing the font size or levelling the font size 
                    
                    elif ("change the font size" or "adjust the font size" or "level the font size" or "want to some font ize settings" or "want to do some size settings") in ch:
                        speak("Ok please do the required settings")
                        app.Notepad.menu_select("Format->Font->Size")
                        app.Size.wait_out('enabled')
                    
                    #process for changing the font face or doing font face settings 
                    
                    elif ("change the font face" or "font face settings" or "font face type" or "want to some font face settings" or "want to do change the font face") in ch:
                        speak("please choose the desired font type or font face")
                        app.Notepad.menu_select("Format->Font->Font")
                        app.Font.wait_out('enabled')
                        
                #process for find and replacing the particular character or word 
                    
                elif ("perform the find and replace command" or "perform find and replace operation" or "find and replace operation" or "find and replace command" or "find and replace a particular word" or "find and replace particular character") in query:
                    app.Notepad.menu_select("Edit->Replace")
                    speak("Please enter the desired word or character in find and replace dialog box")
                    app.Replace.wait_out('enabled')
                    
                #process for terminating the system or application as per the requirment
                elif ("no thanks" or "no" or "terminate") in query:
                    speak(f"Ok {user_name} have a nice day please remember me for further uses")
                    sys.exit()
                    
        #Command for opening the Ms word  as per its version and doing the necessary advancements in it 
        elif ("open msword" or "open ms word" or "open word" or "open word document" or "please open a word document" or "please open msword" or"please open ms word") in query:
            speak(f"OK {user_name} opening ms word")
            apath="C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
            from pywinauto import application 
            app=application.Application()
            app.start(f"{apath}")
            #dlg=app.window(title="Untitled - Notepad")
            
            speak("Ms word is opened succesfully and untitled file is created succesfully")
            
        elif "no thanks" in query:
            speak(f"Ok thank you for using me {user_name}")
            sys.exit()
    
    
    