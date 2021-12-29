import pyttsx3,datetime,speech_recognition as sr,pyaudio,wikipedia,os,webbrowser
from win32com.client import Dispatch
engine=pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
engine.setProperty('voice',voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishme():
    hour=(datetime.datetime.now().hour)
    if hour>=0 and hour<=12:
        speak("Good Morning!")
    elif hour>12 and hour<=18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening")    
    speak("How Are You")

def takeCommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold=1
        audio=r.listen(source)
    try:
        print("Recognizing...")
        query= r.recognize_google(audio,language="en-in")
        print(f"User Said: {query}")
    except Exception as e:
        print("Please Say it Again")  
        return ""
    return query
if __name__=="__main__":
    wishme()
    while(True):
      query=takeCommand().lower()  
      if('wikipedia' in query):
          speak('searching wikipedia...')
          query.replace("wikipedia","")
          webResult=wikipedia.summary(query,sentences=2)
          speak('According to wikipedia')
          print(webResult)
          speak(webResult)
      elif("i am fine" in query):
          speak("nice to hear it")
      elif("open youtube" in query):
          webbrowser.open("youtube.com")
      elif("open stack overflow" in query):
          webbrowser.open("stackoverflow.com")    
      elif("open google" in query):
          webbrowser.open("google.com")
      elif("open github" in query):
          webbrowser.open('github.com')
      elif("the time" in query):
          currtime=datetime.datetime.now().strftime("%H:%M:%S")
          print(currtime)
          speak(f"The time is {currtime}") 
      elif("open code" in query) :
          codePath="C:\\Users\\faiza\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
          os.startfile(codePath)        
      elif("quit" in query or "kuwait" in query):
          exit()    
