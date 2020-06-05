import os
import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import random
import smtplib


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice',voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")
    
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")
    
    speak("I am Jarvis. Please tell me how can I help you")

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login('youremail@email.com','pass')
    server.sendmail('youremail@email.com', to, content)
    server.close()

def takeCommand():
    '''it takes microphone input from the user and returns string output
    '''
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query =  r.recognize_google(audio, language = 'en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)
        print("Say that again please.")
        return "None"
    return query

if __name__ == "__main__":
    wishMe()
    while True:
        query = takeCommand().lower()

        #logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia","")
            results = wikipedia.summary(query, sentences = 2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")
        
        elif'play something' in query:
            music_dir = 'C:\\Users\\91741\\Documents\\Friends'
            songs = os.listdir(music_dir)
            n = random.randint(0,len(songs))
            # print(songs)
            os.startfile(os.path.join(music_dir, songs[n]))

        elif 'time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, The time is {strTime}")

        elif 'code' in query:
            codePath = "C:\\Users\\91741\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)

        elif 'mail' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "receiveremail@email.com"
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry! I cant send it now. Please try again later.")
        elif 'quit' or 'bye' in query:
            speak("Bye, See you later!")
            exit()                
        