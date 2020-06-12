from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from bs4 import BeautifulSoup as soup # for pulling data out of HTML and XML files 
from pyowm.owm import OWM #Open Weather Maps
import urllib.request as urllib2 #for fetching URL
import pafy # To Retrieve YouTube content and metadata
from distutils import spawn #to find .exe files
import os
import re
import pyttsx3 
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import random
import smtplib
import pytz #for accurate and cross platform timezone calculations
import subprocess
import shutil
import youtube_dl
import json
import ctypes
import wolframalpha
from docx import Document # for opening docx



# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"]
DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
DAY_EXTENSIONS = ["rd", "th", "st", "nd"]


def speak(text):
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    engine.setProperty('voice',voices[1].id)
    engine.say(text)
    engine.runAndWait()


def authenticate_google():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    return service

def get_events(day, service):

    if day == None:
        return
    date = datetime.datetime.combine(day, datetime.datetime.min.time())
    end_date = datetime.datetime.combine(day, datetime.datetime.max.time())

    utc = pytz.UTC
    date = date.astimezone(utc)
    end_date = end_date.astimezone(utc)


    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(), timeMax = end_date.isoformat(),
                                        singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        speak('No upcoming events found.')
    else:
        speak(f"You have {len(events)} events on this day.")

        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start, event['summary'])
            start_time = str(start.split("T")[1].split("+")[0])
            if int(start_time.split(":")[0]) < 12:
                start_time = start_time + "am"
            else:
                start_time = str(int(start_time.split(":")[0]) - 12) + start_time.split(":")[1]
                start_time = start_time + "pm"
            
            speak(event["summary"] + " at " + start_time)


def get_date(text):
    # Deriving date from text representing speech
    today = datetime.date.today()

    if text.count("today") > 0:
        return today

    day = -1
    day_of_week = -1
    month = -1
    year = today.year

    for word in text.split():
        if word in MONTHS:
            month = MONTHS.index(word) + 1
        elif word in DAYS:
            day_of_week = DAYS.index(word)
        elif word.isdigit():
            day = int(word)
        else:
            for ext in DAY_EXTENSIONS:
                found = word.find(ext)
                if found > 0:
                    try:
                        day = int(word[:found])
                    except:
                        pass
    if month < today.month and month != -1:
        year += 1
    
    if day < today.day and month == -1 and day != -1:
        month += 1
    
    if month == -1 and day == -1 and day_of_week != -1:
        current_day_of_week = today.weekday()
        dif = day_of_week - current_day_of_week

        if dif < 0:
            dif += 7
            if text.count("next") >= 1:
                dif += 7
        return today + datetime.timedelta(dif)
    
    if month == -1 or day == -1:
        return None
    return datetime.date(month=month, day=day, year=year)


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")
    
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")
    
    speak("I am Jarvis. Please tell me how can I help you")

def goodBye():
    hour = int(datetime.datetime.now().hour)
    if hour > 21:
        speak('Have a good night sir! See you tomorrow')
        exit()
    else:
        speak("Bye, See you later!")
        exit()     

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
        # print("Listening...")
        r.pause_threshold = 1
        text = r.listen(source)

    try:
        # print("Recognizing...")
        text =  r.recognize_google(text, language = 'en-in')
        print(f"User said: {text}\n")

    except Exception as e:
        return "None"
    return text.lower()

def note(text):

    file_name = "voicenote.txt"
    with open(file_name, "w") as f:
        f.write(text)

    subprocess.Popen(["notepad.exe", file_name])
    # os.chdir(current_dir)
    

wake = "jarvis"
service = authenticate_google()
print("Start by saying 'Hey Jarvis'")


while True:
    print("Listening...")
    text = takeCommand()

    if text.count(wake) > 0:
        speak('I am Ready.')
        text = takeCommand()
        
    # Calender
    CALENDAR_STRS = ["what do i have", "do i have plans","am i busy"]
    for phrase in CALENDAR_STRS:
        if phrase in text:
            date = get_date(text)
            if date:
                get_events(date, service)
            else:
                speak("I don't understand.")
                    
    
    #Make a Note
    NOTE_STRS = ["make a note", "write this down", "remember this"]
    # elif 'note' in text:
    for phrase in NOTE_STRS:
        if phrase in text:
            speak('what would you like me to write down?')
            note_text = takeCommand()
            note(note_text)
            speak("I've made a note of that.")

    # #ask me anything
    # if 'about' in text:
    #     reg_ex = re.search('about (.+)', text)
    #     text = text.replace("wikipedia","")
    #     results = wikipedia.summary(text, sentences = 5)
    #     print(results)
    #     speak(results)

    # Greetings
    greetings = ['hello','hi']
    for phrase in greetings:
        if phrase in text:
            wishMe()

    #open website
    if 'open' in text:
        reg_ex = re.search('open (.+)', text)
        if reg_ex:
            domain = reg_ex.group(1)
            print(domain)
            url = 'https://www.' + domain +'.com'
            webbrowser.open(url)
            speak(domain + 'has been opened for you Sir.')
        else:
            pass

    #read today's news
    elif 'news' in text:
        try:
            news_url="https://news.google.com/news/rss"
            Client=urllib2.urlopen(news_url)
            xml_page=Client.read()
            Client.close()
            soup_page=soup(xml_page, "xml")
            news_list=soup_page.findAll("item")
            for news in news_list[:5]:
                print(news.title.text.encode('utf-8'))
                speak(news.title.text.encode('utf-8'))
                
        except Exception as e:
            print(e)

    #playing music on local system
    # elif'' in text:
    #     music_dir = 'C:\\Users\\91741\\Documents\\Friends'
    #     songs = os.listdir(music_dir)
    #     n = random.randint(0,len(songs))
    #     os.startfile(os.path.join(music_dir, songs[n]))

    #play youtube song
    elif 'play' in text:
        speak('What song shall I play Sir?')
        mysong = takeCommand()
        if mysong:
            flag = 0
            url = "https://www.youtube.com/results?search_query=" + mysong.replace(' ', '+')
            response = urllib2.urlopen(url)
            html = response.read()
            soup1 = soup(html,"lxml")
            url_list = []
            for vid in soup1.findAll(attrs={'class':'yt-uix-tile-link'}):
                if ('https://www.youtube.com' + vid['href']).startswith("https://www.youtube.com/watch?v="):
                    flag = 1
                    final_url = 'https://www.youtube.com' + vid['href']
                    url_list.append(final_url)

            video = pafy.new(url_list[0])
            best = video.getbest()
            playurl = best.url
            webbrowser.open(playurl)
            speak('Playing '+ video.title)

            if flag == 0:
                speak('I have not found anything in Youtube ')

    #current time
    # elif 'time' in text:
    #     strTime = datetime.datetime.now().strftime("%H:%M:%S")
    #     speak(f"Sir, The time is {strTime}")

    #current weather
    # elif 'weather' in text:
    #     reg_ex = re.search('weather in (.*)', text)
    #     if reg_ex:
    #         city = reg_ex.group(1)
    #         # api_key = security.encrypt_password()
    #         owm = OWM('df95dca27f55d7d2daa552256306ef52')
    #         mgr = owm.weather_manager()
    #         obs = mgr.weather_at_place(city)
    #         w = obs.weather
    #         k = w.detailed_status
            
    #         sunrise = w.sunrise_time(timeformat='date')
    #         sunrise_date = w.sunrise_time(timeformat='date').strftime("%H:%M:%S")

    #         sunset = w.sunset_time(timeformat='date')
    #         sunset_date = w.sunset_time(timeformat='date').strftime("%H:%M:%S")
    #         x = w.temperature('celsius')

    #         print(f'Current weather in {city} is {k}. Sunrise is at {sunrise_date} UTC  and Sunset will be at {sunset_date} UTC.')
    #         print('The maximum temperature is %0.2f°C and the minimum temperature is %0.2f°C' % (x['temp_max'], x['temp_min']))
    #         speak(f'Current weather in {city} is {k}. Sunrise is at {sunrise.hour} hours {sunrise.minute} minutes and {sunrise.second} seconds  and Sunset will be at {sunset.hour} hours {sunset.minute} minutes and {sunset.second} seconds.')
    #         speak('The maximum temperature is %0.2f degree celsius and the minimum temperature is %0.2f degree celsius.' % (x['temp_max'], x['temp_min']))

    #change wallpaper
    elif 'wallpaper' in text:
        api_key = '1Kr_PV0Bm2uujrFfdWjcFK11ANPKCZpJxwQoTUXfbc0'
        path = "C:/Users/91741/Downloads/Wallpapers/1"
        url = 'https://api.unsplash.com/photos/random?client_id=' + api_key
        f = urllib2.urlopen(url)
        json_string = f.read()
        f.close()
        parsed_json = json.loads(json_string)
        photo = parsed_json['urls']['raw'] + "&w=1500&dpr=2"
        urllib2.urlretrieve(photo, path) # Location where the image is downloaded
        ctypes.windll.user32.SystemParametersInfoW(20, 0, path , 0)
        speak('wallpaper changed successfully')

    #sending an email
    elif 'mail' in text:
        try:
            speak("What should I say?")
            content = takeCommand()
            to = "receiveremail@email.com"
            sendEmail(to, content)
            speak("Email has been sent!")
        except Exception as e:
            print(e)
            speak("Sorry! I cant send it now. Please try again later.")

    #launch any application
    elif 'launch' in text:
        reg_ex = re.search('launch (.+)', text)
        try:
            appname = reg_ex.group(1)
            if appname == "word":
                
                speak('Launching' + appname)
                os.startfile('test.docx')
                speak("what do you want me to write down")

                document = Document('test.docx')
                document.add_paragraph(takeCommand())
                document.save('test.docx')
            else:
                speak('Launching' + appname)
                os.startfile(spawn.find_executable(appname))
        except:
            # speak('I cant find the application')
            pass
    
    elif 'thank you' in text:
        speak('Pleasure is all mine!')
    
    elif 'sorry' in text:
        speak('No Problem sir')
    
    elif 'good job'in text:
        speak('thank you')
    
    elif 'bye' in text:
        goodBye()   

    #search engine
    elif(True):
        try:
            app_id = 'LVGG7X-TXLHLAR25T'
            client = wolframalpha.Client(app_id)    # Instance of wolfram alpha client class

            res = client.query(text)                # Stores the response from wolf ram alpha
                                                
            answer = next(res.results).text         #Include only the text from the response
            print(answer)
            speak(answer)
        except:
            speak('Sorry I couldnt find that one.')
    
    



