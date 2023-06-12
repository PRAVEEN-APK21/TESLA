import subprocess
import googletrans
import gtts
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import psutil
import pyautogui
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
import pywhatkit
import playsound
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
from email.message import EmailMessage
import win32com.client as wincl
from urllib.request import urlopen
# Python package supporting common text-to-speech engines
import pyttsx3

# For understanding speech
import speech_recognition as sr

# For fetching the answers
# to computational queries
import wolframalpha

# for fetching wikipedia articles
import wikipedia

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def cpu():
    usage = str(psutil.cpu_percent())
    speak('CPU usage is at ' + usage)
    print('CPU usage is at ' + usage)
    battery = psutil.sensors_battery()
    speak("Battery is at")
    speak(battery.percent)
    print("battery is at:" + str(battery.percent))


def quote():
    url = 'https://api.quotable.io/random'

    r = requests.get(url)
    quote = r.json()
    # print(quote)
    print(quote['content'])
    print('     -', quote['author'])
    speak(quote['content'])
    speak(quote['author'])


def translate():
    recognizer = sr.Recognizer()
    translator = googletrans.Translator()
    input_lang = 'en'
    output_lang = 'ta'
    try:
        with sr.Microphone() as source:
            print('Speak Now')
            voice = recognizer.listen(source)
            text = recognizer.recognize_google(voice, language=input_lang)
            print(text)
    except:
        pass

    translated = translator.translate(text, dest=output_lang)
    print(translated.text)
    converted_audio = gtts.gTTS(translated.text, lang=output_lang)
    converted_audio.save('apkvoice.mp3')
    playsound.playsound('apkvoice.mp3')
    os.remove('apkvoice.mp3')
    # print(googletrans.LANGUAGES)


def weather():
    api_key = "b35975e18dc93725acb092f7272cc6b8"  # generate your own api key from open weather
    base_url = "http://api.openweathermap.org//data//2.5//weather?"
    speak("tell me which city")
    city_name = takeCommand()
    complete_url = base_url + "appid=" + api_key + "&q=" + city_name
    response = requests.get(complete_url)
    x = response.json()
    if x["cod"] != "404":
        y = x["main"]
        current_temperature = y["temp"]
        current_pressure = y["pressure"]
        current_humidiy = y["humidity"]
        z = x["weather"]
        weather_description = z[0]["description"]
        r = ("in " + city_name + " Temperature is " +
             str(int(current_temperature - 273.15)) + " degree celsius " +
             ", atmospheric pressure " + str(current_pressure) + " hpa unit" +
             ", humidity is " + str(current_humidiy) + " percent"
                                                       " and " + str(weather_description))
        print(r)
        speak(r)
    else:
        speak(" City Not Found ")


def voice_change(v):
    x = int(v)
    engine.setProperty('voice', voices[x].id)
    speak("changed the voice sir")


def screenshot():
    img = pyautogui.screenshot()
    img.save("C:\\Users\\PRAVEEN KUMAR A\\PycharmProjects\\pythonProject4\\ss.png")


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning Sir !")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir !")

    else:
        speak("Good Evening Sir !")

    assname = ("tesla")
    speak("I am your Assistant")
    speak(assname)


def usrname():
    speak("What should i call you sir")
    uname = takeCommand()
    speak("Welcome Mister")
    speak(uname)
    columns = shutil.get_terminal_size().columns

    # print("#####################".center(columns))
    print("Welcome Mr.", uname)  # ter(columns))
    # print("#####################".center(columns))

    speak("How can i Help you, Sir")


def takeCommand():
    r = sr.Recognizer()

    with sr.Microphone() as source:

        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Unable to Recognize your voice.")
        return "None"

    return query


def send_email(receiver, subject, message):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    # Make sure to give app access in your Google account
    server.login('ucs19312@rmd.ac.in', 'apkrmd102')
    email = EmailMessage()
    email['From'] = 'ucs19312@rmd.ac.in'
    email['To'] = receiver
    email['Subject'] = subject
    email.set_content(message)
    server.send_message(email)


email_list = {
    'APK': 'praveensujith212001@gmail.com',
    'Raghav': 'ragavbharadwaj2002@gmail.com',
    'sabari': 'vsabari236@gmail.com',

}


def mail():
    speak('Whom you want to send email sir')
    name = takeCommand()
    receiver = email_list[name]
    print(receiver)
    speak('May i know the subject of your email sir?')
    subject = takeCommand()
    speak('Tell me the text in your email sir')
    message = takeCommand()
    send_email(receiver, subject, message)
    speak('Your email is sent apk')
    print('your email is sent apk')
    speak('Do you want to send more email?')
    send_more = takeCommand()
    if 'yes' in send_more:
        mail()

    else:
        pass


if __name__ == '__main__':
    clear = lambda: os.system('cls')

    # This Function will clean any
    # command before execution of this python file
    clear()
    wishMe()
    usrname()

    while True:

        query = takeCommand().lower()

        # All the commands said by user will be
        # stored here in 'query' and will be
        # converted to lower case for easily
        # recognition of command
        if 'wikipedia' in query:
            webbrowser.open("https://wikipedia.org/")


        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")
            time.sleep(9)

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")
            time.sleep(9)

        elif 'inspiration' in query:
            quote()
            time.sleep(5)

        elif 'translate' in query:
            translate()
            time.sleep(5)

        elif 'open gmail' in query:
            speak("opening gmail\n")
            webbrowser.open("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")
            time.sleep(5)

        elif "open word" in query:
            speak("Opening Microsoft Word")
            os.startfile('"C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"')
            time.sleep(5)

        elif "open linkedin" in query:
            speak("opening linkedin\n")
            webbrowser.open("https://www.linkedin.com/in/praveenkumar-a-8638531a0")
            time.sleep(5)

        elif "open skillrack" in query:
            speak("opening  your skillrack account\n")
            webbrowser.open(
                "https://www.skillrack.com/faces/ui/profile.xhtml;jsessionid=AA598F06F18FD25CD27EFA289A75D288")
            time.sleep(7)


        elif ("cpu and battery" in query or "battery" in query
              or "cpu" in query):
            cpu()

        elif ("screenshot" in query):
            screenshot()
            speak("screenshot taken sir!")


        elif 'music' in query:
            speak("Here you go with your favourite music")
            # music_dir = "G:\\Song"
            music_dir = "C:\\Users\\PRAVEEN KUMAR A\\Music"
            songs = os.listdir(music_dir)
            print(songs)
            random = os.startfile(os.path.join(music_dir, songs[30]))
            time.sleep(10)

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            speak(f"Sir, the time is {strTime}")


        elif 'mail' in query:
            mail()

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif 'i am fine' in query or "i am good" in query:
            speak("It's good to know that your fine")

        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query

        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = takeCommand()
            speak("Thanks for naming me")

        elif 'exit' in query or 'bye' in query:
            speak("Thanks for giving me your time")
            exit()

        elif "who made you" in query or "who created you" in query:
            speak("I have been created by PRAVEEN.")

        elif 'tell me a joke' in query:
            speak(pyjokes.get_joke())

        elif ("voice" in query):
            speak("for female say female and, for male say boy")
            b = takeCommand()
            if ("female" in b):
                voice_change(1)
            elif ("boy" in b):
                voice_change(0)

        elif "calculate" in query:
            speak("what?")
            g = takeCommand().lower().rstrip().strip()
            k = [i for i in g if i != " "]
            k = "".join(k)
            try:
                answer = eval(k)
                print("The answer is ", answer)
                speak("The answer is " + str(answer))
            except:
                speak("Please give the crt question..... Command invalid")


        elif 'search' in query or 'what is' in query or 'who is' in query:
            person = query.replace('what is', '')
            person = query.replace('who is', '')
            query = query.replace('search', '')
            info = wikipedia.summary(person, 2)
            pywhatkit.search(person)
            pywhatkit.search(query)
            print(info)
            # with open ()
            speak(info)

        elif "who am i" in query:
            speak("If you talk then definately your human.")

        elif "why you came to the world" in query:
            speak("Thanks to PRAVEEN. further It's a secret")

        elif 'ppt' in query:
            speak("opening Power Point presentation of project tesla sir")
            power = r"C:\\Users\\PRAVEEN KUMAR A\\Downloads\\PROJECT TESLA FINAL.pptx"
            os.startfile(power)

        elif "who are you" in query:
            speak("I am tesla a virtual assistant created by PRAVEEN")

        elif 'reason for you' in query:
            speak("I was created as a Mini project by Mister PRAVEENKUMAR")

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,
                                                       0,
                                                       "Location of wallpaper",
                                                       0)
            speak("Background changed succesfully")



        elif 'open news' in query:

            speak("Opening news for you")
            webbrowser.open("https://www.bbc.com/news/live/world-47639450")

        if 'play' in query:
            song = query.replace('play', '')
            speak('playing ' + song)
            pywhatkit.playonyt(song)


        elif 'lock window' in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
            speak("Hold On a Sec ! Your system is on its way to shut down")
            os.system("shutdown /p")

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")

        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop tesla from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.com/maps/place/" + location.strip().rstrip())

        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "tesla Camera", "img.jpg")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            os.system("shutdown / h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "write a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('tesla.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)

        elif "show note" in query:
            speak("Showing Notes")
            file = open("tesla.txt", "r")
            print(file.read())
            speak(file.read(6))

        # NPPR9-FWDCX-D2C8J-H872K-2YT43
        elif "tesla" in query:

            wishMe()
            speak("tesla in your service Mister")
            speak(assname)

        elif ("weather" in query or "temperature" in query):
            weather()

        elif "message " in query:
            # You need to create an account on Twilio to use this service
            account_sid = 'Account Sid key'
            auth_token = 'Auth token'
            client = Client(account_sid, auth_token)

            message = client.messages \
                .create(
                body=takeCommand(),
                from_='',
                to=''
            )

            print(message.sid)


        elif "Good Morning" in query:
            speak("A warm" + query)
            speak("How are you Mister")
            speak(assname)

        elif "will you be my friend" in query:
            speak("I'm not sure about, may be you should give me some time")


