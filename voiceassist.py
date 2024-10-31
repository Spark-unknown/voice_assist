# Importing Libraries
import os
import subprocess
import requests
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import shutil
from twilio.rest import Client
from tqdm import tqdm
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import wolframalpha
import pyttsx3
import tkinter
import os
import sys
import keyboard
import threading
import speech_recognition as sr
import pyttsx3
from voice_assistant import VoiceAssistant


# Initializing Text-to-Speech Engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


# Function to convert text to speech
def speak(audio):
    """Converts text to speech"""
    engine.say(audio)
    engine.runAndWait()


# Function to wish the user
def wish_me():
    """Wishes the user based on the current time"""
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning Sir!")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir!")
    else:
        speak("Good Evening Sir!")


# Function to get the username
def username():
    """Asks for and greets the user"""
    speak("What should I call you, sir?")
    uname = take_command()
    speak("Welcome Mister")
    speak(uname)
    columns = shutil.get_terminal_size().columns
    print("#####################".center(columns))
    print("Welcome Mr.", uname.center(columns))
    print("#####################".center(columns))
    speak("How can I help you, Sir?")


def listen_for_cancel(assistant):
    keyboard.wait('esc')  # Press ESC to cancel
    assistant.cancel_task()


def main():
    assistant = VoiceAssistant()
    t = threading.Thread(target=listen_for_cancel, args=(assistant,))
    t.start()
    assistant.start()
# Function to take voice commands
def take_command():
    """Takes voice commands from the user"""
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
        print("Unable to recognize your voice.")
        return "None"
    return query


# Function to send emails
def send_email(to, content):
    """Sends emails using SMTP"""
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('your email id', 'your email password')
    server.sendmail('your email id', to, content)
    server.close()


# Main function
if __name__ == '__main__':
    clear = lambda: os.system('cls')
    clear()
    wish_me()
    username()

    while True:
        query = take_command().lower()

        # Wikipedia search
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        # Open YouTube
        elif 'open youtube' in query:
            speak("Here you go to Youtube")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com")

        # Play Music
        elif 'play music' in query or "play song" in query:
            speak("Here you go with music")
            music_dir = "C:\\Users\\GAURAV\\Music"
            songs = os.listdir(music_dir)
            print(songs) 
            random = os.startfile(os.path.join(music_dir, songs[1]))

        # Time
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S") 
            speak(f"Sir, the time is {strTime}")

        # Open Opera
        elif 'open opera' in query:
            codePath = r"C:\\Users\\GAURAV\\AppData\\Local\\Programs\\Opera\\launcher.exe"
            os.startfile(codePath)

        # Email
                # Email
        elif 'send a mail' in query:
            try:
                speak("What should I say?")
                content = take_command()
                speak("Whome should I send?")
                to = input() 
                send_email(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        # Greetings
        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that you're fine")

        # Change Name
        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query

        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = take_command()
            speak("Thanks for naming me")

        # Introduce
        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me")
            speak(assname)
            print("My friends call me", assname)

        # Exit
        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        # Creator
        elif "who made you" in query or "who created you" in query: 
            speak("I have been created by Gaurav.")

        # Joke
        elif 'joke' in query:
            speak(pyjokes.get_joke())

        # Calculate
        elif "calculate" in query: 
            app_id = "Wolframalpha api id"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate') 
            query = query.split()[indx + 1:] 
            res = client.query(' '.join(query)) 
            answer = next(res.results).text
            print("The answer is " + answer) 
            speak("The answer is " + answer) 

        # Search
        elif 'search' in query or 'play' in query:
            query = query.replace("search", "") 
            query = query.replace("play", "")		 
            webbrowser.open(query) 

        # Human Check
        elif "who i am" in query:
            speak("If you talk then definitely you're human.")

        # Purpose
        elif "why you came to world" in query:
            speak("Thanks to Gaurav. Further, it's a secret")

        # PowerPoint
        elif 'power point presentation' in query:
            speak("Opening Power Point presentation")
            power = r"C:\\Users\\GAURAV\\Desktop\\Minor Project\\Presentation\\Voice Assistant.pptx"
            os.startfile(power)

        # Love
        elif 'is love' in query:
            speak("It is the 7th sense that destroys all other senses")

        # Introduction
        elif "who are you" in query:
            speak("I am your virtual assistant created by Gaurav")

        # Reason
        elif 'reason for you' in query:
            speak("I was created as a Minor project by Mister Gaurav ")

        # Change Background
        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20, 
                                                    0, 
                                                    "Location of wallpaper",
                                                    0)
            speak("Background changed successfully")

        # Open Bluestack
        elif 'open bluestack' in query:
            appli = r"C:\\ProgramData\\BlueStacks\\Client\\Bluestacks.exe"
            os.startfile(appli)

        # News
        elif 'news' in query:
            try: 
                jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
                data = json.load(jsonObj)
                i = 1
                
                speak('Here are some top news from The Times of India')
                print('''=============== TIMES OF INDIA ============'''+ '\n')
                
                for item in data['articles']:
                    
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                
                print(str(e))

        # Lock Window
        elif 'lock window' in query:
            speak("Locking the device")
            ctypes.windll.user32.LockWorkStation()

        # Shutdown
                # Shutdown
        elif 'shutdown system' in query:
            speak("Hold On a Sec! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')
            
        # Empty Recycle Bin
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")

        # Don't Listen
        elif "don't listen" in query or "stop listening" in query:
            speak("For how much time you want to stop jarvis from listening commands")
            a = int(take_command())
            time.sleep(a)
            print(a)

        # Location
        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl / maps / place/" + location + "")

        # Camera
        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Jarvis Camera ", "img.jpg")

        # Restart
        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])
            
        # Hibernate
        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        # Log Off
        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        # Note
        elif "write a note" in query:
            speak("What should I write, sir")
            note = take_command()
            file = open('jarvis.txt', 'w')
            speak("Sir, Should I include date and time")
            snfm = take_command()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
        
        # Show Note
        elif "show note" in query:
            speak("Showing Notes")
            file = open("jarvis.txt", "r") 
            print(file.read())
            speak(file.read(6))

        # Update Assistant
        elif "update assistant" in query:
            speak("After downloading file please replace this file with the downloaded one")
            url = '# url after uploading file'
            r = requests.get(url, stream=True)
            
            with open("Voice.py", "wb") as Pypdf:
                
                total_length = int(r.headers.get('content-length'))
                
                for ch in tqdm(r.iter_content(chunk_size=2391975),
                                expected_size=(total_length / 1024) + 1):
                    if ch:
                        Pypdf.write(ch)
                    
        # Greeting
        elif "jarvis" in query:
            wish_me()
            speak("Jarvis 1 point o in your service Mister")
            speak(assname)

        # Weather
        elif "weather" in query:
            
            api_key = "Api key"
            base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
            speak("City name ")
            print("City name : ")
            city_name = take_command()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url) 
            x = response.json() 
            
            if x["code"] != "404": 
                y = x["main"] 
                current_temperature = y["temp"] 
                current_pressure = y["pressure"] 
                current_humidity = y["humidity"] 
                z = x["weather"] 
                weather_description = z[0]["description"] 
                print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidity) +"\n description = " +str(weather_description)) 
            
            else: 
                speak(" City Not Found ")
            
        # Send Message
        elif "send message " in query:
            account_sid = 'Account Sid key'
            auth_token = 'Auth token'
            client = Client(account_sid, auth_token)

            message = client.messages \
                            .create(
                                body = take_command(),
                                from_='Sender No',
                                to ='Receiver No'
                            )

            print(message.sid)

        # Wikipedia
        elif "wikipedia" in query:
            webbrowser.open("wikipedia.com")

                # Good Morning
        elif "good morning" in query:
            speak("Good morning, Sir!")
            speak("Today's date is " + datetime.datetime.now().strftime("%d-%m-%Y"))
            speak("Today's time is " + datetime.datetime.now().strftime("%H:%M:%S"))
            speak("Have a great day ahead!")

        # Good Afternoon
        elif "good afternoon" in query:
            speak("Good afternoon, Sir!")
            speak("Today's date is " + datetime.datetime.now().strftime("%d-%m-%Y"))
            speak("Today's time is " + datetime.datetime.now().strftime("%H:%M:%S"))
            speak("Have a great day ahead!")

        # Good Evening
        elif "good evening" in query:
            speak("Good evening, Sir!")
            speak("Today's date is " + datetime.datetime.now().strftime("%d-%m-%Y"))
            speak("Today's time is " + datetime.datetime.now().strftime("%H:%M:%S"))
            speak("Have a great evening ahead!")

        # Thanks
        elif "thanks" in query or "thank you" in query:
            speak("You're welcome, Sir!")

        # Exit
        elif "exit" in query or "quit" in query:
            speak("Thank you for using me, Sir!")
            exit()

    if __name__ == "__main__":
        main()