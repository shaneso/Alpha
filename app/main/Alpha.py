from numpy import source
from youtubesearchpython import *
from yahoo_fin import stock_info
from playsound import playsound
import speech_recognition as sr
from word2number import w2n
from random import randint
from pathlib import Path
import num2words as n2w
import tkinter as tk
import webbrowser
import datetime
import win32api
import tkinter
import pyttsx3
import random
import urllib
import time
import pafy
import cv2
import vlc
import os

WAKE = 'alpha'

username = os.getlogin()

app = tk.Tk()
width_px = app.winfo_screenwidth()
height_px = app.winfo_screenheight()

active_status = False

def speak(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

def get_audio():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
        input = ''
        try:
            input = recognizer.recognize_google(audio)
            print(input)
        except Exception:
            pass
        except sr.UnknownValueError:
            pass
        except sr.RequestError:
            pass
    return input.lower()

def greeting():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak('Hello, good morning')
    elif hour >= 12 and hour < 18:
        speak('Hello, good afternoon')
    else:
        speak('Hello, good evening')

greeting()

def open_program(path):
    try:
        open_program(path)
    except Exception:
        pass

while True:
    text = get_audio().lower()

    if text.count(WAKE) > 0:
        speak('Listening')
        text = get_audio().lower()

        if 'time' in text:
            time = datetime.datetime.now()
            time = time.strftime('%I:%M %p')
            speak(f"It's {time}")

        if 'date' in text:
            date = datetime.datetime.now().strftime('%A %d %B %Y')
            speak(f"It's {date}")

        if 'play music' in text:
            speak('What music would you like to stream?')
            text = get_audio().lower()
            if text != '':
                media_value = 0
                video_search = VideosSearch(text)
                while True:
                    value = media_value
                    media_title = video_search.result()['result'][value]['title']
                    speak(f'I found this. {media_title}')
                    speak('Would you like to stream this?')
                    media_text = get_audio().lower()
                    if 'yes' in media_text:
                        title = video_search.result()['result'][value]['title']
                        video = video_search.result()['result'][value]['link']
                        print(title)
                        print(video)
                        url = video
                        video = pafy.new(url)
                        best = video.getbest()
                        media = vlc.MediaPlayer(best.url)
                        media.play()
                        break
                    if 'no' in media_text:
                        media_value += 1
                    if 'stop' in media_text:
                        break

        if 'play video' in text:
            speak('What video would you like to stream?')
            text = get_audio().lower()
            if text != '':
                media_value = 0
                video_search = VideosSearch(text)
                while True:
                    value = media_value
                    media_title = video_search.result()['result'][value]['title']
                    speak(f'I found this. {media_title}')
                    speak('Would you like to stream this?')
                    media_text = get_audio().lower()
                    if 'yes' in media_text:
                        title = video_search.result()['result'][value]['title']
                        video = video_search.result()['result'][value]['link']
                        print(title)
                        print(video)
                        url = video
                        video = pafy.new(url)
                        best = video.getbest()
                        media = vlc.MediaPlayer(best.url)
                        media.play()
                        break
                    if 'no' in media_text:
                        media_value += 1
                    if 'stop' in media_text:
                        break

        if 'play' in text:
            media.play()

        if 'pause' in text:
            media.pause()

        if 'stop' in text:
            media.stop()

        if 'active mode on' in text:
            speak('Active mode on')
            active_status = True
            while active_status is True:
                x = randint(1, width_px)
                y = randint(1, height_px)

                win32api.SetCursorPos((x, y))
                text = get_audio().lower()
                if 'active mode off' in text:
                    speak('Active mode off')
                    active_status = False

        if 'research stock' in text:
            speak('Which stock would you like to research?')
            text = get_audio().lower()
            if len(text) <= 5:
                try:
                    webbrowser.open(f'https://finviz.com/quote.ashx?t={text}&ty=c&ta=1&p=d')
                except Exception:
                    pass
            else:
                webbrowser.open(f'https://www.google.com/search?q={text}+stock')

        if 'research otc stock' in text:
            speak('Which OTC stock would you like to research?')
            text = get_audio().lower()
            if len(text) <= 5:
                try:
                    webbrowser.open(f'https://www.otcmarkets.com/stock/{text}/overview')
                except Exception:
                    pass
            else:
                webbrowser.open(f'https://www.google.com/search?q={text}+stock')

        if 'stock data' in text:
            speak('What is the stock ticker?')
            text = get_audio().lower()
            input_ticker = True
            while input_ticker is True:
                try:
                    if len(text) <= 5:
                        try:
                            stock_data = stock_info.get_quote_table(text)
                            print(stock_data)
                            input_ticker = False
                        except Exception:
                            pass
                    else:
                        speak('Please provide a stock ticker')
                        text = get_audio().lower()
                except Exception:
                    pass

        if 'stock price' in text:
            speak('What is the stock ticker?')
            text = get_audio().lower()
            input_ticker = True
            while input_ticker is True:
                try:
                    if len(text) <= 5:
                        try:
                            stock_price = stock_info.get_live_price(text)
                            speak(f'The price of {text} is {n2w.num2words(round(stock_price, 2))}')
                            input_ticker = False
                        except Exception:
                            pass
                    else:
                        speak('Please provide a stock ticker')
                        text = get_audio().lower()
                except Exception:
                    pass

        if 'research crypto' in text:
            speak('Which cryptocurrency would you like to research?')
            text = get_audio().lower()
            new_text = text.replace(' ', '-')
            if len(text) <= 5:
                webbrowser.open(f'https://www.binance.com/en/trade/{text}_USDT?layout=pro')
            else:
                webbrowser.open(f'https://coinmarketcap.com/currencies/{new_text}/')

        if 'open gmail' in text:
            speak('Opening Gmail')
            webbrowser.open("https://mail.google.com/mail/u/0/#inbox")

        if 'open finviz' in text:
            speak('Opening Finviz')
            webbrowser.open("https://finviz.com/")

        if 'open youtube' in text:
            speak('Opening YouTube')
            webbrowser.open("https://www.youtube.com/")

        if 'open facebook' in text:
            speak('Opening Facebook')
            webbrowser.open('https://www.facebook.com/')

        if 'open netflix' in text:
            speak('Opening Netflix')
            webbrowser.open('https://www.netflix.com/browse')

        if 'search google' in text:
            speak('What would you like to search?')
            text = get_audio().lower()
            webbrowser.open(f'https://www.google.com/search?q={text}')

        if 'open word' in text:
            path = Path("C:/Program Files/Microsoft Office/root/Office16/WINWORD.exe")
            if path.exists():
                speak('Opening Microsoft Word')
                open_program(path)
            
        if 'open powerpoint' in text:
            path = Path("C:/Program Files/Microsoft Office/root/Office16/POWERPNT.exe")
            if path.exists():
                speak('Opening Microsoft PowerPoint')
                open_program(path)
            
        if 'open excel' in text:
            path = Path("C:/Program Files/Microsoft Office/root/Office16/EXCEL.exe")
            if path.exists():
                speak('Opening Microsoft Excel')
                open_program(path)
            
        if 'open outlook' in text:
            path = Path("C:/Program Files/Microsoft Office/root/Office16/OUTLOOK.EXE")
            if path.exists():
                speak('Opening Microsoft Outlook')
                open_program(path)

        if 'open command prompt' in text:
            path = Path("C:\Windows\System32\cmd.exe")
            if path.exists():
                speak('Opening Command Prompt')
                open_program(path)

        if 'open computer management' in text:
            path = Path("C:\Windows\System32\CompMgmtLauncher.exe")
            if path.exists():
                speak('Opening Computer Management')
                open_program(path)

        if 'open system configuration' in text:
            path = Path("C:\Windows\System32\msconfig.exe")
            if path.exists():
                speak('Opening System Configuration')
                open_program(path)

        if 'open system information' in text:
            path = Path("C:\Windows\System32\msinfo32.exe")
            if path.exists():
                speak('Opening System Information')
                open_program(path)

        if 'open control panel' in text:
            path = Path("C:\Windows\System32\control.exe")
            if path.exists():
                speak('Opening Control Panel')
                open_program(path)

        if 'open bluetooth file transfer' in text:
            path = Path("C:/Windows/System32/fsquirt.exe")
            if path.exists():
                speak('Opening Bluetooth File Transfer')
                open_program(path)

        if 'open task manager' in text:
            path = Path("C:\Windows\System32\Taskmgr.exe")
            if path.exists():
                speak('Opening Task Manager')
                open_program(path)
            
        if 'open photos' in text:
            path = Path("C:\Program Files\WindowsApps\Microsoft.Windows.Photos_2021.21120.8011.0_x64__8wekyb3d8bbwe\Microsoft.Photos.exe")
            if path.exists():
                speak('Opening Photos')
                open_program(path)
            
        if 'open visual studio code' in text:
            path = Path(f"C:/Users/{username}/AppData/Local/Programs/Microsoft VS Code/Code.exe")
            if path.exists():
                speak('Opening Visual Studio Code')
                open_program(path)
            
        if 'open android studio' in text:
            path = Path("C:/Program Files/Android/Android Studio/bin/studio64.exe")
            if path.exists():
                speak('Opening Android Studio')
                open_program(path)
            
        if 'open zoom' in text:
            path = Path(f"C:/Users/{username}/AppData/Roaming/Zoom/bin/Zoom.exe")
            if path.exists():
                speak('Opening Zoom')
                open_program(path)

        if 'open file explorer' in text:
            path = Path("C:\Windows\explorer.exe")
            if path.exists():
                speak('Opening File Explorer')
                open_program(path)

        if 'open microsoft teams' in text:
            path = Path(f"C:/Users/{username}/AppData/Local/Microsoft/Teams/current/Teams.exe")
            if path.exists():
                speak('Opening Microsoft Teams')
                open_program(path)
            
        if 'open microsoft store' in text:
            path = Path("C:\Program Files\WindowsApps\Microsoft.WindowsStore_22112.1401.2.0_x64__8wekyb3d8bbwe\WinStore.App.exe")
            if path.exists():
                speak('Opening Microsoft Store')
                open_program(path)
            
        if 'open chrome' in text:
            path = Path("C:\Program Files\Google\Chrome\Application\chrome.exe")
            if path.exists():
                speak('Opening Google Chrome')
                open_program(path)
            
        if 'open windows security' in text:
            path = Path("C:\Program Files\WindowsApps\microsoft.sechealthui_1000.22000.1.0_neutral__8wekyb3d8bbwe\SecHealthUI.exe")
            if path.exists():
                speak('Opening Windows Security')
                open_program(path)
            
        if 'open binance' in text:
            path = Path("C:\Program Files\Binance\Binance.exe")
            if path.exists():
                speak('Opening Binance')
                open_program(path)
            
        if 'open discord' in text:
            path = Path(f"C:/Users/{username}/AppData/Local/Discord/app-1.0.9003/Discord.exe")
            if path.exists():
                speak('Opening Discord')
                open_program(path)

        if 'open resolve' in text:
            path = Path("C:\Program Files\Blackmagic Design\DaVinci Resolve\Resolve.exe")
            if path.exists():
                speak('Opening Davinci Resolve')
                open_program(path)

        if 'power off computer' in text:
            speak('Powering off computer')
            os.system("shutdown /s /t 1")

        if 'restart computer' in text:
            speak('Restarting computer')
            os.system("shutdown /r /t 1")

        if 'sign out' in text:
            speak('Signing out')
            os.system("shutdown -l")

        if 'close program' in text:
            speak('Closing program')
            quit()