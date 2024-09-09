from re import search
import pyautogui
import pyttsx3
import subprocess
import psutil
import pywhatkit
import speech_recognition as sr
import win32com.client
import webbrowser
import datetime
import wikipedia
import os
import urllib.parse
from mtranslate import translate  # Import translation library
from colorama import Fore, Style, init
import requests
import time

init(autoreset=True)


def say(text, voice="Microsoft Zira Desktop - English (United States)"):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    try:
        voices = speaker.GetVoices()
        for v in voices:
            if v.GetDescription().strip() == voice.strip():
                speaker.Voice = v
                break
    except:
        print(f"Voice '{voice}' not found. Using the default voice.")
    speaker.Speak(text)

say("Sara HERE !!")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language="en-in")
        print(f"User said: {query}\n")
    except Exception as e:
        print("Could not understand audio, please try again...")
        return None
    return query


def Translate_hindi_to_english(text):
    english_text = translate(text, "en-us")
    return english_text


def search_on_google(query):
    try:
        encoded_query = urllib.parse.quote(query)  # Encode query for URL
        url = f"https://www.google.com/search?q={encoded_query}"
        webbrowser.open(url)
        say(f"Here are the results for {query}", voice="Microsoft Zira Desktop - English (United States)")
    except Exception as e:
        print(f"Error opening Google search: {e}")
        say("Sorry, I couldn't open the Google search.", voice="Microsoft Zira Desktop - English (United States)")


def list_running_processes():
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            process_name = proc.info['name']
            if 'code' in process_name.lower():
                print(f"Process Name: {process_name}, Process ID: {proc.info['pid']}")
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass


list_running_processes()


def Speech_To_Text_Python():
    recognizer = sr.Recognizer()
    recognizer.dynamic_energy_threshold = False
    recognizer.energy_threshold = 34000
    recognizer.dynamic_energy_adjustment_damping = 0.010
    recognizer.dynamic_energy_ratio = 1.000
    recognizer.pause_threshold = 0.3
    recognizer.operation_timeout = None
    recognizer.pause_threshold = 0.2
    recognizer.non_speaking_duration = 0.2

    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)
        print("Listening....")
        try:
            audio = recognizer.listen(source, timeout=None)
            print("\rRecognizing....", end="", flush=True)
            recognizer_text = recognizer.recognize_google(audio, language="hi-IN")  # Recognizing Hindi
            if recognizer_text:
                trans_text = Translate_hindi_to_english(recognizer_text)
                print("\r" + Fore.BLUE + "SARA : " + trans_text)
                return trans_text
            else:
                return ""
        except sr.UnknownValueError:
            print("\rCould not understand the audio.")
            return ""
        finally:
            print("\r", end="", flush=True)


def close_application(app_name):
    try:
        app_name = app_name.lower()
        found = False
        process_closed = False

        for proc in psutil.process_iter(['pid', 'name']):
            process_name = proc.info['name'].lower()

            if ("word" in app_name and "winword" in process_name) or \
                    (("vs code" in app_name or "visual studio code" in app_name) and "code" in process_name) or \
                    ("power point" in app_name and "PowerPoint" in process_name) or \
                    ("command prompt" in app_name or "cmd" in app_name and "cmd" in process_name) or \
                    ("brave" in app_name and "brave" in process_name) or \
                    ("chrome" in app_name and "chrome" in process_name):
                proc.terminate()
                say(f"Closing {process_name}")
                process_closed = True
                found = True

                proc.wait(timeout=5)  # Wait for up to 5 seconds for termination

        if not found:
            say("Sorry, I can't find that application to close.")
        elif not process_closed:
            say(f"Could not close {app_name}. It may be already closed.")

    except psutil.NoSuchProcess:
        say(f"{app_name} process is not running.")
    except Exception as e:
        say(f"Error occurred while closing the application: {e}")


def retry_request(url, retries=3, backoff_factor=0.5):
    for i in range(retries):
        try:
            response = requests.get(url, timeout=10)
            return response
        except requests.exceptions.ConnectTimeout:
            print(f"Attempt {i+1} failed. Retrying in {backoff_factor * (2 ** i)} seconds.")
            time.sleep(backoff_factor * (2 ** i))
    return None


# Main loop for assistant commands
while True:
    text = Speech_To_Text_Python()
    if text:
        if text.lower() == "stop":
            say("Stopping the program.")
            break

        say(text)

        if "open google" in text.lower():
            say("What would you like to search on Google?")
            search_query = Speech_To_Text_Python()
            if search_query:
                search_on_google(search_query)
            else:
                say("Sorry, I didn't catch that. Please try again.")

        sites = [
            ["YouTube", "https://youtube.com"],
            ["Browser", "https://google.com"],
            ["abc", "https://wikipedia.com"],
            ["Microsoft edge", "https://www.msn.com/en-in/feed"],
            ["Google Map", "https://www.google.com/maps/@18.5597952,73.8394112,12z?entry=ttu"],
            ["chat gpt", "https://chat.openai.com"],
        ]

        for site in sites:
            if f"open {site[0]}".lower() in text.lower():
                say(f"Opening {site[0]} sir..", voice="Microsoft Zira Desktop - English (United States)")
                webbrowser.open(site[1])

        if "play music" in text.lower():
            music_dir = r"C:\Users\harsh\OneDrive\Documents\music"
            songs = os.listdir(music_dir)
            os.startfile(os.path.join(music_dir, songs[0]))

        if "the time" in text.lower():
            hour = datetime.datetime.now().strftime("%H")
            min = datetime.datetime.now().strftime("%M")
            say(f"Time is {hour} hours and {min} minutes")

        if "play youtube" in text.lower():
            say("What would you like to search on YouTube?", voice="Microsoft Zira Desktop - English (United States)")
            search_query = Speech_To_Text_Python()
            if search_query:
                say(f"Searching YouTube for {search_query}", voice="Microsoft Zira Desktop - English (United States)")
                pywhatkit.playonyt(search_query)
                say("Done, sir")

        if "information" in text.lower():
            say("Searching sir")
            text = text.replace("wikipedia", "")
            try:
                results = wikipedia.summary(text, sentences=3)
                say("According to Wikipedia")
                print(results)
                say(results)
            except wikipedia.exceptions.DisambiguationError as e:
                say(f"Disambiguation error: {e}")
            except wikipedia.exceptions.PageError as e:
                say(f"Page error: {e}")

        if "note" in text.lower():
            say("What should I note down?")
            note = Speech_To_Text_Python().lower()
            with open("notes.txt", "a") as f:
                f.write(note + "\n")
            say("Note added successfully.")

        if "launch" in text.lower():
            if "word" in text.lower():
                say("Launching Microsoft Word", voice="Microsoft Zira Desktop - English (United States)")
                os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk")

            elif "vs code" in text.lower() or "visual studio code" in text.lower():
                say("Launching Visual Studio Code", voice="Microsoft Zira Desktop - English (United States)")
                os.startfile(
                    r"C:\Users\Yash\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code\Visual Studio Code.lnk")

            elif "powerpoint" in text.lower():
                say("Launching Microsoft PowerPoint", voice="Microsoft Zira Desktop - English (United States)")
                os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk")

            elif "command prompt" in text.lower() or "cmd" in text.lower():
                say("Launching Command Prompt", voice="Microsoft Zira Desktop - English (United States)")
                os.startfile(r"C:\Windows\System32\cmd.exe")

            elif "brave" in text.lower():
                say("Launching Brave Browser", voice="Microsoft Zira Desktop - English (United States)")
                os.startfile(r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe")

            elif "chrome" in text.lower():
                say("Launching Google Chrome", voice="Microsoft Zira Desktop - English (United States)")
                os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")

        if "close" in text.lower():
            say("Which application should I close?")
            app_name = Speech_To_Text_Python().lower()
            if app_name:
                close_application(app_name)
            else:
                say("Sorry, I didn't catch that. Please try again.")
