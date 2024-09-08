from re import search
import pyautogui
import pyttsx3
import kit
import subprocess
import pywhatkit
import speech_recognition as sr
import win32com.client
import webbrowser
import datetime
import wikipedia
import os
import urllib.parse

def say(text, voice=None):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    if voice:
        try:
            voices = speaker.GetVoices()
            for v in voices:
                if v.GetDescription().strip() == voice.strip():
                    speaker.Voice = v
                    break
        except:
            print(f"Voice '{voice}' not found. Using the default voice.")
    speaker.Speak(text)

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

say("Sara HERE !!", voice="Microsoft Zira Desktop - English (United States)")

def search_on_google(query):
    try:
        encoded_query = urllib.parse.quote(query)  # Encode query for URL
        url = f"https://www.google.com/search?q={encoded_query}"
        webbrowser.open(url)
        say(f"Here are the results for {query}")
    except Exception as e:
        print(f"Error opening Google search: {e}")
        say("Sorry, I couldn't open the Google search.")

# Move the close_application function outside of any blocks
def close_application(app_name):
    try:
        if "word" in app_name.lower():
            os.system("taskkill /f /im WINWORD.EXE")
            say("Closing Microsoft Word", voice="Microsoft Zira Desktop - English (United States)")

        elif "vs code" in app_name.lower() or "visual studio code" in app_name.lower():
            # Try closing both 'Code.exe' and 'code.exe' in case of different process names
            os.system("taskkill /f /im Code.exe")
            os.system("taskkill /f /im code.exe")
            say("Closing Visual Studio Code", voice="Microsoft Zira Desktop - English (United States)")

        elif "powerpoint" in app_name.lower():
            os.system("taskkill /f /im POWERPNT.EXE")
            say("Closing Microsoft PowerPoint", voice="Microsoft Zira Desktop - English (United States)")

        elif "command prompt" in app_name.lower() or "cmd" in app_name.lower():
            os.system("taskkill /f /im cmd.exe")
            say("Closing Command Prompt", voice="Microsoft Zira Desktop - English (United States)")

        elif "brave" in app_name.lower():
            os.system("taskkill /f /im brave.exe")
            say("Closing Brave", voice="Microsoft Zira Desktop - English (United States)")

        elif "chrome" in app_name.lower():
            os.system("taskkill /f /im chrome.exe")
            say("Closing Chrome", voice="Microsoft Zira Desktop - English (United States)")

        else:
            say("Sorry, I can't find that application to close.")
    except Exception as e:
        say(f"Error occurred while closing the application: {e}")

while True:
    text = takeCommand()
    if text:
        if text.lower() == "stop":
            say("Stopping the program.")
            break

        say(text)

        # Feature to open Google and search
        if "open google" in text.lower():
            say("What would you like to search on Google?")
            search_query = takeCommand()
            if search_query:
                search_on_google(search_query)
            else:
                say("Sorry, I didn't catch that. Please try again.")

        # Feature to open specific websites
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

        # Feature to play music
        if "play music" in text.lower():
            music_dir = "C:\\Users\\harsh\\OneDrive\\Documents\\music"
            songs = os.listdir(music_dir)
            os.startfile(os.path.join(music_dir, songs[0]))

        # Feature for time
        if "the time" in text.lower():
            hour = datetime.datetime.now().strftime("%H")
            min = datetime.datetime.now().strftime("%M")
            say(f"Time is {hour} hours and {min} minutes")

        # Feature for playing YouTube videos
        if "play youtube" in text.lower():
            say("What would you like to search on YouTube?", voice="Microsoft Zira Desktop - English (United States)")
            search_query = takeCommand()
            if search_query:
                say(f"Searching YouTube for {search_query}", voice="Microsoft Zira Desktop - English (United States)")
                pywhatkit.playonyt(search_query)
                say("Done, sir")

        if "information" in text.lower():
            say("Searching sir")
            text = text.replace("wikipedia", "")
            results = wikipedia.summary(text, sentences=3)
            say("According to Wikipedia")
            print(results)
            say(results)

        # Feature to take a note
        if "take a note" in text.lower():
            say("What should I note down?")
            note = takeCommand().lower()
            with open("notes.txt", "a") as f:
                f.write(note + "\n")
            say("Note added successfully.")

        # Feature to launch applications
        if "launch" in text.lower():
            if "word" in text.lower():
                say("Launching Microsoft Word", voice="Microsoft Zira Desktop - English (United States)")
                os.startfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk")  # Modify the path if needed

            elif "vs code" in text.lower() or "visual studio code" in text.lower():
                say("Launching Visual Studio Code", voice="Microsoft Zira Desktop - English (United States)")
                os.startfile("C:\\Users\\harsh\\OneDrive\\Desktop\\Visual Studio Code.lnk")  # Modify path as needed

            elif "powerpoint" in text.lower():
                say("Launching Microsoft PowerPoint", voice="Microsoft Zira Desktop - English (United States)")
                os.startfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk")

            elif "command prompt" in text.lower() or "cmd" in text.lower():
                say("Launching Command Prompt", voice="Microsoft Zira Desktop - English (United States)")
                subprocess.Popen("cmd.exe")

            elif "brave" in text.lower():
                say("Launching Brave", voice="Microsoft Zira Desktop - English (United States)")
                os.startfile("C:\\Users\\Public\\Desktop\\Brave.lnk")

            elif "chrome" in text.lower():
                say("Launching Chrome", voice="Microsoft Zira Desktop - English (United States)")
                os.startfile("C:\\Users\\harsh\\OneDrive\\Desktop\\Person 1 - Chrome.lnk")

            else:
                say("Sorry, I can't find that application.")

        # Feature to close applications
        if "close" in text.lower():
            # You can either extract the app name from the text or ask the user
            # For simplicity, let's extract it from the text
            app_name = text.lower().replace("close", "").strip()
            if app_name:
                close_application(app_name)
            else:
                say("Which application would you like to close?")
                app_name = takeCommand()
                if app_name:
                    close_application(app_name)
                else:
                    say("Sorry, I didn't catch that.")

