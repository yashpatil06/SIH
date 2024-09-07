

import speech_recognition as sr
import win32com.client
import webbrowser
import datetime
import wikipedia
import pywhatkit
import os


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

while True:
    text = takeCommand()
    if text:
        if text.lower() == "stop":
            say("Stopping the program.")
            break

        say(text)
        sites = [
            ["YouTube", "https://youtube.com"],
            ["browser", "https://google.com"],
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

        if "open Google" in text.lower():
            print("Google search detected")
            say("Sir, what should I search on Google?")
            cm = takeCommand()
            if cm:
                print(f"Searching for: {cm}")
                webbrowser.open(f"https://www.google.com/search?q={cm}")
            else:
                print("No search term captured.")


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



           
















