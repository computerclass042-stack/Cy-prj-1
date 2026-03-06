import speech_recognition as sr
import webbrowser
import win32com.client
import jarvismusic
import feedparser
import wikipedia  # Sabse upar import karein

speaker = win32com.client.Dispatch("SAPI.SpVoice")


def speak(text):
    print(text)
    speaker.Speak(text)


def processcommand(c):
    c = c.lower()

    # Auto opening tab programme:- 
    if "google" in c:
        speak("Opening Google")
        webbrowser.open("https://google.com")

    elif "youtube" in c:
        speak("Opening Youtube")
        webbrowser.open("https://youtube.com")

    elif "whatsapp" in c:
        speak("Opening Whatsapp")
        webbrowser.open("https://whatsapp.com")

    elif "trade book" in c:
        speak("Opening Trade book")
        webbrowser.open("https://app.tradefxbook.com/")


    # Music playing programme:- 
    elif c.startswith("play"):
        try:
            # Agar sirf "play" bola hai toh 'haqeeqat' chalao,
            # warna gaane ka naam nikalo
            if c.strip() == "play":
                song = "haqeeqat"
            else:
                song = c.lower().replace("play", "").strip()

            speak(f"Playing {song}")
            link = jarvismusic.music[song]
            webbrowser.open(link)
        except KeyError:
            speak("Song not found in music library")

    # Top 5 headlines of the world:- 
    elif "news" in c:
        speak("Fetching world news from BBC")

        feed = feedparser.parse("http://feeds.bbci.co.uk/news/world/rss.xml")

        for i in range(5):
            headline = feed.entries[i].title
            print(f"{i+1}. {headline}")
            speak(headline)

    # search codes available in the downward:- 

    elif "search" in c:
        speak("Yes sir, what should I search for?")
        r = sr.Recognizer()
        with sr.Microphone() as source:
            audio = r.listen(source)
            try:
                search_query = r.recognize_google(audio)
                print(f"Searching for: {search_query}")
                speak(f"Searching Wikipedia for {search_query}...")

                # Wikipedia se 2 lines ki summary nikalna
                results = wikipedia.summary(search_query, sentences=2)
                speak("According to Wikipedia:")
                speak(results)
            except Exception as e:
                speak("Sorry sir, I couldn't find anything on that topic.")

    elif "stop" in c:
        speak("Goodbye sir")
        exit()

    else:
        speak("Command not recognized")


# ✅ MAIN PROGRAM
if __name__ == "__main__":
    speak("Jarvis activated")

    while True:
        r = sr.Recognizer()

        with sr.Microphone() as source:
            print("Listening for wake word...")
            audio = r.listen(source)

        try:
            wake_word = r.recognize_google(audio).lower()
            print("You said:", wake_word)

            if "jarvis" in wake_word:
                speak("Yes sir, tell me the command")

                with sr.Microphone() as source:
                    print("Listening for command...")
                    audio = r.listen(source)

                command = r.recognize_google(audio)
                print("Command:", command)

                processcommand(command)

        except Exception as e:
            print("Error:", e)
