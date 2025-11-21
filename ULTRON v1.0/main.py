# importing modules

from time import strftime
import speech_recognition as sr             #Audio recognition
import win32com.client                      #THE VOICE of ULTRON
import webbrowser                           #opening sites
import os                                   #music opening
import datetime


#Saying the text
def say(text, voice_id = 0):
    speaker = win32com.client.Dispatch(f"Sapi.SpVoice")   #it loads the Windows speech engine, Prepares it to speak
    voices = speaker.GetVoices()                          #get the diff voices
    speaker.Voice = voices.Item(voice_id)                 #Choose voice
    speaker.rate = 1.8
    speaker.Speak(text)                                   #speak the text out loud.

#Taking command
def TakeCommand():
    r = sr.Recognizer()                             #Create a recognizer
    with sr.Microphone() as source:                 #opening mic as source
        r.pause_threshold = 1                       #you stop speaking for 1 second, treat it as the end of your sentence.
        audio = r.listen(source)                    #Listening using the source
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')         #Trying to recognize the audio using google speech recognition API
            print(f"User said: {query}")
            return query.lower()
        except Exception as e:
            return "Sorry, I didn't get that. Please try again."

if __name__ == '__main__':
    say("Initializing system")
    say("Hi I am ULTRON version 1.0")
    say("How can I help You sir")

    while True:
        print("Listening...")
        query = TakeCommand()                                       #user voice input

        if "introduce yourself" in query or "who are you" in query or "what can you do" in query:
            intro = ("Hello sir. I am Ultron, your personal voice assistant. "
                     "I can open websites, launch applications, tell you the time, "
                     "play music, and send WhatsApp messages. "
                     "I listen to your commands and execute them instantly. "
                     "I am always learning and improving to assist you better.")
            say(intro)
            print(intro)
            continue
    #opeing sites Main code
        sites = [["youtube", "https://www.youtube.com/"], ["wikipedia", "https://en.wikipedia.org/"],
                 ["google", "https://www.google.com/"], ["instagram", "https://www.instagram.com/"],
                 ["chat gpt", "https://www.chatgpt.com/"], ["facebook", "https://www.facebook.com/"],
                 ["whatsapp", "https://www.whatsapp.com/"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} Sir...")
                webbrowser.open_new_tab(site[1])
    #playing music
        if "play music" in query.lower():
                say("Playing Music...")
                music_dir = r"E:\ULTRON\ULTRON v1.0\music.mp3"
                os.system(f'start "" "{music_dir}"')
    #playing yt music
        if "play songs on youtube" in query:
            say("Playing songs sir...")
            webbrowser.open_new_tab("https://www.youtube.com/watch?v=CsrsR_sC2jE")
    #telling time
        if "time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"The time is {strfTime}")

    #sleeping system
        if "sleep" in query or "computer off" in query or "pc sleep" in query:
            say("Putting your computer to sleep sir.")
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
    #opeing apps
        apps = [
            ["chrome", "C:/Program Files/Google/Chrome/Application/chrome.exe"],
            ["edge", r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"],
            ["camera", "microsoft.windows.camera:"],
            ["notepad", "notepad"],
            ["calculator", "calc"]
        ]

        for app_name, app_path in apps:
            if f"open {app_name}" in query:
                say(f"Opening {app_name} sir...")
                os.system(f'start \"\" \"{app_path}\"')

    # word for quiting the code
        if "exit" in query or "stop" in query or "sign off" in query or "quit" in query:
            say("Signing off")
            exit()



