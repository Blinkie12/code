import speech_recognition as sr
import pyttsx3
import os
import webbrowser
import subprocess
import datetime
import time
import random  
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

# Paths to programs
game_path = r"D:\Games\fifa 19\fifa 19.exe"
whatsapp_path = r"C:\Users\YourUsername\AppData\Local\WhatsApp\WhatsApp.exe"
vscode_path = r"C:\Users\YourUsername\AppData\Local\Programs\Microsoft VS Code\Code.exe"
vlc_path = r"C:\Program Files\VideoLAN\VLC\vlc.exe"
word_path = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
powerpoint_path = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
ds4windows_path = r"C:\Program Files\DS4Windows\DS4Windows.exe"

# Initialize audio volume control
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

# List of random greetings
GREETINGS = [
    "Hello, {name}! How can I assist you today?",
    "Hi, {name}! What can I do for you?",
    "Good to see you, {name}! How may I help?",
    "Hey, {name}! Ready to assist you.",
    "Greetings, {name}! What's on your mind?",
    "Hello there, {name}! How can I be of service?",
]

# User's name
USER_NAME = "Blessings"  # Change this as needed

def get_random_greeting():
    """Return a random greeting from the list."""
    return random.choice(GREETINGS).format(name=USER_NAME)

def recognize_speech_from_mic(recognizer, microphone):
    """Capture speech from the microphone and convert it to text."""
    with microphone as source:
        recognizer.adjust_for_ambient_noise(source)
        print("Listening...")
        audio = recognizer.listen(source)

    response = {"success": True, "error": None, "transcription": None}

    try:
        response["transcription"] = recognizer.recognize_google(audio).lower()
    except sr.RequestError:
        response["success"], response["error"] = False, "API unavailable"
    except sr.UnknownValueError:
        response["error"] = "Unable to recognize speech"

    return response

def speak_text(text):
    """Convert text to speech using pyttsx3."""
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')

    for voice in voices:
        if "female" in voice.name.lower():
            engine.setProperty('voice', voice.id)
            break

    engine.setProperty('rate', 180)
    engine.setProperty('volume', 1.0)
    engine.say(text)
    engine.runAndWait()

def execute_command(command):
    """Execute a command based on the recognized text."""
    command = command.lower()

    if "open notepad" in command:
        speak_text("Opening Notepad")
        os.system("notepad.exe")
    elif "close notepad" in command:
        speak_text("Closing Notepad")
        os.system('taskkill /IM notepad.exe /F')
    elif "open calculator" in command:
        speak_text("Opening Calculator")
        os.system("calc.exe")
    elif "open browser" in command:
        speak_text("Opening browser")
        webbrowser.open("https://www.google.com")
    elif "close browser" in command:
        speak_text("Closing browser")
        os.system("taskkill /IM chrome.exe /F") 
        os.system("taskkill /IM msedge.exe /F")  
        os.system("taskkill /IM firefox.exe /F")  
    elif "search for" in command:
        query = command.replace("search for", "").strip()
        speak_text(f"Searching for {query}")
        webbrowser.open(f"https://www.google.com/search?q={query}")
    elif "shutdown" or"shut down " in command:
        speak_text("Shutting down the system")
        os.system("shutdown /s /t 10")
    elif "restart" in command:
        speak_text("Restarting the system")
        os.system("shutdown /r /t 10")
    elif "sleep the system" in command:
        speak_text("Putting the system to sleep")
        os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
    elif "open fifa 19" in command:
        speak_text("Opening Fifa 19")
        subprocess.Popen(f'"{game_path}"', shell=True)
    elif "open whatsapp" in command:
        speak_text("Opening WhatsApp")
        subprocess.Popen(f'"{whatsapp_path}"', shell=True)
    elif "open visual studio code" in command:
        speak_text("Opening Visual Studio Code")
        subprocess.Popen(f'"{vscode_path}"', shell=True)
    elif "open vlc" in command:
        speak_text("Opening VLC Media Player")
        subprocess.Popen(f'"{vlc_path}"', shell=True)
    elif "open word" in command:
        speak_text("Opening Microsoft Word")
        subprocess.Popen(f'"{word_path}"', shell=True)
    elif "open powerpoint" in command:
        speak_text("Opening Microsoft PowerPoint")
        subprocess.Popen(f'"{powerpoint_path}"', shell=True)
    elif "turn on wifi" in command:
        speak_text("Turning on WiFi")
        os.system("netsh interface set interface Wi-Fi admin=enable")
    elif "turn off wifi" in command:
        speak_text("Turning off WiFi")
        os.system("netsh interface set interface Wi-Fi admin=disable")
    elif "turn on bluetooth" in command:
        speak_text("Turning on Bluetooth")
        os.system("net start bthserv")
    elif "turn off bluetooth" in command:
        speak_text("Turning off Bluetooth")
        os.system("net stop bthserv")
    elif "open ds4windows" in command:
        speak_text("Opening DS4Windows")
        subprocess.Popen(f'"{ds4windows_path}"', shell=True)
    elif "exit" in command or "quit" in command:
        speak_text(f"See you next time, {USER_NAME}!")
        time.sleep(10)  # Delay before exit
        exit()
    else:
        speak_text("Sorry, I don't understand that command.")

if __name__ == "__main__":
    recognizer = sr.Recognizer()
    microphone = sr.Microphone()

    greeting = get_random_greeting()
    speak_text(greeting)

    while True:
        print("\nPlease speak a command...")
        result = recognize_speech_from_mic(recognizer, microphone)

        if result["success"] and result["transcription"]:
            command = result["transcription"]
            print(f"You said: {command}")
            execute_command(command)
        else:
            print(f"Error: {result['error']}")
