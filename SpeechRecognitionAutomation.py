import os
import webbrowser
import pyautogui
import whisper
import openai
import speech_recognition as sr
import psutil
import winshell
import pyttsx3
import asyncio
import json
import re
import datetime
from EdgeGPT import Chatbot, ConversationStyle
import random
import string
import requests
from urllib import request

# --------------------------- GLOBAL VARIABLES ----------------------------

OPENAI_API_KEY = ''
openai.api_key = OPENAI_API_KEY
model = whisper.load_model('base.en')
cookies = json.loads(open("cookies.json", encoding="utf-8").read())

# --------------------------- INIT VOICE ENGINE ---------------------------

engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speakText(text):
    try:
        engine.say(text)
        engine.runAndWait()
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        engine.say("You didn't enter a string")
        engine.runAndWait()


def internetOn():
    try:
        request.urlopen('https://www.google.com', timeout=1)
        return True
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        speakText("No internet connection!")
        return False



def generateRandomString(extension):
    ext_list = ['.txt', 'csv', '.docx']
    if extension in ext_list:
        letters = string.ascii_letters
        file_name = ''.join(random.choice(letters) for _ in range(10)) + extension
        return file_name


def formatTime():
    now = datetime.datetime.now()
    day = now.strftime("%d").lstrip("0")
    suffix = getDaySuffix(day)
    month = now.strftime("%B")
    year = now.strftime("%Y")
    time = now.strftime("%I:%M %p")

    formatted_time = f"{day}{suffix} of {month} {year} {time}"
    return formatted_time

def getDaySuffix(day):
    if day in {"11", "12", "13"}:
        suffix = "th"
    elif day[-1] == "1":
        suffix = "st"
    elif day[-1] == "2":
        suffix = "nd"
    elif day[-1] == "3":
        suffix = "rd"
    else:
        suffix = "th"
    return suffix


def createFile(file_name):
    try:
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        file_path = os.path.join(desktop_path, file_name)

        with open(file_path, "w") as file:
            file.close()
        print(f"File '{file_name}' created on the desktop.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


def getWeatherAPI(city):
    weather_api_key = ''
    url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={weather_api_key}&units=metric"
    response = requests.get(url)
    data = response.json()
    if response.status_code == 200:
        temperature = data["main"]["temp"]
        weather_description = data["weather"][0]["description"]
        sentence = f"The current weather in {city} is " \
                   f"{weather_description.lower()} with a temperature of {temperature}°C."
        return sentence
    else:
        print("Failed to retrieve weather data.")
        speakText("Failed to retrieve weather data.")
        return None


def removeStringSymbols(text):
    if text is not None:
        clean_text = re.sub(r'[^\w\s]', '', text)
        return clean_text
    else:
        return ""


def exeController(words):
    if "task" and "manager" in words:
        if "open" in words:
            try:
                os.system('start taskmgr.exe')
            except Exception as e:
                print(f"An error occurred: {str(e)}")

        elif "close" in words:
            try:
                os.system("taskkill /f /im  taskmgr.exe")
            except Exception as e:
                print(f"An error occurred: {str(e)}")

    elif "magnifier" in words:
        if "open" in words or "start" in words:
            try:
                os.system("start magnify.exe")
            except Exception as e:
                print(f"An error occurred: {str(e)}")

        elif "close" in words:
            try:
                os.system("taskkill /f /im  magnify.exe")
            except Exception as e:
                print(f"An error occurred: {str(e)}")

    if "log" and "out" and "account" in words:
        import ctypes
        try:
            ctypes.windll.user32.LockWorkStation()
        except Exception as e:
            print(f"An error occurred: {str(e)}")

    elif "create" and "excel" in words:
        extension = ".csv"
        createFile(generateRandomString(extension))

    elif "file" in words:
        if "create" in words:
            if "excel" in words:
                extension = ".csv"
                createFile(generateRandomString(extension))
            if "word" in words:
                extension = ".docx"
                createFile(generateRandomString(extension))
            if "text" in words:
                extension = ".txt"
                createFile(generateRandomString(extension))

    elif "narrator" in words:
        if "open" or "start" in words:
            try:
                os.system("start narrator.exe")
            except Exception as e:
                print(f"An error occurred: {str(e)}")

    elif "microsoft" and "edge" in words:
        if "open" in words:
            try:
                os.system("start msedge.exe")
            except Exception as e:
                print(f"An error occurred: {str(e)}")
        if "close" in words:
            try:
                os.system("taskkill /f /im  msedge.exe")
            except Exception as e:
                print(f"An error occurred: {str(e)}")

    elif "on" and "screen" and "keyboard" in words:
        if "open" in words:
            try:
                os.system('start osk.exe')
            except Exception as e:
                print(f"An error occurred: {str(e)}")
        elif "close" in words:
            try:
                os.system("taskkill /f /im  osk.exe")
            except Exception as e:
                print(f"An error occurred: {str(e)}")

    elif "computer" and "management" in words:
        if "open" in words:
            try:
                os.system('start compmgmt.msc')
            except Exception as e:
                print(f"An error occurred: {str(e)}")
        elif "close" in words:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] == 'mmc.exe':
                    try:
                        proc.kill()
                    except Exception as e:
                        print(f"An error occurred: {str(e)}")

    elif "disk" and "management" in words:
        if "open" in words:
            try:
                os.system('start diskmgmt.msc')
            except Exception as e:
                print(f"An error occurred: {str(e)}")
        elif "close" in words:
            try:
                os.system("taskkill /f /im  mmc.exe")
            except Exception as e:
                print(f"An error occurred: {str(e)}")

    elif "event" and "viewer" in words:
        if "open" in words:
            try:
                os.system('start eventvwr.msc')
            except Exception as e:
                print(f"An error occurred: {str(e)}")
        elif "close" in words:
            try:
                os.system("taskkill /f /im  mmc.exe")
            except Exception as e:
                print(f"An error occurred: {str(e)}")

    elif "performance" and "monitor" in words:
        if "open" in words:
            try:
                os.system('start perfmon.msc')
            except Exception as e:
                print(f"An error occurred: {str(e)}")
        elif "close" in words:
            try:
                os.system("taskkill /f /im  perfmon.exe")
            except Exception as e:
                print(f"An error occurred: {str(e)}")

    elif "empty" and "recycle" and "bin" in words:
        try:
            winshell.recycle_bin().empty(confirm=False,
                                         show_progress=False, sound=True)
            print("Recycle Bin is emptied now!")
            speakText("Recycle Bin is emptied now!")
        except Exception as e:

            print(f"An error occurred: {str(e)}")
            speakText("Recycle Bin is emptied now!")


def tellTime(words):
    if "tell" in words and "time" in words:
        try:
            print(formatTime())
            speakText(formatTime())
        except Exception as e:
            print(f"An error occurred: {str(e)}")

def keywordSearch(words, keyword):
    if words is not  None and keyword is not None:
        if len(words.split(' ')) > 3:

            keyword_index = words.index(keyword)
            words_after_keyword = words[keyword_index + len(keyword):]
            return words_after_keyword
    return ''

def browserFunc(sentence):
    if sentence:
        words = sentence.lower()
        if "google" in words:
            if "search" in words and "on" in words:
                if internetOn():
                    import pywhatkit
                    keyword = "google"
                    search_material = keywordSearch(words, keyword)
                    search_material = search_material.strip(string.punctuation)
                    if search_material:
                        speakText("Searching: " + search_material)
                        try:
                            pywhatkit.search(search_material)
                        except Exception as e:
                            print(f"An error occurred: {str(e)}")
        elif "youtube" in words:
            if internetOn():
                import pywhatkit
                keyword = "youtube"
                if "play" in words:
                    search_material = keywordSearch(words, keyword)
                    search_material = search_material.strip(string.punctuation)
                    if search_material:
                        speakText("Playing" + search_material)
                        try:
                            pywhatkit.playonyt(search_material)
                        except Exception as e:
                            print(f"An error occurred: {str(e)}")
                elif "search" in words:
                    search_material = keywordSearch(words, keyword)
                    search_material = search_material.strip(string.punctuation)
                    url = f"https://www.youtube.com/search?q={search_material}"
                    try:
                        webbrowser.open(url)
                    except Exception as e:
                        print(f"An error occurred: {str(e)}")

def weatherChecker(words):
    if "weather" in words and "in" in words:
        keyword = "weather in"
        chat_completion = openai.ChatCompletion.create(model='gpt-3.5-turbo',
        messages=[{'role': 'user','content': f'You will be given sentences that has citys in them, you need to find the city and return it.'
        f' It does not matter that else is in the sentence. Sentence: {words} . Your asnwer should containg only the name of the city.'
        f' If there is more than one - return only the name of the first one.'}])
        city = chat_completion['choices'][0]['message']['content']
        speakText("The city is "+city)
        weather_today = getWeatherAPI(city.strip(string.punctuation))
        if weather_today is not None:
            print(weather_today)
            speakText(weather_today)

def voiceToTextRecorder(words):
    if "record" in words and "voice" in words:
        speakText("We are in voice recorder!")
        voice_record = speechRecognition()
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        file_path = os.path.join(desktop_path, generateRandomString('.txt'))
        with open(file_path, "w") as file:
            file.write(voice_record)
            file.close()
        print(f" '{voice_record}'")

def keyboardCommands(words):
    if "volume" in words:
        if "up" in words:
            pyautogui.press("volumeup", 10)
        if "down" in words:
            pyautogui.press("volumedown", 10)
        if "mute" in words:
            pyautogui.press("volumedown", 100)
        if "max" in words:
            pyautogui.press("volumeup", 100)
    elif "file" in words and "explorer" in words:
        if "open" in words:
            pyautogui.hotkey("win", "e")
    elif "browser" in words:
        if "open" in words:
            webbrowser.open('https://www.google.com')
        if "forward" in words:
            pyautogui.press("browserforward")
        if "home" in words:
            pyautogui.press("browserhome")
        if "refresh" in words:
            pyautogui.press("browserrefresh")
        if "back" in words:
            pyautogui.press("browserback")
    elif "clear" and "view" in words:
        pyautogui.hotkey("win", "d")
    elif "run" and "prompt" in words:
        pyautogui.hotkey("win", "r")
    elif "take" in words and "screenshot" in words:
        screenshot = pyautogui.screenshot()
        screenshot_name = '\\' + "screenshot.png"
        screenshot_path = os.path.join(os.path.expanduser("~"),
                                       "Desktop") + screenshot_name
        screenshot.save(screenshot_path)

def runCommands():
    while True:
        sentence = speechRecognition()
        if sentence:
            words = sentence.lower()
            if "exit" in words:
                return
            exeController(words)
            browserFunc(sentence)
            keyboardCommands(words)
            weatherChecker(sentence)
            voiceToTextRecorder(words)
            tellTime(words)

def speechRecognition():

    path = 'audio.wav'
    recognizer = sr.Recognizer()

    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=0.5)
        print("Please talk")
        speakText('Please talk!')
        audio = recognizer.listen(source)

        with open(path, "wb") as f:
            f.write(audio.get_wav_data())
            f.close()

    try:
        result = model.transcribe(path, fp16=False)
        text = result['text']
        print(f"You said:{text[1:]}")
        deleteFile(path)
        return text[1:]
    except sr.UnknownValueError:
        speakText("Sorry, I could not understand what you said.")
        print("Sorry, I could not understand what you said.")
        return ''

def chatGPT():
    while True:
        user_input = speechRecognition()
        if user_input is not None:
            exit_keyword = removeStringSymbols(user_input.lower())
            if "please" in exit_keyword and "exit" in exit_keyword:
                return
            chat_completion = openai.ChatCompletion.create(model='gpt-3.5-turbo',
                                                           messages=[{'role': 'user', 'content': user_input}])

            text = chat_completion['choices'][0]['message']['content']
            print(text)
            speakText(text)

async def bingAI():
    while True:
        user_input = speechRecognition()
        exit_keyword = removeStringSymbols(user_input.lower())
        if "please" in exit_keyword and "exit" in exit_keyword:
            return
        elif user_input is not None:
            bot_response = ''
            bot = await Chatbot.create(cookies=cookies)
            response = await bot.ask(prompt=user_input, conversation_style=ConversationStyle.creative)
            for message in response["item"]["messages"]:
                if message["author"] == "bot":
                    bot_response = message["text"]
            bot_response = re.sub("\[\^\d+\^\]", '', bot_response)
            bot_response = bot_response.replace('*', '')
            print(bot_response)
            speakText(bot_response)

def modeSelection(user_input):

    if "chat" in user_input:
        print("You are in ChatGPT mode")
        speakText("You are in ChatGPT mode")
        chatGPT()
        speakText("Еnd of ChatGPT mode")
        print("Еnd of ChatGPT mode")

    elif "task" in user_input:
        print("You are in Task mode")
        speakText("You are in Task mode")
        runCommands()
        print("Еnd of Task mode")
        speakText("Еnd of Task mode")

    elif "bing" in user_input:
        print("You are in BingAi mode")
        speakText("You are in BingAi mode")
        asyncio.run(bingAI())
        print("You are in Task mode")
        speakText("Еnd of BingAi mode")

def startAutomation():
    print("Mode selection!")
    speakText("Please select mode. Task mode. ChatGPT mode. Bing AI mode.")
    while True:
        speakText("Select mode, please!")
        user_input = speechRecognition().lower()
        if "please" and "exit" in user_input:
            speakText("EXITING MODE")
            break
        modeSelection(user_input)

def deleteFile(file_path):
    if os.path.exists(file_path) and os.path.isfile(file_path):
        os.remove(file_path)
    else:
        print("File does not exist.")

def main():
    while True:
        user_input = speechRecognition().lower()
        if user_input is not None:
            if "goodbye" in user_input:
                print("Automation deactivated!")
                speakText("Automation is deactivated. Goodbye.")
                exit()
            elif "hello" in user_input:
                print("Automation activated!")
                speakText("Hello, I'm Jarvis. How can I help you today?")
                startAutomation()

#------------------MAIN------------------
if __name__ == '__main__':
    main()